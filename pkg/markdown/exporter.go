package markdown

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/zerx-lab/wordZero/pkg/document"
)

// WordToMarkdownExporter defines the interface for Word-to-Markdown export.
type WordToMarkdownExporter interface {
	// ExportToFile exports a Word document to a Markdown file.
	ExportToFile(docxPath, mdPath string, options *ExportOptions) error

	// ExportToString exports a Word document to a Markdown string.
	ExportToString(doc *document.Document, options *ExportOptions) (string, error)

	// ExportToBytes exports a Word document to Markdown bytes.
	ExportToBytes(doc *document.Document, options *ExportOptions) ([]byte, error)

	// BatchExport performs batch export.
	BatchExport(inputs []string, outputDir string, options *ExportOptions) error
}

// Exporter is the Word-to-Markdown exporter implementation.
type Exporter struct {
	opts *ExportOptions
}

// NewExporter creates a new exporter instance.
func NewExporter(opts *ExportOptions) *Exporter {
	if opts == nil {
		opts = DefaultExportOptions()
	}
	return &Exporter{opts: opts}
}

// ExportToFile exports a Word document to a Markdown file.
func (e *Exporter) ExportToFile(docxPath, mdPath string, options *ExportOptions) error {
	// Load Word document
	doc, err := document.Open(docxPath)
	if err != nil {
		return NewExportError("DocumentOpen", fmt.Sprintf("failed to open document: %v", err), err)
	}

	// Set image output path
	if options == nil {
		options = e.opts
	}
	if options.ExtractImages && options.ImageOutputDir == "" {
		options.ImageOutputDir = filepath.Dir(mdPath)
	}

	// Convert to Markdown
	markdown, err := e.ExportToString(doc, options)
	if err != nil {
		return err
	}

	// Write to file
	err = os.WriteFile(mdPath, []byte(markdown), 0644)
	if err != nil {
		return NewExportError("FileWrite", fmt.Sprintf("failed to write markdown file: %v", err), err)
	}

	return nil
}

// ExportToString exports a Word document to a Markdown string.
func (e *Exporter) ExportToString(doc *document.Document, options *ExportOptions) (string, error) {
	bytes, err := e.ExportToBytes(doc, options)
	if err != nil {
		return "", err
	}
	return string(bytes), nil
}

// ExportToBytes exports a Word document to Markdown bytes.
func (e *Exporter) ExportToBytes(doc *document.Document, options *ExportOptions) ([]byte, error) {
	if options != nil {
		e.opts = options
	}

	writer := &MarkdownWriter{
		opts:      e.opts,
		doc:       doc,
		imageNum:  0,
		footnotes: make([]string, 0),
	}

	return writer.Write()
}

// BatchExport exports multiple files in batch.
func (e *Exporter) BatchExport(inputs []string, outputDir string, options *ExportOptions) error {
	// Ensure the output directory exists
	err := os.MkdirAll(outputDir, 0755)
	if err != nil {
		return NewExportError("DirectoryCreate", fmt.Sprintf("failed to create output directory: %v", err), err)
	}

	total := len(inputs)
	for i, input := range inputs {
		// Report progress
		if options != nil && options.ProgressCallback != nil {
			options.ProgressCallback(i+1, total)
		}

		// Generate output filename
		base := strings.TrimSuffix(filepath.Base(input), filepath.Ext(input))
		output := filepath.Join(outputDir, base+".md")

		// Export a single file
		err := e.ExportToFile(input, output, options)
		if err != nil {
			if options != nil && options.ErrorCallback != nil {
				options.ErrorCallback(err)
			}
			if options == nil || !options.IgnoreErrors {
				return err
			}
		}
	}

	return nil
}

// DefaultExportOptions returns the default export configuration.
func DefaultExportOptions() *ExportOptions {
	return &ExportOptions{
		UseGFMTables:        true,
		PreserveFootnotes:   true,
		PreserveLineBreaks:  false,
		WrapLongLines:       false,
		MaxLineLength:       80,
		ExtractImages:       true,
		ImageNamePattern:    "image_%d.png",
		ImageRelativePath:   true,
		PreserveBookmarks:   true,
		ConvertHyperlinks:   true,
		PreserveCodeStyle:   true,
		DefaultCodeLang:     "",
		IgnoreUnknownStyles: true,
		PreserveTOC:         false,
		IncludeMetadata:     false,
		StripComments:       true,
		UseSetext:           false,
		BulletListMarker:    "-",
		EmphasisMarker:      "*",
		StrictMode:          false,
		IgnoreErrors:        true,
	}
}

// HighQualityExportOptions returns a high-quality export configuration.
func HighQualityExportOptions() *ExportOptions {
	opts := DefaultExportOptions()
	opts.ExtractImages = true
	opts.PreserveFootnotes = true
	opts.PreserveBookmarks = true
	opts.PreserveTOC = true
	opts.IncludeMetadata = true
	opts.StrictMode = true
	opts.IgnoreErrors = false
	return opts
}

// BidirectionalConverter is a bidirectional converter between Markdown and Word.
type BidirectionalConverter struct {
	mdToWord *Converter
	wordToMd *Exporter
}

// NewBidirectionalConverter creates a new bidirectional converter.
func NewBidirectionalConverter(mdOpts *ConvertOptions, exportOpts *ExportOptions) *BidirectionalConverter {
	return &BidirectionalConverter{
		mdToWord: NewConverter(mdOpts),
		wordToMd: NewExporter(exportOpts),
	}
}

// AutoConvert automatically detects the file type and converts accordingly.
func (bc *BidirectionalConverter) AutoConvert(inputPath, outputPath string) error {
	ext := strings.ToLower(filepath.Ext(inputPath))

	switch ext {
	case ".md", ".markdown":
		return bc.mdToWord.ConvertFile(inputPath, outputPath, nil)
	case ".docx":
		return bc.wordToMd.ExportToFile(inputPath, outputPath, nil)
	default:
		return fmt.Errorf("unsupported file type: %s", ext)
	}
}

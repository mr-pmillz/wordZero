package markdown

import (
	"os"
	"path/filepath"
	"strings"

	"github.com/yuin/goldmark"
	"github.com/yuin/goldmark/extension"
	"github.com/yuin/goldmark/parser"
	"github.com/yuin/goldmark/text"

	mathjax "github.com/litao91/goldmark-mathjax"

	"github.com/zerx-lab/wordZero/pkg/document"
)

// MarkdownConverter defines the interface for Markdown-to-Word conversion.
type MarkdownConverter interface {
	// ConvertFile converts a single file.
	ConvertFile(mdPath, docxPath string, options *ConvertOptions) error

	// ConvertBytes converts byte data.
	ConvertBytes(mdContent []byte, options *ConvertOptions) (*document.Document, error)

	// ConvertString converts a string.
	ConvertString(mdContent string, options *ConvertOptions) (*document.Document, error)

	// BatchConvert performs batch conversion.
	BatchConvert(inputs []string, outputDir string, options *ConvertOptions) error
}

// Converter is the default converter implementation.
type Converter struct {
	md   goldmark.Markdown
	opts *ConvertOptions
}

// NewConverter creates a new converter instance.
func NewConverter(opts *ConvertOptions) *Converter {
	if opts == nil {
		opts = DefaultOptions()
	}

	extensions := []goldmark.Extender{}
	if opts.EnableGFM {
		extensions = append(extensions, extension.GFM)
	}
	if opts.EnableFootnotes {
		extensions = append(extensions, extension.Footnote)
	}
	if opts.EnableMath {
		// Use standard LaTeX math delimiters: $...$ for inline formulas, $$...$$ for block formulas
		extensions = append(extensions, mathjax.NewMathJax(
			mathjax.WithInlineDelim("$", "$"),
			mathjax.WithBlockDelim("$$", "$$"),
		))
	}

	md := goldmark.New(
		goldmark.WithExtensions(extensions...),
		goldmark.WithParserOptions(
			parser.WithAutoHeadingID(),
		),
	)

	return &Converter{md: md, opts: opts}
}

// ConvertString converts string content to a Word document.
func (c *Converter) ConvertString(content string, opts *ConvertOptions) (*document.Document, error) {
	return c.ConvertBytes([]byte(content), opts)
}

// ConvertBytes converts byte data to a Word document.
func (c *Converter) ConvertBytes(content []byte, opts *ConvertOptions) (*document.Document, error) {
	if opts != nil {
		c.opts = opts
	}

	// Create a new Word document
	doc := document.New()

	// Apply page settings
	if c.opts.PageSettings != nil {
		// Can be extended later using the existing page settings API
	}

	// Parse Markdown
	reader := text.NewReader(content)
	astDoc := c.md.Parser().Parse(reader)

	// Create renderer and convert
	renderer := &WordRenderer{
		doc:    doc,
		opts:   c.opts,
		source: content,
	}

	err := renderer.Render(astDoc)
	if err != nil {
		return nil, err
	}

	return doc, nil
}

// ConvertFile converts a Markdown file to a Word document.
func (c *Converter) ConvertFile(mdPath, docxPath string, options *ConvertOptions) error {
	// Read Markdown file
	content, err := os.ReadFile(mdPath)
	if err != nil {
		return NewConversionError("FileRead", "failed to read markdown file", 0, 0, err)
	}

	// Set image base path (if not specified)
	if options == nil {
		options = c.opts
	}
	if options.ImageBasePath == "" {
		options.ImageBasePath = filepath.Dir(mdPath)
	}

	// Convert content
	doc, err := c.ConvertBytes(content, options)
	if err != nil {
		return err
	}

	// Save Word document
	err = doc.Save(docxPath)
	if err != nil {
		return NewConversionError("FileSave", "failed to save word document", 0, 0, err)
	}

	return nil
}

// BatchConvert converts multiple files in batch.
func (c *Converter) BatchConvert(inputs []string, outputDir string, options *ConvertOptions) error {
	// Ensure the output directory exists
	err := os.MkdirAll(outputDir, 0755)
	if err != nil {
		return NewConversionError("DirectoryCreate", "failed to create output directory", 0, 0, err)
	}

	total := len(inputs)
	for i, input := range inputs {
		// Report progress
		if options != nil && options.ProgressCallback != nil {
			options.ProgressCallback(i+1, total)
		}

		// Generate output filename
		base := strings.TrimSuffix(filepath.Base(input), filepath.Ext(input))
		output := filepath.Join(outputDir, base+".docx")

		// Convert a single file
		err := c.ConvertFile(input, output, options)
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

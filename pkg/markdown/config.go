// Package markdown provides Markdown-to-Word document conversion functionality.
package markdown

import "github.com/zerx-lab/wordZero/pkg/document"

// ConvertOptions configures conversion options.
type ConvertOptions struct {
	// Basic configuration
	EnableGFM       bool // enable GitHub Flavored Markdown
	EnableFootnotes bool // enable footnote support
	EnableTables    bool // enable table support
	EnableTaskList  bool // enable task lists
	EnableMath      bool // enable math formula support (LaTeX syntax)

	// Style configuration
	StyleMapping      map[string]string // custom style mapping
	DefaultFontFamily string            // default font family
	DefaultFontSize   float64           // default font size

	// Image handling
	ImageBasePath string  // image base path
	EmbedImages   bool    // whether to embed images
	MaxImageWidth float64 // maximum image width (inches)

	// Link handling
	PreserveLinkStyle  bool // preserve link style
	ConvertToBookmarks bool // convert internal links to bookmarks

	// Document settings
	GenerateTOC  bool                   // generate table of contents
	TOCMaxLevel  int                    // maximum TOC heading level
	PageSettings *document.PageSettings // page settings (using existing struct)

	// Error handling
	StrictMode    bool        // strict mode
	IgnoreErrors  bool        // ignore conversion errors
	ErrorCallback func(error) // error callback

	// Progress reporting
	ProgressCallback func(int, int) // progress callback
}

// DefaultOptions returns the default conversion configuration.
func DefaultOptions() *ConvertOptions {
	return &ConvertOptions{
		EnableGFM:         true,
		EnableFootnotes:   true,
		EnableTables:      true,
		EnableTaskList:    true,
		EnableMath:        true, // math formula support enabled by default
		DefaultFontFamily: "Calibri",
		DefaultFontSize:   11.0,
		EmbedImages:       false,
		MaxImageWidth:     6.0, // inches
		GenerateTOC:       true,
		TOCMaxLevel:       3,
		StrictMode:        false,
		IgnoreErrors:      true,
	}
}

// HighQualityOptions returns a high-quality conversion configuration.
func HighQualityOptions() *ConvertOptions {
	opts := DefaultOptions()
	opts.EmbedImages = true
	opts.PreserveLinkStyle = true
	opts.ConvertToBookmarks = true
	opts.StrictMode = true
	opts.IgnoreErrors = false
	return opts
}

// Package style provides predefined style constants.
package style

// Predefined style ID constants
const (
	// StyleNormal is the normal text style.
	StyleNormal = "Normal"

	// Heading styles
	StyleHeading1 = "Heading1"
	StyleHeading2 = "Heading2"
	StyleHeading3 = "Heading3"
	StyleHeading4 = "Heading4"
	StyleHeading5 = "Heading5"
	StyleHeading6 = "Heading6"
	StyleHeading7 = "Heading7"
	StyleHeading8 = "Heading8"
	StyleHeading9 = "Heading9"

	// Document title styles
	StyleTitle    = "Title"    // Document title
	StyleSubtitle = "Subtitle" // Subtitle

	// Character styles
	StyleEmphasis = "Emphasis" // Emphasis (italic)
	StyleStrong   = "Strong"   // Bold
	StyleCodeChar = "CodeChar" // Code character

	// Paragraph styles
	StyleQuote         = "Quote"         // Quote style
	StyleListParagraph = "ListParagraph" // List paragraph
	StyleCodeBlock     = "CodeBlock"     // Code block
)

// GetPredefinedStyleNames returns a mapping of all predefined style names.
func GetPredefinedStyleNames() map[string]string {
	return map[string]string{
		StyleNormal:        "Normal",
		StyleHeading1:      "Heading 1",
		StyleHeading2:      "Heading 2",
		StyleHeading3:      "Heading 3",
		StyleHeading4:      "Heading 4",
		StyleHeading5:      "Heading 5",
		StyleHeading6:      "Heading 6",
		StyleHeading7:      "Heading 7",
		StyleHeading8:      "Heading 8",
		StyleHeading9:      "Heading 9",
		StyleTitle:         "Document Title",
		StyleSubtitle:      "Subtitle",
		StyleEmphasis:      "Emphasis",
		StyleStrong:        "Strong",
		StyleCodeChar:      "Code Character",
		StyleQuote:         "Quote",
		StyleListParagraph: "List Paragraph",
		StyleCodeBlock:     "Code Block",
	}
}

// StyleConfig is a helper struct for style configuration.
type StyleConfig struct {
	StyleID     string
	Name        string
	Description string
	StyleType   StyleType
}

// GetPredefinedStyleConfigs returns all predefined style configurations.
func GetPredefinedStyleConfigs() []StyleConfig {
	return []StyleConfig{
		{
			StyleID:     StyleNormal,
			Name:        "Normal",
			Description: "Default paragraph style using Calibri font at 11pt",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading1,
			Name:        "Heading 1",
			Description: "Heading level 1, 16pt blue bold, 12pt space before",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading2,
			Name:        "Heading 2",
			Description: "Heading level 2, 13pt blue bold, 6pt space before",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading3,
			Name:        "Heading 3",
			Description: "Heading level 3, 12pt blue bold, 6pt space before",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading4,
			Name:        "Heading 4",
			Description: "Heading level 4, 12pt blue bold, 6pt space before",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading5,
			Name:        "Heading 5",
			Description: "Heading level 5, 12pt blue bold, 6pt space before",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading6,
			Name:        "Heading 6",
			Description: "Heading level 6, 12pt blue bold, 6pt space before",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading7,
			Name:        "Heading 7",
			Description: "Heading level 7, 12pt blue bold, 6pt space before",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading8,
			Name:        "Heading 8",
			Description: "Heading level 8, 12pt blue bold, 6pt space before",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleHeading9,
			Name:        "Heading 9",
			Description: "Heading level 9, 12pt blue bold, 6pt space before",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleTitle,
			Name:        "Document Title",
			Description: "Document title style",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleSubtitle,
			Name:        "Subtitle",
			Description: "Subtitle style",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleEmphasis,
			Name:        "Emphasis",
			Description: "Italic text style",
			StyleType:   StyleTypeCharacter,
		},
		{
			StyleID:     StyleStrong,
			Name:        "Strong",
			Description: "Bold text style",
			StyleType:   StyleTypeCharacter,
		},
		{
			StyleID:     StyleCodeChar,
			Name:        "Code Character",
			Description: "Monospace font, red text, for code snippets",
			StyleType:   StyleTypeCharacter,
		},
		{
			StyleID:     StyleQuote,
			Name:        "Quote",
			Description: "Quote paragraph style, italic gray, indented 0.5 inch on each side",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleListParagraph,
			Name:        "List Paragraph",
			Description: "List paragraph style",
			StyleType:   StyleTypeParagraph,
		},
		{
			StyleID:     StyleCodeBlock,
			Name:        "Code Block",
			Description: "Code block style",
			StyleType:   StyleTypeParagraph,
		},
	}
}

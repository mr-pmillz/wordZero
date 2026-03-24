// Package document page settings tests
package document

import (
	"testing"
)

// TestDefaultPageSettings tests default page settings
func TestDefaultPageSettings(t *testing.T) {
	settings := DefaultPageSettings()

	if settings.Size != PageSizeA4 {
		t.Errorf("default page size should be A4, got: %s", settings.Size)
	}

	if settings.Orientation != OrientationPortrait {
		t.Errorf("default page orientation should be portrait, got: %s", settings.Orientation)
	}

	if settings.MarginTop != 25.4 {
		t.Errorf("default top margin should be 25.4mm, got: %.1fmm", settings.MarginTop)
	}
}

// TestSetPageSize tests setting page size
func TestSetPageSize(t *testing.T) {
	doc := New()

	// Test setting to Letter size
	err := doc.SetPageSize(PageSizeLetter)
	if err != nil {
		t.Errorf("failed to set page size: %v", err)
	}

	settings := doc.GetPageSettings()
	if settings.Size != PageSizeLetter {
		t.Errorf("page size should be Letter, got: %s", settings.Size)
	}
}

// TestSetCustomPageSize tests setting custom page size
func TestSetCustomPageSize(t *testing.T) {
	doc := New()

	// Test valid custom size
	err := doc.SetCustomPageSize(200, 300)
	if err != nil {
		t.Errorf("failed to set custom page size: %v", err)
	}

	settings := doc.GetPageSettings()
	if settings.Size != PageSizeCustom {
		t.Errorf("page size should be Custom, got: %s", settings.Size)
	}

	if abs(settings.CustomWidth-200) > 0.1 {
		t.Errorf("custom width should be 200mm, got: %.1fmm", settings.CustomWidth)
	}

	if abs(settings.CustomHeight-300) > 0.1 {
		t.Errorf("custom height should be 300mm, got: %.1fmm", settings.CustomHeight)
	}

	// Test invalid custom size
	err = doc.SetCustomPageSize(-100, 200)
	if err == nil {
		t.Error("setting negative size should return an error")
	}

	err = doc.SetCustomPageSize(100, 0)
	if err == nil {
		t.Error("setting zero height should return an error")
	}
}

// TestSetPageOrientation tests setting page orientation
func TestSetPageOrientation(t *testing.T) {
	doc := New()

	// Test setting to landscape
	err := doc.SetPageOrientation(OrientationLandscape)
	if err != nil {
		t.Errorf("failed to set page orientation: %v", err)
	}

	settings := doc.GetPageSettings()
	if settings.Orientation != OrientationLandscape {
		t.Errorf("page orientation should be landscape, got: %s", settings.Orientation)
	}
}

// TestSetPageMargins tests setting page margins
func TestSetPageMargins(t *testing.T) {
	doc := New()

	// Test valid margin settings
	err := doc.SetPageMargins(20, 15, 25, 30)
	if err != nil {
		t.Errorf("failed to set page margins: %v", err)
	}

	settings := doc.GetPageSettings()
	if abs(settings.MarginTop-20) > 0.1 {
		t.Errorf("top margin should be 20mm, got: %.1fmm", settings.MarginTop)
	}
	if abs(settings.MarginRight-15) > 0.1 {
		t.Errorf("right margin should be 15mm, got: %.1fmm", settings.MarginRight)
	}
	if abs(settings.MarginBottom-25) > 0.1 {
		t.Errorf("bottom margin should be 25mm, got: %.1fmm", settings.MarginBottom)
	}
	if abs(settings.MarginLeft-30) > 0.1 {
		t.Errorf("left margin should be 30mm, got: %.1fmm", settings.MarginLeft)
	}

	// Test negative margins
	err = doc.SetPageMargins(-10, 15, 25, 30)
	if err == nil {
		t.Error("setting negative margins should return an error")
	}
}

// TestSetHeaderFooterDistance tests setting header/footer distance
func TestSetHeaderFooterDistance(t *testing.T) {
	doc := New()

	// Test valid header/footer distance
	err := doc.SetHeaderFooterDistance(10, 15)
	if err != nil {
		t.Errorf("failed to set header/footer distance: %v", err)
	}

	settings := doc.GetPageSettings()
	if abs(settings.HeaderDistance-10) > 0.1 {
		t.Errorf("header distance should be 10mm, got: %.1fmm", settings.HeaderDistance)
	}
	if abs(settings.FooterDistance-15) > 0.1 {
		t.Errorf("footer distance should be 15mm, got: %.1fmm", settings.FooterDistance)
	}

	// Test negative distance
	err = doc.SetHeaderFooterDistance(-5, 15)
	if err == nil {
		t.Error("setting negative header distance should return an error")
	}
}

// TestSetGutterWidth tests setting gutter width
func TestSetGutterWidth(t *testing.T) {
	doc := New()

	// Test valid gutter width
	err := doc.SetGutterWidth(5)
	if err != nil {
		t.Errorf("failed to set gutter width: %v", err)
	}

	settings := doc.GetPageSettings()
	if abs(settings.GutterWidth-5) > 0.1 {
		t.Errorf("gutter width should be 5mm, got: %.1fmm", settings.GutterWidth)
	}

	// Test negative gutter width
	err = doc.SetGutterWidth(-2)
	if err == nil {
		t.Error("setting negative gutter width should return an error")
	}
}

// TestPageDimensions tests page dimension calculations
func TestPageDimensions(t *testing.T) {
	tests := []struct {
		name      string
		settings  *PageSettings
		expWidth  float64
		expHeight float64
	}{
		{
			name: "A4 portrait",
			settings: &PageSettings{
				Size:        PageSizeA4,
				Orientation: OrientationPortrait,
			},
			expWidth:  210,
			expHeight: 297,
		},
		{
			name: "A4 landscape",
			settings: &PageSettings{
				Size:        PageSizeA4,
				Orientation: OrientationLandscape,
			},
			expWidth:  297,
			expHeight: 210,
		},
		{
			name: "Custom size",
			settings: &PageSettings{
				Size:         PageSizeCustom,
				CustomWidth:  150,
				CustomHeight: 200,
				Orientation:  OrientationPortrait,
			},
			expWidth:  150,
			expHeight: 200,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			width, height := getPageDimensions(tt.settings)

			if width != tt.expWidth {
				t.Errorf("width mismatch, expected: %.1fmm, got: %.1fmm", tt.expWidth, width)
			}

			if height != tt.expHeight {
				t.Errorf("height mismatch, expected: %.1fmm, got: %.1fmm", tt.expHeight, height)
			}
		})
	}
}

// TestIdentifyPageSize tests page size identification
func TestIdentifyPageSize(t *testing.T) {
	tests := []struct {
		name     string
		width    float64
		height   float64
		expected PageSize
	}{
		{
			name:     "A4 portrait",
			width:    210,
			height:   297,
			expected: PageSizeA4,
		},
		{
			name:     "A4 landscape",
			width:    297,
			height:   210,
			expected: PageSizeA4,
		},
		{
			name:     "Letter",
			width:    215.9,
			height:   279.4,
			expected: PageSizeLetter,
		},
		{
			name:     "Custom size",
			width:    100,
			height:   150,
			expected: PageSizeCustom,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := identifyPageSize(tt.width, tt.height)

			if result != tt.expected {
				t.Errorf("page size identification error, expected: %s, got: %s", tt.expected, result)
			}
		})
	}
}

// TestValidatePageSettings tests page settings validation
func TestValidatePageSettings(t *testing.T) {
	// Test valid settings
	validSettings := &PageSettings{
		Size:         PageSizeA4,
		Orientation:  OrientationPortrait,
		CustomWidth:  0,
		CustomHeight: 0,
	}

	err := validatePageSettings(validSettings)
	if err != nil {
		t.Errorf("valid settings should pass validation, error: %v", err)
	}

	// Test invalid custom size
	invalidCustomSize := &PageSettings{
		Size:         PageSizeCustom,
		Orientation:  OrientationPortrait,
		CustomWidth:  -100,
		CustomHeight: 200,
	}

	err = validatePageSettings(invalidCustomSize)
	if err == nil {
		t.Error("negative custom size should fail validation")
	}

	// Test oversized custom dimensions
	oversizeCustom := &PageSettings{
		Size:         PageSizeCustom,
		Orientation:  OrientationPortrait,
		CustomWidth:  600, // exceeds maximum size
		CustomHeight: 200,
	}

	err = validatePageSettings(oversizeCustom)
	if err == nil {
		t.Error("oversized custom dimensions should fail validation")
	}

	// Test invalid orientation
	invalidOrientation := &PageSettings{
		Size:        PageSizeA4,
		Orientation: PageOrientation("invalid"),
	}

	err = validatePageSettings(invalidOrientation)
	if err == nil {
		t.Error("invalid orientation should fail validation")
	}
}

// TestMmToTwips tests millimeter to twips conversion
func TestMmToTwips(t *testing.T) {
	// Test several known conversion values
	tests := []struct {
		mm       float64
		expected float64
	}{
		{25.4, 1440}, // 1 inch = 1440 twips
		{0, 0},       // 0mm = 0 twips
		{10, 566.93}, // approx 567 twips
	}

	for _, tt := range tests {
		result := mmToTwips(tt.mm)
		// Allow decimal point error
		if abs(result-tt.expected) > 1 {
			t.Errorf("mm conversion error, input: %.1fmm, expected: %.0f twips, got: %.0f twips",
				tt.mm, tt.expected, result)
		}
	}
}

// TestTwipsToMM tests twips to millimeter conversion
func TestTwipsToMM(t *testing.T) {
	// Test reverse conversion
	tests := []struct {
		twips    float64
		expected float64
	}{
		{1440, 25.4}, // 1440 twips = 1 inch = 25.4mm
		{0, 0},       // 0 twips = 0mm
		{567, 10.0},  // approx 10mm
	}

	for _, tt := range tests {
		result := twipsToMM(tt.twips)
		// Allow decimal point error
		if abs(result-tt.expected) > 0.1 {
			t.Errorf("twips conversion error, input: %.0f twips, expected: %.1fmm, got: %.1fmm",
				tt.twips, tt.expected, result)
		}
	}
}

// TestCompletePageSettings tests the complete page settings workflow
func TestCompletePageSettings(t *testing.T) {
	doc := New()

	// Create complete page settings
	settings := &PageSettings{
		Size:           PageSizeLetter,
		Orientation:    OrientationLandscape,
		MarginTop:      20,
		MarginRight:    15,
		MarginBottom:   25,
		MarginLeft:     30,
		HeaderDistance: 8,
		FooterDistance: 12,
		GutterWidth:    5,
	}

	// Apply settings
	err := doc.SetPageSettings(settings)
	if err != nil {
		t.Errorf("failed to set page settings: %v", err)
	}

	// Verify settings were correctly applied
	retrieved := doc.GetPageSettings()

	if retrieved.Size != settings.Size {
		t.Errorf("page size mismatch, expected: %s, got: %s", settings.Size, retrieved.Size)
	}

	if retrieved.Orientation != settings.Orientation {
		t.Errorf("page orientation mismatch, expected: %s, got: %s", settings.Orientation, retrieved.Orientation)
	}

	if abs(retrieved.MarginTop-settings.MarginTop) > 0.1 {
		t.Errorf("top margin mismatch, expected: %.1fmm, got: %.1fmm", settings.MarginTop, retrieved.MarginTop)
	}
}

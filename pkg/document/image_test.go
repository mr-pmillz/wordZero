package document

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"os"
	"testing"
)

const testImagePNG = "image/png"

// createTestImage creates a test PNG image
func createTestImage(width, height int) []byte {
	img := image.NewRGBA(image.Rect(0, 0, width, height))

	// Create a simple red rectangle
	for y := 0; y < height; y++ {
		for x := 0; x < width; x++ {
			img.Set(x, y, color.RGBA{255, 0, 0, 255})
		}
	}

	var buf bytes.Buffer
	png.Encode(&buf, img)
	return buf.Bytes()
}

func TestDetectImageFormat(t *testing.T) {
	tests := []struct {
		name     string
		data     []byte
		expected ImageFormat
		hasError bool
	}{
		{
			name:     "PNG format",
			data:     []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A},
			expected: ImageFormatPNG,
			hasError: false,
		},
		{
			name:     "JPEG format",
			data:     []byte{0xFF, 0xD8, 0xFF},
			expected: ImageFormatJPEG,
			hasError: false,
		},
		{
			name:     "GIF87a format",
			data:     []byte("GIF87a"),
			expected: ImageFormatGIF,
			hasError: false,
		},
		{
			name:     "GIF89a format",
			data:     []byte("GIF89a"),
			expected: ImageFormatGIF,
			hasError: false,
		},
		{
			name:     "Data too short",
			data:     []byte{0x89},
			expected: "",
			hasError: true,
		},
		{
			name:     "Unsupported format",
			data:     []byte("INVALID_FORMAT"),
			expected: "",
			hasError: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			format, err := detectImageFormat(tt.data)

			if tt.hasError {
				if err == nil {
					t.Errorf("expected an error, but got none")
				}
			} else {
				if err != nil {
					t.Errorf("did not expect an error, but got: %v", err)
				}
				if format != tt.expected {
					t.Errorf("expected format %v, but got %v", tt.expected, format)
				}
			}
		})
	}
}

func TestGetImageDimensions(t *testing.T) {
	// Create a 100x50 test image
	testImageData := createTestImage(100, 50)

	width, height, err := getImageDimensions(testImageData, ImageFormatPNG)
	if err != nil {
		t.Fatalf("failed to get image dimensions: %v", err)
	}

	if width != 100 {
		t.Errorf("expected width 100, got %d", width)
	}

	if height != 50 {
		t.Errorf("expected height 50, got %d", height)
	}
}

func TestCalculateDisplaySize(t *testing.T) {
	doc := New()

	tests := []struct {
		name      string
		imageInfo *ImageInfo
		expectedW int64
		expectedH int64
	}{
		{
			name: "Default size",
			imageInfo: &ImageInfo{
				Width:  100,
				Height: 50,
				Config: nil,
			},
			expectedW: 100 * 9525, // pixels to EMU
			expectedH: 50 * 9525,
		},
		{
			name: "Specified dimensions",
			imageInfo: &ImageInfo{
				Width:  100,
				Height: 50,
				Config: &ImageConfig{
					Size: &ImageSize{
						Width:  50.0, // 50mm
						Height: 25.0, // 25mm
					},
				},
			},
			expectedW: int64(50.0 * 36000), // mm to EMU
			expectedH: int64(25.0 * 36000),
		},
		{
			name: "Width only, keep aspect ratio",
			imageInfo: &ImageInfo{
				Width:  100,
				Height: 50,
				Config: &ImageConfig{
					Size: &ImageSize{
						Width:           50.0, // 50mm
						KeepAspectRatio: true,
					},
				},
			},
			expectedW: int64(50.0 * 36000),
			expectedH: int64(50.0 * 36000 * 0.5), // keep aspect ratio 2:1
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			width, height := doc.calculateDisplaySize(tt.imageInfo)

			if width != tt.expectedW {
				t.Errorf("expected width %d, got %d", tt.expectedW, width)
			}

			if height != tt.expectedH {
				t.Errorf("expected height %d, got %d", tt.expectedH, height)
			}
		})
	}
}

func TestAddImageFromData(t *testing.T) {
	doc := New()

	// Create test image data
	imageData := createTestImage(100, 50)

	// Add image
	imageInfo, err := doc.AddImageFromData(
		imageData,
		"test.png",
		ImageFormatPNG,
		100,
		50,
		&ImageConfig{
			AltText: "测试图片",
			Title:   "测试标题",
		},
	)

	if err != nil {
		t.Fatalf("failed to add image: %v", err)
	}

	// Verify image info
	if imageInfo.Format != ImageFormatPNG {
		t.Errorf("expected format PNG, got %v", imageInfo.Format)
	}

	if imageInfo.Width != 100 {
		t.Errorf("expected width 100, got %d", imageInfo.Width)
	}

	if imageInfo.Height != 50 {
		t.Errorf("expected height 50, got %d", imageInfo.Height)
	}

	if imageInfo.Config.AltText != "测试图片" {
		t.Errorf("expected alt text '测试图片', got '%s'", imageInfo.Config.AltText)
	}

	// Verify relationship was correctly added
	if len(doc.documentRelationships.Relationships) != 1 {
		t.Errorf("expected 1 relationship, got %d", len(doc.documentRelationships.Relationships))
	}

	// Verify image data is stored (now uses safe filename image0.png)
	if _, exists := doc.parts["word/media/image0.png"]; !exists {
		t.Error("image data was not stored correctly")
	}

	// Verify content type was added
	foundPNG := false
	for _, def := range doc.contentTypes.Defaults {
		if def.Extension == string(ImageFormatPNG) && def.ContentType == testImagePNG {
			foundPNG = true
			break
		}
	}
	if !foundPNG {
		t.Error("PNG content type was not added correctly")
	}
}

func TestResizeImage(t *testing.T) {
	doc := New()

	imageInfo := &ImageInfo{
		Config: &ImageConfig{},
	}

	newSize := &ImageSize{
		Width:           100.0,
		Height:          50.0,
		KeepAspectRatio: true,
	}

	err := doc.ResizeImage(imageInfo, newSize)
	if err != nil {
		t.Fatalf("failed to resize image: %v", err)
	}

	if imageInfo.Config.Size != newSize {
		t.Error("image size was not set correctly")
	}
}

func TestSetImagePosition(t *testing.T) {
	doc := New()

	imageInfo := &ImageInfo{
		Config: &ImageConfig{},
	}

	err := doc.SetImagePosition(imageInfo, ImagePositionFloatLeft, 10.0, 20.0)
	if err != nil {
		t.Fatalf("failed to set image position: %v", err)
	}

	if imageInfo.Config.Position != ImagePositionFloatLeft {
		t.Error("image position was not set correctly")
	}

	if imageInfo.Config.OffsetX != 10.0 {
		t.Error("image X offset was not set correctly")
	}

	if imageInfo.Config.OffsetY != 20.0 {
		t.Error("image Y offset was not set correctly")
	}
}

func TestSetImageWrapText(t *testing.T) {
	doc := New()

	imageInfo := &ImageInfo{
		Config: &ImageConfig{},
	}

	err := doc.SetImageWrapText(imageInfo, ImageWrapSquare)
	if err != nil {
		t.Fatalf("failed to set image text wrapping: %v", err)
	}

	if imageInfo.Config.WrapText != ImageWrapSquare {
		t.Error("image text wrapping was not set correctly")
	}
}

// TestFloatingImageXMLStructure tests the floating image XML structure fix
func TestFloatingImageXMLStructure(t *testing.T) {
	doc := New()

	// Create test image data
	imageData := []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A} // PNG header

	// Test left float + tight wrap
	config := &ImageConfig{
		Position: ImagePositionFloatLeft,
		WrapText: ImageWrapTight,
		Size: &ImageSize{
			Width:  50,
			Height: 37.5,
		},
		AltText: "测试图片",
		Title:   "测试图片",
	}

	imageInfo, err := doc.AddImageFromData(imageData, "test.png", ImageFormatPNG, 100, 75, config)
	if err != nil {
		t.Fatalf("failed to add floating image: %v", err)
	}

	// Verify config was set correctly
	if imageInfo.Config.Position != ImagePositionFloatLeft {
		t.Error("image position was not set to float left")
	}

	if imageInfo.Config.WrapText != ImageWrapTight {
		t.Error("image wrap type was not set to tight wrap")
	}

	// Save document and check for success
	err = doc.Save("test_floating_fix.docx")
	if err != nil {
		t.Fatalf("failed to save document with fixed floating image: %v", err)
	}

	// Clean up test file
	defer func() {
		if err := os.Remove("test_floating_fix.docx"); err != nil {
			t.Logf("failed to clean up test file: %v", err)
		}
	}()

	t.Log("floating image XML structure fix test passed")
}

// TestCreateDefaultWrapPolygon tests default wrap polygon creation
func TestCreateDefaultWrapPolygon(t *testing.T) {
	doc := New()

	polygon := doc.createDefaultWrapPolygon()
	if polygon == nil {
		t.Fatal("failed to create default wrap polygon")
	}

	if polygon.Start == nil {
		t.Error("wrap polygon missing start point")
	}

	if len(polygon.LineTo) == 0 {
		t.Error("wrap polygon missing line segments")
	}

	// Verify start point coordinates
	if polygon.Start.X != "0" || polygon.Start.Y != "0" {
		t.Error("wrap polygon start point coordinates are incorrect")
	}

	// Verify closed path is formed
	expectedPoints := 4 // rectangle should have 4 points
	if len(polygon.LineTo) != expectedPoints {
		t.Errorf("expected %d line segments, got %d", expectedPoints, len(polygon.LineTo))
	}

	t.Log("default wrap polygon creation test passed")
}

func TestSetImageAltText(t *testing.T) {
	doc := New()

	imageInfo := &ImageInfo{
		Config: &ImageConfig{},
	}

	err := doc.SetImageAltText(imageInfo, "新的替代文字")
	if err != nil {
		t.Fatalf("failed to set image alt text: %v", err)
	}

	if imageInfo.Config.AltText != "新的替代文字" {
		t.Error("image alt text was not set correctly")
	}
}

func TestSetImageTitle(t *testing.T) {
	doc := New()

	imageInfo := &ImageInfo{
		Config: &ImageConfig{},
	}

	err := doc.SetImageTitle(imageInfo, "新的标题")
	if err != nil {
		t.Fatalf("failed to set image title: %v", err)
	}

	if imageInfo.Config.Title != "新的标题" {
		t.Error("image title was not set correctly")
	}
}

func TestAddImageContentType(t *testing.T) {
	doc := New()

	// Test adding PNG content type
	doc.addImageContentType(ImageFormatPNG)

	found := false
	for _, def := range doc.contentTypes.Defaults {
		if def.Extension == string(ImageFormatPNG) && def.ContentType == testImagePNG {
			found = true
			break
		}
	}

	if !found {
		t.Error("PNG content type was not added correctly")
	}

	// Test adding the same type again
	originalCount := len(doc.contentTypes.Defaults)
	doc.addImageContentType(ImageFormatPNG)

	if len(doc.contentTypes.Defaults) != originalCount {
		t.Error("duplicate content type addition should be ignored")
	}
}

func TestSetImageAlignment(t *testing.T) {
	doc := New()

	imageInfo := &ImageInfo{
		Config: &ImageConfig{},
	}

	// Test various alignment types
	alignments := []AlignmentType{
		AlignLeft,
		AlignCenter,
		AlignRight,
		AlignJustify,
	}

	for _, alignment := range alignments {
		err := doc.SetImageAlignment(imageInfo, alignment)
		if err != nil {
			t.Fatalf("failed to set image alignment: %v, alignment: %s", err, alignment)
		}

		if imageInfo.Config.Alignment != alignment {
			t.Errorf("image alignment not set correctly, expected: %s, got: %s", alignment, imageInfo.Config.Alignment)
		}
	}
}

func TestImageParagraphAlignment(t *testing.T) {
	doc := New()

	// Create test image data
	imageData := []byte{
		0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG header
		0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR
		0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1 pixel
		0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
		0x89, 0x00, 0x00, 0x00, 0x0B, 0x49, 0x44, 0x41,
		0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
		0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
		0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
		0x42, 0x60, 0x82,
	}

	// Test center-aligned image
	imageInfo, err := doc.AddImageFromData(
		imageData,
		"test.png",
		ImageFormatPNG,
		1, 1,
		&ImageConfig{
			Position:  ImagePositionInline,
			Alignment: AlignCenter,
		},
	)
	if err != nil {
		t.Fatalf("failed to add image: %v", err)
	}

	// Verify paragraph alignment of the image
	if len(doc.Body.Elements) == 0 {
		t.Fatal("no paragraphs were added to the document")
	}

	paragraph, ok := doc.Body.Elements[0].(*Paragraph)
	if !ok {
		t.Fatal("first element is not a paragraph")
	}

	if paragraph.Properties == nil {
		t.Fatal("paragraph properties are nil")
	}

	if paragraph.Properties.Justification == nil {
		t.Fatal("paragraph justification property is nil")
	}

	if paragraph.Properties.Justification.Val != string(AlignCenter) {
		t.Errorf("paragraph alignment is incorrect, expected: %s, got: %s",
			AlignCenter, paragraph.Properties.Justification.Val)
	}

	// Test modifying alignment
	err = doc.SetImageAlignment(imageInfo, AlignRight)
	if err != nil {
		t.Fatalf("failed to modify image alignment: %v", err)
	}

	if imageInfo.Config.Alignment != AlignRight {
		t.Error("image alignment modification failed")
	}
}

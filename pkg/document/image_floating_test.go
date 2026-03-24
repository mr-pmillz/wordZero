package document

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"os"
	"testing"
)

// TestFloatingImageLeftWithTightWrap tests left-floating image + tight wrap
func TestFloatingImageLeftWithTightWrap(t *testing.T) {
	doc := New()

	// Create test image
	imageData := createTestImageRGBA(100, 100)

	config := &ImageConfig{
		Position: ImagePositionFloatLeft,
		WrapText: ImageWrapTight,
		Size: &ImageSize{
			Width:  23.6,
			Height: 13,
		},
		AltText: "左浮动测试图片",
		Title:   "测试",
	}

	_, err := doc.AddImageFromData(imageData, "test.png", ImageFormatPNG, 100, 100, config)
	if err != nil {
		t.Fatalf("failed to add left-floating image: %v", err)
	}

	// Save and verify
	filename := "test_float_left_tight.docx"
	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(filename)

	// Reopen document to verify
	doc2, err := Open(filename)
	if err != nil {
		t.Fatalf("failed to open saved document: %v", err)
	}

	if len(doc2.Body.Elements) == 0 {
		t.Fatal("document has no elements")
	}

	t.Logf("left float + tight wrap image test passed")
}

// TestFloatingImageRightWithSquareWrap tests right-floating image + square wrap
func TestFloatingImageRightWithSquareWrap(t *testing.T) {
	doc := New()

	imageData := createTestImageRGBA(100, 100)

	config := &ImageConfig{
		Position: ImagePositionFloatRight,
		WrapText: ImageWrapSquare,
		Size: &ImageSize{
			Width:  30,
			Height: 20,
		},
		AltText: "右浮动测试图片",
		Title:   "测试",
	}

	_, err := doc.AddImageFromData(imageData, "test.png", ImageFormatPNG, 100, 100, config)
	if err != nil {
		t.Fatalf("failed to add right-floating image: %v", err)
	}

	filename := "test_float_right_square.docx"
	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(filename)

	// Reopen document to verify
	doc2, err := Open(filename)
	if err != nil {
		t.Fatalf("failed to open saved document: %v", err)
	}

	if len(doc2.Body.Elements) == 0 {
		t.Fatal("document has no elements")
	}

	t.Logf("right float + square wrap image test passed")
}

// TestFloatingImageWithTopAndBottomWrap tests floating image + top and bottom wrap
func TestFloatingImageWithTopAndBottomWrap(t *testing.T) {
	doc := New()

	imageData := createTestImageRGBA(100, 100)

	config := &ImageConfig{
		Position: ImagePositionFloatLeft,
		WrapText: ImageWrapTopAndBottom,
		Size: &ImageSize{
			Width:  40,
			Height: 30,
		},
		AltText: "上下环绕测试图片",
		Title:   "测试",
	}

	_, err := doc.AddImageFromData(imageData, "test.png", ImageFormatPNG, 100, 100, config)
	if err != nil {
		t.Fatalf("failed to add floating image: %v", err)
	}

	filename := "test_float_topbottom.docx"
	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(filename)

	// Reopen document to verify
	doc2, err := Open(filename)
	if err != nil {
		t.Fatalf("failed to open saved document: %v", err)
	}

	if len(doc2.Body.Elements) == 0 {
		t.Fatal("document has no elements")
	}

	t.Logf("float + top and bottom wrap image test passed")
}

// TestFloatingImageWithNoWrap tests floating image + no wrap
func TestFloatingImageWithNoWrap(t *testing.T) {
	doc := New()

	imageData := createTestImageRGBA(100, 100)

	config := &ImageConfig{
		Position: ImagePositionFloatRight,
		WrapText: ImageWrapNone,
		Size: &ImageSize{
			Width:  25,
			Height: 15,
		},
		AltText: "无环绕测试图片",
		Title:   "测试",
	}

	_, err := doc.AddImageFromData(imageData, "test.png", ImageFormatPNG, 100, 100, config)
	if err != nil {
		t.Fatalf("failed to add floating image: %v", err)
	}

	filename := "test_float_nowrap.docx"
	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(filename)

	// Reopen document to verify
	doc2, err := Open(filename)
	if err != nil {
		t.Fatalf("failed to open saved document: %v", err)
	}

	if len(doc2.Body.Elements) == 0 {
		t.Fatal("document has no elements")
	}

	t.Logf("float + no wrap image test passed")
}

// TestMultipleFloatingImages tests multiple floating images
func TestMultipleFloatingImages(t *testing.T) {
	doc := New()

	// Add text paragraph
	doc.AddParagraph("这是一个包含多个浮动图片的文档测试。")

	imageData := createTestImageRGBA(80, 80)

	// Add left-floating image
	config1 := &ImageConfig{
		Position: ImagePositionFloatLeft,
		WrapText: ImageWrapSquare,
		Size: &ImageSize{
			Width:  20,
			Height: 20,
		},
	}
	_, err := doc.AddImageFromData(imageData, "test1.png", ImageFormatPNG, 80, 80, config1)
	if err != nil {
		t.Fatalf("failed to add first floating image: %v", err)
	}

	// Add more text
	doc.AddParagraph("第一个图片已添加。")

	// Add right-floating image
	config2 := &ImageConfig{
		Position: ImagePositionFloatRight,
		WrapText: ImageWrapTight,
		Size: &ImageSize{
			Width:  20,
			Height: 20,
		},
	}
	_, err = doc.AddImageFromData(imageData, "test2.png", ImageFormatPNG, 80, 80, config2)
	if err != nil {
		t.Fatalf("failed to add second floating image: %v", err)
	}

	doc.AddParagraph("第二个图片也已添加。")

	filename := "test_multiple_float.docx"
	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(filename)

	// Verify document can be opened normally
	doc2, err := Open(filename)
	if err != nil {
		t.Fatalf("failed to open saved document: %v", err)
	}

	// Should have 5 elements (3 text paragraphs + 2 image paragraphs)
	const expectedElements = 5 // 3 text paragraphs + 2 image paragraphs
	if len(doc2.Body.Elements) != expectedElements {
		t.Errorf("expected %d elements (3 text paragraphs + 2 image paragraphs), got %d", expectedElements, len(doc2.Body.Elements))
	}

	t.Logf("multiple floating images test passed")
}

// createTestImageRGBA creates a colorful test image
func createTestImageRGBA(width, height int) []byte {
	img := image.NewRGBA(image.Rect(0, 0, width, height))

	// Create gradient colors
	for y := 0; y < height; y++ {
		for x := 0; x < width; x++ {
			r := uint8(x * 255 / width)   //nolint:gosec
			g := uint8(y * 255 / height)  //nolint:gosec
			b := uint8((x + y) * 255 / (width + height)) //nolint:gosec
			img.SetRGBA(x, y, color.RGBA{r, g, b, 255})
		}
	}

	var buf bytes.Buffer
	png.Encode(&buf, img)
	return buf.Bytes()
}

// TestInlineImageNotAffected ensures inline images are not affected by the fix
func TestInlineImageNotAffected(t *testing.T) {
	doc := New()

	imageData := createTestImageRGBA(100, 100)

	// Test inline image (default)
	config := &ImageConfig{
		Position:  ImagePositionInline,
		Alignment: AlignCenter,
		Size: &ImageSize{
			Width:  30,
			Height: 30,
		},
	}

	_, err := doc.AddImageFromData(imageData, "test.png", ImageFormatPNG, 100, 100, config)
	if err != nil {
		t.Fatalf("failed to add inline image: %v", err)
	}

	filename := "test_inline_image.docx"
	err = doc.Save(filename)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(filename)

	// Reopen to verify
	doc2, err := Open(filename)
	if err != nil {
		t.Fatalf("failed to open saved document: %v", err)
	}

	if len(doc2.Body.Elements) == 0 {
		t.Fatal("document has no elements")
	}

	t.Logf("inline image not affected test passed")
}

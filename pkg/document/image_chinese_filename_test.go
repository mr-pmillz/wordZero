package document

import (
	"os"
	"strconv"
	"strings"
	"testing"
)

const testImageRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

// TestChineseFilename tests whether images with Chinese filenames can be saved and opened correctly
func TestChineseFilename(t *testing.T) {
	doc := New()
	doc.AddParagraph("测试中文文件名")

	// Create test image
	imageData := createTestImage(100, 75)

	// Add image with Chinese filename
	_, err := doc.AddImageFromData(
		imageData,
		"测试图片.png", // Chinese filename
		ImageFormatPNG,
		100, 75,
		&ImageConfig{
			Position:  ImagePositionInline,
			Alignment: AlignCenter,
			AltText:   "测试图片",
			Title:     "测试图片标题",
		},
	)
	if err != nil {
		t.Fatalf("failed to add image with Chinese filename: %v", err)
	}

	doc.AddParagraph("图片下方的文字")

	// Save document
	testFile := "test_chinese_filename.docx"
	err = doc.Save(testFile)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(testFile)

	// Verify image is stored with safe filename (image0.png instead of 测试图片.png)
	foundSafeFilename := false
	foundChineseFilename := false
	for partName := range doc.parts {
		if strings.Contains(partName, "word/media/") {
			if strings.Contains(partName, "image0.png") {
				foundSafeFilename = true
			}
			if strings.Contains(partName, "测试") {
				foundChineseFilename = true
			}
			t.Logf("found image: %s", partName)
		}
	}

	if !foundSafeFilename {
		t.Error("safe filename (image0.png) not found, Chinese filename conversion failed")
	}

	if foundChineseFilename {
		t.Error("Chinese filename found, should have been converted to safe ASCII filename")
	}

	// Verify relationships also use safe filename
	foundImageRelationship := false
	for _, rel := range doc.documentRelationships.Relationships {
		if rel.Type == testImageRelType {
			foundImageRelationship = true
			if !strings.Contains(rel.Target, "image0.png") {
				t.Errorf("image relationship not using safe filename, Target=%s", rel.Target)
			}
			if strings.Contains(rel.Target, "测试") {
				t.Errorf("image relationship contains Chinese characters, Target=%s", rel.Target)
			}
			t.Logf("image relationship: ID=%s, Target=%s", rel.ID, rel.Target)
			break
		}
	}

	if !foundImageRelationship {
		t.Error("image relationship not found")
	}

	// Open document to verify
	doc2, err := Open(testFile)
	if err != nil {
		t.Fatalf("failed to open document: %v", err)
	}

	// Verify image data exists
	if _, exists := doc2.parts["word/media/image0.png"]; !exists {
		t.Error("image data not found in opened document")
	}

	t.Log("Chinese filename test passed: automatically converted to safe ASCII filename")
}

// TestMultipleNonASCIIFilenames tests multiple images with non-ASCII filenames
func TestMultipleNonASCIIFilenames(t *testing.T) {
	doc := New()
	doc.AddParagraph("测试多个非ASCII文件名")

	imageData := createTestImage(50, 50)

	// Add multiple images with filenames in different languages
	testFilenames := []string{
		"中文图片.png",    // Chinese
		"日本語.png",     // Japanese
		"한국어.png",     // Korean
		"Русский.png", // Russian
		"العربية.png", // Arabic
	}

	for i, filename := range testFilenames {
		_, err := doc.AddImageFromData(imageData, filename, ImageFormatPNG, 50, 50, nil)
		if err != nil {
			t.Fatalf("failed to add image %s: %v", filename, err)
		}

		// Verify each image uses a safe filename
		expectedSafeFilename := "image" + strconv.Itoa(i) + ".png"
		if _, exists := doc.parts["word/media/"+expectedSafeFilename]; !exists {
			t.Errorf("image %s did not use safe filename %s", filename, expectedSafeFilename)
		}
	}

	// Save and reopen
	testFile := "test_multiple_nonascii_filenames.docx"
	err := doc.Save(testFile)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(testFile)

	doc2, err := Open(testFile)
	if err != nil {
		t.Fatalf("failed to open document: %v", err)
	}

	// Verify all images exist
	for i := 0; i < len(testFilenames); i++ {
		expectedSafeFilename := "image" + strconv.Itoa(i) + ".png"
		if _, exists := doc2.parts["word/media/"+expectedSafeFilename]; !exists {
			t.Errorf("image %s not found in opened document", expectedSafeFilename)
		}
	}

	t.Logf("multi-language filename test passed: all %d images converted correctly", len(testFilenames))
}

package document

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"os"
	"testing"
)

// TestImagePersistenceAfterOpenAndSave tests that images persist after opening and re-saving a document
func TestImagePersistenceAfterOpenAndSave(t *testing.T) {
	// Step 1: Create a document containing an image
	doc1 := New()
	doc1.AddParagraph("测试文档 - 图片持久性测试")

	// Create test image
	imageData := createTestImageForPersistence(100, 75, color.RGBA{255, 100, 100, 255})

	// Add image (filename will be automatically converted to safe image0.png)
	imageInfo, err := doc1.AddImageFromData(
		imageData,
		"test_image.png",
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
		t.Fatalf("failed to add image: %v", err)
	}

	doc1.AddParagraph("图片下方的文字")

	// Save the first document
	testFile1 := "test_image_persistence_1.docx"
	err = doc1.Save(testFile1)
	if err != nil {
		t.Fatalf("failed to save first document: %v", err)
	}
	defer os.Remove(testFile1)

	// Verify the first document contains image data (using safe filename image0.png)
	if _, exists := doc1.parts["word/media/image0.png"]; !exists {
		t.Fatal("image data not found in first document")
	}

	// Step 2: Open the saved document
	doc2, err := Open(testFile1)
	if err != nil {
		t.Fatalf("failed to open document: %v", err)
	}

	// Verify the opened document contains image data
	if _, exists := doc2.parts["word/media/image0.png"]; !exists {
		t.Fatal("image data not found in opened document")
	}

	// Verify document relationships contain image relationship
	foundImageRelationship := false
	for _, rel := range doc2.documentRelationships.Relationships {
		if rel.Type == testImageRelType {
			foundImageRelationship = true
			t.Logf("found image relationship: ID=%s, Target=%s", rel.ID, rel.Target)
			break
		}
	}
	if !foundImageRelationship {
		t.Fatal("image relationship not found in opened document")
	}

	// Step 3: Modify document and save as new file
	doc2.AddParagraph("这是新添加的段落")

	testFile2 := "test_image_persistence_2.docx"
	err = doc2.Save(testFile2)
	if err != nil {
		t.Fatalf("failed to save second document: %v", err)
	}
	defer os.Remove(testFile2)

	// Step 4: Open the second document, verify image still exists
	doc3, err := Open(testFile2)
	if err != nil {
		t.Fatalf("failed to open second document: %v", err)
	}

	// Verify image data still exists
	if _, exists := doc3.parts["word/media/image0.png"]; !exists {
		t.Fatal("[issue] image data not found in second document - image lost after save!")
	}

	// Verify image relationship still exists
	foundImageRelationship = false
	for _, rel := range doc3.documentRelationships.Relationships {
		if rel.Type == testImageRelType {
			foundImageRelationship = true
			t.Logf("found image relationship in second document: ID=%s, Target=%s", rel.ID, rel.Target)
			break
		}
	}
	if !foundImageRelationship {
		t.Fatal("[issue] image relationship not found in second document - image relationship lost after save!")
	}

	// Verify image data integrity
	originalImageData := doc1.parts["word/media/image0.png"]
	finalImageData := doc3.parts["word/media/image0.png"]

	if !bytes.Equal(originalImageData, finalImageData) {
		t.Fatal("image data changed after save and reopen")
	}

	t.Log("image persistence test passed: image still exists after modifying and saving document")
	t.Logf("original image info: ID=%s, format=%s, dimensions=%dx%d",
		imageInfo.ID, imageInfo.Format, imageInfo.Width, imageInfo.Height)
}

// TestAddImageToOpenedDocument tests adding new images to an opened document
func TestAddImageToOpenedDocument(t *testing.T) {
	// Step 1: Create a document containing one image
	doc1 := New()
	doc1.AddParagraph("原始文档")

	// Add first image (red) - will be saved as image0.png
	imageData1 := createTestImageForPersistence(100, 75, color.RGBA{255, 0, 0, 255})
	_, err := doc1.AddImageFromData(
		imageData1,
		"image1.png",
		ImageFormatPNG,
		100, 75,
		&ImageConfig{
			Position:  ImagePositionInline,
			Alignment: AlignCenter,
		},
	)
	if err != nil {
		t.Fatalf("failed to add first image: %v", err)
	}

	// Save document
	testFile1 := "test_add_image_to_opened_1.docx"
	err = doc1.Save(testFile1)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(testFile1)

	// Step 2: Open document and add second image
	doc2, err := Open(testFile1)
	if err != nil {
		t.Fatalf("failed to open document: %v", err)
	}

	doc2.AddParagraph("添加第二张图片")

	// Add second image (blue) - will be saved as image1.png
	imageData2 := createTestImageForPersistence(100, 75, color.RGBA{0, 0, 255, 255})
	_, err = doc2.AddImageFromData(
		imageData2,
		"image2.png",
		ImageFormatPNG,
		100, 75,
		&ImageConfig{
			Position:  ImagePositionInline,
			Alignment: AlignCenter,
		},
	)
	if err != nil {
		t.Fatalf("failed to add second image: %v", err)
	}

	// Save document
	testFile2 := "test_add_image_to_opened_2.docx"
	err = doc2.Save(testFile2)
	if err != nil {
		t.Fatalf("failed to save document with two images: %v", err)
	}
	defer os.Remove(testFile2)

	// Step 3: Open document, verify both images exist
	doc3, err := Open(testFile2)
	if err != nil {
		t.Fatalf("failed to open document with two images: %v", err)
	}

	// Verify both image data exist (now using safe filenames image0.png and image1.png)
	if _, exists := doc3.parts["word/media/image0.png"]; !exists {
		t.Fatal("[issue] first image data lost")
	}

	if _, exists := doc3.parts["word/media/image1.png"]; !exists {
		t.Fatal("[issue] second image data lost")
	}

	// Verify image relationship count
	imageRelCount := 0
	for _, rel := range doc3.documentRelationships.Relationships {
		if rel.Type == testImageRelType {
			imageRelCount++
			t.Logf("found image relationship: ID=%s, Target=%s", rel.ID, rel.Target)
		}
	}

	if imageRelCount != 2 {
		t.Fatalf("expected 2 image relationships, got %d", imageRelCount)
	}

	t.Log("add image to opened document test passed: both images saved correctly")
}

// TestImageIDCounterAfterOpen tests that the image ID counter is correctly updated after opening a document
func TestImageIDCounterAfterOpen(t *testing.T) {
	// Step 1: Create a document containing two images
	doc1 := New()
	doc1.AddParagraph("测试图片ID计数器")

	// Add two images (will be saved as image0.png and image1.png)
	imageData := createTestImageForPersistence(50, 50, color.RGBA{255, 0, 0, 255})

	_, err := doc1.AddImageFromData(imageData, "img1.png", ImageFormatPNG, 50, 50, nil)
	if err != nil {
		t.Fatalf("failed to add first image: %v", err)
	}

	_, err = doc1.AddImageFromData(imageData, "img2.png", ImageFormatPNG, 50, 50, nil)
	if err != nil {
		t.Fatalf("failed to add second image: %v", err)
	}

	// Save document
	testFile := "test_image_id_counter.docx"
	err = doc1.Save(testFile)
	if err != nil {
		t.Fatalf("failed to save document: %v", err)
	}
	defer os.Remove(testFile)

	// Step 2: Open document
	doc2, err := Open(testFile)
	if err != nil {
		t.Fatalf("failed to open document: %v", err)
	}

	// Verify nextImageID was correctly updated
	// doc1 has two images, using rId2 and rId3 (rId1 is styles.xml)
	// So after opening, nextImageID should be at least 2 (max rId is 3)
	if doc2.nextImageID < 2 {
		t.Fatalf("nextImageID not correctly updated: expected >= 2, got = %d", doc2.nextImageID)
	}

	t.Logf("nextImageID after opening document = %d (as expected)", doc2.nextImageID)

	// Step 3: Add third image
	_, err = doc2.AddImageFromData(imageData, "img3.png", ImageFormatPNG, 50, 50, nil)
	if err != nil {
		t.Fatalf("failed to add third image: %v", err)
	}

	// Save and reopen, verify all three images exist
	testFile2 := "test_image_id_counter_2.docx"
	err = doc2.Save(testFile2)
	if err != nil {
		t.Fatalf("failed to save document with three images: %v", err)
	}
	defer os.Remove(testFile2)

	doc3, err := Open(testFile2)
	if err != nil {
		t.Fatalf("failed to open document with three images: %v", err)
	}

	// Verify all three images exist (using safe filenames image0.png, image1.png, image2.png)
	images := []string{"image0.png", "image1.png", "image2.png"}
	for _, imgName := range images {
		if _, exists := doc3.parts["word/media/"+imgName]; !exists {
			t.Fatalf("[issue] image %s lost", imgName)
		}
	}

	// Verify image relationship count
	imageRelCount := 0
	for _, rel := range doc3.documentRelationships.Relationships {
		if rel.Type == testImageRelType {
			imageRelCount++
		}
	}

	if imageRelCount != 3 {
		t.Fatalf("expected 3 image relationships, got %d", imageRelCount)
	}

	t.Log("image ID counter test passed: all image IDs are correct with no conflicts")
}

// createTestImageForPersistence creates an image for persistence testing
func createTestImageForPersistence(width, height int, bgColor color.RGBA) []byte {
	img := image.NewRGBA(image.Rect(0, 0, width, height))

	// Fill background color
	for y := 0; y < height; y++ {
		for x := 0; x < width; x++ {
			img.Set(x, y, bgColor)
		}
	}

	// Add border
	borderColor := color.RGBA{0, 0, 0, 255}
	for x := 0; x < width; x++ {
		img.Set(x, 0, borderColor)
		img.Set(x, height-1, borderColor)
	}
	for y := 0; y < height; y++ {
		img.Set(0, y, borderColor)
		img.Set(width-1, y, borderColor)
	}

	var buf bytes.Buffer
	png.Encode(&buf, img)
	return buf.Bytes()
}

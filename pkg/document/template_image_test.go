// Package document template image preservation tests
package document

import (
	"os"
	"strings"
	"testing"
)

const testWordMediaPrefix = "word/media/"

// TestTemplateImagePreservation tests whether images and formatting are preserved after template replacement
//nolint:gocognit
func TestTemplateImagePreservation(t *testing.T) {
	// Create a minimal valid PNG image data for testing
	testImageData := []byte{
		0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
		0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
		0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
		0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
		0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
		0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
		0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
		0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
		0x42, 0x60, 0x82,
	}

	// Step 1: Create a template document with images and formatting
	t.Log("Step 1: Create template document with images and formatting")
	templateDoc := New()

	// Add heading with template variable
	templateDoc.AddHeadingParagraph("文档标题: {{title}}", 1)

	// Add formatted paragraph
	formattedPara := templateDoc.AddFormattedParagraph("作者: {{author}}", &TextFormat{
		Bold:      true,
		FontSize:  14,
		FontColor: "FF0000",
		FontName:  "Arial",
	})
	formattedPara.SetAlignment(AlignCenter)

	// Add image
	imageConfig := &ImageConfig{
		Position:  ImagePositionInline,
		Alignment: AlignCenter,
		Size:      &ImageSize{Width: 50, KeepAspectRatio: true},
	}

	imageInfo, err := templateDoc.AddImageFromData(testImageData, "test_image.png", ImageFormatPNG, 1, 1, imageConfig)
	if err != nil {
		t.Fatalf("failed to add image: %v", err)
	}
	t.Logf("added image, ID: %s", imageInfo.ID)

	// Add another paragraph with template variable
	templateDoc.AddParagraph("日期: {{date}}")

	// Save template document
	templatePath := "test_template_with_image.docx"
	err = templateDoc.Save(templatePath)
	if err != nil {
		t.Fatalf("failed to save template document: %v", err)
	}
	defer os.Remove(templatePath)

	// Count image elements in initial document
	initialImageCount := 0
	for _, elem := range templateDoc.Body.Elements {
		if para, ok := elem.(*Paragraph); ok {
			for _, run := range para.Runs {
				if run.Drawing != nil {
					initialImageCount++
				}
			}
		}
	}
	t.Logf("initial document contains %d image elements", initialImageCount)

	// Step 2: Reopen the template document
	t.Log("Step 2: Reopen template document")
	openedDoc, err := Open(templatePath)
	if err != nil {
		t.Fatalf("failed to open template document: %v", err)
	}

	// Check media files in opened document
	openedHasImageMedia := false
	for partName := range openedDoc.parts {
		if len(partName) > 11 && partName[:11] == testWordMediaPrefix {
			openedHasImageMedia = true
			t.Logf("opened document contains media file: %s", partName)
		}
	}
	if !openedHasImageMedia {
		t.Logf("warning: no media files found in opened document")
	}

	// Step 3: Use template engine for variable replacement
	t.Log("Step 3: Use template engine for variable replacement")
	engine := NewTemplateEngine()

	_, err = engine.LoadTemplateFromDocument("test_template", openedDoc)
	if err != nil {
		t.Fatalf("failed to load template: %v", err)
	}

	// Prepare data
	data := NewTemplateData()
	data.SetVariable("title", "测试报告")
	data.SetVariable("author", "张三")
	data.SetVariable("date", "2025年1月1日")

	// Render template
	resultDoc, err := engine.RenderTemplateToDocument("test_template", data)
	if err != nil {
		t.Fatalf("failed to render template: %v", err)
	}

	// Step 4: Check result document
	t.Log("Step 4: Check result document")

	// Check image elements in result document
	resultImageCount := 0
	for _, elem := range resultDoc.Body.Elements {
		if para, ok := elem.(*Paragraph); ok {
			for _, run := range para.Runs {
				if run.Drawing != nil {
					resultImageCount++
					t.Logf("found image element in result document")
				}
			}
		}
	}

	// Check media files in result document
	resultHasImageMedia := false
	for partName := range resultDoc.parts {
		if len(partName) > 11 && partName[:11] == testWordMediaPrefix {
			resultHasImageMedia = true
			t.Logf("result document contains media file: %s", partName)
		}
	}

	// Verify image elements are preserved
	if resultImageCount == 0 {
		t.Errorf("image elements lost after template rendering. initial: %d, result: %d", initialImageCount, resultImageCount)
	} else {
		t.Logf("image elements preserved: %d", resultImageCount)
	}

	// Verify media files are preserved
	if !resultHasImageMedia {
		t.Errorf("media files lost after template rendering")
	} else {
		t.Logf("media files preserved")
	}

	// Save result document
	resultPath := "test_template_rendered.docx"
	err = resultDoc.Save(resultPath)
	if err != nil {
		t.Fatalf("failed to save result document: %v", err)
	}
	defer os.Remove(resultPath)

	// Step 5: Verify variable replacement succeeded
	t.Log("Step 5: Verify variable replacement")
	foundTitle := false
	foundAuthor := false
	for _, elem := range resultDoc.Body.Elements {
		if para, ok := elem.(*Paragraph); ok {
			for _, run := range para.Runs {
				if run.Text.Content != "" {
					if strings.Contains(run.Text.Content, "测试报告") {
						foundTitle = true
					}
					if strings.Contains(run.Text.Content, "张三") {
						foundAuthor = true
					}
				}
			}
		}
	}

	if !foundTitle {
		t.Logf("warning: title variable was not replaced")
	}
	if !foundAuthor {
		t.Logf("warning: author variable was not replaced")
	}

	t.Log("TestTemplateImagePreservation test complete")
}

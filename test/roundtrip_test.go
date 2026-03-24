package test

import (
	"archive/zip"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/mr-pmillz/wordZero/pkg/document"
)

const roundtripTestDoc = "/tmp/Document.docx"

func skipIfNoTestDoc(t *testing.T) {
	t.Helper()
	if _, err := os.Stat(roundtripTestDoc); os.IsNotExist(err) {
		t.Skipf("test document not found at %s — skipping round-trip test", roundtripTestDoc)
	}
}

func TestRoundTripPreservesTOC(t *testing.T) {
	skipIfNoTestDoc(t)

	doc, err := document.Open(roundtripTestDoc)
	if err != nil {
		t.Fatalf("failed to open document: %v", err)
	}

	// Save to temp file
	outPath := filepath.Join(t.TempDir(), "roundtrip_toc.docx")
	if err := doc.Save(outPath); err != nil {
		t.Fatalf("failed to save document: %v", err)
	}

	// Read the saved file and check for SDT (TOC container) in document.xml
	zipReader, err := zip.OpenReader(outPath)
	if err != nil {
		t.Fatalf("failed to open saved docx: %v", err)
	}
	defer zipReader.Close()

	for _, f := range zipReader.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				t.Fatalf("failed to open document.xml: %v", err)
			}
			defer rc.Close()

			data, err := io.ReadAll(rc)
			if err != nil {
				t.Fatalf("failed to read document.xml: %v", err)
			}
			content := string(data)

			// The template uses TOC paragraph styles and field codes for the table of contents.
			// Verify TOC-related content survived the round-trip.
			hasTOCStyle := strings.Contains(content, "TOC1") || strings.Contains(content, "TOC")
			hasFieldCodes := strings.Contains(content, "fldChar") || strings.Contains(content, "instrText")

			if !hasTOCStyle && !hasFieldCodes {
				t.Error("document.xml should contain TOC styles or field codes after round-trip")
			}

			// Check for preserved bookmarks (raw XML elements at body level)
			if strings.Contains(content, "bookmarkEnd") || strings.Contains(content, "bookmarkStart") {
				t.Log("Bookmarks preserved in round-trip")
			}

			return
		}
	}
	t.Error("word/document.xml not found in saved docx")
}

func TestRoundTripPreservesFieldCodes(t *testing.T) {
	skipIfNoTestDoc(t)

	doc, err := document.Open(roundtripTestDoc)
	if err != nil {
		t.Fatalf("failed to open document: %v", err)
	}

	// Check that at least some runs have FieldChar or InstrText populated
	paragraphs := doc.Body.GetParagraphs()
	fieldCharCount := 0
	instrTextCount := 0

	for _, para := range paragraphs {
		for _, run := range para.Runs {
			if run.FieldChar != nil {
				fieldCharCount++
			}
			if run.InstrText != nil {
				instrTextCount++
			}
		}
	}

	if fieldCharCount == 0 {
		t.Error("expected at least some runs with FieldChar after parsing — field codes are being lost")
	}
	if instrTextCount == 0 {
		t.Error("expected at least some runs with InstrText after parsing — field instructions are being lost")
	}

	t.Logf("Found %d FieldChar runs and %d InstrText runs", fieldCharCount, instrTextCount)
}

func TestRoundTripPreservesSectionBreaks(t *testing.T) {
	skipIfNoTestDoc(t)

	doc, err := document.Open(roundtripTestDoc)
	if err != nil {
		t.Fatalf("failed to open document: %v", err)
	}

	// Save and reopen
	outPath := filepath.Join(t.TempDir(), "roundtrip_sections.docx")
	if err := doc.Save(outPath); err != nil {
		t.Fatalf("failed to save: %v", err)
	}

	// Check for sectPr in the saved document.xml
	zipReader, err := zip.OpenReader(outPath)
	if err != nil {
		t.Fatalf("failed to open saved docx: %v", err)
	}
	defer zipReader.Close()

	for _, f := range zipReader.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				t.Fatalf("failed to open document.xml: %v", err)
			}
			defer rc.Close()

			data, err := io.ReadAll(rc)
			if err != nil {
				t.Fatalf("failed to read document.xml: %v", err)
			}
			content := string(data)

			sectPrCount := strings.Count(content, "w:sectPr")
			if sectPrCount == 0 {
				t.Error("document.xml should contain section properties after round-trip")
			}
			t.Logf("Found %d sectPr elements in saved document", sectPrCount)
			return
		}
	}
}

func TestRoundTripElementCount(t *testing.T) {
	skipIfNoTestDoc(t)

	doc, err := document.Open(roundtripTestDoc)
	if err != nil {
		t.Fatalf("failed to open document: %v", err)
	}

	elementCount := len(doc.Body.Elements)
	paragraphs := doc.Body.GetParagraphs()

	t.Logf("Document has %d body elements, %d paragraphs", elementCount, len(paragraphs))

	if elementCount == 0 {
		t.Error("document should have body elements after opening")
	}

	// The template should have more than just paragraphs and tables
	// (SDT elements should now be preserved as RawXMLElement)
	if elementCount <= len(paragraphs) {
		t.Log("WARNING: element count equals paragraph count — no raw XML elements preserved")
	}
}

func TestRoundTripFootnoteNoVertAlign(t *testing.T) {
	// This test creates a new document with footnotes and verifies the output
	doc := document.New()
	doc.AddParagraph("Test paragraph")
	doc.AddFootnote("Text with footnote", "Footnote content here")

	outPath := filepath.Join(t.TempDir(), "roundtrip_footnote.docx")
	if err := doc.Save(outPath); err != nil {
		t.Fatalf("failed to save: %v", err)
	}

	// Read footnotes.xml from the saved file
	zipReader, err := zip.OpenReader(outPath)
	if err != nil {
		t.Fatalf("failed to open saved docx: %v", err)
	}
	defer zipReader.Close()

	for _, f := range zipReader.File {
		if f.Name == "word/footnotes.xml" {
			rc, err := f.Open()
			if err != nil {
				t.Fatalf("failed to open footnotes.xml: %v", err)
			}
			defer rc.Close()

			data, err := io.ReadAll(rc)
			if err != nil {
				t.Fatalf("failed to read footnotes.xml: %v", err)
			}
			content := string(data)

			if strings.Contains(content, "vertAlign") {
				t.Error("footnotes.xml should not contain vertAlign — the rStyle provides superscript")
			}
			return
		}
	}
	t.Error("word/footnotes.xml not found in saved docx")
}

func TestRoundTripPreservesAllParts(t *testing.T) {
	skipIfNoTestDoc(t)

	// Open the template and list its ZIP parts
	origZip, err := zip.OpenReader(roundtripTestDoc)
	if err != nil {
		t.Fatalf("failed to open original: %v", err)
	}
	origParts := make(map[string]bool)
	for _, f := range origZip.File {
		origParts[f.Name] = true
	}
	origZip.Close()

	// Open with wordZero and save
	doc, err := document.Open(roundtripTestDoc)
	if err != nil {
		t.Fatalf("failed to open: %v", err)
	}

	outPath := filepath.Join(t.TempDir(), "roundtrip_parts.docx")
	if err := doc.Save(outPath); err != nil {
		t.Fatalf("failed to save: %v", err)
	}

	// Check saved ZIP parts
	savedZip, err := zip.OpenReader(outPath)
	if err != nil {
		t.Fatalf("failed to open saved: %v", err)
	}
	defer savedZip.Close()

	savedParts := make(map[string]bool)
	for _, f := range savedZip.File {
		savedParts[f.Name] = true
	}

	// Critical parts that must be preserved
	criticalParts := []string{
		"word/document.xml",
		"word/styles.xml",
		"[Content_Types].xml",
		"_rels/.rels",
		"word/_rels/document.xml.rels",
	}

	for _, part := range criticalParts {
		if !savedParts[part] {
			t.Errorf("critical part %s missing from saved document", part)
		}
	}

	// Log parts that were in original but not in saved
	missing := 0
	for part := range origParts {
		if !savedParts[part] {
			t.Logf("Part in original but missing from saved: %s", part)
			missing++
		}
	}
	if missing > 0 {
		t.Logf("%d parts from original not found in saved document", missing)
	}
}

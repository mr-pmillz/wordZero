package document

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"path/filepath"
	"strings"
	"testing"
)

const (
	testFootnoteReferenceStyle = "FootnoteReference"
	testFootnotesXMLPath       = "word/footnotes.xml"
)

// --- Basic Add/Remove ---

func TestAddFootnote_Basic(t *testing.T) {
	doc := New()

	err := doc.AddFootnote("body text", "footnote content")
	if err != nil {
		t.Fatalf("AddFootnote failed: %v", err)
	}

	if count := doc.GetFootnoteCount(); count != 1 {
		t.Errorf("expected footnote count 1, got %d", count)
	}

	footnotesXML, exists := doc.parts[testFootnotesXMLPath]
	if !exists {
		t.Fatal("word/footnotes.xml not created")
	}

	if !strings.Contains(string(footnotesXML), "footnote content") {
		t.Error("footnotes.xml does not contain the footnote text")
	}

	// Verify the paragraph contains the body text and a correct footnote reference run
	paragraphs := doc.Body.GetParagraphs()
	if len(paragraphs) == 0 {
		t.Fatal("no paragraphs added to body")
	}

	lastPara := paragraphs[len(paragraphs)-1]
	if len(lastPara.Runs) < 2 {
		t.Fatalf("expected at least 2 runs (text + reference), got %d", len(lastPara.Runs))
	}

	// First run should contain body text
	if lastPara.Runs[0].Text.Content != "body text" {
		t.Errorf("first run should contain body text, got: %s", lastPara.Runs[0].Text.Content)
	}

	// Second run should contain FootnoteReference
	refRun := lastPara.Runs[1]
	if refRun.FootnoteReference == nil {
		t.Fatal("reference run should have FootnoteReference set")
	}
	if refRun.FootnoteReference.ID != "1" {
		t.Errorf("expected footnote reference ID '1', got '%s'", refRun.FootnoteReference.ID)
	}
}

func TestAddEndnote_Basic(t *testing.T) {
	doc := New()

	err := doc.AddEndnote("body text", "endnote content")
	if err != nil {
		t.Fatalf("AddEndnote failed: %v", err)
	}

	if count := doc.GetEndnoteCount(); count != 1 {
		t.Errorf("expected endnote count 1, got %d", count)
	}

	endnotesXML, exists := doc.parts["word/endnotes.xml"]
	if !exists {
		t.Fatal("word/endnotes.xml not created")
	}

	if !strings.Contains(string(endnotesXML), "endnote content") {
		t.Error("endnotes.xml does not contain the endnote text")
	}

	// Verify the endnote reference uses EndnoteReference
	paragraphs := doc.Body.GetParagraphs()
	if len(paragraphs) == 0 {
		t.Fatal("no paragraphs added to body")
	}
	lastPara := paragraphs[len(paragraphs)-1]
	if len(lastPara.Runs) < 2 {
		t.Fatalf("expected at least 2 runs, got %d", len(lastPara.Runs))
	}

	refRun := lastPara.Runs[1]
	if refRun.EndnoteReference == nil {
		t.Fatal("reference run should have EndnoteReference set")
	}
	if refRun.EndnoteReference.ID != "1" {
		t.Errorf("expected endnote reference ID '1', got '%s'", refRun.EndnoteReference.ID)
	}
}

func TestAddFootnoteToRun(t *testing.T) {
	doc := New()

	para := doc.AddParagraph("some text")
	if len(para.Runs) == 0 {
		t.Fatal("paragraph has no runs")
	}
	run := &para.Runs[0]

	err := doc.AddFootnoteToRun(run, "run footnote")
	if err != nil {
		t.Fatalf("AddFootnoteToRun failed: %v", err)
	}

	if count := doc.GetFootnoteCount(); count != 1 {
		t.Errorf("expected footnote count 1, got %d", count)
	}

	// AddFootnoteToRun sets FootnoteReference on the run
	if run.FootnoteReference == nil {
		t.Fatal("run should have FootnoteReference set")
	}
	if run.FootnoteReference.ID != "1" {
		t.Errorf("expected ID '1', got '%s'", run.FootnoteReference.ID)
	}
}

func TestAddFootnoteToParagraph(t *testing.T) {
	doc := New()

	para := doc.AddParagraph("text with footnote")

	err := doc.AddFootnoteToParagraph(para, "paragraph footnote content")
	if err != nil {
		t.Fatalf("AddFootnoteToParagraph failed: %v", err)
	}

	if count := doc.GetFootnoteCount(); count != 1 {
		t.Errorf("expected footnote count 1, got %d", count)
	}

	// Should have appended a reference run
	if len(para.Runs) < 2 {
		t.Fatalf("expected at least 2 runs, got %d", len(para.Runs))
	}

	refRun := para.Runs[len(para.Runs)-1]
	if refRun.FootnoteReference == nil {
		t.Fatal("last run should have FootnoteReference")
	}
	if refRun.Properties == nil || refRun.Properties.RunStyle == nil {
		t.Fatal("reference run should have RunStyle property")
	}
	if refRun.Properties.RunStyle.Val != testFootnoteReferenceStyle {
		t.Errorf("expected RunStyle 'FootnoteReference', got '%s'", refRun.Properties.RunStyle.Val)
	}
}

func TestAddEndnoteToParagraph(t *testing.T) {
	doc := New()

	para := doc.AddParagraph("text with endnote")

	err := doc.AddEndnoteToParagraph(para, "paragraph endnote content")
	if err != nil {
		t.Fatalf("AddEndnoteToParagraph failed: %v", err)
	}

	if count := doc.GetEndnoteCount(); count != 1 {
		t.Errorf("expected endnote count 1, got %d", count)
	}

	refRun := para.Runs[len(para.Runs)-1]
	if refRun.EndnoteReference == nil {
		t.Fatal("last run should have EndnoteReference")
	}
}

// --- Multiple Notes ---

func TestMultipleFootnotes_IncrementingIDs(t *testing.T) {
	doc := New()

	for i := 0; i < 3; i++ {
		err := doc.AddFootnote("text", "footnote")
		if err != nil {
			t.Fatalf("AddFootnote %d failed: %v", i, err)
		}
	}

	if count := doc.GetFootnoteCount(); count != 3 {
		t.Errorf("expected 3 footnotes, got %d", count)
	}

	// Verify paragraph references contain incrementing IDs
	paragraphs := doc.Body.GetParagraphs()
	expectedIDs := []string{"1", "2", "3"}
	for i := 0; i < 3; i++ {
		if i >= len(paragraphs) {
			t.Fatalf("missing paragraph %d", i)
		}
		// Reference run is the last run in the paragraph
		runs := paragraphs[i].Runs
		refRun := runs[len(runs)-1]
		if refRun.FootnoteReference == nil {
			t.Fatalf("paragraph %d: missing FootnoteReference", i)
		}
		if refRun.FootnoteReference.ID != expectedIDs[i] {
			t.Errorf("paragraph %d: expected ID '%s', got '%s'", i, expectedIDs[i], refRun.FootnoteReference.ID)
		}
	}
}

func TestMultipleEndnotes_IncrementingIDs(t *testing.T) {
	doc := New()

	for i := 0; i < 3; i++ {
		err := doc.AddEndnote("text", "endnote")
		if err != nil {
			t.Fatalf("AddEndnote %d failed: %v", i, err)
		}
	}

	if count := doc.GetEndnoteCount(); count != 3 {
		t.Errorf("expected 3 endnotes, got %d", count)
	}
}

func TestMixedFootnotesAndEndnotes(t *testing.T) {
	doc := New()

	doc.AddFootnote("text1", "fn1")
	doc.AddFootnote("text2", "fn2")
	doc.AddEndnote("text3", "en1")
	doc.AddEndnote("text4", "en2")

	if count := doc.GetFootnoteCount(); count != 2 {
		t.Errorf("expected 2 footnotes, got %d", count)
	}
	if count := doc.GetEndnoteCount(); count != 2 {
		t.Errorf("expected 2 endnotes, got %d", count)
	}

	if _, exists := doc.parts[testFootnotesXMLPath]; !exists {
		t.Error("word/footnotes.xml not created")
	}
	if _, exists := doc.parts["word/endnotes.xml"]; !exists {
		t.Error("word/endnotes.xml not created")
	}
}

// --- Remove ---

func TestRemoveFootnote_Success(t *testing.T) {
	doc := New()
	doc.AddFootnote("text1", "fn1")
	doc.AddFootnote("text2", "fn2")

	err := doc.RemoveFootnote("1")
	if err != nil {
		t.Fatalf("RemoveFootnote failed: %v", err)
	}

	if count := doc.GetFootnoteCount(); count != 1 {
		t.Errorf("expected 1 footnote after removal, got %d", count)
	}
}

func TestRemoveFootnote_NonExistent(t *testing.T) {
	doc := New()
	doc.AddFootnote("text", "fn")

	err := doc.RemoveFootnote("999")
	if err == nil {
		t.Fatal("expected error when removing non-existent footnote")
	}
}

func TestRemoveEndnote_Success(t *testing.T) {
	doc := New()
	doc.AddEndnote("text1", "en1")
	doc.AddEndnote("text2", "en2")

	err := doc.RemoveEndnote("1")
	if err != nil {
		t.Fatalf("RemoveEndnote failed: %v", err)
	}

	if count := doc.GetEndnoteCount(); count != 1 {
		t.Errorf("expected 1 endnote after removal, got %d", count)
	}
}

func TestRemoveEndnote_NonExistent(t *testing.T) {
	doc := New()
	doc.AddEndnote("text", "en")

	err := doc.RemoveEndnote("999")
	if err == nil {
		t.Fatal("expected error when removing non-existent endnote")
	}
}

// --- Counts ---

func TestGetFootnoteCount_Empty(t *testing.T) {
	doc := New()
	if count := doc.GetFootnoteCount(); count != 0 {
		t.Errorf("expected 0 footnotes on new document, got %d", count)
	}
}

func TestGetEndnoteCount_Empty(t *testing.T) {
	doc := New()
	if count := doc.GetEndnoteCount(); count != 0 {
		t.Errorf("expected 0 endnotes on new document, got %d", count)
	}
}

// --- Config ---

func TestSetFootnoteConfig_AllFormats(t *testing.T) {
	formats := []FootnoteNumberFormat{
		FootnoteFormatDecimal, FootnoteFormatLowerRoman, FootnoteFormatUpperRoman,
		FootnoteFormatLowerLetter, FootnoteFormatUpperLetter, FootnoteFormatSymbol,
	}

	for _, format := range formats {
		doc := New()
		config := &FootnoteConfig{
			NumberFormat: format, StartNumber: 1,
			RestartEach: FootnoteRestartContinuous, Position: FootnotePositionPageBottom,
		}
		err := doc.SetFootnoteConfig(config)
		if err != nil {
			t.Fatalf("SetFootnoteConfig with format %s failed: %v", format, err)
		}
		settingsXML, exists := doc.parts["word/settings.xml"]
		if !exists {
			t.Fatalf("settings.xml not created for format %s", format)
		}
		if !strings.Contains(string(settingsXML), string(format)) {
			t.Errorf("settings.xml does not contain format %s", format)
		}
	}
}

func TestSetFootnoteConfig_AllPositions(t *testing.T) {
	positions := []FootnotePosition{
		FootnotePositionPageBottom, FootnotePositionBeneathText,
		FootnotePositionSectionEnd, FootnotePositionDocumentEnd,
	}
	for _, position := range positions {
		doc := New()
		config := &FootnoteConfig{
			NumberFormat: FootnoteFormatDecimal, StartNumber: 1,
			RestartEach: FootnoteRestartContinuous, Position: position,
		}
		err := doc.SetFootnoteConfig(config)
		if err != nil {
			t.Fatalf("SetFootnoteConfig with position %s failed: %v", position, err)
		}
		settingsXML := doc.parts["word/settings.xml"]
		if !strings.Contains(string(settingsXML), string(position)) {
			t.Errorf("settings.xml does not contain position %s", position)
		}
	}
}

func TestSetFootnoteConfig_AllRestartRules(t *testing.T) {
	restarts := []FootnoteRestart{
		FootnoteRestartContinuous, FootnoteRestartEachSection, FootnoteRestartEachPage,
	}
	for _, restart := range restarts {
		doc := New()
		config := &FootnoteConfig{
			NumberFormat: FootnoteFormatDecimal, StartNumber: 1,
			RestartEach: restart, Position: FootnotePositionPageBottom,
		}
		err := doc.SetFootnoteConfig(config)
		if err != nil {
			t.Fatalf("SetFootnoteConfig with restart %s failed: %v", restart, err)
		}
		settingsXML := doc.parts["word/settings.xml"]
		if !strings.Contains(string(settingsXML), string(restart)) {
			t.Errorf("settings.xml does not contain restart rule %s", restart)
		}
	}
}

func TestSetFootnoteConfig_NilDefaults(t *testing.T) {
	doc := New()
	err := doc.SetFootnoteConfig(nil)
	if err != nil {
		t.Fatalf("SetFootnoteConfig(nil) failed: %v", err)
	}
	if _, exists := doc.parts["word/settings.xml"]; !exists {
		t.Error("settings.xml not created with nil config")
	}
}

func TestDefaultFootnoteConfig_Values(t *testing.T) {
	config := DefaultFootnoteConfig()
	if config.NumberFormat != FootnoteFormatDecimal {
		t.Errorf("expected decimal format, got %s", config.NumberFormat)
	}
	if config.StartNumber != 1 {
		t.Errorf("expected start number 1, got %d", config.StartNumber)
	}
	if config.RestartEach != FootnoteRestartContinuous {
		t.Errorf("expected continuous restart, got %s", config.RestartEach)
	}
	if config.Position != FootnotePositionPageBottom {
		t.Errorf("expected page bottom position, got %s", config.Position)
	}
}

// --- Initialization ---

func TestFootnoteInitialization_ContentType(t *testing.T) {
	doc := New()
	doc.AddFootnote("text", "footnote")

	found := false
	for _, override := range doc.contentTypes.Overrides {
		if override.PartName == "/word/footnotes.xml" {
			found = true
			if override.ContentType != "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml" {
				t.Errorf("wrong content type: %s", override.ContentType)
			}
			break
		}
	}
	if !found {
		t.Error("footnotes content type override not found")
	}
}

func TestEndnoteInitialization_ContentType(t *testing.T) {
	doc := New()
	doc.AddEndnote("text", "endnote")

	found := false
	for _, override := range doc.contentTypes.Overrides {
		if override.PartName == "/word/endnotes.xml" {
			found = true
			break
		}
	}
	if !found {
		t.Error("endnotes content type override not found")
	}
}

func TestFootnoteInitialization_Relationship(t *testing.T) {
	doc := New()
	doc.AddFootnote("text", "footnote")

	found := false
	for _, rel := range doc.documentRelationships.Relationships {
		if strings.Contains(rel.Type, "footnotes") {
			found = true
			if rel.Target != "footnotes.xml" {
				t.Errorf("wrong relationship target: %s", rel.Target)
			}
			break
		}
	}
	if !found {
		t.Error("footnotes relationship not found in document relationships")
	}
}

func TestEndnoteInitialization_Relationship(t *testing.T) {
	doc := New()
	doc.AddEndnote("text", "endnote")

	found := false
	for _, rel := range doc.documentRelationships.Relationships {
		if strings.Contains(rel.Type, "endnotes") {
			found = true
			break
		}
	}
	if !found {
		t.Error("endnotes relationship not found in document relationships")
	}
}

func TestFootnote_IdempotentInitialization(t *testing.T) {
	doc := New()
	doc.AddFootnote("text1", "fn1")
	doc.AddFootnote("text2", "fn2")
	doc.AddFootnote("text3", "fn3")

	count := 0
	for _, override := range doc.contentTypes.Overrides {
		if override.PartName == "/word/footnotes.xml" {
			count++
		}
	}
	if count != 1 {
		t.Errorf("expected exactly 1 footnotes content type override, got %d", count)
	}
}

// --- XML Structure ---

func TestFootnoteSeparatorInXML(t *testing.T) {
	doc := New()
	doc.AddFootnote("text", "footnote")

	footnotesXML := string(doc.parts[testFootnotesXMLPath])

	if !strings.Contains(footnotesXML, `w:type="separator"`) {
		t.Error("footnotes.xml missing separator footnote type")
	}
	if !strings.Contains(footnotesXML, `w:id="-1"`) {
		t.Error("footnotes.xml missing separator footnote with id=-1")
	}
	// Should use w:separator element, not w:footnoteRef
	if !strings.Contains(footnotesXML, "w:separator") {
		t.Error("separator footnote should use w:separator element")
	}
	// Should have continuation separator with id=0
	if !strings.Contains(footnotesXML, `w:type="continuationSeparator"`) {
		t.Error("footnotes.xml missing continuationSeparator")
	}
	if !strings.Contains(footnotesXML, `w:id="0"`) {
		t.Error("footnotes.xml missing continuation separator with id=0")
	}
}

func TestFootnoteXML_ValidStructure(t *testing.T) {
	doc := New()
	doc.AddFootnote("text", "my footnote content")

	footnotesXML := doc.parts[testFootnotesXMLPath]

	// Verify the XML is well-formed
	decoder := xml.NewDecoder(bytes.NewReader(footnotesXML))
	for {
		_, err := decoder.Token()
		if err != nil {
			if err.Error() == "EOF" {
				break
			}
			t.Fatalf("footnotes.xml is not valid XML: %v", err)
		}
	}

	xmlStr := string(footnotesXML)
	if !strings.Contains(xmlStr, "my footnote content") {
		t.Error("missing user footnote content")
	}
}

func TestFootnoteXML_ContainsFootnoteRef(t *testing.T) {
	doc := New()
	doc.AddFootnote("text", "footnote content")

	xmlStr := string(doc.parts[testFootnotesXMLPath])

	// Verify it contains the w:footnoteRef self-reference element
	if !strings.Contains(xmlStr, "w:footnoteRef") {
		t.Error("footnotes.xml should contain w:footnoteRef element")
	}
}

func TestFootnoteXML_ContainsFootnoteTextStyle(t *testing.T) {
	doc := New()
	doc.AddFootnote("text", "footnote content")

	xmlStr := string(doc.parts[testFootnotesXMLPath])

	// Verify it contains the FootnoteText paragraph style
	if !strings.Contains(xmlStr, "FootnoteText") {
		t.Error("footnotes.xml should contain FootnoteText paragraph style")
	}
}

func TestFootnoteReferenceHasRunStyle(t *testing.T) {
	doc := New()
	doc.AddFootnote("text", "footnote")

	paragraphs := doc.Body.GetParagraphs()
	lastPara := paragraphs[len(paragraphs)-1]
	refRun := lastPara.Runs[len(lastPara.Runs)-1]

	if refRun.Properties == nil {
		t.Fatal("reference run should have Properties")
	}
	// Body reference runs use rStyle only — the style itself handles superscript
	if refRun.Properties.RunStyle == nil {
		t.Fatal("reference run should have RunStyle")
	}
	if refRun.Properties.RunStyle.Val != testFootnoteReferenceStyle {
		t.Errorf("expected RunStyle 'FootnoteReference', got '%s'", refRun.Properties.RunStyle.Val)
	}
	// VerticalAlign should NOT be set on body references (the style provides it)
	if refRun.Properties.VerticalAlign != nil {
		t.Error("body reference run should not have VerticalAlign (style provides superscript)")
	}
}

// --- Save and Verify ---

func TestFootnote_SaveAndVerifyZip(t *testing.T) {
	doc := New()
	doc.AddParagraph("Introduction")
	doc.AddFootnote("This has a footnote", "The footnote content")
	doc.AddEndnote("This has an endnote", "The endnote content")

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "footnotes_test.docx")

	err := doc.Save(path)
	if err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	zipReader, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("failed to open saved docx as zip: %v", err)
	}
	defer zipReader.Close()

	expectedParts := map[string]bool{
		testFootnotesXMLPath:           false,
		"word/endnotes.xml":            false,
		"[Content_Types].xml":          false,
		"word/_rels/document.xml.rels": false,
	}

	for _, f := range zipReader.File {
		if _, ok := expectedParts[f.Name]; ok {
			expectedParts[f.Name] = true
		}
	}

	for part, found := range expectedParts {
		if !found {
			t.Errorf("expected part %s not found in saved docx", part)
		}
	}
}

func TestFootnote_SaveToBytes(t *testing.T) {
	doc := New()
	doc.AddFootnote("text", "footnote content")

	data, err := doc.ToBytes()
	if err != nil {
		t.Fatalf("ToBytes failed: %v", err)
	}

	reader := bytes.NewReader(data)
	zipReader, err := zip.NewReader(reader, int64(len(data)))
	if err != nil {
		t.Fatalf("output is not a valid zip: %v", err)
	}

	foundFootnotes := false
	for _, f := range zipReader.File {
		if f.Name == testFootnotesXMLPath {
			foundFootnotes = true
			break
		}
	}
	if !foundFootnotes {
		t.Error("word/footnotes.xml not found in output bytes")
	}
}

// --- Document Isolation ---

func TestDocumentIsolation(t *testing.T) {
	doc1 := New()
	doc2 := New()

	doc1.AddFootnote("doc1 text", "doc1 footnote 1")
	doc1.AddFootnote("doc1 text", "doc1 footnote 2")

	if count := doc2.GetFootnoteCount(); count != 0 {
		t.Errorf("doc2 should have 0 footnotes, got %d (state leaked from doc1)", count)
	}

	doc2.AddFootnote("doc2 text", "doc2 footnote 1")

	if count := doc1.GetFootnoteCount(); count != 2 {
		t.Errorf("doc1 should have 2 footnotes, got %d", count)
	}
	if count := doc2.GetFootnoteCount(); count != 1 {
		t.Errorf("doc2 should have 1 footnote, got %d", count)
	}

	// doc2 footnote IDs should start from 1
	paragraphs := doc2.Body.GetParagraphs()
	lastPara := paragraphs[len(paragraphs)-1]
	refRun := lastPara.Runs[len(lastPara.Runs)-1]
	if refRun.FootnoteReference == nil {
		t.Fatal("doc2 reference run missing FootnoteReference")
	}
	if refRun.FootnoteReference.ID != "1" {
		t.Errorf("doc2 footnote should have ID '1', got '%s'", refRun.FootnoteReference.ID)
	}
}

func TestDocumentIsolation_Endnotes(t *testing.T) {
	doc1 := New()
	doc2 := New()

	doc1.AddEndnote("doc1 text", "doc1 endnote")

	if count := doc2.GetEndnoteCount(); count != 0 {
		t.Errorf("doc2 should have 0 endnotes, got %d", count)
	}
}

// --- Edge Cases ---

func TestAddFootnote_EmptyText(t *testing.T) {
	doc := New()
	err := doc.AddFootnote("", "footnote content")
	if err != nil {
		t.Fatalf("AddFootnote with empty body text failed: %v", err)
	}
	if count := doc.GetFootnoteCount(); count != 1 {
		t.Errorf("expected 1 footnote, got %d", count)
	}
}

func TestAddFootnote_EmptyFootnoteText(t *testing.T) {
	doc := New()
	err := doc.AddFootnote("body text", "")
	if err != nil {
		t.Fatalf("AddFootnote with empty footnote text failed: %v", err)
	}
	if count := doc.GetFootnoteCount(); count != 1 {
		t.Errorf("expected 1 footnote, got %d", count)
	}
}

// --- Settings ---

func TestFootnoteConfig_CreatesSettingsRelationship(t *testing.T) {
	doc := New()
	doc.SetFootnoteConfig(DefaultFootnoteConfig())

	found := false
	for _, rel := range doc.documentRelationships.Relationships {
		if strings.Contains(rel.Type, "settings") {
			found = true
			break
		}
	}
	if !found {
		t.Error("settings relationship not found in document relationships after SetFootnoteConfig")
	}
}

func TestFootnoteConfig_CreatesSettingsContentType(t *testing.T) {
	doc := New()
	doc.SetFootnoteConfig(DefaultFootnoteConfig())

	found := false
	for _, override := range doc.contentTypes.Overrides {
		if override.PartName == "/word/settings.xml" {
			found = true
			break
		}
	}
	if !found {
		t.Error("settings content type override not found after SetFootnoteConfig")
	}
}

package document

import (
	"bytes"
	"strings"
	"testing"
)

func TestDefaultLanguage_IsEnglish(t *testing.T) {
	var buf bytes.Buffer
	logger := NewLogger(LogLevelDebug, &buf)

	logger.DebugMsg(MsgCreatingNewDocument)

	output := buf.String()
	if !strings.Contains(output, "Creating new document") {
		t.Errorf("default language should produce English output, got: %s", output)
	}
}

func TestSetLanguage_English(t *testing.T) {
	var buf bytes.Buffer
	logger := NewLogger(LogLevelDebug, &buf)
	logger.SetLanguage(LogLanguageEN)

	logger.DebugMsg(MsgCreatingNewDocument)

	output := buf.String()
	if !strings.Contains(output, "Creating new document") {
		t.Errorf("English language should produce English output, got: %s", output)
	}
}

func TestSetLanguage_SwitchAtRuntime(t *testing.T) {
	var buf bytes.Buffer
	logger := NewLogger(LogLevelDebug, &buf)

	// Default is English
	logger.DebugMsg(MsgCreatingNewDocument)
	if !strings.Contains(buf.String(), "Creating new document") {
		t.Error("expected English output by default")
	}

	// Switch to Chinese
	buf.Reset()
	logger.SetLanguage(LogLanguageZH)
	logger.DebugMsg(MsgCreatingNewDocument)
	if !strings.Contains(buf.String(), "创建新文档") {
		t.Error("expected Chinese output after switching language")
	}

	// Switch back to English
	buf.Reset()
	logger.SetLanguage(LogLanguageEN)
	logger.DebugMsg(MsgCreatingNewDocument)
	if !strings.Contains(buf.String(), "Creating new document") {
		t.Error("expected English output after switching back")
	}
}

func TestMsgf_WithFormatArgs(t *testing.T) {
	var buf bytes.Buffer
	logger := NewLogger(LogLevelInfo, &buf)

	// Default English with format args
	logger.InfoMsgf(MsgOpeningDocumentPath, "test.docx")
	if !strings.Contains(buf.String(), "test.docx") {
		t.Errorf("format args not resolved in English, got: %s", buf.String())
	}
	if !strings.Contains(buf.String(), "Opening document") {
		t.Errorf("expected English message with args, got: %s", buf.String())
	}

	// Chinese with format args
	buf.Reset()
	logger.SetLanguage(LogLanguageZH)
	logger.InfoMsgf(MsgOpeningDocumentPath, "test.docx")
	if !strings.Contains(buf.String(), "test.docx") {
		t.Errorf("format args not resolved in Chinese, got: %s", buf.String())
	}
	if !strings.Contains(buf.String(), "正在打开文档") {
		t.Errorf("expected Chinese message with args, got: %s", buf.String())
	}
}

func TestMsgf_MultipleArgs(t *testing.T) {
	var buf bytes.Buffer
	logger := NewLogger(LogLevelDebug, &buf)
	logger.SetLanguage(LogLanguageEN)

	logger.DebugMsgf(MsgAddingHeadingParagraph, "Title", 1, "Heading1", "bookmark1")

	output := buf.String()
	if !strings.Contains(output, "Title") {
		t.Errorf("missing arg 'Title' in output: %s", output)
	}
	if !strings.Contains(output, "Adding heading paragraph") {
		t.Errorf("expected English heading message, got: %s", output)
	}
}

func TestUnknownMsgKey_FallsBackToKey(t *testing.T) {
	var buf bytes.Buffer
	logger := NewLogger(LogLevelDebug, &buf)

	unknownKey := MsgKey("nonexistent_key")
	logger.DebugMsg(unknownKey)

	output := buf.String()
	if !strings.Contains(output, "nonexistent_key") {
		t.Errorf("unknown key should fall back to key string, got: %s", output)
	}
}

func TestAllLogLevels_WithMsgKey(t *testing.T) {
	var buf bytes.Buffer
	logger := NewLogger(LogLevelDebug, &buf)
	logger.SetLanguage(LogLanguageEN)

	logger.DebugMsg(MsgCreatingNewDocument)
	if !strings.Contains(buf.String(), "DEBUG") {
		t.Error("DebugMsg should produce DEBUG level")
	}

	buf.Reset()
	logger.InfoMsg(MsgTableContentCleared)
	if !strings.Contains(buf.String(), "INFO") {
		t.Error("InfoMsg should produce INFO level")
	}

	buf.Reset()
	logger.WarnMsg(MsgOutlineLevelAdjusted)
	if !strings.Contains(buf.String(), "WARN") {
		t.Error("WarnMsg should produce WARN level")
	}

	buf.Reset()
	logger.ErrorMsg(MsgFailedToSerializeDocument)
	if !strings.Contains(buf.String(), "ERROR") {
		t.Error("ErrorMsg should produce ERROR level")
	}
}

func TestSetGlobalLanguage(t *testing.T) {
	// Save original state
	origLang := defaultLogger.language
	defer func() { defaultLogger.language = origLang }()

	var buf bytes.Buffer
	origOutput := defaultLogger.output
	defer func() { defaultLogger.SetOutput(origOutput) }()
	defaultLogger.SetOutput(&buf)

	// Ensure starting at Debug level
	origLevel := defaultLogger.level
	defer func() { defaultLogger.level = origLevel }()
	defaultLogger.SetLevel(LogLevelDebug)

	// Test using global functions
	SetGlobalLanguage(LogLanguageEN)
	DebugMsg(MsgCreatingNewDocument)

	output := buf.String()
	if !strings.Contains(output, "Creating new document") {
		t.Errorf("SetGlobalLanguage should affect global logger, got: %s", output)
	}
}

func TestLogLevel_FiltersMsgCalls(t *testing.T) {
	var buf bytes.Buffer
	logger := NewLogger(LogLevelInfo, &buf)

	// Debug should be filtered
	logger.DebugMsg(MsgCreatingNewDocument)
	if buf.Len() > 0 {
		t.Error("Debug message should be filtered at Info level")
	}

	// Info should pass through
	logger.InfoMsg(MsgTableContentCleared)
	if buf.Len() == 0 {
		t.Error("Info message should pass at Info level")
	}
}

func TestMessageCatalog_Completeness(t *testing.T) {
	// Verify every ZH message has a corresponding EN message
	for key := range messagesZH {
		if _, exists := messagesEN[key]; !exists {
			t.Errorf("message key %s exists in ZH but not EN", key)
		}
	}

	// Verify every EN message has a corresponding ZH message
	for key := range messagesEN {
		if _, exists := messagesZH[key]; !exists {
			t.Errorf("message key %s exists in EN but not ZH", key)
		}
	}
}

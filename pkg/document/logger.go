// Package document provides the logging system.
package document

import (
	"fmt"
	"io"
	"log"
	"os"
	"time"
)

// LogLevel represents the logging level.
type LogLevel int

const (
	// LogLevelDebug is the debug level.
	LogLevelDebug LogLevel = iota
	// LogLevelInfo is the info level.
	LogLevelInfo
	// LogLevelWarn is the warn level.
	LogLevelWarn
	// LogLevelError is the error level.
	LogLevelError
	// LogLevelSilent is the silent level.
	LogLevelSilent
)

// String returns the string representation of the log level.
func (l LogLevel) String() string {
	switch l {
	case LogLevelDebug:
		return "DEBUG"
	case LogLevelInfo:
		return "INFO"
	case LogLevelWarn:
		return "WARN"
	case LogLevelError:
		return "ERROR"
	case LogLevelSilent:
		return "SILENT"
	default:
		return "UNKNOWN"
	}
}

// Logger is the log recorder.
type Logger struct {
	level    LogLevel    // log level
	output   io.Writer   // output destination
	logger   *log.Logger // internal logger
	language LogLanguage // log language (default English)
}

// defaultLogger is the default global logger.
var defaultLogger = NewLogger(LogLevelInfo, os.Stdout)

// NewLogger creates a new logger.
func NewLogger(level LogLevel, output io.Writer) *Logger {
	return &Logger{
		level:    level,
		output:   output,
		logger:   log.New(output, "", 0),
		language: LogLanguageEN,
	}
}

// SetLevel sets the log level.
func (l *Logger) SetLevel(level LogLevel) {
	l.level = level
}

// SetOutput sets the output destination.
func (l *Logger) SetOutput(output io.Writer) {
	l.output = output
	l.logger.SetOutput(output)
}

// SetLanguage sets the log message language.
func (l *Logger) SetLanguage(lang LogLanguage) {
	l.language = lang
}

// logf performs formatted log output.
func (l *Logger) logf(level LogLevel, format string, args ...interface{}) {
	if l.level > level {
		return
	}

	timestamp := time.Now().Format("2006-01-02 15:04:05")
	message := fmt.Sprintf(format, args...)
	l.logger.Printf("[%s] %s - %s", timestamp, level.String(), message)
}

// logMsgf performs formatted log output using a message key (supports multiple languages).
func (l *Logger) logMsgf(level LogLevel, key MsgKey, args ...interface{}) {
	if l.level > level {
		return
	}

	format := getMessage(key, l.language)
	timestamp := time.Now().Format("2006-01-02 15:04:05")
	message := fmt.Sprintf(format, args...)
	l.logger.Printf("[%s] %s - %s", timestamp, level.String(), message)
}

// --- Raw format methods (retained for backward compatibility) ---

// Debugf outputs a debug log message.
func (l *Logger) Debugf(format string, args ...interface{}) {
	l.logf(LogLevelDebug, format, args...)
}

// Infof outputs an info log message.
func (l *Logger) Infof(format string, args ...interface{}) {
	l.logf(LogLevelInfo, format, args...)
}

// Warnf outputs a warning log message.
func (l *Logger) Warnf(format string, args ...interface{}) {
	l.logf(LogLevelWarn, format, args...)
}

// Errorf outputs an error log message.
func (l *Logger) Errorf(format string, args ...interface{}) {
	l.logf(LogLevelError, format, args...)
}

// Debug outputs a debug log message.
func (l *Logger) Debug(msg string) {
	l.Debugf("%s", msg)
}

// Info outputs an info log message.
func (l *Logger) Info(msg string) {
	l.Infof("%s", msg)
}

// Warn outputs a warning log message.
func (l *Logger) Warn(msg string) {
	l.Warnf("%s", msg)
}

// Error outputs an error log message.
func (l *Logger) Error(msg string) {
	l.Errorf("%s", msg)
}

// --- Message key methods (supports multiple languages) ---

// DebugMsg outputs a debug log using a message key.
func (l *Logger) DebugMsg(key MsgKey) {
	l.logMsgf(LogLevelDebug, key)
}

// InfoMsg outputs an info log using a message key.
func (l *Logger) InfoMsg(key MsgKey) {
	l.logMsgf(LogLevelInfo, key)
}

// WarnMsg outputs a warning log using a message key.
func (l *Logger) WarnMsg(key MsgKey) {
	l.logMsgf(LogLevelWarn, key)
}

// ErrorMsg outputs an error log using a message key.
func (l *Logger) ErrorMsg(key MsgKey) {
	l.logMsgf(LogLevelError, key)
}

// DebugMsgf outputs a formatted debug log using a message key.
func (l *Logger) DebugMsgf(key MsgKey, args ...interface{}) {
	l.logMsgf(LogLevelDebug, key, args...)
}

// InfoMsgf outputs a formatted info log using a message key.
func (l *Logger) InfoMsgf(key MsgKey, args ...interface{}) {
	l.logMsgf(LogLevelInfo, key, args...)
}

// WarnMsgf outputs a formatted warning log using a message key.
func (l *Logger) WarnMsgf(key MsgKey, args ...interface{}) {
	l.logMsgf(LogLevelWarn, key, args...)
}

// ErrorMsgf outputs a formatted error log using a message key.
func (l *Logger) ErrorMsgf(key MsgKey, args ...interface{}) {
	l.logMsgf(LogLevelError, key, args...)
}

// --- Global functions ---

// SetGlobalLevel sets the global log level.
func SetGlobalLevel(level LogLevel) {
	defaultLogger.SetLevel(level)
}

// SetGlobalOutput sets the global log output destination.
func SetGlobalOutput(output io.Writer) {
	defaultLogger.SetOutput(output)
}

// SetGlobalLanguage sets the global log language.
func SetGlobalLanguage(lang LogLanguage) {
	defaultLogger.SetLanguage(lang)
}

// --- Global raw format functions (retained for backward compatibility) ---

// Debugf outputs a global debug log message.
func Debugf(format string, args ...interface{}) {
	defaultLogger.Debugf(format, args...)
}

// Infof outputs a global info log message.
func Infof(format string, args ...interface{}) {
	defaultLogger.Infof(format, args...)
}

// Warnf outputs a global warning log message.
func Warnf(format string, args ...interface{}) {
	defaultLogger.Warnf(format, args...)
}

// Errorf outputs a global error log message.
func Errorf(format string, args ...interface{}) {
	defaultLogger.Errorf(format, args...)
}

// Debug outputs a global debug log message.
func Debug(msg string) {
	defaultLogger.Debug(msg)
}

// Info outputs a global info log message.
func Info(msg string) {
	defaultLogger.Info(msg)
}

// Warn outputs a global warning log message.
func Warn(msg string) {
	defaultLogger.Warn(msg)
}

// Error outputs a global error log message.
func Error(msg string) {
	defaultLogger.Error(msg)
}

// --- Global message key functions ---

// DebugMsg outputs a global debug log using a message key.
func DebugMsg(key MsgKey) {
	defaultLogger.DebugMsg(key)
}

// InfoMsg outputs a global info log using a message key.
func InfoMsg(key MsgKey) {
	defaultLogger.InfoMsg(key)
}

// WarnMsg outputs a global warning log using a message key.
func WarnMsg(key MsgKey) {
	defaultLogger.WarnMsg(key)
}

// ErrorMsg outputs a global error log using a message key.
func ErrorMsg(key MsgKey) {
	defaultLogger.ErrorMsg(key)
}

// DebugMsgf outputs a global formatted debug log using a message key.
func DebugMsgf(key MsgKey, args ...interface{}) {
	defaultLogger.DebugMsgf(key, args...)
}

// InfoMsgf outputs a global formatted info log using a message key.
func InfoMsgf(key MsgKey, args ...interface{}) {
	defaultLogger.InfoMsgf(key, args...)
}

// WarnMsgf outputs a global formatted warning log using a message key.
func WarnMsgf(key MsgKey, args ...interface{}) {
	defaultLogger.WarnMsgf(key, args...)
}

// ErrorMsgf outputs a global formatted error log using a message key.
func ErrorMsgf(key MsgKey, args ...interface{}) {
	defaultLogger.ErrorMsgf(key, args...)
}

// Package document bilingual message catalog for the logging system.
package document

// LogLanguage represents the language used for log messages.
type LogLanguage int

const (
	// LogLanguageZH Chinese log messages (default).
	LogLanguageZH LogLanguage = iota
	// LogLanguageEN English log messages.
	LogLanguageEN
)

// MsgKey is a typed key for looking up bilingual log messages.
type MsgKey string

// ---------------------------------------------------------------------------
// Message key constants – document.go
// ---------------------------------------------------------------------------

const (
	// Simple messages (no format args)
	MsgCreatingNewDocument           MsgKey = "creating_new_document"
	MsgAddingParagraph               MsgKey = "adding_paragraph"
	MsgAddingFormattedParagraph      MsgKey = "adding_formatted_paragraph"
	MsgSettingParagraphAlignment     MsgKey = "setting_paragraph_alignment"
	MsgAddingFormattedText           MsgKey = "adding_formatted_text"
	MsgAddingPageBreakToParagraph    MsgKey = "adding_page_break_to_paragraph"
	MsgAddingHeadingParagraph        MsgKey = "adding_heading_paragraph"
	MsgAddingPageBreak               MsgKey = "adding_page_break"
	MsgSettingParagraphStyle         MsgKey = "setting_paragraph_style"
	MsgSettingParagraphIndent        MsgKey = "setting_paragraph_indent"
	MsgSettingParagraphOutlineLevel  MsgKey = "setting_paragraph_outline_level"
	MsgApplyingParagraphFormat       MsgKey = "applying_paragraph_format"
	MsgSettingParagraphBorder        MsgKey = "setting_paragraph_border"
	MsgSettingHorizontalRule         MsgKey = "setting_horizontal_rule"
	MsgSettingParagraphUnderline     MsgKey = "setting_paragraph_underline"
	MsgSettingParagraphBold          MsgKey = "setting_paragraph_bold"
	MsgSettingParagraphItalic        MsgKey = "setting_paragraph_italic"
	MsgSettingParagraphStrikethrough MsgKey = "setting_paragraph_strikethrough"
	MsgSettingParagraphHighlight     MsgKey = "setting_paragraph_highlight"
	MsgSettingParagraphFont          MsgKey = "setting_paragraph_font"
	MsgSettingParagraphFontSize      MsgKey = "setting_paragraph_font_size"
	MsgSettingParagraphColor         MsgKey = "setting_paragraph_color"
	MsgParsingDocumentContent        MsgKey = "parsing_document_content"
	MsgSerializingDocument           MsgKey = "serializing_document"
	MsgDocumentSerializationComplete MsgKey = "document_serialization_complete"
	MsgSerializingStyles             MsgKey = "serializing_styles"
	MsgStyleSerializationComplete    MsgKey = "style_serialization_complete"
	MsgParsingContentTypes           MsgKey = "parsing_content_types"
	MsgContentTypesParsed            MsgKey = "content_types_parsed"
	MsgParsingRelationships          MsgKey = "parsing_relationships"
	MsgRelationshipsParsed           MsgKey = "relationships_parsed"
	MsgParsingStyles                 MsgKey = "parsing_styles"
	MsgStylesParsed                  MsgKey = "styles_parsed"
	MsgParsingDocumentRelationships  MsgKey = "parsing_document_relationships"
	MsgDocumentRelationshipsParsed   MsgKey = "document_relationships_parsed"
	MsgUpdatingImageIDCounter        MsgKey = "updating_image_id_counter"
	MsgParagraphToDeleteNotFound     MsgKey = "paragraph_to_delete_not_found"
	MsgParagraphIndexOutOfRange      MsgKey = "paragraph_index_out_of_range"
	MsgDeletingElement               MsgKey = "deleting_element"

	// Info-level messages
	MsgOpeningDocumentPath MsgKey = "opening_document_path"
	MsgDocumentOpenedPath  MsgKey = "document_opened_path"
	MsgOpeningDocument     MsgKey = "opening_document"
	MsgDocumentOpened      MsgKey = "document_opened"
	MsgSavingDocument      MsgKey = "saving_document"
	MsgDocumentSaved       MsgKey = "document_saved"
	MsgParsingComplete     MsgKey = "parsing_complete"

	// Error-level messages
	MsgFailedToOpenFile          MsgKey = "failed_to_open_file"
	MsgFailedToOpenFileSimple    MsgKey = "failed_to_open_file_simple"
	MsgFailedToParseDocument     MsgKey = "failed_to_parse_document"
	MsgFailedToCreateDirectory   MsgKey = "failed_to_create_directory"
	MsgFailedToCreateFile        MsgKey = "failed_to_create_file"
	MsgFailedToSerializeDocument MsgKey = "failed_to_serialize_document"
	MsgFailedToSerializeStyles   MsgKey = "failed_to_serialize_styles"
	MsgXMLSerializationFailed    MsgKey = "xml_serialization_failed"
	MsgFailedToOpenFilePart      MsgKey = "failed_to_open_file_part"
	MsgFailedToReadFilePart      MsgKey = "failed_to_read_file_part"
	MsgFailedToCreateZIPEntry    MsgKey = "failed_to_create_zip_entry"
	MsgFailedToWriteZIPEntry     MsgKey = "failed_to_write_zip_entry"

	// Nested Debug messages
	MsgReadFilePart                      MsgKey = "read_file_part"
	MsgFailedToParseContentTypesDefault  MsgKey = "failed_to_parse_content_types_default"
	MsgFailedToParseRelationshipsDefault MsgKey = "failed_to_parse_relationships_default"
	MsgFailedToParseStylesDefault        MsgKey = "failed_to_parse_styles_default"
	MsgFailedToParseDocRelDefault        MsgKey = "failed_to_parse_doc_rel_default"
	MsgWrittenZIPEntry                   MsgKey = "written_zip_entry"
	MsgSettingParagraphSpacing           MsgKey = "setting_paragraph_spacing"
	MsgHeadingLevelOutOfRange            MsgKey = "heading_level_out_of_range"
	MsgStyleNotFoundUsingDefault         MsgKey = "style_not_found_using_default"
	MsgAddingBookmarkStart               MsgKey = "adding_bookmark_start"
	MsgAddingBookmarkEnd                 MsgKey = "adding_bookmark_end"
	MsgSettingKeepWithNext               MsgKey = "setting_keep_with_next"
	MsgUnsettingKeepWithNext             MsgKey = "unsetting_keep_with_next"
	MsgSettingKeepLinesTogether          MsgKey = "setting_keep_lines_together"
	MsgUnsettingKeepLinesTogether        MsgKey = "unsetting_keep_lines_together"
	MsgSettingPageBreakBefore            MsgKey = "setting_page_break_before"
	MsgUnsettingPageBreakBefore          MsgKey = "unsetting_page_break_before"
	MsgEnablingWidowOrphanControl        MsgKey = "enabling_widow_orphan_control"
	MsgDisablingWidowOrphanControl       MsgKey = "disabling_widow_orphan_control"
	MsgOutlineLevelAdjusted              MsgKey = "outline_level_adjusted"
	MsgDisablingParagraphGrid            MsgKey = "disabling_paragraph_grid"
	MsgEnablingParagraphGrid             MsgKey = "enabling_paragraph_grid"
	MsgSkippingUnknownElement            MsgKey = "skipping_unknown_element"
	MsgPreservingUnknownElement          MsgKey = "preserving_unknown_element"
	MsgExistingStylesDetected            MsgKey = "existing_styles_detected"
	MsgDocRelFileNotFound                MsgKey = "doc_rel_file_not_found"
	MsgParagraphIndexNegative            MsgKey = "paragraph_index_negative"
	MsgElementIndexOutOfRange            MsgKey = "element_index_out_of_range"
)

// ---------------------------------------------------------------------------
// Message key constants – table.go
// ---------------------------------------------------------------------------

const (
	MsgTableCreated                  MsgKey = "table_created"
	MsgTableAddedToDocument          MsgKey = "table_added_to_document"
	MsgRowInserted                   MsgKey = "row_inserted"
	MsgRowDeleted                    MsgKey = "row_deleted"
	MsgRowsDeleted                   MsgKey = "rows_deleted"
	MsgColumnInserted                MsgKey = "column_inserted"
	MsgColumnDeleted                 MsgKey = "column_deleted"
	MsgColumnsDeleted                MsgKey = "columns_deleted"
	MsgTableContentCleared           MsgKey = "table_content_cleared"
	MsgTableCopied                   MsgKey = "table_copied"
	MsgCellFormatSet                 MsgKey = "cell_format_set"
	MsgCellRichTextSet               MsgKey = "cell_rich_text_set"
	MsgFormattedTextAddedToCell      MsgKey = "formatted_text_added_to_cell"
	MsgHorizontalMerge               MsgKey = "horizontal_merge"
	MsgVerticalMerge                 MsgKey = "vertical_merge"
	MsgRangeMerge                    MsgKey = "range_merge"
	MsgCellMergeCancelled            MsgKey = "cell_merge_cancelled"
	MsgCellContentCleared            MsgKey = "cell_content_cleared"
	MsgCellFormatCleared             MsgKey = "cell_format_cleared"
	MsgCellPaddingSet                MsgKey = "cell_padding_set"
	MsgCellTextDirectionSet          MsgKey = "cell_text_direction_set"
	MsgRowHeightSet                  MsgKey = "row_height_set"
	MsgRowsHeightSet                 MsgKey = "rows_height_set"
	MsgTableLayoutSet                MsgKey = "table_layout_set"
	MsgTableFloatingMode             MsgKey = "table_floating_mode"
	MsgRowPageSplitSet               MsgKey = "row_page_split_set"
	MsgRowSetAsHeader                MsgKey = "row_set_as_header"
	MsgRowsSetAsHeaders              MsgKey = "rows_set_as_headers"
	MsgTablePagination               MsgKey = "table_pagination"
	MsgRowKeepWithNextSet            MsgKey = "row_keep_with_next_set" //nolint:gosec // G101: this is a message key, not a credential
	MsgTableStyleApplied             MsgKey = "table_style_applied"
	MsgTableBorderSet                MsgKey = "table_border_set"
	MsgTableBackgroundSet            MsgKey = "table_background_set"
	MsgCellBorderSet                 MsgKey = "cell_border_set"
	MsgCellBackgroundSet             MsgKey = "cell_background_set"
	MsgAlternatingRowColorsSet       MsgKey = "alternating_row_colors_set"
	MsgCustomTableStyleCreated       MsgKey = "custom_table_style_created"
	MsgParagraphAddedToCell          MsgKey = "paragraph_added_to_cell"
	MsgFormattedParagraphAddedToCell MsgKey = "formatted_paragraph_added_to_cell"
	MsgCellParagraphsCleared         MsgKey = "cell_paragraphs_cleared"
	MsgNestedTableAddedToCell        MsgKey = "nested_table_added_to_cell"
	MsgListAddedToCell               MsgKey = "list_added_to_cell"
)

// ---------------------------------------------------------------------------
// Message key constants – image.go
// ---------------------------------------------------------------------------

const (
	MsgAddingImageFile        MsgKey = "adding_image_file"
	MsgImageRead              MsgKey = "image_read"
	MsgImageAddedToCell       MsgKey = "image_added_to_cell"
	MsgImageParagraphNotFound MsgKey = "image_paragraph_not_found"
)

// ---------------------------------------------------------------------------
// Message key constants – math.go
// ---------------------------------------------------------------------------

const (
	MsgAddingMathFormula       MsgKey = "adding_math_formula"
	MsgAddingInlineMathFormula MsgKey = "adding_inline_math_formula"
)

// ---------------------------------------------------------------------------
// Message key constants – page.go
// ---------------------------------------------------------------------------

const (
	MsgPageSettingsUpdated MsgKey = "page_settings_updated"
)

// ---------------------------------------------------------------------------
// Chinese message map
// ---------------------------------------------------------------------------

//nolint:dupl,gosec
var messagesZH = map[MsgKey]string{
	// document.go – Debug
	MsgCreatingNewDocument:           "创建新文档",
	MsgAddingParagraph:               "添加段落: %s",
	MsgAddingFormattedParagraph:      "添加格式化段落: %s",
	MsgSettingParagraphAlignment:     "设置段落对齐方式: %s",
	MsgAddingFormattedText:           "向段落添加格式化文本: %s",
	MsgAddingPageBreakToParagraph:    "向段落添加分页符",
	MsgAddingHeadingParagraph:        "添加标题段落: %s (级别: %d, 样式: %s, 书签: %s)",
	MsgAddingPageBreak:               "添加分页符",
	MsgSettingParagraphStyle:         "设置段落样式: %s",
	MsgSettingParagraphIndent:        "设置段落缩进: 首行=%.2fcm, 左=%.2fcm, 右=%.2fcm",
	MsgSettingParagraphOutlineLevel:  "设置段落大纲级别: %d",
	MsgApplyingParagraphFormat:       "应用段落格式配置: 对齐=%s, 样式=%s, 行距=%.1f, 段前=%d, 段后=%d",
	MsgSettingParagraphBorder:        "设置段落边框: 上=%v, 左=%v, 下=%v, 右=%v",
	MsgSettingHorizontalRule:         "设置水平分割线: 样式=%s, 粗细=%d, 颜色=%s",
	MsgSettingParagraphUnderline:     "设置段落下划线: %v",
	MsgSettingParagraphBold:          "设置段落粗体: %v",
	MsgSettingParagraphItalic:        "设置段落斜体: %v",
	MsgSettingParagraphStrikethrough: "设置段落删除线: %v",
	MsgSettingParagraphHighlight:     "设置段落高亮: %s",
	MsgSettingParagraphFont:          "设置段落字体: %s",
	MsgSettingParagraphFontSize:      "设置段落字体大小: %d",
	MsgSettingParagraphColor:         "设置段落颜色: %s",
	MsgParsingDocumentContent:        "开始解析文档内容",
	MsgSerializingDocument:           "开始序列化文档",
	MsgDocumentSerializationComplete: "文档序列化完成",
	MsgSerializingStyles:             "开始序列化样式",
	MsgStyleSerializationComplete:    "样式序列化完成",
	MsgParsingContentTypes:           "开始解析内容类型文件",
	MsgContentTypesParsed:            "内容类型解析完成",
	MsgParsingRelationships:          "开始解析关系文件",
	MsgRelationshipsParsed:           "关系解析完成",
	MsgParsingStyles:                 "开始解析样式文件",
	MsgStylesParsed:                  "样式解析完成",
	MsgParsingDocumentRelationships:  "开始解析文档关系文件",
	MsgDocumentRelationshipsParsed:   "文档关系解析完成，共 %d 个关系",
	MsgUpdatingImageIDCounter:        "更新图片ID计数器: nextImageID = %d",
	MsgParagraphToDeleteNotFound:     "警告：未找到要删除的段落",
	MsgParagraphIndexOutOfRange:      "错误：段落索引 %d 超出范围 [0, %d)",
	MsgDeletingElement:               "删除元素: 索引 %d",

	// document.go – Info
	MsgOpeningDocumentPath: "正在打开文档: %s",
	MsgDocumentOpenedPath:  "成功打开文档: %s",
	MsgOpeningDocument:     "正在打开文档",
	MsgDocumentOpened:      "成功打开文档",
	MsgSavingDocument:      "正在保存文档: %s",
	MsgDocumentSaved:       "成功保存文档: %s",
	MsgParsingComplete:     "解析完成，共 %d 个元素",

	// document.go – Error
	MsgFailedToOpenFile:          "无法打开文件: %s",
	MsgFailedToOpenFileSimple:    "无法打开文件",
	MsgFailedToParseDocument:     "解析文档失败: %s",
	MsgFailedToCreateDirectory:   "无法创建目录: %s",
	MsgFailedToCreateFile:        "无法创建文件: %s",
	MsgFailedToSerializeDocument: "序列化文档失败",
	MsgFailedToSerializeStyles:   "序列化样式失败",
	MsgXMLSerializationFailed:    "XML序列化失败: %v",
	MsgFailedToOpenFilePart:      "无法打开文件部件: %s",
	MsgFailedToReadFilePart:      "无法读取文件部件: %s",
	MsgFailedToCreateZIPEntry:    "无法创建ZIP条目: %s",
	MsgFailedToWriteZIPEntry:     "无法写入ZIP条目: %s",

	// document.go – Nested Debug
	MsgReadFilePart:                      "已读取文件部件: %s (%d 字节)",
	MsgFailedToParseContentTypesDefault:  "解析内容类型失败，使用默认值: %v",
	MsgFailedToParseRelationshipsDefault: "解析关系失败，使用默认值: %v",
	MsgFailedToParseStylesDefault:        "解析样式失败，使用默认样式: %v",
	MsgFailedToParseDocRelDefault:        "解析文档关系失败，使用默认值: %v",
	MsgWrittenZIPEntry:                   "已写入ZIP条目: %s (%d 字节)",
	MsgSettingParagraphSpacing:           "设置段落间距: 段前=%d, 段后=%d, 行距=%.1f, 首行缩进=%d",
	MsgHeadingLevelOutOfRange:            "标题级别 %d 超出范围，使用默认级别 1",
	MsgStyleNotFoundUsingDefault:         "警告：找不到样式 %s，使用默认样式",
	MsgAddingBookmarkStart:               "添加书签开始: ID=%s, Name=%s",
	MsgAddingBookmarkEnd:                 "添加书签结束: ID=%s",
	MsgSettingKeepWithNext:               "设置段落与下一段保持在一起",
	MsgUnsettingKeepWithNext:             "取消段落与下一段保持在一起",
	MsgSettingKeepLinesTogether:          "设置段落行保持在一起",
	MsgUnsettingKeepLinesTogether:        "取消段落行保持在一起",
	MsgSettingPageBreakBefore:            "设置段前分页",
	MsgUnsettingPageBreakBefore:          "取消段前分页",
	MsgEnablingWidowOrphanControl:        "启用段落孤行控制",
	MsgDisablingWidowOrphanControl:       "禁用段落孤行控制",
	MsgOutlineLevelAdjusted:              "大纲级别应在0-8之间，已调整为有效范围",
	MsgDisablingParagraphGrid:            "禁用段落网格对齐",
	MsgEnablingParagraphGrid:             "启用段落网格对齐（默认）",
	MsgSkippingUnknownElement:            "跳过未知元素: %s",
	MsgPreservingUnknownElement:          "保留未知元素: %s",
	MsgExistingStylesDetected:            "检测到已有 styles.xml，跳过样式重建以保留模板默认样式",
	MsgDocRelFileNotFound:                "文档关系文件不存在，文档可能不包含图片等资源",
	MsgParagraphIndexNegative:            "错误：段落索引不能为负数: %d",
	MsgElementIndexOutOfRange:            "错误：元素索引 %d 超出范围 [0, %d)",

	// table.go
	MsgTableCreated:                  "创建表格成功：%d行 x %d列",
	MsgTableAddedToDocument:          "表格已添加到文档，当前文档包含%d个表格",
	MsgRowInserted:                   "在位置%d插入行成功",
	MsgRowDeleted:                    "删除第%d行成功",
	MsgRowsDeleted:                   "删除第%d到%d行成功",
	MsgColumnInserted:                "在位置%d插入列成功",
	MsgColumnDeleted:                 "删除第%d列成功",
	MsgColumnsDeleted:                "删除第%d到%d列成功",
	MsgTableContentCleared:           "表格内容已清空",
	MsgTableCopied:                   "表格复制成功",
	MsgCellFormatSet:                 "设置单元格(%d,%d)格式成功",
	MsgCellRichTextSet:               "设置单元格(%d,%d)富文本内容成功",
	MsgFormattedTextAddedToCell:      "添加格式化文本到单元格(%d,%d)成功",
	MsgHorizontalMerge:               "水平合并单元格：行%d，列%d到%d",
	MsgVerticalMerge:                 "垂直合并单元格：行%d到%d，列%d",
	MsgRangeMerge:                    "合并单元格区域：行%d到%d，列%d到%d",
	MsgCellMergeCancelled:            "取消单元格(%d,%d)合并成功",
	MsgCellContentCleared:            "清空单元格(%d,%d)内容成功",
	MsgCellFormatCleared:             "清空单元格(%d,%d)格式成功",
	MsgCellPaddingSet:                "设置单元格(%d,%d)内边距为%d磅",
	MsgCellTextDirectionSet:          "设置单元格(%d,%d)文字方向为%s",
	MsgRowHeightSet:                  "设置第%d行高度为%d磅，规则为%s",
	MsgRowsHeightSet:                 "批量设置第%d到%d行高度成功",
	MsgTableLayoutSet:                "设置表格布局：对齐=%s，环绕=%s，定位=%s",
	MsgTableFloatingMode:             "设置表格为浮动定位模式",
	MsgRowPageSplitSet:               "设置第%d行跨页分割为：%t",
	MsgRowSetAsHeader:                "设置第%d行为标题行：%t",
	MsgRowsSetAsHeaders:              "设置第%d到%d行为标题行",
	MsgTablePagination:               "设置表格分页控制：保持与下一段落=%t，保持行=%t，段前分页=%t，孤行控制=%t",
	MsgRowKeepWithNextSet:            "设置第%d行与下一行保持在同一页：%t",
	MsgTableStyleApplied:             "应用表格样式成功：%s",
	MsgTableBorderSet:                "设置表格边框成功",
	MsgTableBackgroundSet:            "设置表格背景成功",
	MsgCellBorderSet:                 "设置单元格(%d,%d)边框成功",
	MsgCellBackgroundSet:             "设置单元格(%d,%d)背景成功",
	MsgAlternatingRowColorsSet:       "设置奇偶行颜色交替成功",
	MsgCustomTableStyleCreated:       "创建自定义表格样式成功：%s",
	MsgParagraphAddedToCell:          "向单元格(%d,%d)添加段落成功",
	MsgFormattedParagraphAddedToCell: "向单元格(%d,%d)添加格式化段落成功",
	MsgCellParagraphsCleared:         "清空单元格(%d,%d)段落成功",
	MsgNestedTableAddedToCell:        "向单元格(%d,%d)添加嵌套表格成功：%d行 x %d列",
	MsgListAddedToCell:               "向单元格(%d,%d)添加列表成功：%d个列表项",

	// image.go
	MsgAddingImageFile:        "开始添加图片文件: %s",
	MsgImageRead:              "成功读取图片: %s (格式: %s, 尺寸: %dx%d, 大小: %d字节)",
	MsgImageAddedToCell:       "向表格单元格(%d,%d)添加图片成功: ID=%s",
	MsgImageParagraphNotFound: "未找到包含图片ID %s 的段落，已仅更新配置中的对齐方式",

	// math.go
	MsgAddingMathFormula:       "添加数学公式: %s (块级: %v)",
	MsgAddingInlineMathFormula: "向段落添加行内数学公式",

	// page.go
	MsgPageSettingsUpdated: "页面设置已更新: 尺寸=%s, 方向=%s",
}

// ---------------------------------------------------------------------------
// English message map
// ---------------------------------------------------------------------------

//nolint:dupl,gosec
var messagesEN = map[MsgKey]string{
	// document.go – Debug
	MsgCreatingNewDocument:           "Creating new document",
	MsgAddingParagraph:               "Adding paragraph: %s",
	MsgAddingFormattedParagraph:      "Adding formatted paragraph: %s",
	MsgSettingParagraphAlignment:     "Setting paragraph alignment: %s",
	MsgAddingFormattedText:           "Adding formatted text to paragraph: %s",
	MsgAddingPageBreakToParagraph:    "Adding page break to paragraph",
	MsgAddingHeadingParagraph:        "Adding heading paragraph: %s (level: %d, style: %s, bookmark: %s)",
	MsgAddingPageBreak:               "Adding page break",
	MsgSettingParagraphStyle:         "Setting paragraph style: %s",
	MsgSettingParagraphIndent:        "Setting paragraph indent: firstLine=%.2fcm, left=%.2fcm, right=%.2fcm",
	MsgSettingParagraphOutlineLevel:  "Setting paragraph outline level: %d",
	MsgApplyingParagraphFormat:       "Applying paragraph format: align=%s, style=%s, lineSpacing=%.1f, before=%d, after=%d",
	MsgSettingParagraphBorder:        "Setting paragraph border: top=%v, left=%v, bottom=%v, right=%v",
	MsgSettingHorizontalRule:         "Setting horizontal rule: style=%s, size=%d, color=%s",
	MsgSettingParagraphUnderline:     "Setting paragraph underline: %v",
	MsgSettingParagraphBold:          "Setting paragraph bold: %v",
	MsgSettingParagraphItalic:        "Setting paragraph italic: %v",
	MsgSettingParagraphStrikethrough: "Setting paragraph strikethrough: %v",
	MsgSettingParagraphHighlight:     "Setting paragraph highlight: %s",
	MsgSettingParagraphFont:          "Setting paragraph font: %s",
	MsgSettingParagraphFontSize:      "Setting paragraph font size: %d",
	MsgSettingParagraphColor:         "Setting paragraph color: %s",
	MsgParsingDocumentContent:        "Parsing document content",
	MsgSerializingDocument:           "Serializing document",
	MsgDocumentSerializationComplete: "Document serialization complete",
	MsgSerializingStyles:             "Serializing styles",
	MsgStyleSerializationComplete:    "Style serialization complete",
	MsgParsingContentTypes:           "Parsing content types",
	MsgContentTypesParsed:            "Content types parsed",
	MsgParsingRelationships:          "Parsing relationships",
	MsgRelationshipsParsed:           "Relationships parsed",
	MsgParsingStyles:                 "Parsing styles",
	MsgStylesParsed:                  "Styles parsed",
	MsgParsingDocumentRelationships:  "Parsing document relationships",
	MsgDocumentRelationshipsParsed:   "Document relationships parsed, total: %d",
	MsgUpdatingImageIDCounter:        "Updating image ID counter: nextImageID = %d",
	MsgParagraphToDeleteNotFound:     "Warning: paragraph to delete not found",
	MsgParagraphIndexOutOfRange:      "Error: paragraph index %d out of range [0, %d)",
	MsgDeletingElement:               "Deleting element: index %d",

	// document.go – Info
	MsgOpeningDocumentPath: "Opening document: %s",
	MsgDocumentOpenedPath:  "Document opened: %s",
	MsgOpeningDocument:     "Opening document",
	MsgDocumentOpened:      "Document opened",
	MsgSavingDocument:      "Saving document: %s",
	MsgDocumentSaved:       "Document saved: %s",
	MsgParsingComplete:     "Parsing complete, total elements: %d",

	// document.go – Error
	MsgFailedToOpenFile:          "Failed to open file: %s",
	MsgFailedToOpenFileSimple:    "Failed to open file",
	MsgFailedToParseDocument:     "Failed to parse document: %s",
	MsgFailedToCreateDirectory:   "Failed to create directory: %s",
	MsgFailedToCreateFile:        "Failed to create file: %s",
	MsgFailedToSerializeDocument: "Failed to serialize document",
	MsgFailedToSerializeStyles:   "Failed to serialize styles",
	MsgXMLSerializationFailed:    "XML serialization failed: %v",
	MsgFailedToOpenFilePart:      "Failed to open file part: %s",
	MsgFailedToReadFilePart:      "Failed to read file part: %s",
	MsgFailedToCreateZIPEntry:    "Failed to create ZIP entry: %s",
	MsgFailedToWriteZIPEntry:     "Failed to write ZIP entry: %s",

	// document.go – Nested Debug
	MsgReadFilePart:                      "Read file part: %s (%d bytes)",
	MsgFailedToParseContentTypesDefault:  "Failed to parse content types, using defaults: %v",
	MsgFailedToParseRelationshipsDefault: "Failed to parse relationships, using defaults: %v",
	MsgFailedToParseStylesDefault:        "Failed to parse styles, using defaults: %v",
	MsgFailedToParseDocRelDefault:        "Failed to parse document relationships, using defaults: %v",
	MsgWrittenZIPEntry:                   "Written ZIP entry: %s (%d bytes)",
	MsgSettingParagraphSpacing:           "Setting paragraph spacing: before=%d, after=%d, lineSpacing=%.1f, firstIndent=%d",
	MsgHeadingLevelOutOfRange:            "Heading level %d out of range, using default level 1",
	MsgStyleNotFoundUsingDefault:         "Warning: style %s not found, using default",
	MsgAddingBookmarkStart:               "Adding bookmark start: ID=%s, Name=%s",
	MsgAddingBookmarkEnd:                 "Adding bookmark end: ID=%s",
	MsgSettingKeepWithNext:               "Setting keep with next paragraph",
	MsgUnsettingKeepWithNext:             "Unsetting keep with next paragraph",
	MsgSettingKeepLinesTogether:          "Setting keep lines together",
	MsgUnsettingKeepLinesTogether:        "Unsetting keep lines together",
	MsgSettingPageBreakBefore:            "Setting page break before",
	MsgUnsettingPageBreakBefore:          "Unsetting page break before",
	MsgEnablingWidowOrphanControl:        "Enabling widow/orphan control",
	MsgDisablingWidowOrphanControl:       "Disabling widow/orphan control",
	MsgOutlineLevelAdjusted:              "Outline level should be 0-8, adjusted to valid range",
	MsgDisablingParagraphGrid:            "Disabling paragraph grid alignment",
	MsgEnablingParagraphGrid:             "Enabling paragraph grid alignment (default)",
	MsgSkippingUnknownElement:            "Skipping unknown element: %s",
	MsgPreservingUnknownElement:          "Preserving unknown element: %s",
	MsgExistingStylesDetected:            "Existing styles.xml detected, skipping style rebuild to preserve template styles",
	MsgDocRelFileNotFound:                "Document relationships file not found, document may not contain images or resources",
	MsgParagraphIndexNegative:            "Error: paragraph index cannot be negative: %d",
	MsgElementIndexOutOfRange:            "Error: element index %d out of range [0, %d)",

	// table.go
	MsgTableCreated:                  "Table created: %d rows x %d cols",
	MsgTableAddedToDocument:          "Table added to document, current table count: %d",
	MsgRowInserted:                   "Row inserted at position %d",
	MsgRowDeleted:                    "Row %d deleted",
	MsgRowsDeleted:                   "Rows %d to %d deleted",
	MsgColumnInserted:                "Column inserted at position %d",
	MsgColumnDeleted:                 "Column %d deleted",
	MsgColumnsDeleted:                "Columns %d to %d deleted",
	MsgTableContentCleared:           "Table content cleared",
	MsgTableCopied:                   "Table copied",
	MsgCellFormatSet:                 "Cell (%d,%d) format set",
	MsgCellRichTextSet:               "Cell (%d,%d) rich text content set",
	MsgFormattedTextAddedToCell:      "Formatted text added to cell (%d,%d)",
	MsgHorizontalMerge:               "Horizontal merge: row %d, cols %d to %d",
	MsgVerticalMerge:                 "Vertical merge: rows %d to %d, col %d",
	MsgRangeMerge:                    "Range merge: rows %d to %d, cols %d to %d",
	MsgCellMergeCancelled:            "Cell (%d,%d) merge cancelled",
	MsgCellContentCleared:            "Cell (%d,%d) content cleared",
	MsgCellFormatCleared:             "Cell (%d,%d) format cleared",
	MsgCellPaddingSet:                "Cell (%d,%d) padding set to %d pt",
	MsgCellTextDirectionSet:          "Cell (%d,%d) text direction set to %s",
	MsgRowHeightSet:                  "Row %d height set to %d pt, rule: %s",
	MsgRowsHeightSet:                 "Rows %d to %d height set",
	MsgTableLayoutSet:                "Table layout set: align=%s, wrap=%s, position=%s",
	MsgTableFloatingMode:             "Table set to floating position mode",
	MsgRowPageSplitSet:               "Row %d page split set to: %t",
	MsgRowSetAsHeader:                "Row %d set as header: %t",
	MsgRowsSetAsHeaders:              "Rows %d to %d set as headers",
	MsgTablePagination:               "Table pagination: keepWithNext=%t, keepLines=%t, pageBreakBefore=%t, widowControl=%t",
	MsgRowKeepWithNextSet:            "Row %d keep with next set to: %t",
	MsgTableStyleApplied:             "Table style applied: %s",
	MsgTableBorderSet:                "Table border set",
	MsgTableBackgroundSet:            "Table background set",
	MsgCellBorderSet:                 "Cell (%d,%d) border set",
	MsgCellBackgroundSet:             "Cell (%d,%d) background set",
	MsgAlternatingRowColorsSet:       "Alternating row colors set",
	MsgCustomTableStyleCreated:       "Custom table style created: %s",
	MsgParagraphAddedToCell:          "Paragraph added to cell (%d,%d)",
	MsgFormattedParagraphAddedToCell: "Formatted paragraph added to cell (%d,%d)",
	MsgCellParagraphsCleared:         "Cell (%d,%d) paragraphs cleared",
	MsgNestedTableAddedToCell:        "Nested table added to cell (%d,%d): %d rows x %d cols",
	MsgListAddedToCell:               "List added to cell (%d,%d): %d items",

	// image.go
	MsgAddingImageFile:        "Adding image file: %s",
	MsgImageRead:              "Image read: %s (format: %s, size: %dx%d, %d bytes)",
	MsgImageAddedToCell:       "Image added to cell (%d,%d): ID=%s",
	MsgImageParagraphNotFound: "Paragraph with image ID %s not found, alignment updated in config only",

	// math.go
	MsgAddingMathFormula:       "Adding math formula: %s (block: %v)",
	MsgAddingInlineMathFormula: "Adding inline math formula to paragraph",

	// page.go
	MsgPageSettingsUpdated: "Page settings updated: size=%s, orientation=%s",
}

// getMessage returns the localized format string for the given message key
// and language. If the key is not found in the requested language map, it
// falls back to the key string itself so that callers always receive a
// usable value.
func getMessage(key MsgKey, lang LogLanguage) string {
	var msgs map[MsgKey]string
	switch lang {
	case LogLanguageEN:
		msgs = messagesEN
	default:
		msgs = messagesZH
	}
	if msg, ok := msgs[key]; ok {
		return msg
	}
	return string(key)
}

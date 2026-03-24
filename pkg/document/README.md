# Document 包 API 文档

本文档记录了 `pkg/document` 包中所有可用的公开方法和功能。

## 核心类型

### Document 文档
- [Document](document.go) - Word文档的核心结构
- [Body](document.go) - 文档主体
- [Paragraph](document.go) - 段落结构
- [Table](table.go) - 表格结构

## 文档操作方法

### 文档创建与加载
- [`New()`](document.go#L232) - 创建新的Word文档
- [`Open(filename string)`](document.go#L269) - 打开现有Word文档 ✨ **重大改进**
  
#### 文档解析功能重大升级 ✨
`Open` 方法现在支持完整的文档结构解析，包括：

**动态元素解析支持**：
- **段落解析** (`<w:p>`): 完整解析段落内容、属性、运行和格式
- **表格解析** (`<w:tbl>`): 支持表格结构、网格、行列、单元格内容
- **节属性解析** (`<w:sectPr>`): 页面设置、边距、分栏等属性
- **扩展性设计**: 新的解析架构可轻松添加更多元素类型

**解析器特性**：
- **流式解析**: 使用XML流式解析器，内存效率高，适用于大型文档
- **结构保持**: 完整保留文档元素的原始顺序和层次结构
- **错误恢复**: 智能跳过未知或损坏的元素，确保解析过程稳定
- **深度解析**: 支持嵌套结构（如表格中的段落、段落中的运行等）

**解析的内容包括**：
- 段落文本内容和所有格式属性（字体、大小、颜色、样式等）
- 表格完整结构（行列定义、单元格内容、表格属性）
- 页面设置信息（页面尺寸、方向、边距等）
- 样式引用和属性继承关系

### 文档保存与导出
- [`Save(filename string)`](document.go#L337) - 保存文档到文件
- [`ToBytes()`](document.go#L1107) - 将文档转换为字节数组

### 文档内容操作
- [`AddParagraph(text string)`](document.go#L420) - 添加简单段落
- [`AddFormattedParagraph(text string, format *TextFormat)`](document.go#L459) - 添加格式化段落
- [`AddHeadingParagraph(text string, level int)`](document.go#L682) - 添加标题段落
- [`AddHeadingParagraphWithBookmark(text string, level int, bookmarkName string)`](document.go#L747) - 添加带书签的标题段落 ✨ **新增功能**
- [`AddPageBreak()`](document.go#L1185) - 添加分页符

#### 分页符功能 ✨

WordZero提供多种方式添加分页符（页面分页符）：

**方法一：文档级分页符**
```go
doc := document.New()
doc.AddParagraph("第一页内容")
doc.AddPageBreak()  // 添加分页符
doc.AddParagraph("第二页内容")
```

**方法二：段落内分页符**
```go
para := doc.AddParagraph("第一页内容")
para.AddPageBreak()  // 在段落内添加分页符
para.AddFormattedText("第二页内容", nil)
```

**方法三：段前分页**
```go
para := doc.AddParagraph("第二章标题")
para.SetPageBreakBefore(true)  // 设置段落前自动分页
```

**分页功能特性**：
- **独立分页符**: `Document.AddPageBreak()` 创建独立的分页段落
- **段落内分页**: `Paragraph.AddPageBreak()` 在当前段落内添加分页符
- **段前分页**: `Paragraph.SetPageBreakBefore(true)` 设置段落前自动分页
- **表格分页控制**: 支持表格的分页控制设置

#### 标题段落书签功能 ✨
`AddHeadingParagraphWithBookmark` 方法现在支持为标题段落添加书签：

**书签功能特性**：
- **自动书签生成**: 为标题段落创建唯一的书签标识
- **灵活命名**: 支持自定义书签名称或留空不添加书签
- **目录兼容**: 生成的书签与目录功能完美兼容，支持导航和超链接
- **Word标准**: 符合Microsoft Word的书签格式规范

**书签生成规则**：
- 书签ID自动生成为 `bookmark_{元素索引}_{书签名称}` 格式
- 书签开始标记插入在段落之前
- 书签结束标记插入在段落之后
- 支持空书签名称以跳过书签创建

### 样式管理
- [`GetStyleManager()`](document.go#L791) - 获取样式管理器

### 页面设置 ✨ 新增功能
- [`SetPageSettings(settings *PageSettings)`](page.go) - 设置完整页面属性
- [`GetPageSettings()`](page.go) - 获取当前页面设置
- [`SetPageSize(size PageSize)`](page.go) - 设置页面尺寸
- [`SetCustomPageSize(width, height float64)`](page.go) - 设置自定义页面尺寸（毫米）
- [`SetPageOrientation(orientation PageOrientation)`](page.go) - 设置页面方向
- [`SetPageMargins(top, right, bottom, left float64)`](page.go) - 设置页面边距（毫米）
- [`SetHeaderFooterDistance(header, footer float64)`](page.go) - 设置页眉页脚距离（毫米）
- [`SetGutterWidth(width float64)`](page.go) - 设置装订线宽度（毫米）
- [`DefaultPageSettings()`](page.go) - 获取默认页面设置（A4纵向）

### 页眉页脚操作 ✨ 新增功能
- [`AddHeader(headerType HeaderFooterType, text string)`](header_footer.go) - 添加页眉
- [`AddFooter(footerType HeaderFooterType, text string)`](header_footer.go) - 添加页脚
- [`AddHeaderWithPageNumber(headerType HeaderFooterType, text string, showPageNum bool)`](header_footer.go) - 添加带页码的页眉
- [`AddFooterWithPageNumber(footerType HeaderFooterType, text string, showPageNum bool)`](header_footer.go) - 添加带页码的页脚
- [`SetDifferentFirstPage(different bool)`](header_footer.go) - 设置首页不同

### 目录功能 ✨ 新增功能
- [`GenerateTOC(config *TOCConfig)`](toc.go) - 生成目录
- [`UpdateTOC()`](toc.go) - 更新目录
- [`AddHeadingWithBookmark(text string, level int, bookmarkName string)`](toc.go) - 添加带书签的标题
- [`AutoGenerateTOC(config *TOCConfig)`](toc.go) - 自动生成目录
- [`GetHeadingCount()`](toc.go) - 获取标题统计
- [`ListHeadings()`](toc.go) - 列出所有标题
- [`SetTOCStyle(level int, style *TextFormat)`](toc.go) - 设置目录样式

### 脚注与尾注功能 ✨ 新增功能
- [`AddFootnote(text string, footnoteText string)`](footnotes.go) - 添加脚注
- [`AddEndnote(text string, endnoteText string)`](footnotes.go) - 添加尾注
- [`AddFootnoteToRun(run *Run, footnoteText string)`](footnotes.go) - 为运行添加脚注
- [`SetFootnoteConfig(config *FootnoteConfig)`](footnotes.go) - 设置脚注配置
- [`GetFootnoteCount()`](footnotes.go) - 获取脚注数量
- [`GetEndnoteCount()`](footnotes.go) - 获取尾注数量
- [`RemoveFootnote(footnoteID string)`](footnotes.go) - 移除脚注
- [`RemoveEndnote(endnoteID string)`](footnotes.go) - 移除尾注

### 列表与编号功能 ✨ 新增功能
- [`AddListItem(text string, config *ListConfig)`](numbering.go) - 添加列表项
- [`AddBulletList(text string, level int, bulletType BulletType)`](numbering.go) - 添加无序列表
- [`AddNumberedList(text string, level int, numType ListType)`](numbering.go) - 添加有序列表
- [`CreateMultiLevelList(items []ListItem)`](numbering.go) - 创建多级列表
- [`RestartNumbering(numID string)`](numbering.go) - 重启编号

### 结构化文档标签 ✨ 新增功能
- [`CreateTOCSDT(title string, maxLevel int)`](sdt.go) - 创建目录SDT结构

### 模板功能 ✨ 新增功能

#### 模板渲染器（推荐使用）✨
- [`NewTemplateRenderer()`](template_engine.go) - 创建新的模板渲染器（推荐）
- [`SetLogging(enabled bool)`](template_engine.go) - 设置日志记录
- [`LoadTemplateFromFile(name, filePath string)`](template_engine.go) - 从DOCX文件加载模板
- [`RenderTemplate(templateName string, data *TemplateData)`](template_engine.go) - 渲染模板（最推荐方法）
- [`AnalyzeTemplate(templateName string)`](template_engine.go) - 分析模板结构

#### 模板引擎（底层API）
- [`NewTemplateEngine()`](template.go) - 创建新的模板引擎
- [`LoadTemplate(name, content string)`](template.go) - 从字符串加载模板
- [`LoadTemplateFromDocument(name string, doc *Document)`](template.go) - 从现有文档创建模板
- [`GetTemplate(name string)`](template.go) - 获取缓存的模板
- [`RenderTemplateToDocument(templateName string, data *TemplateData)`](template.go) - 渲染模板到新文档（推荐方法）
- [`RenderToDocument(templateName string, data *TemplateData)`](template.go) - 渲染模板到新文档（传统方法）
- [`ValidateTemplate(template *Template)`](template.go) - 验证模板语法
- [`ClearCache()`](template.go) - 清空模板缓存
- [`RemoveTemplate(name string)`](template.go) - 移除指定模板

#### 模板引擎功能特性 ✨
**变量替换**: 支持 `{{变量名}}` 语法进行动态内容替换
**条件语句**: 支持 `{{#if 条件}}...{{/if}}` 语法进行条件渲染
**循环语句**: 支持 `{{#each 列表}}...{{/each}}` 语法进行列表渲染
**模板继承**: 支持 `{{extends "基础模板"}}` 语法和 `{{#block "块名"}}...{{/block}}` 块重写机制，实现真正的模板继承
  - **块定义**: 在基础模板中定义可重写的内容块
  - **块重写**: 在子模板中选择性重写特定块，未重写的块保持父模板默认内容
  - **多级继承**: 支持模板的多层继承关系
  - **完整保留**: 未重写的块完整保留父模板的默认内容和格式
**循环内条件**: 完美支持循环内部的条件表达式，如 `{{#each items}}{{#if isActive}}...{{/if}}{{/each}}`
**数据类型支持**: 支持字符串、数字、布尔值、对象等多种数据类型
**结构体绑定**: 支持从Go结构体自动生成模板数据
**模板分析**: ✨ **新增功能** 自动分析模板结构，提取变量、列表、条件和表格信息
  - **结构分析**: 识别模板中使用的所有变量、列表和条件
  - **表格分析**: 专门分析表格中的模板语法和循环结构
  - **依赖检查**: 检查模板的数据依赖关系
  - **示例数据生成**: 根据分析结果自动生成示例数据结构
**日志记录**: ✨ **新增功能** 完善的日志系统，支持模板加载、渲染和分析过程的详细记录
**数据验证**: ✨ **新增功能** 自动验证模板数据的完整性和格式正确性
**DOCX模板支持**: ✨ **新增功能** 直接从现有DOCX文件加载模板
**页眉页脚模板支持**: ✨ **新增功能** 完整支持页眉页脚中的模板变量
  - **变量识别**: 自动识别页眉页脚中的 `{{变量名}}` 语法
  - **变量替换**: 渲染时自动替换页眉页脚中的模板变量
  - **条件语句**: 支持页眉页脚中的条件渲染
  - **模板分析**: `AnalyzeTemplate` 会自动分析页眉页脚中的变量

### 模板数据操作
- [`NewTemplateData()`](template.go) - 创建新的模板数据
- [`SetVariable(name string, value interface{})`](template.go) - 设置变量
- [`SetList(name string, list []interface{})`](template.go) - 设置列表
- [`SetCondition(name string, value bool)`](template.go) - 设置条件
- [`SetVariables(variables map[string]interface{})`](template.go) - 批量设置变量
- [`GetVariable(name string)`](template.go) - 获取变量
- [`GetList(name string)`](template.go) - 获取列表
- [`GetCondition(name string)`](template.go) - 获取条件
- [`Merge(other *TemplateData)`](template.go) - 合并模板数据
- [`Clear()`](template.go) - 清空模板数据
- [`FromStruct(data interface{})`](template.go) - 从结构体生成模板数据

### 模板继承详细使用说明 ✨ **新增功能**

模板继承是WordZero模板引擎的高级功能，允许基于现有模板创建扩展模板，通过块定义和重写机制实现模板的复用和扩展。

#### 基础语法

**1. 基础模板块定义**
```go
// 定义带有可重写块的基础模板
baseTemplate := `{{companyName}} 报告

{{#block "header"}}
默认标题内容
日期：{{reportDate}}
{{/block}}

{{#block "summary"}}
默认摘要内容
{{/block}}

{{#block "main_content"}}
默认主要内容
{{/block}}

{{#block "footer"}}
报告人：{{reporterName}}
{{/block}}`

engine.LoadTemplate("base_report", baseTemplate)
```

**2. 子模板继承和块重写**
```go
// 创建继承基础模板的子模板
childTemplate := `{{extends "base_report"}}

{{#block "summary"}}
销售业绩摘要
本月销售目标已达成 {{achievementRate}}%
{{/block}}

{{#block "main_content"}}
详细销售数据：
- 总销售额：{{totalSales}}
- 新增客户：{{newCustomers}}
- 成交订单：{{orders}}
{{/block}}`

engine.LoadTemplate("sales_report", childTemplate)
```

#### 继承特性

**块重写策略**：
- 重写的块完全替换父模板中的对应块
- 未重写的块保持父模板的默认内容
- 支持选择性重写，灵活性极高

**多级继承**：
```go
// 第一级：基础模板
baseTemplate := `{{#block "document"}}基础文档{{/block}}`

// 第二级：业务模板
businessTemplate := `{{extends "base"}}
{{#block "document"}}
{{#block "business_header"}}业务标题{{/block}}
{{#block "business_content"}}业务内容{{/block}}
{{/block}}`

// 第三级：具体业务模板
salesTemplate := `{{extends "business"}}
{{#block "business_header"}}销售报告{{/block}}
{{#block "business_content"}}销售数据分析{{/block}}`
```

#### 实际应用示例

```go
func demonstrateTemplateInheritance() {
    engine := document.NewTemplateEngine()
    
    // 基础报告模板
    baseTemplate := `{{companyName}} 工作报告
报告日期：{{reportDate}}

{{#block "summary"}}
默认摘要内容
{{/block}}

{{#block "main_content"}}
默认主要内容
{{/block}}

{{#block "conclusion"}}
默认结论
{{/block}}

{{#block "signature"}}
报告人：{{reporterName}}
部门：{{department}}
{{/block}}`
    
    engine.LoadTemplate("base_report", baseTemplate)
    
    // 销售报告模板（重写部分块）
    salesTemplate := `{{extends "base_report"}}

{{#block "summary"}}
销售业绩摘要
本月销售目标已达成 {{achievementRate}}%
{{/block}}

{{#block "main_content"}}
销售数据分析
- 总销售额：{{totalSales}}
- 新增客户：{{newCustomers}}
- 成交订单：{{orders}}

{{#each channels}}
- {{name}}：{{sales}} ({{percentage}}%)
{{/each}}
{{/block}}`
    
    engine.LoadTemplate("sales_report", salesTemplate)
    
    // 准备数据并渲染
    data := document.NewTemplateData()
    data.SetVariable("companyName", "WordZero科技")
    data.SetVariable("reportDate", "2024年12月01日")
    data.SetVariable("reporterName", "张三")
    data.SetVariable("department", "销售部")
    data.SetVariable("achievementRate", "125")
    data.SetVariable("totalSales", "1,850,000")
    data.SetVariable("newCustomers", "45")
    data.SetVariable("orders", "183")
    
    channels := []interface{}{
        map[string]interface{}{"name": "线上电商", "sales": "742,000", "percentage": "40.1"},
        map[string]interface{}{"name": "直销团队", "sales": "555,000", "percentage": "30.0"},
    }
    data.SetList("channels", channels)
    
    // 渲染并保存（推荐方法）
    doc, _ := engine.RenderTemplateToDocument("sales_report", data)
    doc.Save("sales_report.docx")
}
```

#### 优势与应用场景

**主要优势**：
- **代码复用**：避免重复定义相同的模板结构
- **维护性**：修改基础模板自动影响所有子模板
- **灵活性**：可选择性重写需要的部分，保留其他默认内容
- **扩展性**：支持多级继承，构建复杂的模板层次结构

**典型应用场景**：
- **企业报告体系**：基础报告模板+各部门专用模板
- **文档标准化**：统一格式的不同类型文档（合同、发票、通知等）
- **多语言文档**：相同结构不同语言的文档模板
- **品牌一致性**：保持企业品牌元素的统一性

### 图片操作功能 ✨ 新增功能
- [`AddImageFromFile(filePath string, config *ImageConfig)`](image.go) - 从文件添加图片
- [`AddImageFromData(imageData []byte, fileName string, format ImageFormat, width, height int, config *ImageConfig)`](image.go) - 从数据添加图片
- [`ResizeImage(imageInfo *ImageInfo, size *ImageSize)`](image.go) - 调整图片大小
- [`SetImagePosition(imageInfo *ImageInfo, position ImagePosition, offsetX, offsetY float64)`](image.go) - 设置图片位置
- [`SetImageWrapText(imageInfo *ImageInfo, wrapText ImageWrapText)`](image.go) - 设置图片文字环绕
- [`SetImageAltText(imageInfo *ImageInfo, altText string)`](image.go) - 设置图片替代文字
- [`SetImageTitle(imageInfo *ImageInfo, title string)`](image.go) - 设置图片标题

## 段落操作方法

### 段落格式设置
- [`SetAlignment(alignment AlignmentType)`](document.go) - 设置段落对齐方式
- [`SetSpacing(config *SpacingConfig)`](document.go) - 设置段落间距
- [`SetStyle(styleID string)`](document.go) - 设置段落样式
- [`SetIndentation(firstLineCm, leftCm, rightCm float64)`](document.go) - 设置段落缩进 ✨ **已完善**
- [`SetKeepWithNext(keep bool)`](document.go) - 设置与下一段落保持在同一页 ✨ **新增**
- [`SetKeepLines(keep bool)`](document.go) - 设置段落所有行保持在同一页 ✨ **新增**
- [`SetPageBreakBefore(pageBreak bool)`](document.go) - 设置段前分页 ✨ **新增**
- [`SetWidowControl(control bool)`](document.go) - 设置孤行控制 ✨ **新增**
- [`SetOutlineLevel(level int)`](document.go) - 设置大纲级别 ✨ **新增**
- [`SetParagraphFormat(config *ParagraphFormatConfig)`](document.go) - 一次性设置所有段落格式属性 ✨ **新增**

#### 段落格式高级功能 ✨ **新增功能**

WordZero现在支持完整的段落格式自定义功能，提供与Microsoft Word相同的高级段落控制选项。

**分页控制功能**：
- **SetKeepWithNext** - 确保段落与下一段落保持在同一页，避免标题单独出现在页面底部
- **SetKeepLines** - 防止段落被分页拆分，保持段落完整性
- **SetPageBreakBefore** - 在段落前强制插入分页符，常用于章节开始

**孤行控制**：
- **SetWidowControl** - 防止段落第一行或最后一行单独出现在页面顶部或底部，提升排版质量

**大纲级别**：
- **SetOutlineLevel** - 设置段落的大纲级别（0-8），用于文档导航窗格显示和目录生成

**综合格式设置**：
- **SetParagraphFormat** - 使用`ParagraphFormatConfig`结构一次性设置所有段落属性
  - 基础格式：对齐方式、样式
  - 间距设置：行间距、段前段后间距、首行缩进
  - 缩进设置：首行缩进、左右缩进（支持悬挂缩进）
  - 分页控制：与下段保持、行保持、段前分页、孤行控制
  - 大纲级别：0-8级别设置

**使用示例**：

```go
// 方法1：使用单独的方法设置
title := doc.AddParagraph("第一章 概述")
title.SetAlignment(document.AlignCenter)
title.SetKeepWithNext(true)
title.SetPageBreakBefore(true)
title.SetOutlineLevel(0)

// 方法2：使用SetParagraphFormat一次性设置
para := doc.AddParagraph("重要内容")
para.SetParagraphFormat(&document.ParagraphFormatConfig{
    Alignment:       document.AlignJustify,
    Style:           "Normal",
    LineSpacing:     1.5,
    BeforePara:      12,
    AfterPara:       6,
    FirstLineCm:     0.5,
    KeepWithNext:    true,
    KeepLines:       true,
    WidowControl:    true,
    OutlineLevel:    0,
})
```

**应用场景**：
- **文档结构化** - 使用大纲级别创建清晰的文档层次结构
- **专业排版** - 使用分页控制确保标题和内容的关联性
- **内容保护** - 使用行保持防止重要段落被分页
- **章节管理** - 使用段前分页实现章节的页面独立性

### 段落内容操作
- [`AddFormattedText(text string, format *TextFormat)`](document.go) - 添加格式化文本
- [`AddPageBreak()`](document.go) - 向段落添加分页符 ✨ **新增**
- [`ElementType()`](document.go) - 获取段落元素类型

## 文档主体操作方法

### 元素查询
- [`GetParagraphs()`](document.go) - 获取所有段落
- [`GetTables()`](document.go) - 获取所有表格

### 元素添加
- [`AddElement(element interface{})`](document.go) - 添加元素到文档主体

## 表格操作方法

### 表格创建
- [`CreateTable(config *TableConfig)`](table.go#L161) - 创建新表格（✨ 新增：默认包含单线边框样式）
- [`AddTable(config *TableConfig)`](table.go#L257) - 添加表格到文档

### 行操作
- [`InsertRow(position int, data []string)`](table.go#L271) - 在指定位置插入行
- [`AppendRow(data []string)`](table.go#L329) - 在表格末尾添加行
- [`DeleteRow(rowIndex int)`](table.go#L334) - 删除指定行
- [`DeleteRows(startIndex, endIndex int)`](table.go#L351) - 删除多行
- [`GetRowCount()`](table.go#L562) - 获取行数

### 列操作
- [`InsertColumn(position int, data []string, width int)`](table.go#L369) - 在指定位置插入列
- [`AppendColumn(data []string, width int)`](table.go#L438) - 在表格末尾添加列
- [`DeleteColumn(colIndex int)`](table.go#L447) - 删除指定列
- [`DeleteColumns(startIndex, endIndex int)`](table.go#L474) - 删除多列
- [`GetColumnCount()`](table.go#L567) - 获取列数

### 单元格操作
- [`GetCell(row, col int)`](table.go#L502) - 获取指定单元格
- [`SetCellText(row, col int, text string)`](table.go#L515) - 设置单元格文本
- [`GetCellText(row, col int)`](table.go#L623) - 获取单元格文本（已升级：返回单元格内所有段落与 Run 的完整内容，段落之间使用 `\n` 分隔）
    - 旧行为：仅返回第一个段落的第一个 Run 文本，导致多行/软换行内容丢失
    - 新行为：遍历所有段落与其下所有 Run，拼接文本；空段落跳过内容但仍产生段落换行（除末尾）
    - 注意：如果未来需要保留 Word 中 `<w:br/>`（同一段落内的手动软换行），需扩展解析逻辑；当前仅按段落分隔
- [`SetCellFormat(row, col int, format *CellFormat)`](table.go#L691) - 设置单元格格式
- [`GetCellFormat(row, col int)`](table.go#L1238) - 获取单元格格式

### 单元格文本格式化
- [`SetCellFormattedText(row, col int, text string, format *TextFormat)`](table.go#L780) - 设置格式化文本
- [`AddCellFormattedText(row, col int, text string, format *TextFormat)`](table.go#L833) - 添加格式化文本

### 单元格合并
- [`MergeCellsHorizontal(row, startCol, endCol int)`](table.go#L887) - 水平合并单元格
- [`MergeCellsVertical(startRow, endRow, col int)`](table.go#L924) - 垂直合并单元格
- [`MergeCellsRange(startRow, endRow, startCol, endCol int)`](table.go#L971) - 范围合并单元格
- [`UnmergeCells(row, col int)`](table.go#L1004) - 取消合并单元格
- [`IsCellMerged(row, col int)`](table.go#L1074) - 检查单元格是否已合并
- [`GetMergedCellInfo(row, col int)`](table.go#L1098) - 获取合并单元格信息

### 单元格特殊属性
- [`SetCellPadding(row, col int, padding int)`](table.go#L1189) - 设置单元格内边距
- [`SetCellTextDirection(row, col int, direction CellTextDirection)`](table.go#L1202) - 设置文字方向
- [`GetCellTextDirection(row, col int)`](table.go#L1223) - 获取文字方向
- [`ClearCellContent(row, col int)`](table.go#L1138) - 清除单元格内容
- [`ClearCellFormat(row, col int)`](table.go#L1156) - 清除单元格格式

### 表格整体操作
- [`ClearTable()`](table.go#L575) - 清空表格内容
- [`CopyTable()`](table.go#L593) - 复制表格
- [`ElementType()`](table.go#L66) - 获取表格元素类型

### 行高设置
- [`SetRowHeight(rowIndex int, config *RowHeightConfig)`](table.go#L1318) - 设置行高
- [`GetRowHeight(rowIndex int)`](table.go#L1339) - 获取行高
- [`SetRowHeightRange(startRow, endRow int, config *RowHeightConfig)`](table.go#L1371) - 设置多行行高

### 表格布局与对齐
- [`SetTableLayout(config *TableLayoutConfig)`](table.go#L1447) - 设置表格布局
- [`GetTableLayout()`](table.go#L1473) - 获取表格布局
- [`SetTableAlignment(alignment TableAlignment)`](table.go#L1488) - 设置表格对齐

### 行属性设置
- [`SetRowKeepTogether(rowIndex int, keepTogether bool)`](table.go#L1529) - 设置行保持完整
- [`SetRowAsHeader(rowIndex int, isHeader bool)`](table.go#L1552) - 设置行为标题行
- [`SetHeaderRows(startRow, endRow int)`](table.go#L1575) - 设置多行为标题行
- [`IsRowHeader(rowIndex int)`](table.go#L1600) - 检查是否为标题行
- [`IsRowKeepTogether(rowIndex int)`](table.go#L1614) - 检查行是否保持完整
- [`SetRowKeepWithNext(rowIndex int, keepWithNext bool)`](table.go#L1645) - 设置与下一行保持在一起

### 表格分页设置
- [`SetTablePageBreak(config *TablePageBreakConfig)`](table.go#L1636) - 设置表格分页
- [`GetTableBreakInfo()`](table.go#L1657) - 获取分页信息

### 表格样式
- [`ApplyTableStyle(config *TableStyleConfig)`](table.go#L1956) - 应用表格样式
- [`CreateCustomTableStyle(styleID, styleName string, borderConfig *TableBorderConfig, shadingConfig *ShadingConfig, firstRowBold bool)`](table.go#L2213) - 创建自定义表格样式

### 边框设置
- [`SetTableBorders(config *TableBorderConfig)`](table.go#L2038) - 设置表格边框
- [`SetCellBorders(row, col int, config *CellBorderConfig)`](table.go#L2085) - 设置单元格边框
- [`RemoveTableBorders()`](table.go#L2168) - 移除表格边框
- [`RemoveCellBorders(row, col int)`](table.go#L2194) - 移除单元格边框

### 背景与阴影
- [`SetTableShading(config *ShadingConfig)`](table.go#L2069) - 设置表格底纹
- [`SetCellShading(row, col int, config *ShadingConfig)`](table.go#L2121) - 设置单元格底纹
- [`SetAlternatingRowColors(evenRowColor, oddRowColor string)`](table.go#L2142) - 设置交替行颜色

### 单元格图片功能 ✨ **新功能**

支持向表格单元格中添加图片：

- [`AddCellImage(table *Table, row, col int, config *CellImageConfig)`](image.go#L1106) - 向单元格添加图片（完整配置）
- [`AddCellImageFromFile(table *Table, row, col int, filePath string, widthMM float64)`](image.go#L1214) - 从文件向单元格添加图片
- [`AddCellImageFromData(table *Table, row, col int, data []byte, widthMM float64)`](image.go#L1236) - 从二进制数据向单元格添加图片

#### CellImageConfig - 单元格图片配置
```go
type CellImageConfig struct {
    FilePath        string      // 图片文件路径
    Data            []byte      // 图片二进制数据（与FilePath二选一）
    Format          ImageFormat // 图片格式（当使用Data时需要指定）
    Width           float64     // 图片宽度（毫米），0表示自动
    Height          float64     // 图片高度（毫米），0表示自动
    KeepAspectRatio bool        // 是否保持宽高比
    AltText         string      // 图片替代文字
    Title           string      // 图片标题
}
```

#### 表格单元格图片使用示例
```go
// 创建表格
table, err := doc.AddTable(&document.TableConfig{
    Rows:  2,
    Cols:  2,
    Width: 8000,
})

// 方式1：从文件添加图片到单元格
imageInfo, err := doc.AddCellImageFromFile(table, 0, 0, "logo.png", 30) // 30mm宽度

// 方式2：从二进制数据添加图片
imageData := []byte{...} // 图片二进制数据
imageInfo, err := doc.AddCellImageFromData(table, 0, 1, imageData, 25) // 25mm宽度

// 方式3：使用完整配置
config := &document.CellImageConfig{
    FilePath:        "product.jpg",
    Width:           50,     // 50mm宽度
    Height:          40,     // 40mm高度
    KeepAspectRatio: false,  // 不保持宽高比
    AltText:         "产品图片",
    Title:           "产品展示",
}
imageInfo, err := doc.AddCellImage(table, 1, 0, config)
```

**注意事项**：
- 图片通过 `Document` 对象的方法添加，因为图片资源需要在文档级别管理
- 支持 PNG、JPEG、GIF 格式的图片
- 宽度/高度单位为毫米，设置为0时使用原始尺寸
- 当设置 `KeepAspectRatio` 为 `true` 时，只需设置宽度或高度其中之一

### 单元格遍历迭代器 ✨ **新功能**

提供强大的单元格遍历和查找功能：

##### CellIterator - 单元格迭代器
```go
// 创建迭代器
iterator := table.NewCellIterator()

// 遍历所有单元格
for iterator.HasNext() {
    cellInfo, err := iterator.Next()
    if err != nil {
        break
    }
    fmt.Printf("单元格[%d,%d]: %s\n", cellInfo.Row, cellInfo.Col, cellInfo.Text)
}

// 获取进度
progress := iterator.Progress() // 0.0 - 1.0

// 重置迭代器
iterator.Reset()
```

##### ForEach 批量处理
```go
// 遍历所有单元格
err := table.ForEach(func(row, col int, cell *TableCell, text string) error {
    // 处理每个单元格
    return nil
})

// 按行遍历
err := table.ForEachInRow(rowIndex, func(col int, cell *TableCell, text string) error {
    // 处理行中的每个单元格
    return nil
})

// 按列遍历
err := table.ForEachInColumn(colIndex, func(row int, cell *TableCell, text string) error {
    // 处理列中的每个单元格
    return nil
})
```

##### 范围操作
```go
// 获取指定范围的单元格
cells, err := table.GetCellRange(startRow, startCol, endRow, endCol)
for _, cellInfo := range cells {
    fmt.Printf("单元格[%d,%d]: %s\n", cellInfo.Row, cellInfo.Col, cellInfo.Text)
}
```

##### 查找功能
```go
// 自定义条件查找
cells, err := table.FindCells(func(row, col int, cell *TableCell, text string) bool {
    return strings.Contains(text, "关键词")
})

// 按文本查找
exactCells, err := table.FindCellsByText("精确匹配", true)
fuzzyCells, err := table.FindCellsByText("模糊", false)
```

##### CellInfo 结构
```go
type CellInfo struct {
    Row    int        // 行索引
    Col    int        // 列索引
    Cell   *TableCell // 单元格引用
    Text   string     // 单元格文本
    IsLast bool       // 是否为最后一个单元格
}
```

## 工具函数

### 日志系统
- [`NewLogger(level LogLevel, output io.Writer)`](logger.go#L56) - 创建新的日志记录器
- [`SetGlobalLevel(level LogLevel)`](logger.go#L129) - 设置全局日志级别
- [`SetGlobalOutput(output io.Writer)`](logger.go#L134) - 设置全局日志输出
- [`Debug(msg string)`](logger.go#L159) - 输出调试信息
- [`Info(msg string)`](logger.go#L164) - 输出信息
- [`Warn(msg string)`](logger.go#L169) - 输出警告
- [`Error(msg string)`](logger.go#L174) - 输出错误

### 错误处理
- [`NewDocumentError(operation string, cause error, context string)`](errors.go#L47) - 创建文档错误
- [`WrapError(operation string, err error)`](errors.go#L56) - 包装错误
- [`WrapErrorWithContext(operation string, err error, context string)`](errors.go#L64) - 带上下文包装错误
- [`NewValidationError(field, value, message string)`](errors.go#L84) - 创建验证错误

### 域字段工具 ✨ 新增功能
- [`CreateHyperlinkField(anchor string)`](field.go) - 创建超链接域
- [`CreatePageRefField(anchor string)`](field.go) - 创建页码引用域

## 常用配置结构

### 文本格式
- `TextFormat` - 文本格式配置
- `AlignmentType` - 对齐类型
- `SpacingConfig` - 间距配置

### 表格配置
- `TableConfig` - 表格基础配置
- `CellFormat` - 单元格格式
- `RowHeightConfig` - 行高配置
- `TableLayoutConfig` - 表格布局配置
- `TableStyleConfig` - 表格样式配置
- `BorderConfig` - 边框配置
- `ShadingConfig` - 底纹配置

### 页面设置配置 ✨ 新增
- `PageSettings` - 页面设置配置
- `PageSize` - 页面尺寸类型（A4、Letter、Legal、A3、A5、Custom）
- `PageOrientation` - 页面方向（Portrait纵向、Landscape横向）
- `SectionProperties` - 节属性（包含页面设置信息）

### 页眉页脚配置 ✨ 新增
- `HeaderFooterType` - 页眉页脚类型（Default、First、Even）
- `Header` - 页眉结构
- `Footer` - 页脚结构
- `HeaderFooterReference` - 页眉页脚引用
- `PageNumber` - 页码字段

### 目录配置 ✨ 新增
- `TOCConfig` - 目录配置
- `TOCEntry` - 目录条目
- `Bookmark` - 书签结构
- `BookmarkEnd` - 书签结束标记

### 脚注尾注配置 ✨ 新增
- `FootnoteConfig` - 脚注配置
- `FootnoteType` - 脚注类型（Footnote脚注、Endnote尾注）
- `FootnoteNumberFormat` - 脚注编号格式
- `FootnoteRestart` - 脚注重新开始规则
- `FootnotePosition` - 脚注位置
- `Footnote` - 脚注结构
- `Endnote` - 尾注结构

### 列表编号配置 ✨ 新增
- `ListConfig` - 列表配置
- `ListType` - 列表类型（Bullet无序、Number有序等）
- `BulletType` - 项目符号类型
- `ListItem` - 列表项结构
- `Numbering` - 编号定义
- `AbstractNum` - 抽象编号定义
- `Level` - 编号级别

### 结构化文档标签配置 ✨ 新增
- `SDT` - 结构化文档标签
- `SDTProperties` - SDT属性
- `SDTContent` - SDT内容

### 域字段配置 ✨ 新增
- `FieldChar` - 域字符
- `InstrText` - 域指令文本
- `HyperlinkField` - 超链接域
- `PageRefField` - 页码引用域

### 图片配置 ✨ 新增
- `ImageConfig` - 图片配置
- `ImageSize` - 图片尺寸配置
- `ImageFormat` - 图片格式（PNG、JPEG、GIF）
- `ImagePosition` - 图片位置（inline、floatLeft、floatRight）
- `ImageWrapText` - 文字环绕类型（none、square、tight、topAndBottom）
- `ImageInfo` - 图片信息结构
- `AlignmentType` - 对齐方式（left、center、right、justify）

## 使用示例

```go
// 创建新文档
doc := document.New()

// ✨ 新增：页面设置示例
// 设置页面为A4横向
doc.SetPageOrientation(document.OrientationLandscape)

// 设置自定义边距（上下左右：25mm）
doc.SetPageMargins(25, 25, 25, 25)

// 设置自定义页面尺寸（200mm x 300mm）
doc.SetCustomPageSize(200, 300)

// 或者使用完整页面设置
pageSettings := &document.PageSettings{
    Size:           document.PageSizeLetter,
    Orientation:    document.OrientationPortrait,
    MarginTop:      30,
    MarginRight:    20,
    MarginBottom:   30,
    MarginLeft:     20,
    HeaderDistance: 15,
    FooterDistance: 15,
    GutterWidth:    0,
}
doc.SetPageSettings(pageSettings)

// ✨ 新增：页眉页脚示例
// 添加页眉
doc.AddHeader(document.HeaderFooterTypeDefault, "这是页眉")

// 添加带页码的页脚
doc.AddFooterWithPageNumber(document.HeaderFooterTypeDefault, "第", true)

// 设置首页不同
doc.SetDifferentFirstPage(true)

// ✨ 新增：目录示例
// 添加带书签的标题
doc.AddHeadingWithBookmark("第一章 概述", 1, "chapter1")
doc.AddHeadingWithBookmark("1.1 背景", 2, "section1_1")

// 生成目录
tocConfig := document.DefaultTOCConfig()
tocConfig.Title = "目录"
tocConfig.MaxLevel = 3
doc.GenerateTOC(tocConfig)

// ✨ 新增：脚注示例
// 添加脚注
doc.AddFootnote("这是正文内容", "这是脚注内容")

// 添加尾注
doc.AddEndnote("更多说明", "这是尾注内容")

// ✨ 新增：列表示例
// 添加无序列表
doc.AddBulletList("列表项1", 0, document.BulletTypeDot)
doc.AddBulletList("列表项2", 1, document.BulletTypeCircle)

// 添加有序列表
doc.AddNumberedList("编号项1", 0, document.ListTypeNumber)

// ✨ 新增：图片示例
// 从文件添加图片
imageInfo, err := doc.AddImageFromFile("path/to/image.png", &document.ImageConfig{
    Size: &document.ImageSize{
        Width:  100.0, // 100毫米宽度
        Height: 75.0,  // 75毫米高度
    },
    Position: document.ImagePositionInline,
    WrapText: document.ImageWrapNone,
    AltText:  "示例图片",
    Title:    "这是一个示例图片",
})

// 从数据添加图片
imageData := []byte{...} // 图片二进制数据
imageInfo2, err := doc.AddImageFromData(
    imageData,
    "example.png",
    document.ImageFormatPNG,
    200, 150, // 原始像素尺寸
    &document.ImageConfig{
        Size: &document.ImageSize{
            Width:           60.0, // 只设置宽度
            KeepAspectRatio: true, // 保持长宽比
        },
        AltText: "数据图片",
    },
)

// 调整图片大小
err = doc.ResizeImage(imageInfo, &document.ImageSize{
    Width:  80.0,
    Height: 60.0,
})

// 设置图片属性
err = doc.SetImagePosition(imageInfo, document.ImagePositionFloatLeft, 5.0, 0.0)
err = doc.SetImageWrapText(imageInfo, document.ImageWrapSquare)
err = doc.SetImageAltText(imageInfo, "更新的替代文字")
err = doc.SetImageTitle(imageInfo, "更新的标题")

// ✨ 新增：设置图片对齐方式（仅适用于嵌入式图片）
err = doc.SetImageAlignment(imageInfo, document.AlignCenter)  // 居中对齐
err = doc.SetImageAlignment(imageInfo, document.AlignLeft)    // 左对齐
err = doc.SetImageAlignment(imageInfo, document.AlignRight)   // 右对齐
doc.AddNumberedList("第一项", 0, document.ListTypeDecimal)
doc.AddNumberedList("第二项", 0, document.ListTypeDecimal)

// 添加段落
para := doc.AddParagraph("这是一个段落")
para.SetAlignment(document.AlignCenter)

// 创建表格
table := doc.CreateTable(&document.TableConfig{
    Rows:  3,
    Cols:  3,
    Width: 5000,
})

// 设置单元格内容
table.SetCellText(0, 0, "标题")

// 保存文档
doc.Save("example.docx")
```

## 注意事项

1. 所有位置索引都是从0开始
2. 宽度单位使用磅（pt），1磅 = 20twips
3. 颜色使用十六进制格式，如 "FF0000" 表示红色
4. 在操作表格前请确保行列索引有效，否则可能返回错误
5. 页眉页脚类型包括：Default（默认）、First（首页）、Even（偶数页）
6. 目录功能需要先添加带书签的标题，然后调用生成目录方法
7. 脚注和尾注会自动编号，支持多种编号格式和重启规则
8. 列表支持多级嵌套，最多支持9级缩进
9. 结构化文档标签主要用于目录等特殊功能的实现
10. 图片支持PNG、JPEG、GIF格式，会自动嵌入到文档中
11. 图片尺寸可以用毫米或像素指定，支持保持长宽比的缩放
12. 图片位置支持嵌入式、左浮动、右浮动等多种布局方式
13. 图片对齐功能仅适用于嵌入式图片（ImagePositionInline），浮动图片请使用位置控制

## Markdown转Word功能 ✨ **新增功能**

WordZero现在支持将Markdown文档转换为Word格式，基于goldmark解析引擎实现，提供高质量的转换效果。

### Markdown包API

#### 转换器接口
- [`NewConverter(options *ConvertOptions)`](../markdown/converter.go) - 创建新的Markdown转换器
- [`DefaultOptions()`](../markdown/config.go) - 获取默认转换选项
- [`HighQualityOptions()`](../markdown/config.go) - 获取高质量转换选项

#### 转换方法
- [`ConvertString(content string, options *ConvertOptions)`](../markdown/converter.go) - 转换Markdown字符串为Word文档
- [`ConvertBytes(content []byte, options *ConvertOptions)`](../markdown/converter.go) - 转换Markdown字节数组为Word文档
- [`ConvertFile(mdPath, docxPath string, options *ConvertOptions)`](../markdown/converter.go) - 转换Markdown文件为Word文件
- [`BatchConvert(inputs []string, outputDir string, options *ConvertOptions)`](../markdown/converter.go) - 批量转换Markdown文件

#### 配置选项 (`ConvertOptions`)
- `EnableGFM` - 启用GitHub Flavored Markdown支持
- `EnableFootnotes` - 启用脚注支持
- `EnableTables` - 启用表格支持
- `EnableTaskList` - 启用任务列表支持
- `StyleMapping` - 自定义样式映射
- `DefaultFontFamily` - 默认字体族
- `DefaultFontSize` - 默认字体大小
- `ImageBasePath` - 图片基础路径
- `EmbedImages` - 是否嵌入图片
- `MaxImageWidth` - 最大图片宽度（英寸）
- `PreserveLinkStyle` - 保留链接样式
- `ConvertToBookmarks` - 内部链接转书签
- `GenerateTOC` - 生成目录
- `TOCMaxLevel` - 目录最大级别
- `PageSettings` - 页面设置
- `StrictMode` - 严格模式
- `IgnoreErrors` - 忽略转换错误
- `ErrorCallback` - 错误回调函数
- `ProgressCallback` - 进度回调函数

### 支持的Markdown语法

#### 基础语法
- **标题** (`# ## ### #### ##### ######`) - 转换为Word标题样式1-6
- **段落** - 转换为Word正文段落
- **粗体** (`**文本**`) - 转换为粗体格式
- **斜体** (`*文本*`) - 转换为斜体格式
- **行内代码** (`` `代码` ``) - 转换为等宽字体
- **代码块** (``` ```) - 转换为代码块样式

#### 列表支持
- **无序列表** (`- * +`) - 转换为Word项目符号列表
- **有序列表** (`1. 2. 3.`) - 转换为Word编号列表
- **多级列表** - 支持嵌套列表结构

#### GitHub Flavored Markdown扩展 ✨ **新增**
- **表格** (`| 列1 | 列2 |`) - 转换为Word表格
  - 支持表头自动识别和样式设置
  - 支持对齐控制（左对齐 `:---`、居中 `:---:`、右对齐 `---:`）
  - 自动设置表格边框和单元格格式
- **任务列表** (`- [x] 已完成` / `- [ ] 未完成`) - 转换为复选框符号
  - ☑ 表示已完成任务
  - ☐ 表示未完成任务
  - 支持嵌套任务列表
  - 支持混合格式（粗体、斜体、代码等）

#### 其他元素
- **引用块** (`> 引用文本`) - 转换为斜体引用样式
- **分割线** (`---`) - 转换为水平线
- **链接** (`[文本](URL)`) - 转换为蓝色文本（后续支持超链接）
- **图片** (`![alt](src)`) - 转换为图片占位符（后续支持图片嵌入）

### 使用示例

#### 基础字符串转换
```go
import "github.com/mr-pmillz/wordZero/pkg/markdown"

// 创建转换器
converter := markdown.NewConverter(markdown.DefaultOptions())

// 转换Markdown字符串
markdownText := `# 标题

这是一个包含**粗体**和*斜体*的段落。

## 子标题

- 列表项1
- 列表项2

> 引用文本

` + "`" + `代码示例` + "`" + `
`

doc, err := converter.ConvertString(markdownText, nil)
if err != nil {
    log.Fatal(err)
}

// 保存Word文档
err = doc.Save("output.docx")
```

#### 表格和任务列表示例 ✨ **新增**
```go
// 启用表格和任务列表功能
options := markdown.DefaultOptions()
options.EnableTables = true
options.EnableTaskList = true

converter := markdown.NewConverter(options)

// 包含表格和任务列表的Markdown
markdownWithTable := `# 项目进度表

## 功能实现状态

| 功能名称 | 状态 | 负责人 |
|:---------|:----:|-------:|
| 表格转换 | ✅ | 张三 |
| 任务列表 | ✅ | 李四 |
| 图片处理 | 🚧 | 王五 |

## 待办事项

- [x] 实现表格转换功能
  - [x] 基础表格支持
  - [x] 对齐方式处理
  - [x] 表头样式设置
- [ ] 完善任务列表功能
  - [x] 复选框显示
  - [ ] 交互功能
- [ ] 图片嵌入支持
  - [ ] PNG格式
  - [ ] JPEG格式

## 备注

> 表格支持**左对齐**、` + "`" + `居中对齐` + "`" + `和***右对齐***三种方式
`

doc, err := converter.ConvertString(markdownWithTable, options)
if err != nil {
    log.Fatal(err)
}

err = doc.Save("project_status.docx")
```

#### 高级配置示例
```go
// 创建高质量转换配置
options := &markdown.ConvertOptions{
    EnableGFM:         true,
    EnableFootnotes:   true,
    EnableTables:      true,
    GenerateTOC:       true,
    TOCMaxLevel:       3,
    DefaultFontFamily: "Calibri",
    DefaultFontSize:   11.0,
    EmbedImages:       true,
    MaxImageWidth:     6.0,
    PageSettings: &document.PageSettings{
        Size:        document.PageSizeA4,
        Orientation: document.OrientationPortrait,
        MarginTop:   25,
        MarginRight: 20,
        MarginBottom: 25,
        MarginLeft:  20,
    },
    ProgressCallback: func(current, total int) {
        fmt.Printf("转换进度: %d/%d\n", current, total)
    },
}

converter := markdown.NewConverter(options)
```

#### 文件转换示例
```go
// 单文件转换
err := converter.ConvertFile("input.md", "output.docx", nil)

// 批量文件转换
files := []string{"doc1.md", "doc2.md", "doc3.md"}
err := converter.BatchConvert(files, "output/", options)
```

#### 自定义样式映射
```go
options := markdown.DefaultOptions()
options.StyleMapping = map[string]string{
    "heading1": "CustomTitle",
    "heading2": "CustomSubtitle", 
    "quote":    "CustomQuote",
    "code":     "CustomCode",
}

converter := markdown.NewConverter(options)
```

## Word转Markdown功能 ✨ **新增功能**

WordZero现在支持将Word文档反向转换为Markdown格式，提供完整的双向转换能力。

### Word导出器API

#### 导出器接口
- [`NewExporter(options *ExportOptions)`](../markdown/exporter.go) - 创建新的Word导出器
- [`DefaultExportOptions()`](../markdown/exporter.go) - 获取默认导出选项
- [`HighQualityExportOptions()`](../markdown/exporter.go) - 获取高质量导出选项

#### 导出方法
- [`ExportToFile(docxPath, mdPath string, options *ExportOptions)`](../markdown/exporter.go) - 导出Word文档到Markdown文件
- [`ExportToString(doc *Document, options *ExportOptions)`](../markdown/exporter.go) - 导出Word文档到Markdown字符串
- [`ExportToBytes(doc *Document, options *ExportOptions)`](../markdown/exporter.go) - 导出Word文档到Markdown字节数组
- [`BatchExport(inputs []string, outputDir string, options *ExportOptions)`](../markdown/exporter.go) - 批量导出Word文档

#### 导出配置选项 (`ExportOptions`)
- `UseGFMTables` - 使用GitHub风味Markdown表格
- `PreserveFootnotes` - 保留脚注
- `PreserveLineBreaks` - 保留换行符
- `WrapLongLines` - 自动换行
- `MaxLineLength` - 最大行长度
- `ExtractImages` - 导出图片文件
- `ImageOutputDir` - 图片输出目录
- `ImageNamePattern` - 图片命名模式
- `ImageRelativePath` - 使用相对路径
- `PreserveBookmarks` - 保留书签
- `ConvertHyperlinks` - 转换超链接
- `PreserveCodeStyle` - 保留代码样式
- `DefaultCodeLang` - 默认代码语言
- `IgnoreUnknownStyles` - 忽略未知样式
- `PreserveTOC` - 保留目录
- `IncludeMetadata` - 包含文档元数据
- `StripComments` - 删除注释
- `UseSetext` - 使用Setext样式标题
- `BulletListMarker` - 项目符号标记
- `EmphasisMarker` - 强调标记
- `StrictMode` - 严格模式
- `IgnoreErrors` - 忽略错误
- `ErrorCallback` - 错误回调函数
- `ProgressCallback` - 进度回调函数

### Word→Markdown转换映射

| Word元素 | Markdown语法 | 说明 |
|----------|-------------|------|
| Heading1-6 | `# ## ### #### ##### ######` | 标题级别对应 |
| 粗体 | `**粗体**` | 文本格式 |
| 斜体 | `*斜体*` | 文本格式 |
| 删除线 | `~~删除线~~` | 文本格式 |
| 行内代码 | `` `代码` `` | 代码格式 |
| 代码块 | ```` 代码块 ```` | 代码块 |
| 超链接 | `[链接文本](URL)` | 链接转换 |
| 图片 | `![图片](路径)` | 图片引用 |
| 表格 | `\| 表格 \|` | GFM表格格式 |
| 无序列表 | `- 项目` | 列表项 |
| 有序列表 | `1. 项目` | 编号列表 |
| 引用块 | `> 引用内容` | 引用格式 |

### Word转Markdown使用示例

#### 基础文件导出
```go
import "github.com/mr-pmillz/wordZero/pkg/markdown"

// 创建导出器
exporter := markdown.NewExporter(markdown.DefaultExportOptions())

// 导出Word文档为Markdown
err := exporter.ExportToFile("document.docx", "output.md", nil)
if err != nil {
    log.Fatal(err)
}
```

#### 导出为字符串
```go
// 打开Word文档
doc, err := document.Open("document.docx")
if err != nil {
    log.Fatal(err)
}

// 导出为Markdown字符串
exporter := markdown.NewExporter(markdown.DefaultExportOptions())
markdownText, err := exporter.ExportToString(doc, nil)
if err != nil {
    log.Fatal(err)
}

fmt.Println(markdownText)
```

#### 高质量导出配置
```go
// 高质量导出配置
options := &markdown.ExportOptions{
    UseGFMTables:      true,              // 使用GFM表格
    ExtractImages:     true,              // 导出图片
    ImageOutputDir:    "./images",        // 图片目录
    PreserveFootnotes: true,              // 保留脚注
    IncludeMetadata:   true,              // 包含元数据
    ConvertHyperlinks: true,              // 转换超链接
    PreserveCodeStyle: true,              // 保留代码样式
    UseSetext:         false,             // 使用ATX标题
    BulletListMarker:  "-",              // 使用短横线
    EmphasisMarker:    "*",              // 使用星号
    ProgressCallback: func(current, total int) {
        fmt.Printf("导出进度: %d/%d\n", current, total)
    },
}

exporter := markdown.NewExporter(options)
err := exporter.ExportToFile("complex_document.docx", "output.md", options)
```

#### 批量导出示例
```go
// 批量导出Word文档
files := []string{"doc1.docx", "doc2.docx", "doc3.docx"}

options := &markdown.ExportOptions{
    ExtractImages:     true,
    ImageOutputDir:    "extracted_images/",
    UseGFMTables:      true,
    ProgressCallback: func(current, total int) {
        fmt.Printf("批量导出进度: %d/%d\n", current, total)
    },
}

exporter := markdown.NewExporter(options)
err := exporter.BatchExport(files, "markdown_output/", options)
```

## 双向转换器 ✨ **统一接口**

### 双向转换器API
- [`NewBidirectionalConverter(mdOpts *ConvertOptions, exportOpts *ExportOptions)`](../markdown/exporter.go) - 创建双向转换器
- [`AutoConvert(inputPath, outputPath string)`](../markdown/exporter.go) - 自动检测文件类型并转换

### 双向转换使用示例

#### 自动转换
```go
import "github.com/mr-pmillz/wordZero/pkg/markdown"

// 创建双向转换器
converter := markdown.NewBidirectionalConverter(
    markdown.HighQualityOptions(),        // Markdown→Word选项
    markdown.HighQualityExportOptions(),  // Word→Markdown选项
)

// 自动检测文件类型并转换
err := converter.AutoConvert("input.docx", "output.md")     // Word→Markdown
err = converter.AutoConvert("input.md", "output.docx")     // Markdown→Word
```

#### 配置独立的转换方向
```go
// Markdown转Word配置
mdToWordOpts := &markdown.ConvertOptions{
    EnableGFM:         true,
    EnableTables:      true,
    GenerateTOC:       true,
    DefaultFontFamily: "Calibri",
    DefaultFontSize:   11.0,
}

// Word转Markdown配置
wordToMdOpts := &markdown.ExportOptions{
    UseGFMTables:      true,
    ExtractImages:     true,
    ImageOutputDir:    "./images",
    PreserveFootnotes: true,
    ConvertHyperlinks: true,
}

// 创建双向转换器
converter := markdown.NewBidirectionalConverter(mdToWordOpts, wordToMdOpts)

// 执行转换
err := converter.AutoConvert("document.docx", "document.md")
```

### 技术特性

#### 架构设计
- **goldmark集成** - 使用高性能的goldmark解析引擎
- **AST遍历** - 基于抽象语法树的转换处理
- **API复用** - 充分复用现有WordZero document API
- **向后兼容** - 不影响现有document包功能

#### 性能优势  
- **流式处理** - 支持大型文档的流式转换
- **内存效率** - 优化的内存使用模式
- **并发支持** - 批量转换支持并发处理
- **错误恢复** - 智能错误处理和恢复机制

#### 扩展性
- **插件架构** - 支持自定义渲染器扩展
- **配置驱动** - 丰富的配置选项支持不同需求
- **样式系统** - 灵活的样式映射和自定义能力
- **回调机制** - 进度和错误回调支持

### 注意事项

1. **兼容性** - 基于CommonMark 0.31.2标准，与GitHub Markdown高度兼容
2. **图片处理** - 当前版本图片转换为占位符，完整图片支持在规划中
3. **表格支持** ✨ **已完善** - 支持完整的GFM表格语法，包括对齐控制和表头样式
4. **任务列表** ✨ **已实现** - 支持任务复选框，显示为Unicode符号（☑/☐）
5. **链接处理** - 当前转换为蓝色文本，超链接功能在开发中
6. **样式映射** - 可通过StyleMapping自定义Markdown元素到Word样式的映射
7. **错误处理** - 建议在生产环境中启用错误回调，监控转换质量
8. **性能考虑** - 批量转换大量文件时建议分批处理，避免内存压力
9. **编码支持** - 完全支持UTF-8编码，包括中文等多字节字符
10. **配置要求** - 表格和任务列表功能需要在ConvertOptions中显式启用
11. **向后兼容** - 新功能不会影响现有的document包API，保持完全兼容 
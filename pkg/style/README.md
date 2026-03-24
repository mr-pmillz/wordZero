# Style Package - WordZero 样式管理系统

WordZero 的样式管理包提供了完整的 Word 文档样式系统实现，支持预定义样式、自定义样式和样式继承机制。

## 🌟 核心特性

### 🎨 完整的预定义样式库
- **标题样式**: Heading1-Heading9，支持完整的标题层次结构和导航窗格识别
- **文档样式**: Title（文档标题）、Subtitle（副标题）
- **段落样式**: Normal（正文）、Quote（引用）、ListParagraph（列表段落）、CodeBlock（代码块）
- **字符样式**: Emphasis（强调）、Strong（加粗）、CodeChar（代码字符）

### 🔧 高级样式管理
- **样式继承**: 完整的样式继承机制，自动合并父样式属性
- **自定义样式**: 快速创建和管理自定义样式
- **样式验证**: 样式存在性检查和错误处理
- **类型分类**: 按样式类型（段落、字符、表格等）管理和查询

### 🚀 便捷API接口
- **StyleManager**: 核心样式管理器，提供底层样式操作
- **QuickStyleAPI**: 高级样式操作接口，简化常用操作
- **样式信息查询**: 获取样式详情、按类型筛选、批量操作

## 📦 安装使用

```go
import "github.com/mr-pmillz/wordZero/pkg/style"
```

## 🚀 快速开始

### 创建样式管理器

```go
// 创建样式管理器（自动加载预定义样式）
styleManager := style.NewStyleManager()

// 创建快速API（推荐方式）
quickAPI := style.NewQuickStyleAPI(styleManager)

// 获取所有可用样式
allStyles := quickAPI.GetAllStylesInfo()
fmt.Printf("加载了 %d 个样式\n", len(allStyles))
```

### 使用预定义样式

```go
// 获取特定样式
heading1 := styleManager.GetStyle("Heading1")
if heading1 != nil {
    fmt.Printf("找到样式: %s\n", heading1.Name.Val)
}

// 获取所有标题样式
headingStyles := styleManager.GetHeadingStyles()
fmt.Printf("标题样式数量: %d\n", len(headingStyles))

// 获取样式详细信息
styleInfo, err := quickAPI.GetStyleInfo("Heading1")
if err == nil {
    fmt.Printf("样式名称: %s\n", styleInfo.Name)
    fmt.Printf("样式类型: %s\n", styleInfo.Type)
    fmt.Printf("样式描述: %s\n", styleInfo.Description)
}
```

### 在文档中应用样式

```go
import "github.com/mr-pmillz/wordZero/pkg/document"

// 创建文档
doc := document.New()

// 使用AddHeadingParagraph方法（推荐）
doc.AddHeadingParagraph("第一章：概述", 1)        // 自动应用Heading1样式
doc.AddHeadingParagraph("1.1 背景介绍", 2)       // 自动应用Heading2样式

// 或手动设置样式
para := doc.AddParagraph("这是引用文本")
para.SetStyle("Quote")  // 应用Quote样式

// 保存文档
doc.Save("styled_document.docx")
```

## 📋 预定义样式详细列表

### 段落样式 (Paragraph Styles)

| 样式ID | 中文名称 | 英文名称 | 描述 |
|--------|----------|----------|------|
| Normal | 普通文本 | Normal | 默认段落样式，Calibri 11磅，1.15倍行距 |
| Heading1 | 标题 1 | Heading 1 | 一级标题，16磅蓝色粗体，支持导航窗格 |
| Heading2 | 标题 2 | Heading 2 | 二级标题，13磅蓝色粗体，支持导航窗格 |
| Heading3 | 标题 3 | Heading 3 | 三级标题，12磅蓝色粗体，支持导航窗格 |
| Heading4 | 标题 4 | Heading 4 | 四级标题，11磅蓝色粗体 |
| Heading5 | 标题 5 | Heading 5 | 五级标题，11磅蓝色 |
| Heading6 | 标题 6 | Heading 6 | 六级标题，11磅蓝色 |
| Heading7 | 标题 7 | Heading 7 | 七级标题，11磅斜体 |
| Heading8 | 标题 8 | Heading 8 | 八级标题，10磅灰色 |
| Heading9 | 标题 9 | Heading 9 | 九级标题，10磅斜体灰色 |
| Title | 文档标题 | Title | 28磅居中标题样式 |
| Subtitle | 副标题 | Subtitle | 15磅居中副标题样式 |
| Quote | 引用 | Quote | 斜体灰色，左右缩进720TWIPs |
| ListParagraph | 列表段落 | List Paragraph | 带左缩进的列表样式 |
| CodeBlock | 代码块 | Code Block | 等宽字体，灰色背景效果 |

### 字符样式 (Character Styles)

| 样式ID | 中文名称 | 英文名称 | 描述 |
|--------|----------|----------|------|
| Emphasis | 强调 | Emphasis | 斜体文本 |
| Strong | 加粗 | Strong | 粗体文本 |
| CodeChar | 代码字符 | Code Character | 红色等宽字体 |

## 🔧 自定义样式创建

### 使用QuickStyleConfig快速创建

```go
// 创建自定义段落样式
config := style.QuickStyleConfig{
    ID:      "MyTitle",
    Name:    "我的标题样式",
    Type:    style.StyleTypeParagraph,
    BasedOn: "Normal",  // 基于Normal样式
    ParagraphConfig: &style.QuickParagraphConfig{
        Alignment:       "center",
        LineSpacing:     1.5,
        SpaceBefore:     15,
        SpaceAfter:      10,
        FirstLineIndent: 0,
        LeftIndent:      0,
        RightIndent:     0,
    },
    RunConfig: &style.QuickRunConfig{
        FontName:  "华文中宋",
        FontSize:  18,
        FontColor: "2F5496",  // 深蓝色
        Bold:      true,
        Italic:    false,
        Underline: false,
    },
}

// 创建样式
customStyle, err := quickAPI.CreateQuickStyle(config)
if err != nil {
    log.Printf("创建样式失败: %v", err)
} else {
    fmt.Printf("成功创建样式: %s\n", customStyle.Name.Val)
}
```

### 创建字符样式

```go
// 创建自定义字符样式
charConfig := style.QuickStyleConfig{
    ID:   "Highlight",
    Name: "高亮文本",
    Type: style.StyleTypeCharacter,
    RunConfig: &style.QuickRunConfig{
        FontColor: "FF0000",  // 红色
        Bold:      true,
        Highlight: "yellow",  // 黄色高亮
    },
}

highlightStyle, err := quickAPI.CreateQuickStyle(charConfig)
if err != nil {
    log.Printf("创建字符样式失败: %v", err)
}
```

### 高级自定义样式

```go
// 使用完整的Style结构创建复杂样式
complexStyle := &style.Style{
    Type:    string(style.StyleTypeParagraph),
    StyleID: "ComplexTitle",
    Name:    &style.StyleName{Val: "复杂标题样式"},
    BasedOn: &style.BasedOn{Val: "Heading1"},
    Next:    &style.Next{Val: "Normal"},
    ParagraphPr: &style.ParagraphProperties{
        Spacing: &style.Spacing{
            Before: "240",  // 12磅
            After:  "120",  // 6磅
            Line:   "276",  // 1.15倍行距
        },
        Justification: &style.Justification{Val: "center"},
        Indentation: &style.Indentation{
            FirstLine: "0",
            Left:      "0",
        },
    },
    RunPr: &style.RunProperties{
        FontFamily: &style.FontFamily{ASCII: "Times New Roman"},
        FontSize:   &style.FontSize{Val: "32"},  // 16磅
        Color:      &style.Color{Val: "1F4E79"},
        Bold:       &style.Bold{},
    },
}

styleManager.AddStyle(complexStyle)
```

## 🔍 样式查询和管理

### 按类型查询样式

```go
// 获取所有段落样式信息
paragraphStyles := quickAPI.GetParagraphStylesInfo()
fmt.Printf("段落样式数量: %d\n", len(paragraphStyles))

// 获取所有字符样式信息
characterStyles := quickAPI.GetCharacterStylesInfo()
fmt.Printf("字符样式数量: %d\n", len(characterStyles))

// 获取所有标题样式信息
headingStyles := quickAPI.GetHeadingStylesInfo()
fmt.Printf("标题样式数量: %d\n", len(headingStyles))

// 打印样式详情
for _, styleInfo := range headingStyles {
    fmt.Printf("- %s (%s): %s\n", 
        styleInfo.Name, styleInfo.ID, styleInfo.Description)
}
```

### 样式存在性检查

```go
// 检查样式是否存在
if styleManager.StyleExists("Heading1") {
    fmt.Println("Heading1 样式存在")
}

// 验证样式并获取详情
styleInfo, err := quickAPI.GetStyleInfo("CustomStyle")
if err != nil {
    fmt.Printf("样式不存在: %v\n", err)
} else {
    fmt.Printf("找到样式: %s\n", styleInfo.Name)
}
```

### 样式管理操作

```go
// 获取所有样式
allStyles := styleManager.GetAllStyles()
fmt.Printf("总样式数: %d\n", len(allStyles))

// 移除自定义样式
styleManager.RemoveStyle("MyCustomStyle")

// 清空所有样式（注意：这会删除预定义样式）
// styleManager.ClearStyles()

// 重新加载预定义样式
styleManager.LoadPredefinedStyles()
```

## 🔄 样式继承机制

### 理解样式继承

```go
// 获取带继承的完整样式
fullStyle := styleManager.GetStyleWithInheritance("Heading2")

// Heading2 基于 Normal 样式
// GetStyleWithInheritance 会自动合并：
// 1. Normal 样式的所有属性
// 2. Heading2 样式的覆盖属性
// 3. 返回完整的合并样式

if fullStyle.BasedOn != nil {
    fmt.Printf("Heading2 基于样式: %s\n", fullStyle.BasedOn.Val)
}

// 检查继承的属性
if fullStyle.RunPr != nil && fullStyle.RunPr.FontSize != nil {
    fmt.Printf("继承的字体大小: %s\n", fullStyle.RunPr.FontSize.Val)
}
```

### 创建继承样式

```go
// 创建基于Heading1的自定义样式
customHeading := style.QuickStyleConfig{
    ID:      "MyHeading",
    Name:    "我的标题",
    Type:    style.StyleTypeParagraph,
    BasedOn: "Heading1",  // 继承Heading1的所有属性
    // 只覆盖需要修改的属性
    RunConfig: &style.QuickRunConfig{
        FontColor: "8B0000",  // 改为深红色
        // 其他属性（字体大小、粗体等）从Heading1继承
    },
}

inheritedStyle, _ := quickAPI.CreateQuickStyle(customHeading)
```

## 🎯 样式属性配置详解

### ParagraphConfig 段落属性

```go
type QuickParagraphConfig struct {
    Alignment       string  // 对齐方式
    LineSpacing     float64 // 行间距倍数
    SpaceBefore     int     // 段前间距（磅）
    SpaceAfter      int     // 段后间距（磅）
    FirstLineIndent int     // 首行缩进（磅）
    LeftIndent      int     // 左缩进（磅）
    RightIndent     int     // 右缩进（磅）
}
```

**对齐方式选项:**
- `"left"` - 左对齐
- `"center"` - 居中对齐
- `"right"` - 右对齐
- `"justify"` - 两端对齐

**间距和缩进单位:**
- 所有数值单位为磅（Point）
- 1磅 = 1/72英寸 = 20TWIPs

### RunConfig 字符属性

```go
type QuickRunConfig struct {
    FontName  string // 字体名称
    FontSize  int    // 字体大小（磅）
    FontColor string // 字体颜色（十六进制）
    Bold      bool   // 粗体
    Italic    bool   // 斜体
    Underline bool   // 下划线
    Strike    bool   // 删除线
    Highlight string // 高亮颜色
}
```

**字体颜色格式:**
- 十六进制RGB格式，如 `"FF0000"` (红色)
- 不需要 `#` 前缀

**高亮颜色选项:**
- `"yellow"` - 黄色
- `"green"` - 绿色
- `"cyan"` - 青色
- `"magenta"` - 洋红色
- `"blue"` - 蓝色
- `"red"` - 红色
- `"darkBlue"` - 深蓝色
- `"darkCyan"` - 深青色
- `"darkGreen"` - 深绿色
- `"darkMagenta"` - 深洋红色
- `"darkRed"` - 深红色
- `"darkYellow"` - 深黄色
- `"darkGray"` - 深灰色
- `"lightGray"` - 浅灰色
- `"black"` - 黑色

## 📋 完整使用示例

### 创建带样式的完整文档

```go
package main

import (
    "fmt"
    "log"
    "github.com/mr-pmillz/wordZero/pkg/document"
    "github.com/mr-pmillz/wordZero/pkg/style"
)

func main() {
    // 创建文档和样式管理器
    doc := document.New()
    styleManager := doc.GetStyleManager()
    quickAPI := style.NewQuickStyleAPI(styleManager)

    // 创建自定义样式
    createCustomStyles(quickAPI)

    // 构建文档内容
    buildDocumentContent(doc)

    // 保存文档
    err := doc.Save("styled_document_complete.docx")
    if err != nil {
        log.Fatal(err)
    }

    fmt.Println("文档创建完成：styled_document_complete.docx")
}

func createCustomStyles(quickAPI *style.QuickStyleAPI) {
    // 创建自定义标题样式
    titleConfig := style.QuickStyleConfig{
        ID:      "CustomTitle",
        Name:    "自定义文档标题",
        Type:    style.StyleTypeParagraph,
        BasedOn: "Title",
        ParagraphConfig: &style.QuickParagraphConfig{
            Alignment:   "center",
            SpaceBefore: 24,
            SpaceAfter:  18,
        },
        RunConfig: &style.QuickRunConfig{
            FontName:  "华文中宋",
            FontSize:  20,
            FontColor: "1F4E79",
            Bold:      true,
        },
    }

    // 创建高亮文本样式
    highlightConfig := style.QuickStyleConfig{
        ID:   "ImportantText",
        Name: "重要文本",
        Type: style.StyleTypeCharacter,
        RunConfig: &style.QuickRunConfig{
            FontColor: "C00000",
            Bold:      true,
            Highlight: "yellow",
        },
    }

    quickAPI.CreateQuickStyle(titleConfig)
    quickAPI.CreateQuickStyle(highlightConfig)
}

func buildDocumentContent(doc *document.Document) {
    // 使用自定义标题样式
    title := doc.AddParagraph("WordZero 样式系统使用指南")
    title.SetStyle("CustomTitle")

    // 使用标题样式（支持导航窗格）
    doc.AddHeadingParagraph("1. 样式系统概述", 1)
    doc.AddParagraph("WordZero 提供了完整的样式管理系统，支持预定义样式和自定义样式。")

    doc.AddHeadingParagraph("1.1 预定义样式", 2)
    para := doc.AddParagraph("系统预置了18种常用样式，包括：")
    para.AddFormattedText("标题样式", &document.TextFormat{Bold: true})
    para.AddFormattedText("、", nil)
    para.AddFormattedText("段落样式", &document.TextFormat{Bold: true})
    para.AddFormattedText("和", nil)
    para.AddFormattedText("字符样式", &document.TextFormat{Bold: true})
    para.AddFormattedText("。", nil)

    doc.AddHeadingParagraph("1.2 自定义样式", 2)
    doc.AddParagraph("用户可以基于现有样式创建自定义样式，实现个性化的文档格式。")

    doc.AddHeadingParagraph("2. 实际应用", 1)
    
    // 使用引用样式
    quote := doc.AddParagraph("样式是文档格式化的核心，它决定了文档的外观和专业程度。")
    quote.SetStyle("Quote")

    // 使用代码块样式
    code := doc.AddParagraph("doc.AddHeadingParagraph(\"标题\", 1)")
    code.SetStyle("CodeBlock")

    doc.AddParagraph("更多详细信息请参考API文档。")
}
```

## 🧪 测试

详细的测试示例请参考：

```bash
# 运行样式系统测试
go test ./pkg/style/

# 运行带覆盖率的测试
go test -cover ./pkg/style/

# 运行样式演示程序
go run ./examples/style_demo/
```

## 📚 相关文档

- [项目主README](../../README.md) - 完整项目介绍
- [文档操作API](../document/) - 核心文档操作功能
- [使用示例](../../examples/) - 完整的使用示例

## 🤝 贡献

欢迎提交样式相关的改进建议和代码！请确保：

1. 新增样式遵循Word标准规范
2. 提供完整的测试用例
3. 更新相关文档

## 📄 许可证

本包遵循项目的 MIT 许可证。 
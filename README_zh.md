<div align="center">
  <img src="docs/logo-banner.svg" alt="WordZero Logo" width="400"/>
  
  <h1>WordZero - Golang Word操作库</h1>
</div>

<div align="center">
  
[![Go Version](https://img.shields.io/badge/Go-1.19+-00ADD8?style=flat&logo=go)](https://golang.org)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Tests](https://img.shields.io/badge/Tests-Passing-green.svg)](#测试)
[![Benchmark](https://img.shields.io/badge/Benchmark-Go%202.62ms%20%7C%20JS%209.63ms%20%7C%20Python%2055.98ms-success.svg)](https://github.com/mr-pmillz/wordZero/wiki/13-%E6%80%A7%E8%83%BD%E5%9F%BA%E5%87%86%E6%B5%8B%E8%AF%95)
[![Performance](https://img.shields.io/badge/Performance-Golang%20优胜-brightgreen.svg)](https://github.com/mr-pmillz/wordZero/wiki/13-%E6%80%A7%E8%83%BD%E5%9F%BA%E5%87%86%E6%B5%8B%E8%AF%95)
[![Ask DeepWiki](https://deepwiki.com/badge.svg)](https://deepwiki.com/mr-pmillz/wordZero)

</div>

[English](README.md) | **中文**

## 项目介绍

WordZero 是一个使用 Golang 实现的 Word 文档操作库，提供基础的文档创建、修改等操作功能。该库遵循最新的 Office Open XML (OOXML) 规范，专注于现代 Word 文档格式（.docx）的支持。

### 核心特性

- 🚀 **完整的文档操作**: 创建、读取、修改 Word 文档
- 🎨 **丰富的样式系统**: 18种预定义样式，支持自定义样式和样式继承
- 📝 **文本格式化**: 字体、大小、颜色、粗体、斜体等完整支持
- 📐 **段落格式**: 对齐、间距、缩进、分页控制、行控制、孤行控制、大纲级别等全面段落属性设置 ✨ **已完善**
- 🏷️ **标题导航**: 完整支持Heading1-9样式，可被Word导航窗格识别
- 📊 **表格功能**: 完整的表格创建、编辑、样式设置和迭代器支持
- 📄 **页面设置**: 页面尺寸、边距、页眉页脚等专业排版功能
- 🔧 **高级功能**: 目录生成、脚注尾注、列表编号、模板引擎等
- 🎯 **模板继承**: 支持基础模板和块重写机制，实现模板复用和扩展
- 📝 **页眉页脚模板**: 支持在页眉页脚中使用模板变量进行动态内容替换
- ⚡ **卓越性能**: 零依赖的纯Go实现，平均2.62ms处理速度，比JavaScript快3.7倍，比Python快21倍
- 🔧 **易于使用**: 简洁的API设计，链式调用支持

## 相关推荐项目

### Excel文档操作推荐 - Excelize

如果您需要处理Excel文档，我们强烈推荐使用 [**Excelize**](https://github.com/qax-os/excelize) —— 最受欢迎的Go语言Excel操作库：

- ⭐ **GitHub 19.2k+ 星标** - Go生态系统中最受欢迎的Excel处理库
- 📊 **完整Excel支持** - 支持XLAM/XLSM/XLSX/XLTM/XLTX等所有现代Excel格式
- 🎯 **功能丰富** - 图表、数据透视表、图片、流式API等完整功能
- 🚀 **高性能** - 专为大数据集处理优化的流式读写API
- 🔧 **易于集成** - 与WordZero完美互补，构建完整的Office文档处理解决方案

**完美搭配**: WordZero负责Word文档处理，Excelize负责Excel文档处理，共同为您的Go项目提供完整的Office文档操作能力。

```go
// WordZero + Excelize 组合示例
import (
    "github.com/mr-pmillz/wordZero/pkg/document"
    "github.com/xuri/excelize/v2"
)

// 创建Word报告
doc := document.New()
doc.AddParagraph("数据分析报告").SetStyle(style.StyleHeading1)

// 创建Excel数据表
xlsx := excelize.NewFile()
xlsx.SetCellValue("Sheet1", "A1", "数据项")
xlsx.SetCellValue("Sheet1", "B1", "数值")
```

## 安装

```bash
go get github.com/mr-pmillz/wordZero
```

### 版本说明

推荐使用带版本号的安装方式：

```bash
# 安装最新版本
go get github.com/mr-pmillz/wordZero@latest

# 安装指定版本
go get github.com/mr-pmillz/wordZero@v1.6.0
```

## 快速开始

```go
package main

import (
    "log"
    "github.com/mr-pmillz/wordZero/pkg/document"
    "github.com/mr-pmillz/wordZero/pkg/style"
)

func main() {
    // 创建新文档
    doc := document.New()
    
    // 添加标题
    titlePara := doc.AddParagraph("WordZero 使用示例")
    titlePara.SetStyle(style.StyleHeading1)
    
    // 添加正文段落
    para := doc.AddParagraph("这是一个使用 WordZero 创建的文档示例。")
    para.SetFontFamily("宋体")
    para.SetFontSize(12)
    para.SetColor("333333")
    
    // 创建表格
    tableConfig := &document.TableConfig{
        Rows:    3,
        Columns: 3,
    }
    table := doc.AddTable(tableConfig)
    table.SetCellText(0, 0, "表头1")
    table.SetCellText(0, 1, "表头2")
    table.SetCellText(0, 2, "表头3")
    
    // 保存文档
    if err := doc.Save("example.docx"); err != nil {
        log.Fatal(err)
    }
}
```

### 模板继承功能示例

```go
// 创建基础模板
engine := document.NewTemplateEngine()
baseTemplate := `{{companyName}} 工作报告

{{#block "summary"}}
默认摘要内容
{{/block}}

{{#block "content"}}
默认主要内容
{{/block}}`

engine.LoadTemplate("base_report", baseTemplate)

// 创建扩展模板，重写特定块
salesTemplate := `{{extends "base_report"}}

{{#block "summary"}}
销售业绩摘要：本月达成 {{achievement}}%
{{/block}}

{{#block "content"}}
销售详情：
- 总销售额：{{totalSales}}
- 新增客户：{{newCustomers}}
{{/block}}`

engine.LoadTemplate("sales_report", salesTemplate)

// 渲染模板
data := document.NewTemplateData()
data.SetVariable("companyName", "WordZero科技")
data.SetVariable("achievement", "125")
data.SetVariable("totalSales", "1,850,000")
data.SetVariable("newCustomers", "45")

doc, _ := engine.RenderTemplateToDocument("sales_report", data)
doc.Save("sales_report.docx")
```

### 图片占位符模板功能示例 ✨ **新增**

```go
package main

import (
    "log"
    "github.com/mr-pmillz/wordZero/pkg/document"
)

func main() {
    // 创建包含图片占位符的模板
    engine := document.NewTemplateEngine()
    template := `公司：{{companyName}}

{{#image companyLogo}}

项目报告：{{projectName}}

状态：{{#if isCompleted}}已完成{{else}}进行中{{/if}}

{{#image statusChart}}

团队成员：
{{#each teamMembers}}
- {{name}}：{{role}}
{{/each}}`

    engine.LoadTemplate("project_report", template)

    // 准备模板数据
    data := document.NewTemplateData()
    data.SetVariable("companyName", "WordZero科技")
    data.SetVariable("projectName", "文档处理系统")
    data.SetCondition("isCompleted", true)
    
    // 设置团队成员列表
    data.SetList("teamMembers", []interface{}{
        map[string]interface{}{"name": "张三", "role": "首席开发"},
        map[string]interface{}{"name": "李四", "role": "前端开发"},
    })
    
    // 配置并设置图片
    logoConfig := &document.ImageConfig{
        Width:     100,
        Height:    50,
        Alignment: document.AlignCenter,
    }
    data.SetImage("companyLogo", "assets/logo.png", logoConfig)
    
    chartConfig := &document.ImageConfig{
        Width:       200,
        Height:      150,
        Alignment:   document.AlignCenter,
        AltText:     "项目状态图表",
        Title:       "当前项目状态",
    }
    data.SetImage("statusChart", "assets/chart.png", chartConfig)
    
    // 渲染模板到文档
    doc, err := engine.RenderTemplateToDocument("project_report", data)
    if err != nil {
        log.Fatal(err)
    }
    
    // 保存文档
    err = doc.Save("project_report.docx")
    if err != nil {
        log.Fatal(err)
    }
}
```

### Markdown转Word功能示例 ✨ **新增**

```go
package main

import (
    "log"
    "github.com/mr-pmillz/wordZero/pkg/markdown"
)

func main() {
    // 创建Markdown转换器
    converter := markdown.NewConverter(markdown.DefaultOptions())
    
    // Markdown内容
    markdownText := `# WordZero Markdown转换示例

欢迎使用WordZero的**Markdown到Word**转换功能！

## 支持的语法

### 文本格式
- **粗体文本**
- *斜体文本*
- ` + "`行内代码`" + `

### 列表
1. 有序列表项1
2. 有序列表项2

- 无序列表项A
- 无序列表项B

### 引用和代码

> 这是引用块内容
> 支持多行引用

` + "```" + `go
// 代码块示例
func main() {
    fmt.Println("Hello, WordZero!")
}
` + "```" + `

---

转换完成！`

    // 转换为Word文档
    doc, err := converter.ConvertString(markdownText, nil)
    if err != nil {
        log.Fatal(err)
    }
    
    // 保存Word文档
    err = doc.Save("markdown_example.docx")
    if err != nil {
        log.Fatal(err)
    }
    
    // 文件转换
    err = converter.ConvertFile("input.md", "output.docx", nil)
    if err != nil {
        log.Fatal(err)
    }
}
```

### 文档分页和段落删除功能示例 ✨ **新增**

```go
package main

import (
    "log"
    "github.com/mr-pmillz/wordZero/pkg/document"
)

func main() {
    doc := document.New()
    
    // 添加第一页内容
    doc.AddHeadingParagraph("第一章：引言", 1)
    doc.AddParagraph("这是第一章的内容。")
    
    // 添加分页符，开始新的一页
    doc.AddPageBreak()
    
    // 添加第二页内容
    doc.AddHeadingParagraph("第二章：正文", 1)
    tempPara := doc.AddParagraph("这是一个临时段落。")
    doc.AddParagraph("这是第二章的内容。")
    
    // 删除临时段落
    doc.RemoveParagraph(tempPara)
    
    // 也可以按索引删除段落
    // doc.RemoveParagraphAt(1)  // 删除第二个段落
    
    // 保存文档
    if err := doc.Save("example.docx"); err != nil {
        log.Fatal(err)
    }
}
```

## 文档和示例

### 📚 完整文档

**多语言文档支持**:
- **中文**: [📖 中文文档](https://github.com/mr-pmillz/wordZero/wiki)
- **English**: [📖 Wiki Documentation](https://github.com/mr-pmillz/wordZero/wiki/en-Home)

**核心文档**:
- [**🚀 快速开始**](https://github.com/mr-pmillz/wordZero/wiki/01-快速开始) - 新手入门指南
- [**⚡ 功能特性详览**](https://github.com/mr-pmillz/wordZero/wiki/14-功能特性详览) - 所有功能的详细说明
- [**📊 性能基准测试**](https://github.com/mr-pmillz/wordZero/wiki/13-性能基准测试) - 跨语言性能对比分析
- [**🏗️ 项目结构详解**](https://github.com/mr-pmillz/wordZero/wiki/15-项目结构详解) - 项目架构和代码组织

### 💡 使用示例
查看 `examples/` 目录下的示例代码：

- `examples/basic/` - 基础功能演示
- `examples/style_demo/` - 样式系统演示  
- `examples/table/` - 表格功能演示
- `examples/formatting/` - 格式化演示
- `examples/page_settings/` - 页面设置演示
- `examples/advanced_features/` - 高级功能综合演示
- `examples/template_demo/` - 模板功能演示
- `examples/template_inheritance_demo/` - 模板继承功能演示 ✨ **新增**
- `examples/template_image_demo/` - 图片占位符模板演示 ✨ **新增**
- `examples/markdown_conversion/` - Markdown转Word功能演示 ✨ **新增**
- `examples/pagination_deletion_demo/` - 分页和段落删除功能演示 ✨ **新增**
- `examples/paragraph_format_demo/` - 段落格式自定义功能演示 ✨ **新增**

运行示例：
```bash
# 运行基础功能演示
go run ./examples/basic/

# 运行样式演示
go run ./examples/style_demo/

# 运行表格演示
go run ./examples/table/

# 运行模板继承演示
go run ./examples/template_inheritance_demo/

# 运行图片占位符模板演示
go run ./examples/template_image_demo/

# 运行Markdown转Word演示
go run ./examples/markdown_conversion/

# 运行段落格式自定义演示
go run ./examples/paragraph_format_demo/
```

## 主要功能

### ✅ 已实现功能
- **文档操作**: 创建、读取、保存、解析DOCX文档
- **文本格式化**: 字体、大小、颜色、粗体、斜体等
- **样式系统**: 18种预定义样式 + 自定义样式支持
- **段落格式**: 对齐、间距、缩进、分页控制、行控制、孤行控制、大纲级别等完整支持 ✨ **已完善**
- **段落管理**: 段落删除、按索引删除、元素删除 ✨ **新增**
- **文档分页**: 分页符插入，支持多页文档结构 ✨ **新增**
- **表格功能**: 完整的表格操作、样式设置、单元格迭代器
- **页面设置**: 页面尺寸、边距、页眉页脚等
- **高级功能**: 目录生成、脚注尾注、列表编号、模板引擎（含模板继承）
- **图片功能**: 图片插入、大小调整、位置设置
- **Markdown转Word**: 基于goldmark的高质量Markdown到Word转换

### 🚧 规划中功能
- 表格排序和高级操作
- 书签和交叉引用
- 文档批注和修订
- 图形绘制功能
- 多语言和国际化支持

👉 **查看完整功能列表**: [功能特性详览](https://github.com/mr-pmillz/wordZero/wiki/14-功能特性详览)

## 性能表现

WordZero 在性能方面表现卓越，通过完整的基准测试验证：

| 语言 | 平均执行时间 | 相对性能 |
|------|-------------|----------|
| **Golang** | **2.62ms** | **1.00×** |
| JavaScript | 9.63ms | 3.67× |
| Python | 55.98ms | 21.37× |

👉 **查看详细性能分析**: [性能基准测试](https://github.com/mr-pmillz/wordZero/wiki/13-性能基准测试)

## 项目结构

```
wordZero/
├── pkg/                    # 核心库代码
│   ├── document/          # 文档操作功能
│   └── style/             # 样式管理系统
├── examples/              # 使用示例
├── test/                  # 集成测试
├── benchmark/             # 性能基准测试
├── docs/                  # 文档和资源文件
│   ├── logo.svg           # 主Logo带性能指标
│   ├── logo-banner.svg    # 横幅版本用于README标题
│   └── logo-simple.svg    # 简化图标版本
└── wordZero.wiki/         # 完整文档
```

👉 **查看详细结构说明**: [项目结构详解](https://github.com/mr-pmillz/wordZero/wiki/15-项目结构详解)

### Logo设计

项目包含多种Logo变体，适用于不同使用场景：

<div align="center">

| Logo类型 | 使用场景 | 预览 |
|----------|----------|------|
| **横幅版** | README标题、文档头部 | <img src="docs/logo-banner.svg" alt="横幅Logo" width="200"/> |
| **主版本** | 通用品牌展示 | <img src="docs/logo.svg" alt="主Logo" width="120"/> |
| **简化版** | 图标、网站标识 | <img src="docs/logo-simple.svg" alt="简化Logo" width="32"/> |

</div>

## 贡献指南

欢迎提交 Issue 和 Pull Request！在提交代码前请确保：

1. 代码符合 Go 代码规范
2. 添加必要的测试用例
3. 更新相关文档
4. 确保所有测试通过

## 许可证

本项目采用 MIT 许可证。详见 [LICENSE](LICENSE) 文件。

---

**更多资源**:
- 📖 [完整文档](https://github.com/mr-pmillz/wordZero/wiki)
- 🔧 [API参考](https://github.com/mr-pmillz/wordZero/wiki/10-API参考)
- 💡 [最佳实践](https://github.com/mr-pmillz/wordZero/wiki/09-最佳实践)
- 📝 [更新日志](CHANGELOG.md) 
package main

import (
	"fmt"
	"log"

	"github.com/mr-pmillz/wordZero/pkg/markdown"
)

func main() {
	// 全面的Markdown内容示例，涵盖所有支持的功能
	markdownContent := `# WordZero Markdown 功能演示文档

这是一个全面展示 WordZero Markdown 转换功能的演示文档。

## 1. 标题演示

### 1.1 标题级别测试

# 一级标题 (H1)
## 二级标题 (H2)  
### 三级标题 (H3)
#### 四级标题 (H4)
##### 五级标题 (H5)
###### 六级标题 (H6)

## 2. 文本格式化

### 2.1 基础格式

这里有**粗体文本**和*斜体文本*。

你也可以组合使用***粗斜体文本***。

还支持行内代码 ` + "`var x = \"hello world\"`" + `。

### 2.2 链接演示

这是一个[外部链接](https://github.com)的示例。

这是一个[WordZero项目链接](https://github.com/mr-pmillz/wordZero)。

## 3. 列表功能

### 3.1 无序列表

- 第一个项目
- 第二个项目
  - 嵌套项目 1
  - 嵌套项目 2
    - 更深层嵌套
- 第三个项目

### 3.2 有序列表

1. 步骤一：准备工作
2. 步骤二：执行操作
   1. 子步骤 2.1
   2. 子步骤 2.2
3. 步骤三：验证结果

### 3.3 任务列表（GitHub Flavored Markdown）

**项目待办事项清单：**

- [x] ✅ 完成项目需求分析
- [x] ✅ 设计系统架构  
- [ ] ⏳ 实现核心功能
  - [x] ✅ 用户管理模块
  - [x] ✅ 文档处理模块
  - [ ] ⏳ 权限控制模块
  - [ ] ⏳ 数据存储模块
- [x] ✅ 编写测试用例
- [ ] ⏳ 部署到生产环境
- [ ] ⏳ 编写用户文档

## 4. 表格功能

### 4.1 基础表格

| 功能特性 | 状态 | 优先级 | 备注 |
|----------|------|--------|------|
| 标题转换 | ✅ 完成 | 高 | 支持1-6级标题 |
| 文本格式 | ✅ 完成 | 高 | 粗体、斜体、代码 |
| 表格转换 | ✅ 完成 | 中 | 支持对齐方式 |
| 任务列表 | ✅ 完成 | 中 | GFM扩展功能 |
| 图片处理 | 🔄 开发中 | 低 | 基础支持 |

### 4.2 对齐方式表格

| 左对齐 | 居中对齐 | 右对齐 | 默认对齐 |
|:-------|:--------:|-------:|----------|
| 内容1 | 居中内容 | 右对齐内容 | 默认内容 |
| 较长的内容文本 | 短内容 | 数字123 | 普通文本 |
| Left | Center | Right | Normal |

### 4.3 复杂表格

| 序号 | 模块名称 | 实现状态 | 功能描述 | 测试覆盖率 |
|------|----------|----------|----------|------------|
| 1 | **核心解析器** | ✅ 已完成 | 基于goldmark的MD解析 | 95% |
| 2 | **文本渲染** | ✅ 已完成 | 支持格式化文本输出 | 90% |
| 3 | **表格处理** | ✅ 已完成 | *支持GFM表格规范* | 85% |
| 4 | **列表渲染** | ✅ 已完成 | 包含任务列表支持 | 88% |
| 5 | **图片处理** | 🔄 进行中 | 图片插入和路径处理 | 60% |

## 5. 代码块演示

### 5.1 内联代码

在Go语言中，你可以使用 ` + "`fmt.Println(\"Hello, World!\")`" + ` 来输出文本。

创建变量：` + "`var name string = \"WordZero\"`" + `

### 5.2 代码块

` + "```" + `go
package main

import (
    "fmt"
    "github.com/mr-pmillz/wordZero/pkg/markdown"
)

func main() {
    // 创建转换器
    converter := markdown.NewConverter(markdown.HighQualityOptions())
    
    // 转换文档
    doc, err := converter.ConvertString(content, nil)
    if err != nil {
        log.Fatal(err)
    }
    
    // 保存文档
    err = doc.Save("output.docx")
    if err != nil {
        log.Fatal(err)
    }
    
    fmt.Println("转换完成！")
}
` + "```" + `

` + "```" + `javascript
// JavaScript示例
function convertMarkdown(content) {
    const options = {
        enableGFM: true,
        enableTables: true,
        enableTaskList: true
    };
    
    return markdownConverter.convert(content, options);
}
` + "```" + `

` + "```" + `json
{
    "project": "WordZero",
    "version": "1.0.0",
    "features": [
        "markdown_conversion",
        "table_support", 
        "task_lists",
        "gfm_support"
    ],
    "status": "active"
}
` + "```" + `

## 6. 引用块

### 6.1 简单引用

> 这是一个简单的引用块示例。
> 它可以包含多行内容。

### 6.2 嵌套引用

> 外层引用内容
> 
> > 这是嵌套的引用内容
> > 可以包含更深层的内容
> 
> 回到外层引用

### 6.3 引用中的格式

> **重要提示：** 在使用WordZero时，请确保：
> 
> - 使用 *最新版本* 的库
> - 遵循 **官方文档** 的指导
> - 测试你的 ` + "`代码实现`" + `

## 7. 分割线

下面是一条分割线：

---

上面和下面都有分割线。

***

### 8.2 网络图片

![GitHub Logo](https://github.githubassets.com/images/modules/logos_page/GitHub-Mark.png)

## 9. 混合内容示例

这个段落包含**多种格式**，包括*斜体*、` + "`代码`" + `和[链接](https://example.com)。

### 9.1 复杂列表混合

1. **第一步：环境准备**
   - 安装Go语言环境（版本 >= 1.19）
   - 克隆项目：` + "`git clone https://github.com/mr-pmillz/wordZero.git`" + `
   - 下载依赖：` + "`go mod download`" + `

2. **第二步：配置设置**
   ` + "```" + `bash
   # 设置环境变量
   export WORDZERO_CONFIG=./config.json
   
   # 运行测试
   go test ./...
   ` + "```" + `

3. **第三步：运行示例**
   - [ ] 基础文档创建示例
   - [x] Markdown转换示例  
   - [ ] 高级功能演示
   - [x] 表格和列表演示

## 10. 功能特性总结

| 分类 | 功能点 | Markdown语法 | Word输出 | 状态 |
|------|--------|-------------|----------|------|
| **标题** | 1-6级标题 | ` + "`# ## ### #### ##### ######`" + ` | Heading样式 | ✅ |
| **格式** | 粗体 | ` + "`**text**`" + ` | Bold格式 | ✅ |
| **格式** | 斜体 | ` + "`*text*`" + ` | Italic格式 | ✅ |
| **格式** | 行内代码 | ` + "```text```" + ` | 等宽字体 | ✅ |
| **链接** | 超链接 | ` + "`[text](url)`" + ` | 蓝色文本 | ✅ |
| **列表** | 无序列表 | ` + "`- item`" + ` | 项目符号 | ✅ |
| **列表** | 有序列表 | ` + "`1. item`" + ` | 编号列表 | ✅ |
| **列表** | 任务列表 | ` + "`- [x] done`" + ` | 复选框 | ✅ |
| **表格** | GFM表格 | ` + "`| cell |`" + ` | Word表格 | ✅ |
| **代码** | 代码块 | ` + "```code```" + ` | 等宽字体块 | ✅ |
| **引用** | 引用块 | ` + "`> quote`" + ` | 斜体格式 | ✅ |
| **分割** | 分割线 | ` + "`---`" + ` | 横线 | ✅ |
| **图片** | 图片引用 | ` + "`![alt](src)`" + ` | 文本占位 | 🔄 |

---

**文档生成时间：** 2025年
**功能完整性：** 核心功能100%实现

> 📝 **说明：** 这个文档展示了WordZero Markdown转换器的所有主要功能。
> 所有标记为✅的功能都已完全实现并可正常使用。`

	fmt.Println("🚀 开始创建全面的Markdown功能演示...")

	// 创建高质量转换器配置
	opts := markdown.HighQualityOptions()
	opts.EnableTables = true
	opts.EnableTaskList = true
	opts.EnableGFM = true
	opts.EnableFootnotes = true
	opts.GenerateTOC = true
	opts.TOCMaxLevel = 3

	converter := markdown.NewConverter(opts)

	// 转换为Word文档
	doc, err := converter.ConvertString(markdownContent, opts)
	if err != nil {
		log.Fatalf("❌ 转换失败: %v", err)
	}

	// 保存文档
	outputPath := "examples/output/comprehensive_markdown_demo.docx"
	err = doc.Save(outputPath)
	if err != nil {
		log.Fatalf("❌ 保存文档失败: %v", err)
	}

	fmt.Printf("✅ 全面Markdown功能演示已保存到: %s\n", outputPath)
	fmt.Println("\n📋 演示包含以下功能特性:")
	fmt.Println("   🔸 标题转换 (H1-H6)")
	fmt.Println("   🔸 文本格式化 (粗体、斜体、行内代码)")
	fmt.Println("   🔸 链接处理")
	fmt.Println("   🔸 列表支持 (有序、无序、嵌套)")
	fmt.Println("   🔸 任务列表 (GitHub Flavored Markdown)")
	fmt.Println("   🔸 表格转换 (支持对齐方式)")
	fmt.Println("   🔸 代码块 (多语言语法)")
	fmt.Println("   🔸 引用块 (支持嵌套)")
	fmt.Println("   🔸 分割线")
	fmt.Println("   🔸 图片引用 (基础支持)")
	fmt.Println("   🔸 混合内容处理")
	fmt.Println("\n🎯 配置选项:")
	fmt.Printf("   • GitHub Flavored Markdown: %v\n", opts.EnableGFM)
	fmt.Printf("   • 表格支持: %v\n", opts.EnableTables)
	fmt.Printf("   • 任务列表: %v\n", opts.EnableTaskList)
	fmt.Printf("   • 脚注支持: %v\n", opts.EnableFootnotes)
	fmt.Printf("   • 生成目录: %v\n", opts.GenerateTOC)
	fmt.Printf("   • 目录最大级别: %d\n", opts.TOCMaxLevel)
	fmt.Printf("   • 默认字体: %s\n", opts.DefaultFontFamily)
	fmt.Printf("   • 默认字号: %.1f\n", opts.DefaultFontSize)
}

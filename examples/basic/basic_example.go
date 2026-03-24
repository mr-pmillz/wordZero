// Package main 展示WordZero基础功能使用示例
package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/mr-pmillz/wordZero/pkg/document"
	"github.com/mr-pmillz/wordZero/pkg/style"
)

func main() {
	fmt.Println("WordZero 基础功能演示")
	fmt.Println("====================")

	// 创建新文档
	doc := document.New()

	// 获取样式管理器
	styleManager := doc.GetStyleManager()

	// 1. 创建标题
	fmt.Println("📋 创建文档标题...")
	titlePara := doc.AddParagraph("WordZero 使用指南")
	titlePara.SetStyle(style.StyleTitle)

	// 2. 创建副标题
	fmt.Println("📋 创建副标题...")
	subtitlePara := doc.AddParagraph("一个简单、强大的Go语言Word文档操作库")
	subtitlePara.SetStyle(style.StyleSubtitle)

	// 3. 创建各级标题
	fmt.Println("📋 创建章节标题...")
	chapter1 := doc.AddParagraph("第一章 快速开始")
	chapter1.SetStyle(style.StyleHeading1)

	section1 := doc.AddParagraph("1.1 安装")
	section1.SetStyle(style.StyleHeading2)

	subsection1 := doc.AddParagraph("1.1.1 Go模块安装")
	subsection1.SetStyle(style.StyleHeading3)

	// 4. 添加普通文本段落
	fmt.Println("📋 添加正文内容...")
	normalText := "WordZero是一个专门为Go语言设计的Word文档操作库。它提供了简洁的API，让您能够轻松创建、编辑和保存Word文档。"
	normalPara := doc.AddParagraph(normalText)
	normalPara.SetStyle(style.StyleNormal)

	// 5. 添加代码块
	fmt.Println("📋 添加代码示例...")
	codeTitle := doc.AddParagraph("代码示例")
	codeTitle.SetStyle(style.StyleHeading3)

	codeExample := `go get github.com/mr-pmillz/wordZero

// 使用示例
import "github.com/mr-pmillz/wordZero/pkg/document"

doc := document.New()
doc.AddParagraph("Hello, WordZero!")
doc.Save("example.docx")`

	codePara := doc.AddParagraph(codeExample)
	codePara.SetStyle(style.StyleCodeBlock)

	// 6. 添加引用
	fmt.Println("📋 添加引用...")
	quoteText := "简单的API设计是WordZero的核心理念。我们相信强大的功能不应该以复杂的使用方式为代价。"
	quotePara := doc.AddParagraph(quoteText)
	quotePara.SetStyle(style.StyleQuote)

	// 7. 添加格式化文本
	fmt.Println("📋 添加格式化文本...")
	mixedPara := doc.AddParagraph("")
	mixedPara.AddFormattedText("WordZero支持多种文本格式：", nil)
	mixedPara.AddFormattedText("粗体", &document.TextFormat{Bold: true})
	mixedPara.AddFormattedText("、", nil)
	mixedPara.AddFormattedText("斜体", &document.TextFormat{Italic: true})
	mixedPara.AddFormattedText("、", nil)
	mixedPara.AddFormattedText("彩色文本", &document.TextFormat{FontColor: "FF0000"})
	mixedPara.AddFormattedText("以及", nil)
	mixedPara.AddFormattedText("不同字体", &document.TextFormat{FontFamily: "Times New Roman", FontSize: 14})
	mixedPara.AddFormattedText("。", nil)

	// 8. 创建列表
	fmt.Println("📋 创建列表...")
	listTitle := doc.AddParagraph("WordZero主要特性：")
	listTitle.SetStyle(style.StyleNormal)

	features := []string{
		"• 简洁易用的API设计",
		"• 完整的样式系统支持",
		"• 符合OOXML规范",
		"• 无外部依赖",
		"• 跨平台兼容",
	}

	for _, feature := range features {
		featurePara := doc.AddParagraph(feature)
		featurePara.SetStyle(style.StyleListParagraph)
	}

	// 9. 演示样式信息
	fmt.Println("📋 显示样式信息...")
	quickAPI := style.NewQuickStyleAPI(styleManager)
	allStyles := quickAPI.GetAllStylesInfo()

	stylesInfo := doc.AddParagraph(fmt.Sprintf("本文档使用了%d种预定义样式。", len(allStyles)))
	stylesInfo.SetStyle(style.StyleNormal)

	// 确保输出目录存在
	outputFile := "examples/output/basic_example.docx"
	outputDir := filepath.Dir(outputFile)

	fmt.Printf("📁 检查输出目录: %s\n", outputDir)
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		log.Printf("创建输出目录失败: %v", err)
		// 尝试当前目录
		outputFile = "basic_example.docx"
		fmt.Printf("📁 改为保存到当前目录: %s\n", outputFile)
	}

	fmt.Printf("📁 保存文档到: %s\n", outputFile)

	err := doc.Save(outputFile)
	if err != nil {
		log.Printf("保存文档失败: %v", err)
		fmt.Printf("❌ 文档保存失败，但演示程序已成功运行！\n")
		fmt.Printf("🔍 错误信息: %v\n", err)
		return
	}

	fmt.Println("✅ 基础示例文档创建完成！")
	fmt.Printf("🎉 您可以在 %s 查看生成的文档\n", outputFile)
}

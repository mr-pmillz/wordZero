// Package main 演示分页和段落删除功能
package main

import (
	"fmt"
	"log"
	"os"

	"github.com/mr-pmillz/wordZero/pkg/document"
)

func main() {
	fmt.Println("=== 文档分页和段落删除功能演示 ===\n")

	// 确保输出目录存在
	err := os.MkdirAll("examples/output", 0755)
	if err != nil {
		log.Fatalf("创建输出目录失败: %v", err)
	}

	// 演示1: 分页功能
	demonstratePageBreaks()

	// 演示2: 段落删除功能
	demonstrateParagraphDeletion()

	// 演示3: 组合使用分页和删除
	demonstrateCombinedUsage()

	fmt.Println("\n✅ 演示完成！")
}

// demonstratePageBreaks 演示分页符功能
func demonstratePageBreaks() {
	fmt.Println("📄 演示1: 分页符功能")

	doc := document.New()

	// 第一页内容
	doc.AddHeadingParagraph("第一页：项目概述", 1)
	doc.AddParagraph("这是项目的概述内容。")
	doc.AddParagraph("本文档演示了如何使用分页符来组织文档结构。")

	// 添加分页符，开始新的一页
	doc.AddPageBreak()

	// 第二页内容
	doc.AddHeadingParagraph("第二页：技术架构", 1)
	doc.AddParagraph("这是技术架构的详细说明。")
	doc.AddParagraph("通过分页符，我们可以将不同的章节分布在不同的页面上。")

	// 再添加一个分页符
	doc.AddPageBreak()

	// 第三页内容
	doc.AddHeadingParagraph("第三页：实施计划", 1)
	doc.AddParagraph("这是项目的实施计划和时间表。")

	// 保存文档
	filename := "examples/output/pagination_demo.docx"
	err := doc.Save(filename)
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Printf("✓ 分页演示文档已保存: %s\n", filename)
	fmt.Println("  - 文档包含3页内容")
	fmt.Println("  - 使用分页符分隔不同章节\n")
}

// demonstrateParagraphDeletion 演示段落删除功能
func demonstrateParagraphDeletion() {
	fmt.Println("🗑️  演示2: 段落删除功能")

	doc := document.New()

	// 添加多个段落
	doc.AddHeadingParagraph("文档编辑演示", 1)
	doc.AddParagraph("这是第一段，将被保留。")
	para2 := doc.AddParagraph("这是第二段，将被删除。")
	doc.AddParagraph("这是第三段，将被保留。")
	para4 := doc.AddParagraph("这是第四段，也将被删除。")
	doc.AddParagraph("这是第五段，将被保留。")

	fmt.Println("\n  原始文档包含以下段落:")
	fmt.Println("  1. 标题段落")
	fmt.Println("  2. 第一段（保留）")
	fmt.Println("  3. 第二段（删除）")
	fmt.Println("  4. 第三段（保留）")
	fmt.Println("  5. 第四段（删除）")
	fmt.Println("  6. 第五段（保留）")

	// 方法1: 使用 RemoveParagraph 直接删除段落对象
	fmt.Println("\n  执行删除操作:")
	if doc.RemoveParagraph(para2) {
		fmt.Println("  ✓ 删除第二段（使用 RemoveParagraph）")
	}

	// 方法2: 使用 RemoveParagraph 删除第四段
	if doc.RemoveParagraph(para4) {
		fmt.Println("  ✓ 删除第四段（使用 RemoveParagraph）")
	}

	// 验证剩余的段落
	paragraphs := doc.Body.GetParagraphs()
	fmt.Printf("\n  删除后文档包含 %d 个段落:\n", len(paragraphs))
	for i, p := range paragraphs {
		if len(p.Runs) > 0 {
			content := p.Runs[0].Text.Content
			if content != "" {
				fmt.Printf("  %d. %s\n", i+1, content)
			}
		}
	}

	// 保存文档
	filename := "examples/output/deletion_demo.docx"
	err := doc.Save(filename)
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Printf("\n✓ 段落删除演示文档已保存: %s\n", filename)
}

// demonstrateCombinedUsage 演示组合使用分页和删除功能
func demonstrateCombinedUsage() {
	fmt.Println("\n📝 演示3: 组合使用分页和删除功能")

	doc := document.New()

	// 创建包含分页符的文档
	doc.AddHeadingParagraph("第一章：引言", 1)
	doc.AddParagraph("引言内容...")
	tempPara := doc.AddParagraph("这是一个临时段落，稍后会被删除。")
	doc.AddParagraph("引言结论...")

	doc.AddPageBreak()

	doc.AddHeadingParagraph("第二章：正文", 1)
	doc.AddParagraph("正文内容...")

	doc.AddPageBreak()

	doc.AddHeadingParagraph("第三章：总结", 1)
	doc.AddParagraph("总结内容...")

	// 删除临时段落
	if doc.RemoveParagraph(tempPara) {
		fmt.Println("  ✓ 删除了临时段落")
	}

	// 保存文档
	filename := "examples/output/combined_demo.docx"
	err := doc.Save(filename)
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Printf("✓ 组合演示文档已保存: %s\n", filename)
	fmt.Println("  - 文档包含3个章节，使用分页符分隔")
	fmt.Println("  - 临时段落已被删除")
}

// Package main 演示UpdateTOC功能
package main

import (
	"fmt"
	"log"

	"github.com/mr-pmillz/wordZero/pkg/document"
)

func main() {
	fmt.Println("正在创建目录更新演示文档...")

	// 创建新文档
	doc := document.New()

	// 配置目录
	tocConfig := &document.TOCConfig{
		Title:       "目录", // 目录标题
		MaxLevel:    3,      // 包含到哪个标题级别
		ShowPageNum: true,   // 是否显示页码
		DotLeader:   true,   // 是否使用点状引导线
	}

	// 添加封面
	doc.AddParagraph("封面示例")

	// 生成目录（此时还没有标题）
	fmt.Println("生成初始目录...")
	err := doc.GenerateTOC(tocConfig)
	if err != nil {
		log.Fatalf("GenerateTOC失败: %v", err)
	}

	// 添加标题
	fmt.Println("添加标题...")
	doc.AddHeadingParagraph("第一章", 1)
	doc.AddParagraph("这是第一章的内容。")
	
	doc.AddHeadingParagraph("1.1 第一节", 2)
	doc.AddParagraph("这是第一节的内容。")
	
	doc.AddHeadingParagraph("1.1.1 第一小节", 3)
	doc.AddParagraph("这是第一小节的内容。")
	
	doc.AddHeadingParagraph("第二章", 1)
	doc.AddParagraph("这是第二章的内容。")
	
	doc.AddHeadingParagraph("2.1 第二节", 2)
	doc.AddParagraph("这是第二节的内容。")

	// 更新目录
	fmt.Println("更新目录...")
	err = doc.UpdateTOC()
	if err != nil {
		log.Fatalf("UpdateTOC失败: %v", err)
	}

	// 保存文档
	filename := "examples/output/toc_update_demo.docx"
	fmt.Printf("正在保存文档到: %s\n", filename)

	err = doc.Save(filename)
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Println("✅ 目录更新演示文档创建完成！")
	fmt.Println("📊 文档统计信息:")

	// 获取标题统计
	headingCount := doc.GetHeadingCount()
	for level := 1; level <= 3; level++ {
		if count, exists := headingCount[level]; exists {
			fmt.Printf("   - %d级标题: %d个\n", level, count)
		}
	}

	// 列出所有标题
	fmt.Println("📋 标题列表:")
	headings := doc.ListHeadings()
	for _, heading := range headings {
		indent := ""
		for i := 1; i < heading.Level; i++ {
			indent += "  "
		}
		fmt.Printf("   %s%d. %s\n", indent, heading.Level, heading.Text)
	}

	fmt.Printf("\n🎉 文档已成功保存到: %s\n", filename)
	fmt.Println("💡 提示：打开Word文档后，右键点击目录选择'更新域'，查看目录是否正确显示所有标题！")
}

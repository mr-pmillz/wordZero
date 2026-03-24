// Package main 测试与issue描述完全一致的场景
package main

import (
	"fmt"
	"log"

	"github.com/mr-pmillz/wordZero/pkg/document"
)

func main() {
	// 创建文档
	doc := document.New()

	// 配置目录 - 完全按照issue中的代码
	tocConfig := &document.TOCConfig{
		Title:       "目录", // 目录标题
		MaxLevel:    3,    // 包含到哪个标题级别
		ShowPageNum: true, // 是否显示页码
		DotLeader:   true, // 是否使用点状引导线
	}

	// 添加段落
	doc.AddParagraph("封面示例")

	// 生成目录
	err := doc.GenerateTOC(tocConfig)
	if err != nil {
		log.Fatalf("GenerateTOC失败: %v", err)
	}

	// 添加标题 - 完全按照issue中的代码
	doc.AddHeadingParagraph("第一章", 1)
	doc.AddHeadingParagraph("1.1", 2)
	doc.AddHeadingParagraph("第二章", 1)

	// 更新目录 - 这是issue中失败的调用
	err = doc.UpdateTOC()
	if err != nil {
		log.Fatalf("UpdateTOC失败: %v", err)
	}

	// 保存文档
	filename := "examples/output/issue_scenario_test.docx"
	err = doc.Save(filename)
	if err != nil {
		log.Fatalf("保存失败: %v", err)
	}

	fmt.Println("✅ 成功！issue场景测试通过！")
	fmt.Printf("文档已保存到: %s\n", filename)

	// 验证标题被正确收集
	headings := doc.ListHeadings()
	fmt.Printf("\n收集到 %d 个标题:\n", len(headings))
	for _, h := range headings {
		fmt.Printf("  - [级别%d] %s\n", h.Level, h.Text)
	}

	fmt.Println("\n💡 在Word中打开文档，目录应该显示:")
	fmt.Println("   目录")
	fmt.Println("   第一章 .............. 1")
	fmt.Println("     1.1 ............... 1")
	fmt.Println("   第二章 .............. 1")
}

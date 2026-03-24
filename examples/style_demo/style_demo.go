// Package main 展示WordZero完整样式系统的使用示例
package main

import (
	"fmt"
	"log"

	"github.com/mr-pmillz/wordZero/pkg/document"
	"github.com/mr-pmillz/wordZero/pkg/style"
)

func main() {
	// 创建新文档
	doc := document.New()

	// 获取样式管理器并创建快速API
	styleManager := doc.GetStyleManager()
	quickAPI := style.NewQuickStyleAPI(styleManager)

	fmt.Println("WordZero 完整样式系统演示")
	fmt.Println("==========================")

	// 1. 展示所有预定义样式
	demonstratePredefinedStyles(quickAPI)

	// 2. 演示样式继承机制
	demonstrateStyleInheritance(styleManager)

	// 3. 创建和使用自定义样式
	demonstrateCustomStyles(quickAPI)

	// 4. 创建样式化文档内容
	createStyledDocument(doc, styleManager, quickAPI)

	// 5. 演示样式查询和管理功能
	demonstrateStyleManagement(quickAPI)

	// 保存文档
	outputFile := "examples/output/styled_document_demo.docx"
	err := doc.Save(outputFile)
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Printf("\n✅ 样式化文档已保存到: %s\n", outputFile)
	fmt.Println("\n🎉 样式系统演示完成！")
}

// demonstratePredefinedStyles 展示预定义样式系统
func demonstratePredefinedStyles(api *style.QuickStyleAPI) {
	fmt.Println("\n📋 1. 预定义样式系统展示")
	fmt.Println("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

	// 显示所有样式信息
	allStyles := api.GetAllStylesInfo()
	fmt.Printf("总共有 %d 个预定义样式\n\n", len(allStyles))

	// 按类型显示样式
	fmt.Println("🏷️  段落样式:")
	paragraphStyles := api.GetParagraphStylesInfo()
	for _, info := range paragraphStyles {
		fmt.Printf("   %-15s | %-12s | %s\n", info.ID, info.Name, info.Description)
	}

	fmt.Println("\n🔤 字符样式:")
	characterStyles := api.GetCharacterStylesInfo()
	for _, info := range characterStyles {
		fmt.Printf("   %-15s | %-12s | %s\n", info.ID, info.Name, info.Description)
	}

	fmt.Println("\n📊 标题样式系列:")
	headingStyles := api.GetHeadingStylesInfo()
	for _, info := range headingStyles {
		basedOn := ""
		if info.BasedOn != "" {
			basedOn = fmt.Sprintf(" (基于: %s)", info.BasedOn)
		}
		fmt.Printf("   %-10s | %s%s\n", info.ID, info.Name, basedOn)
	}
}

// demonstrateStyleInheritance 演示样式继承机制
func demonstrateStyleInheritance(sm *style.StyleManager) {
	fmt.Println("\n🔗 2. 样式继承机制演示")
	fmt.Println("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

	// 演示标题样式的继承
	heading2Style := sm.GetStyleWithInheritance(style.StyleHeading2)
	if heading2Style != nil {
		fmt.Println("标题2样式继承分析:")

		if heading2Style.BasedOn != nil {
			fmt.Printf("   📍 基于样式: %s\n", heading2Style.BasedOn.Val)

			// 获取基础样式
			baseStyle := sm.GetStyle(heading2Style.BasedOn.Val)
			if baseStyle != nil {
				fmt.Println("   📋 继承的属性:")
				if baseStyle.RunPr != nil && baseStyle.RunPr.FontFamily != nil {
					fmt.Printf("      字体系列: %s (从 %s 继承)\n",
						baseStyle.RunPr.FontFamily.ASCII, heading2Style.BasedOn.Val)
				}
			}
		}

		if heading2Style.RunPr != nil {
			fmt.Println("   🎨 自有属性:")
			if heading2Style.RunPr.Bold != nil {
				fmt.Println("      加粗: 是")
			}
			if heading2Style.RunPr.FontSize != nil {
				fmt.Printf("      字体大小: %s (半磅单位)\n", heading2Style.RunPr.FontSize.Val)
			}
			if heading2Style.RunPr.Color != nil {
				fmt.Printf("      颜色: #%s\n", heading2Style.RunPr.Color.Val)
			}
		}
	}

	// 演示XML转换
	fmt.Println("\n   🔄 样式XML转换:")
	xmlData, err := sm.ApplyStyleToXML(style.StyleHeading2)
	if err == nil {
		fmt.Printf("      样式ID: %v\n", xmlData["styleId"])
		fmt.Printf("      类型: %v\n", xmlData["type"])
		if runProps, ok := xmlData["runProperties"]; ok {
			fmt.Printf("      字符属性: %+v\n", runProps)
		}
	}
}

// demonstrateCustomStyles 演示自定义样式创建
func demonstrateCustomStyles(api *style.QuickStyleAPI) {
	fmt.Println("\n🎨 3. 自定义样式创建演示")
	fmt.Println("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

	// 创建自定义标题样式
	titleConfig := style.QuickStyleConfig{
		ID:      "CustomDocTitle",
		Name:    "自定义文档标题",
		Type:    style.StyleTypeParagraph,
		BasedOn: style.StyleTitle,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "center",
			LineSpacing: 1.2,
			SpaceBefore: 24,
			SpaceAfter:  12,
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "微软雅黑",
			FontSize:  20,
			FontColor: "2E8B57",
			Bold:      true,
		},
	}

	customTitle, err := api.CreateQuickStyle(titleConfig)
	if err != nil {
		log.Printf("创建自定义标题样式失败: %v", err)
	} else {
		fmt.Printf("✅ 创建自定义标题样式: %s\n", customTitle.Name.Val)
		fmt.Printf("   ID: %s, 基于: %s\n", customTitle.StyleID, customTitle.BasedOn.Val)
	}

	// 创建自定义高亮样式
	highlightConfig := style.QuickStyleConfig{
		ID:   "ImportantHighlight",
		Name: "重要高亮",
		Type: style.StyleTypeCharacter,
		RunConfig: &style.QuickRunConfig{
			FontColor: "FF0000",
			Bold:      true,
			Highlight: "yellow",
		},
	}

	customHighlight, err := api.CreateQuickStyle(highlightConfig)
	if err != nil {
		log.Printf("创建高亮样式失败: %v", err)
	} else {
		fmt.Printf("✅ 创建字符高亮样式: %s\n", customHighlight.Name.Val)
	}

	// 创建自定义代码段落样式
	codeBlockConfig := style.QuickStyleConfig{
		ID:      "CustomCodeBlock",
		Name:    "自定义代码块",
		Type:    style.StyleTypeParagraph,
		BasedOn: style.StyleCodeBlock,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "left",
			LineSpacing: 1.0,
			SpaceBefore: 6,
			SpaceAfter:  6,
			LeftIndent:  20,
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "JetBrains Mono",
			FontSize:  9,
			FontColor: "000080",
		},
	}

	customCodeBlock, err := api.CreateQuickStyle(codeBlockConfig)
	if err != nil {
		log.Printf("创建代码块样式失败: %v", err)
	} else {
		fmt.Printf("✅ 创建自定义代码块样式: %s\n", customCodeBlock.Name.Val)
	}

	fmt.Printf("\n📊 当前样式总数: %d 个\n", len(api.GetAllStylesInfo()))
}

// createStyledDocument 创建样式化文档
func createStyledDocument(doc *document.Document, sm *style.StyleManager, api *style.QuickStyleAPI) {
	fmt.Println("\n📝 4. 创建样式化文档")
	fmt.Println("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

	// 使用自定义文档标题
	fmt.Println("   📋 添加自定义文档标题")
	titlePara := doc.AddParagraph("WordZero 样式系统完整指南")
	titlePara.SetStyle("CustomDocTitle")

	// 使用副标题样式
	fmt.Println("   📋 添加副标题")
	subtitlePara := doc.AddParagraph("全面展示预定义样式、自定义样式和样式继承")
	subtitlePara.SetStyle(style.StyleSubtitle)

	// 使用各级标题
	fmt.Println("   📋 添加多级标题结构")
	h1Para := doc.AddParagraph("第一章：样式系统概述")
	h1Para.SetStyle(style.StyleHeading1)

	h2Para := doc.AddParagraph("1.1 预定义样式")
	h2Para.SetStyle(style.StyleHeading2)

	h3Para := doc.AddParagraph("1.1.1 标题样式系列")
	h3Para.SetStyle(style.StyleHeading3)

	h4Para := doc.AddParagraph("Heading4 示例")
	h4Para.SetStyle(style.StyleHeading4)

	h5Para := doc.AddParagraph("Heading5 示例")
	h5Para.SetStyle(style.StyleHeading5)

	// 添加普通内容
	fmt.Println("   📋 添加正文内容")
	normalText := "WordZero 提供了完整的样式管理系统，支持18种预定义样式，包括9个标题层级、文档标题样式、引用样式、代码样式等。这些样式遵循Microsoft Word的OOXML规范，确保生成的文档具有专业的外观。"
	normalPara := doc.AddParagraph(normalText)
	normalPara.SetStyle(style.StyleNormal)

	// 使用引用样式
	fmt.Println("   📋 添加引用段落")
	quoteText := "样式是文档格式化的灵魂。通过合理使用样式，我们不仅能确保文档外观的一致性，还能提高文档的可维护性和专业性。—— WordZero设计理念"
	quotePara := doc.AddParagraph(quoteText)
	quotePara.SetStyle(style.StyleQuote)

	// 添加列表段落
	fmt.Println("   📋 添加列表内容")
	listTitle := doc.AddParagraph("样式系统的核心特性：")
	listTitle.SetStyle(style.StyleNormal)

	listItems := []string{
		"• 18种预定义样式，覆盖常用文档需求",
		"• 完整的样式继承机制，支持属性合并",
		"• 灵活的自定义样式创建接口",
		"• 类型安全的API设计",
		"• 符合OOXML规范的XML结构",
	}

	for _, item := range listItems {
		listPara := doc.AddParagraph(item)
		listPara.SetStyle(style.StyleListParagraph)
	}

	// 使用代码块样式
	fmt.Println("   📋 添加代码示例")
	codeTitle := doc.AddParagraph("代码示例：创建自定义样式")
	codeTitle.SetStyle(style.StyleHeading3)

	codeContent := `// 创建自定义样式
config := style.QuickStyleConfig{
    ID:      "MyStyle",
    Name:    "我的样式",
    Type:    style.StyleTypeParagraph,
    BasedOn: "Normal",
    RunConfig: &style.QuickRunConfig{
        FontName:  "微软雅黑",
        FontSize:  12,
        Bold:      true,
    },
}

style, err := quickAPI.CreateQuickStyle(config)`

	// 使用自定义代码块样式
	codePara := doc.AddParagraph(codeContent)
	codePara.SetStyle("CustomCodeBlock")

	// 演示混合格式段落
	fmt.Println("   📋 添加混合格式段落")
	mixedPara := doc.AddParagraph("")

	mixedPara.AddFormattedText("本段落演示了多种字符样式的组合使用：", nil)
	mixedPara.AddFormattedText("普通文本，", nil)
	mixedPara.AddFormattedText("粗体文本", &document.TextFormat{Bold: true})
	mixedPara.AddFormattedText("，", nil)
	mixedPara.AddFormattedText("斜体文本", &document.TextFormat{Italic: true})
	mixedPara.AddFormattedText("，", nil)
	mixedPara.AddFormattedText("代码文本", &document.TextFormat{
		FontFamily: "Consolas", FontColor: "E7484F", FontSize: 10})
	mixedPara.AddFormattedText("，以及", nil)
	mixedPara.AddFormattedText("重要高亮文本", &document.TextFormat{
		Bold: true, FontColor: "FF0000"})
	mixedPara.AddFormattedText("。", nil)

	// 总结段落
	fmt.Println("   📋 添加总结")
	summaryTitle := doc.AddParagraph("第二章：使用建议")
	summaryTitle.SetStyle(style.StyleHeading1)

	summaryText := "通过WordZero的样式系统，您可以轻松创建专业、美观、结构清晰的Word文档。建议在文档创建初期就规划好样式体系，这样能够大大提高文档制作效率。"
	summaryPara := doc.AddParagraph(summaryText)
	summaryPara.SetStyle(style.StyleNormal)

	fmt.Println("   ✅ 文档内容创建完成")
}

// demonstrateStyleManagement 演示样式查询和管理功能
func demonstrateStyleManagement(api *style.QuickStyleAPI) {
	fmt.Println("\n🔍 5. 样式管理功能演示")
	fmt.Println("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

	// 样式信息查询
	fmt.Println("📊 样式统计信息:")
	allStyles := api.GetAllStylesInfo()
	paragraphCount := len(api.GetParagraphStylesInfo())
	characterCount := len(api.GetCharacterStylesInfo())
	headingCount := len(api.GetHeadingStylesInfo())

	fmt.Printf("   总样式数: %d\n", len(allStyles))
	fmt.Printf("   段落样式: %d 个\n", paragraphCount)
	fmt.Printf("   字符样式: %d 个\n", characterCount)
	fmt.Printf("   标题样式: %d 个\n", headingCount)

	// 样式详情查询
	fmt.Println("\n🔍 样式详情查询示例:")
	styles := []string{style.StyleHeading1, style.StyleQuote, "CustomDocTitle"}
	for _, styleID := range styles {
		info, err := api.GetStyleInfo(styleID)
		if err == nil {
			fmt.Printf("   %s:\n", styleID)
			fmt.Printf("      名称: %s\n", info.Name)
			fmt.Printf("      类型: %s\n", info.Type)
			fmt.Printf("      内置: %v\n", info.IsBuiltIn)
			if info.BasedOn != "" {
				fmt.Printf("      基于: %s\n", info.BasedOn)
			}
			fmt.Printf("      描述: %s\n", info.Description)
		}
	}

	// 自定义样式列表
	fmt.Println("\n🎨 自定义样式列表:")
	customCount := 0
	for _, info := range allStyles {
		if !info.IsBuiltIn {
			fmt.Printf("   - %s (%s)\n", info.Name, info.ID)
			customCount++
		}
	}
	fmt.Printf("   共 %d 个自定义样式\n", customCount)

	fmt.Println("\n✨ 样式管理演示完成！")
}

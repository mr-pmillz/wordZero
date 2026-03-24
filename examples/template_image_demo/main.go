// Package main 图片占位符模板功能演示示例
package main

import (
	"bytes"
	"fmt"
	"image"
	"image/color"
	"image/png"
	"log"
	"os"

	"github.com/mr-pmillz/wordZero/pkg/document"
)

// createSampleImageWithColor 创建指定颜色的示例图片数据
func createSampleImageWithColor(width, height int, bgColor color.RGBA, text string) []byte {
	// 创建图片
	img := image.NewRGBA(image.Rect(0, 0, width, height))

	// 填充背景色
	for y := 0; y < height; y++ {
		for x := 0; x < width; x++ {
			img.Set(x, y, bgColor)
		}
	}

	// 添加边框
	borderColor := color.RGBA{0, 0, 0, 255} // 黑色边框
	for x := 0; x < width; x++ {
		img.Set(x, 0, borderColor)        // 上边框
		img.Set(x, height-1, borderColor) // 下边框
	}
	for y := 0; y < height; y++ {
		img.Set(0, y, borderColor)       // 左边框
		img.Set(width-1, y, borderColor) // 右边框
	}

	// 添加中心标记点（简单的十字）
	centerX := width / 2
	centerY := height / 2
	markColor := color.RGBA{0, 0, 0, 255} // 黑色标记

	// 画水平线
	for x := centerX - 10; x <= centerX+10; x++ {
		if x >= 0 && x < width {
			img.Set(x, centerY, markColor)
		}
	}

	// 画垂直线
	for y := centerY - 10; y <= centerY+10; y++ {
		if y >= 0 && y < height {
			img.Set(centerX, y, markColor)
		}
	}

	// 转换为PNG字节数组
	buf := new(bytes.Buffer)
	png.Encode(buf, img)
	return buf.Bytes()
}

func main() {
	fmt.Println("WordZero 图片占位符模板功能演示")
	fmt.Println("=====================================")

	// 确保输出目录存在
	if _, err := os.Stat("examples/output"); os.IsNotExist(err) {
		os.MkdirAll("examples/output", 0755)
	}

	// 演示1: 基础图片占位符
	fmt.Println("\n1. 基础图片占位符演示")
	demonstrateBasicImagePlaceholder()

	// 演示2: 配置图片样式的占位符
	fmt.Println("\n2. 配置图片样式演示")
	demonstrateStyledImagePlaceholder()

	// 演示3: 图片与文本混合模板
	fmt.Println("\n3. 图片与文本混合模板演示")
	demonstrateMixedContentTemplate()

	// 演示4: 从现有文档创建带图片的模板
	fmt.Println("\n4. 从现有文档创建图片模板演示")
	demonstrateDocumentImageTemplate()

	// 演示5: 二进制数据图片占位符
	fmt.Println("\n5. 二进制数据图片占位符演示")
	demonstrateBinaryImagePlaceholder()

	fmt.Println("\n=====================================")
	fmt.Println("图片占位符模板功能演示完成！")
	fmt.Println("生成的文档保存在 examples/output/ 目录下")
}

// demonstrateBasicImagePlaceholder 演示基础图片占位符功能
func demonstrateBasicImagePlaceholder() {
	// 创建模板引擎
	engine := document.NewTemplateEngine()

	// 创建包含图片占位符的模板
	templateContent := `产品介绍文档

产品名称：{{productName}}

产品图片：
{{#image productImage}}

产品描述：{{productDescription}}

技术规格：
- 尺寸：{{dimensions}}
- 重量：{{weight}}
- 颜色：{{color}}

联系我们：{{contactInfo}}`

	// 加载模板
	_, err := engine.LoadTemplate("product_intro", templateContent)
	if err != nil {
		log.Fatalf("加载模板失败: %v", err)
	}

	// 创建模板数据
	data := document.NewTemplateData()
	data.SetVariable("productName", "智能手表 Pro")
	data.SetVariable("productDescription", "这是一款功能强大的智能手表，具有健康监测、运动跟踪等多种功能。")
	data.SetVariable("dimensions", "45mm x 38mm x 10.7mm")
	data.SetVariable("weight", "32g")
	data.SetVariable("color", "太空灰")
	data.SetVariable("contactInfo", "电话：400-123-4567 | 邮箱：support@example.com")

	// 创建图片配置（默认居中显示）
	imageConfig := &document.ImageConfig{
		Position:  document.ImagePositionInline,
		Alignment: document.AlignCenter,
		Size: &document.ImageSize{
			Width:           100, // 100mm宽度
			KeepAspectRatio: true,
		},
		AltText: "智能手表产品图片",
		Title:   "智能手表 Pro",
	}

	// 设置图片数据（创建蓝色背景的产品图片）
	imageData := createSampleImageWithColor(200, 150, color.RGBA{100, 150, 255, 255}, "产品图片")
	data.SetImageFromData("productImage", imageData, imageConfig)

	// 渲染模板
	doc, err := engine.RenderToDocument("product_intro", data)
	if err != nil {
		log.Fatalf("渲染模板失败: %v", err)
	}

	// 保存文档
	err = doc.Save("examples/output/template_image_basic_demo.docx")
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Println("✓ 基础图片占位符演示完成，文档已保存为 template_image_basic_demo.docx")
}

// demonstrateStyledImagePlaceholder 演示配置图片样式的占位符
func demonstrateStyledImagePlaceholder() {
	engine := document.NewTemplateEngine()

	templateContent := `公司年度报告

{{companyName}} 2024年度报告

公司标志：
{{#image companyLogo}}

首席执行官致辞：
{{ceoMessage}}

核心团队：
{{#image teamPhoto}}

业绩数据：
销售额：{{revenue}}
增长率：{{growthRate}}

展望未来：
{{futureOutlook}}`

	_, err := engine.LoadTemplate("annual_report", templateContent)
	if err != nil {
		log.Fatalf("加载模板失败: %v", err)
	}

	data := document.NewTemplateData()
	data.SetVariable("companyName", "WordZero科技")
	data.SetVariable("ceoMessage", "过去的一年，我们在技术创新和市场拓展方面取得了显著成就...")
	data.SetVariable("revenue", "5000万元")
	data.SetVariable("growthRate", "25%")
	data.SetVariable("futureOutlook", "我们将继续专注于技术创新，为客户提供更优质的服务。")

	// 公司标志配置 - 小尺寸，右对齐，橙色背景
	logoConfig := &document.ImageConfig{
		Position:  document.ImagePositionInline,
		Alignment: document.AlignRight,
		Size: &document.ImageSize{
			Width:  50, // 50mm宽度
			Height: 20, // 20mm高度
		},
		AltText: "公司标志",
		Title:   "WordZero科技标志",
	}

	// 团队照片配置 - 大尺寸，居中，绿色背景
	teamConfig := &document.ImageConfig{
		Position:  document.ImagePositionInline,
		Alignment: document.AlignCenter,
		Size: &document.ImageSize{
			Width:           150, // 150mm宽度
			KeepAspectRatio: true,
		},
		AltText: "核心团队合影",
		Title:   "WordZero科技核心团队",
	}

	// 设置图片（使用不同颜色的图片）
	logoImageData := createSampleImageWithColor(150, 60, color.RGBA{255, 200, 100, 255}, "LOGO")
	teamImageData := createSampleImageWithColor(300, 200, color.RGBA{100, 255, 150, 255}, "团队照片")

	data.SetImageFromData("companyLogo", logoImageData, logoConfig)
	data.SetImageFromData("teamPhoto", teamImageData, teamConfig)

	doc, err := engine.RenderToDocument("annual_report", data)
	if err != nil {
		log.Fatalf("渲染模板失败: %v", err)
	}

	err = doc.Save("examples/output/template_image_styled_demo.docx")
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Println("✓ 配置图片样式演示完成，文档已保存为 template_image_styled_demo.docx")
}

// demonstrateMixedContentTemplate 演示图片与文本混合模板
func demonstrateMixedContentTemplate() {
	engine := document.NewTemplateEngine()

	templateContent := `技术文档：{{title}}

概述：{{overview}}

步骤1：{{step1Description}}
{{#image step1Image}}

步骤2：{{step2Description}}
{{#image step2Image}}

步骤3：{{step3Description}}
{{#image step3Image}}

{{#if hasWarning}}
⚠️ 注意事项：
{{warningText}}
{{/if}}

{{#each tips}}
💡 提示 {{@index}}：{{this}}
{{/each}}

结论：{{conclusion}}`

	_, err := engine.LoadTemplate("tech_doc", templateContent)
	if err != nil {
		log.Fatalf("加载模板失败: %v", err)
	}

	data := document.NewTemplateData()
	data.SetVariable("title", "智能设备安装指南")
	data.SetVariable("overview", "本文档将指导您完成智能设备的安装过程。")
	data.SetVariable("step1Description", "首先，打开包装盒并取出所有组件。")
	data.SetVariable("step2Description", "将设备连接到电源，等待指示灯亮起。")
	data.SetVariable("step3Description", "使用手机应用程序完成设备配置。")
	data.SetVariable("conclusion", "安装完成！设备现在可以正常使用了。")

	// 设置条件和列表
	data.SetCondition("hasWarning", true)
	data.SetVariable("warningText", "请确保在干燥环境中操作，避免水分接触设备。")

	tips := []interface{}{
		"确保Wi-Fi信号稳定",
		"保持手机和设备距离在3米以内",
		"如遇问题，请重启设备重试",
	}
	data.SetList("tips", tips)

	// 为每个步骤配置图片 - 使用不同颜色
	stepImageConfig := &document.ImageConfig{
		Position:  document.ImagePositionInline,
		Alignment: document.AlignCenter,
		Size: &document.ImageSize{
			Width:           80,
			KeepAspectRatio: true,
		},
	}

	// 创建不同颜色的步骤图片
	step1ImageData := createSampleImageWithColor(160, 120, color.RGBA{255, 180, 180, 255}, "步骤1")
	step2ImageData := createSampleImageWithColor(160, 120, color.RGBA{180, 255, 180, 255}, "步骤2")
	step3ImageData := createSampleImageWithColor(160, 120, color.RGBA{180, 180, 255, 255}, "步骤3")

	data.SetImageFromData("step1Image", step1ImageData, stepImageConfig)
	data.SetImageFromData("step2Image", step2ImageData, stepImageConfig)
	data.SetImageFromData("step3Image", step3ImageData, stepImageConfig)

	doc, err := engine.RenderToDocument("tech_doc", data)
	if err != nil {
		log.Fatalf("渲染模板失败: %v", err)
	}

	err = doc.Save("examples/output/template_image_mixed_demo.docx")
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Println("✓ 图片与文本混合模板演示完成，文档已保存为 template_image_mixed_demo.docx")
}

// demonstrateDocumentImageTemplate 演示从现有文档创建带图片的模板
func demonstrateDocumentImageTemplate() {
	// 首先创建一个基础文档作为模板
	baseDoc := document.New()

	// 添加标题
	title := baseDoc.AddParagraph("{{companyName}} 产品目录")
	title.SetAlignment(document.AlignCenter)

	// 添加介绍段落
	baseDoc.AddParagraph("欢迎浏览我们的产品目录。以下是我们的明星产品：")

	// 添加产品信息段落（包含图片占位符）
	baseDoc.AddParagraph("产品名称：{{productName}}")
	baseDoc.AddParagraph("{{#image productImage}}")
	baseDoc.AddParagraph("产品价格：{{price}}")
	baseDoc.AddParagraph("产品特色：{{features}}")

	// 添加联系信息
	contact := baseDoc.AddParagraph("联系我们：{{contactInfo}}")
	contact.SetAlignment(document.AlignCenter)

	// 从基础文档创建模板
	engine := document.NewTemplateEngine()
	template, err := engine.LoadTemplateFromDocument("product_catalog", baseDoc)
	if err != nil {
		log.Fatalf("从文档创建模板失败: %v", err)
	}

	fmt.Printf("从文档创建的模板包含 %d 个变量\n", len(template.Variables))

	// 准备数据
	data := document.NewTemplateData()
	data.SetVariable("companyName", "创新科技")
	data.SetVariable("productName", "智能音箱 X1")
	data.SetVariable("price", "￥299")
	data.SetVariable("features", "AI语音助手、高保真音质、智能家居控制")
	data.SetVariable("contactInfo", "官网：www.example.com | 热线：400-888-9999")

	// 配置产品图片 - 紫色背景
	productImageConfig := &document.ImageConfig{
		Position:  document.ImagePositionInline,
		Alignment: document.AlignCenter,
		Size: &document.ImageSize{
			Width:           100,
			KeepAspectRatio: true,
		},
		AltText: "智能音箱产品图片",
		Title:   "智能音箱 X1",
	}

	imageData := createSampleImageWithColor(200, 150, color.RGBA{200, 150, 255, 255}, "音箱图片")
	data.SetImageFromData("productImage", imageData, productImageConfig)

	// 渲染模板
	doc, err := engine.RenderTemplateToDocument("product_catalog", data)
	if err != nil {
		log.Fatalf("渲染模板失败: %v", err)
	}

	err = doc.Save("examples/output/template_image_from_doc_demo.docx")
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Println("✓ 从现有文档创建图片模板演示完成，文档已保存为 template_image_from_doc_demo.docx")
}

// demonstrateBinaryImagePlaceholder 演示二进制数据图片占位符
func demonstrateBinaryImagePlaceholder() {
	engine := document.NewTemplateEngine()

	templateContent := `数据分析报告

报告标题：{{reportTitle}}
生成时间：{{generateTime}}

关键指标图表：
{{#image chartImage}}

数据摘要：
{{summary}}

详细分析：
{{analysis}}

结论与建议：
{{conclusion}}`

	_, err := engine.LoadTemplate("data_report", templateContent)
	if err != nil {
		log.Fatalf("加载模板失败: %v", err)
	}

	data := document.NewTemplateData()
	data.SetVariable("reportTitle", "2024年第三季度销售数据分析")
	data.SetVariable("generateTime", "2024年10月15日")
	data.SetVariable("summary", "本季度销售额较上季度增长15%，各产品线表现良好。")
	data.SetVariable("analysis", "移动端销售占比持续提升，达到总销售额的60%。华东地区仍是最大市场。")
	data.SetVariable("conclusion", "建议继续加强移动端渠道建设，并在华南地区投入更多营销资源。")

	// 模拟图表数据 - 黄色背景的图表
	chartImageData := createSampleImageWithColor(300, 200, color.RGBA{255, 255, 150, 255}, "数据图表")

	chartConfig := &document.ImageConfig{
		Position:  document.ImagePositionInline,
		Alignment: document.AlignCenter,
		Size: &document.ImageSize{
			Width:           120,
			KeepAspectRatio: true,
		},
		AltText: "销售数据图表",
		Title:   "2024 Q3 销售数据图表",
	}

	// 使用二进制数据设置图片
	data.SetImageFromData("chartImage", chartImageData, chartConfig)

	doc, err := engine.RenderToDocument("data_report", data)
	if err != nil {
		log.Fatalf("渲染模板失败: %v", err)
	}

	err = doc.Save("examples/output/template_image_binary_demo.docx")
	if err != nil {
		log.Fatalf("保存文档失败: %v", err)
	}

	fmt.Println("✓ 二进制数据图片占位符演示完成，文档已保存为 template_image_binary_demo.docx")
}

// createSampleImageData 创建示例图片数据（为了向后兼容保留，但现在使用红色背景）
func createSampleImageData() []byte {
	return createSampleImageWithColor(100, 100, color.RGBA{255, 100, 100, 255}, "示例图片")
}

# WordZero Multi-Language Documentation Guide

This document outlines the multi-language structure and navigation for WordZero documentation.

## 📚 Documentation Structure

### Main Project Documentation

```
wordZero/
├── README.md              # English (default)
├── README_zh.md           # Chinese
└── docs/
    ├── MULTILINGUAL_GUIDE.md  # This file
    ├── en/                    # English docs (future expansion)
    └── zh/                    # Chinese docs (future expansion)
```

### Wiki Documentation Structure

The GitHub Wiki uses filename prefixes to organize multi-language content:

```
wordZero.wiki/
├── Home.md                # Chinese (default)
├── en-Home.md             # English
├── _Sidebar.md            # Multi-language navigation
├── 
├── Chinese Documents (default):
│   ├── 01-快速开始.md
│   ├── 02-基础功能.md
│   ├── 03-样式系统.md
│   ├── 04-文本格式化.md
│   ├── 05-表格操作.md
│   ├── 06-页面设置.md
│   ├── 07-图片操作.md
│   ├── 08-高级功能.md
│   ├── 09-最佳实践.md
│   ├── 10-API参考.md
│   ├── 11-示例项目.md
│   ├── 12-模板功能.md
│   ├── 13-性能基准测试.md
│   ├── 14-功能特性详览.md
│   ├── 15-项目结构详解.md
│   └── 16-Markdown双向转换.md
│
└── English Documents (en- prefix):
    ├── en-Quick-Start.md
    ├── en-Basic-Features.md
    ├── en-Style-System.md
    ├── en-Text-Formatting.md
    ├── en-Table-Operations.md
    ├── en-Page-Settings.md
    ├── en-Image-Operations.md
    ├── en-Advanced-Features.md
    ├── en-Best-Practices.md
    ├── en-API-Reference.md
    ├── en-Example-Projects.md
    ├── en-Template-Features.md
    ├── en-Performance-Benchmarks.md
    ├── en-Feature-Overview.md
    ├── en-Project-Structure.md
    └── en-Markdown-Conversion.md
```

## 🌍 Language Navigation

### Primary Navigation

Each document page includes language switching links at the top:

```markdown
[**中文文档**](Chinese-Page) | **English Documentation**
```

or

```markdown
**中文文档** | [English Documentation](en-English-Page)
```

### Sidebar Navigation

The `_Sidebar.md` file contains organized navigation for both languages:

- **Language switcher** at the top
- **Chinese section** with all Chinese documents
- **English section** with all English documents
- **External links** section

## 🔄 Content Synchronization

### Document Mapping

| Chinese Document | English Document | Status |
|------------------|------------------|--------|
| Home.md | en-Home.md | ✅ Complete |
| 01-快速开始.md | en-Quick-Start.md | ✅ Complete |
| 02-基础功能.md | en-Basic-Features.md | 🚧 In Progress |
| 03-样式系统.md | en-Style-System.md | 🚧 In Progress |
| 04-文本格式化.md | en-Text-Formatting.md | 🚧 In Progress |
| 05-表格操作.md | en-Table-Operations.md | 🚧 In Progress |
| 06-页面设置.md | en-Page-Settings.md | 🚧 In Progress |
| 07-图片操作.md | en-Image-Operations.md | 🚧 In Progress |
| 08-高级功能.md | en-Advanced-Features.md | 🚧 In Progress |
| 09-最佳实践.md | en-Best-Practices.md | 🚧 In Progress |
| 10-API参考.md | en-API-Reference.md | 🚧 In Progress |
| 11-示例项目.md | en-Example-Projects.md | 🚧 In Progress |
| 12-模板功能.md | en-Template-Features.md | 🚧 In Progress |
| 13-性能基准测试.md | en-Performance-Benchmarks.md | ✅ Complete |
| 14-功能特性详览.md | en-Feature-Overview.md | ✅ Complete |
| 15-项目结构详解.md | en-Project-Structure.md | 🚧 In Progress |
| 16-Markdown双向转换.md | en-Markdown-Conversion.md | 🚧 In Progress |

### Content Maintenance

1. **Source of Truth**: Chinese documents are the primary source
2. **Translation Process**: English documents are translated from Chinese
3. **Version Control**: Both versions should be updated simultaneously
4. **Link Consistency**: Cross-references should point to appropriate language versions

## 📝 Writing Guidelines

### Chinese Documents

- Use simplified Chinese characters
- Follow Chinese technical writing conventions
- Include Chinese-specific examples where relevant
- Use Chinese font examples (宋体, 微软雅黑, etc.)

### English Documents

- Use clear, technical English
- Follow standard technical documentation practices
- Use Western font examples (Arial, Times New Roman, etc.)
- Consider international audience (avoid US-specific references)

### Code Examples

- Keep code examples identical between languages
- Translate comments and string literals
- Use appropriate locale-specific examples
- Maintain consistent variable naming

Example:
```go
// Chinese version
doc := document.New()
标题 := doc.AddParagraph("WordZero 使用示例")
标题.SetStyle(style.StyleHeading1)

// English version  
doc := document.New()
title := doc.AddParagraph("WordZero Usage Example")
title.SetStyle(style.StyleHeading1)
```

## 🛠️ Maintenance Tasks

### Regular Updates

1. **Content Sync**: Ensure feature parity between languages
2. **Link Validation**: Check all cross-references work correctly
3. **Example Updates**: Keep code examples current with API changes
4. **Navigation Updates**: Maintain sidebar and cross-links

### New Document Process

1. **Create Chinese version** first (source of truth)
2. **Create English version** with `en-` prefix
3. **Update sidebar navigation** for both languages
4. **Add cross-language links** in both documents
5. **Update this guide** with new document mapping

### Quality Assurance

- **Language consistency** within each version
- **Technical accuracy** across both languages
- **Link integrity** between language versions
- **Navigation completeness** in sidebar

## 🔗 External References

### Project Links

- **GitHub Repository**: https://github.com/mr-pmillz/wordZero
- **Chinese Wiki**: https://github.com/mr-pmillz/wordZero/wiki
- **English Wiki**: https://github.com/mr-pmillz/wordZero/wiki/en-Home

### Related Documentation

- **API Documentation** (auto-generated, English)
- **Go Package Documentation** (godoc, English)
- **Example Code** (comments in both languages)

## 📊 Analytics and Feedback

### Usage Tracking

- Monitor which language documentation is accessed more
- Track user navigation patterns between languages
- Identify commonly requested translations

### Community Feedback

- GitHub Issues for documentation improvements
- Language-specific feedback channels
- Translation quality assessments

---

This guide will be updated as the multi-language documentation structure evolves. 
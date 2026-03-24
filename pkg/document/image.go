// Package document provides image operations for Word documents.
package document

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"image"
	"image/gif"
	"image/jpeg"
	"image/png"
	"os"
	"path/filepath"
	"strconv"
)

// ImageFormat represents the image format type.
type ImageFormat string

const (
	// Supported image formats
	ImageFormatJPEG ImageFormat = "jpeg"
	ImageFormatPNG  ImageFormat = "png"
	ImageFormatGIF  ImageFormat = "gif"
)

// ImagePosition represents the image position type.
type ImagePosition string

const (
	// Image position options
	ImagePositionInline     ImagePosition = "inline"     // Inline (default)
	ImagePositionFloatLeft  ImagePosition = "floatLeft"  // Float left
	ImagePositionFloatRight ImagePosition = "floatRight" // Float right
)

// ImageWrapText represents the text wrapping type.
type ImageWrapText string

const (
	// Text wrapping options
	ImageWrapNone         ImageWrapText = "none"         // No wrapping
	ImageWrapSquare       ImageWrapText = "square"       // Square wrapping
	ImageWrapTight        ImageWrapText = "tight"        // Tight wrapping
	ImageWrapTopAndBottom ImageWrapText = "topAndBottom" // Top and bottom wrapping
)

// ImageSize represents the image size configuration.
type ImageSize struct {
	Width  float64 // Width in millimeters
	Height float64 // Height in millimeters
	// Whether to keep the aspect ratio when only one dimension is set
	KeepAspectRatio bool
}

// ImageConfig represents the image configuration.
type ImageConfig struct {
	// Image size
	Size *ImageSize
	// Image position
	Position ImagePosition
	// Image alignment (for inline images)
	Alignment AlignmentType
	// Text wrapping
	WrapText ImageWrapText
	// Image description (alt text)
	AltText string
	// Image title
	Title string
	// Horizontal offset in millimeters
	OffsetX float64
	// Vertical offset in millimeters
	OffsetY float64
}

// ImageInfo represents image information.
type ImageInfo struct {
	ID         string       // Image ID
	RelationID string       // Relationship ID
	Format     ImageFormat  // Image format
	Width      int          // Original width in pixels
	Height     int          // Original height in pixels
	Data       []byte       // Image data
	Config     *ImageConfig // Image configuration
}

// DrawingElement represents a drawing element containing an image.
type DrawingElement struct {
	XMLName xml.Name       `xml:"w:drawing"`
	Inline  *InlineDrawing `xml:"wp:inline,omitempty"`
	Anchor  *AnchorDrawing `xml:"wp:anchor,omitempty"`
}

// InlineDrawing represents an inline drawing element.
type InlineDrawing struct {
	XMLName xml.Name        `xml:"wp:inline"`
	DistT   string          `xml:"distT,attr,omitempty"`
	DistB   string          `xml:"distB,attr,omitempty"`
	DistL   string          `xml:"distL,attr,omitempty"`
	DistR   string          `xml:"distR,attr,omitempty"`
	Extent  *DrawingExtent  `xml:"wp:extent"`
	DocPr   *DrawingDocPr   `xml:"wp:docPr"`
	Graphic *DrawingGraphic `xml:"a:graphic"`
}

// AnchorDrawing represents a floating (anchored) drawing element.
type AnchorDrawing struct {
	XMLName           xml.Name            `xml:"wp:anchor"`
	DistT             string              `xml:"distT,attr,omitempty"`
	DistB             string              `xml:"distB,attr,omitempty"`
	DistL             string              `xml:"distL,attr,omitempty"`
	DistR             string              `xml:"distR,attr,omitempty"`
	SimplePos         string              `xml:"simplePos,attr,omitempty"`
	RelativeHeight    string              `xml:"relativeHeight,attr,omitempty"`
	BehindDoc         string              `xml:"behindDoc,attr,omitempty"`
	Locked            string              `xml:"locked,attr,omitempty"`
	LayoutInCell      string              `xml:"layoutInCell,attr,omitempty"`
	AllowOverlap      string              `xml:"allowOverlap,attr,omitempty"`
	SimplePosition    *SimplePosition     `xml:"wp:simplePos,omitempty"`
	PositionH         *HorizontalPosition `xml:"wp:positionH,omitempty"`
	PositionV         *VerticalPosition   `xml:"wp:positionV,omitempty"`
	Extent            *DrawingExtent      `xml:"wp:extent"`
	EffectExtent      *EffectExtent       `xml:"wp:effectExtent,omitempty"`
	WrapNone          *WrapNone           `xml:"wp:wrapNone,omitempty"`
	WrapSquare        *WrapSquare         `xml:"wp:wrapSquare,omitempty"`
	WrapTight         *WrapTight          `xml:"wp:wrapTight,omitempty"`
	WrapThrough       *WrapThrough        `xml:"wp:wrapThrough,omitempty"`
	WrapTopAndBottom  *WrapTopAndBottom   `xml:"wp:wrapTopAndBottom,omitempty"`
	DocPr             *DrawingDocPr       `xml:"wp:docPr"`
	CNvGraphicFramePr *CNvGraphicFramePr  `xml:"wp:cNvGraphicFramePr,omitempty"`
	Graphic           *DrawingGraphic     `xml:"a:graphic"`
}

// SimplePosition represents a simple position.
type SimplePosition struct {
	XMLName xml.Name `xml:"wp:simplePos"`
	X       string   `xml:"x,attr"`
	Y       string   `xml:"y,attr"`
}

// HorizontalPosition represents a horizontal position.
type HorizontalPosition struct {
	XMLName      xml.Name   `xml:"wp:positionH"`
	RelativeFrom string     `xml:"relativeFrom,attr"`
	Align        *PosAlign  `xml:"wp:align,omitempty"`
	PosOffset    *PosOffset `xml:"wp:posOffset,omitempty"`
}

// VerticalPosition represents a vertical position.
type VerticalPosition struct {
	XMLName      xml.Name   `xml:"wp:positionV"`
	RelativeFrom string     `xml:"relativeFrom,attr"`
	Align        *PosAlign  `xml:"wp:align,omitempty"`
	PosOffset    *PosOffset `xml:"wp:posOffset,omitempty"`
}

// PosAlign represents a position alignment.
type PosAlign struct {
	XMLName xml.Name `xml:"wp:align"`
	Value   string   `xml:",chardata"`
}

// PosOffset represents a position offset.
type PosOffset struct {
	XMLName xml.Name `xml:"wp:posOffset"`
	Value   string   `xml:",chardata"`
}

// EffectExtent represents the effect extent.
type EffectExtent struct {
	XMLName xml.Name `xml:"wp:effectExtent"`
	L       string   `xml:"l,attr,omitempty"`
	T       string   `xml:"t,attr,omitempty"`
	R       string   `xml:"r,attr,omitempty"`
	B       string   `xml:"b,attr,omitempty"`
}

// WrapNone represents no text wrapping.
type WrapNone struct {
	XMLName xml.Name `xml:"wp:wrapNone"`
}

// WrapSquare represents square text wrapping.
type WrapSquare struct {
	XMLName  xml.Name `xml:"wp:wrapSquare"`
	WrapText string   `xml:"wrapText,attr,omitempty"`
	DistT    string   `xml:"distT,attr,omitempty"`
	DistB    string   `xml:"distB,attr,omitempty"`
	DistL    string   `xml:"distL,attr,omitempty"`
	DistR    string   `xml:"distR,attr,omitempty"`
}

// WrapTight represents tight text wrapping.
type WrapTight struct {
	XMLName     xml.Name     `xml:"wp:wrapTight"`
	WrapText    string       `xml:"wrapText,attr,omitempty"`
	DistL       string       `xml:"distL,attr,omitempty"`
	DistR       string       `xml:"distR,attr,omitempty"`
	WrapPolygon *WrapPolygon `xml:"wp:wrapPolygon,omitempty"`
}

// WrapThrough represents through text wrapping.
type WrapThrough struct {
	XMLName     xml.Name     `xml:"wp:wrapThrough"`
	WrapText    string       `xml:"wrapText,attr,omitempty"`
	DistL       string       `xml:"distL,attr,omitempty"`
	DistR       string       `xml:"distR,attr,omitempty"`
	WrapPolygon *WrapPolygon `xml:"wp:wrapPolygon,omitempty"`
}

// WrapTopAndBottom represents top and bottom text wrapping.
type WrapTopAndBottom struct {
	XMLName      xml.Name      `xml:"wp:wrapTopAndBottom"`
	DistT        string        `xml:"distT,attr,omitempty"`
	DistB        string        `xml:"distB,attr,omitempty"`
	EffectExtent *EffectExtent `xml:"wp:effectExtent,omitempty"`
}

// WrapPolygon represents a wrapping polygon.
type WrapPolygon struct {
	XMLName xml.Name        `xml:"wp:wrapPolygon"`
	Start   *PolygonStart   `xml:"wp:start"`
	LineTo  []PolygonLineTo `xml:"wp:lineTo"`
}

// PolygonStart represents the starting point of a polygon.
type PolygonStart struct {
	XMLName xml.Name `xml:"wp:start"`
	X       string   `xml:"x,attr"`
	Y       string   `xml:"y,attr"`
}

// PolygonLineTo represents a line segment of a polygon.
type PolygonLineTo struct {
	XMLName xml.Name `xml:"wp:lineTo"`
	X       string   `xml:"x,attr"`
	Y       string   `xml:"y,attr"`
}

// CNvGraphicFramePr represents non-visual graphic frame properties.
type CNvGraphicFramePr struct {
	XMLName           xml.Name           `xml:"wp:cNvGraphicFramePr"`
	GraphicFrameLocks *GraphicFrameLocks `xml:"a:graphicFrameLocks,omitempty"`
}

// GraphicFrameLocks represents graphic frame lock settings.
type GraphicFrameLocks struct {
	XMLName        xml.Name `xml:"a:graphicFrameLocks"`
	Xmlns          string   `xml:"xmlns:a,attr,omitempty"`
	NoChangeAspect string   `xml:"noChangeAspect,attr,omitempty"`
	NoCrop         string   `xml:"noCrop,attr,omitempty"`
	NoMove         string   `xml:"noMove,attr,omitempty"`
	NoResize       string   `xml:"noResize,attr,omitempty"`
	NoRot          string   `xml:"noRot,attr,omitempty"`
	NoSelect       string   `xml:"noSelect,attr,omitempty"`
}

// DrawingExtent represents the drawing dimensions.
type DrawingExtent struct {
	XMLName xml.Name `xml:"wp:extent"`
	Cx      string   `xml:"cx,attr"`
	Cy      string   `xml:"cy,attr"`
}

// DrawingDocPr represents drawing document properties.
type DrawingDocPr struct {
	XMLName xml.Name `xml:"wp:docPr"`
	ID      string   `xml:"id,attr"`
	Name    string   `xml:"name,attr"`
	Descr   string   `xml:"descr,attr,omitempty"`
	Title   string   `xml:"title,attr,omitempty"`
}

// DrawingGraphic represents a graphic element.
type DrawingGraphic struct {
	XMLName     xml.Name     `xml:"a:graphic"`
	Xmlns       string       `xml:"xmlns:a,attr"`
	GraphicData *GraphicData `xml:"a:graphicData"`
}

// GraphicData represents graphic data.
type GraphicData struct {
	XMLName xml.Name    `xml:"a:graphicData"`
	URI     string      `xml:"uri,attr"`
	Pic     *PicElement `xml:"pic:pic"`
}

// PicElement represents a picture element.
type PicElement struct {
	XMLName  xml.Name  `xml:"pic:pic"`
	Xmlns    string    `xml:"xmlns:pic,attr"`
	NvPicPr  *NvPicPr  `xml:"pic:nvPicPr"`
	BlipFill *BlipFill `xml:"pic:blipFill"`
	SpPr     *SpPr     `xml:"pic:spPr"`
}

// NvPicPr represents non-visual picture properties.
type NvPicPr struct {
	XMLName  xml.Name  `xml:"pic:nvPicPr"`
	CNvPr    *CNvPr    `xml:"pic:cNvPr"`
	CNvPicPr *CNvPicPr `xml:"pic:cNvPicPr"`
}

// CNvPr represents common non-visual properties.
type CNvPr struct {
	XMLName xml.Name `xml:"pic:cNvPr"`
	ID      string   `xml:"id,attr"`
	Name    string   `xml:"name,attr"`
	Descr   string   `xml:"descr,attr,omitempty"`
	Title   string   `xml:"title,attr,omitempty"`
}

// CNvPicPr represents picture-specific non-visual properties.
type CNvPicPr struct {
	XMLName  xml.Name  `xml:"pic:cNvPicPr"`
	PicLocks *PicLocks `xml:"a:picLocks,omitempty"`
}

// PicLocks represents picture lock properties.
type PicLocks struct {
	XMLName            xml.Name `xml:"a:picLocks"`
	NoChangeAspect     string   `xml:"noChangeAspect,attr,omitempty"`
	NoChangeArrowheads string   `xml:"noChangeArrowheads,attr,omitempty"`
}

// BlipFill represents a picture fill.
type BlipFill struct {
	XMLName xml.Name `xml:"pic:blipFill"`
	Blip    *Blip    `xml:"a:blip"`
	Stretch *Stretch `xml:"a:stretch"`
}

// Blip represents a binary large image or picture.
type Blip struct {
	XMLName xml.Name `xml:"a:blip"`
	Embed   string   `xml:"r:embed,attr"`
}

// Stretch represents a stretch fill.
type Stretch struct {
	XMLName  xml.Name  `xml:"a:stretch"`
	FillRect *FillRect `xml:"a:fillRect"`
}

// FillRect represents a fill rectangle.
type FillRect struct {
	XMLName xml.Name `xml:"a:fillRect"`
}

// SpPr represents shape properties.
type SpPr struct {
	XMLName  xml.Name  `xml:"pic:spPr"`
	Xfrm     *Xfrm     `xml:"a:xfrm"`
	PrstGeom *PrstGeom `xml:"a:prstGeom"`
}

// Xfrm represents a 2D transform.
type Xfrm struct {
	XMLName xml.Name `xml:"a:xfrm"`
	Off     *Off     `xml:"a:off,omitempty"`
	Ext     *Ext     `xml:"a:ext"`
}

// Off represents an offset.
type Off struct {
	XMLName xml.Name `xml:"a:off"`
	X       string   `xml:"x,attr"`
	Y       string   `xml:"y,attr"`
}

// Ext represents an extent (size).
type Ext struct {
	XMLName xml.Name `xml:"a:ext"`
	Cx      string   `xml:"cx,attr"`
	Cy      string   `xml:"cy,attr"`
}

// PrstGeom represents a preset geometry.
type PrstGeom struct {
	XMLName xml.Name `xml:"a:prstGeom"`
	Prst    string   `xml:"prst,attr"`
	AvLst   *AvLst   `xml:"a:avLst"`
}

// AvLst represents an adjustment value list.
type AvLst struct {
	XMLName xml.Name `xml:"a:avLst"`
}

// AddImageFromFile adds an image to the document from a file.
func (d *Document) AddImageFromFile(filePath string, config *ImageConfig) (*ImageInfo, error) {
	filePath = filepath.Clean(filePath)
	DebugMsgf(MsgAddingImageFile, filePath)

	// Read image file
	imageData, err := os.ReadFile(filePath)
	if err != nil {
		Errorf("failed to read image file %s: %v", filePath, err)
		return nil, fmt.Errorf("failed to read image file: %w", err)
	}

	// Detect image format
	format, err := detectImageFormat(imageData)
	if err != nil {
		Errorf("failed to detect image format %s: %v", filePath, err)
		return nil, fmt.Errorf("failed to detect image format: %w", err)
	}

	// Get image dimensions
	width, height, err := getImageDimensions(imageData, format)
	if err != nil {
		Errorf("failed to get image dimensions %s: %v", filePath, err)
		return nil, fmt.Errorf("failed to get image dimensions: %w", err)
	}

	fileName := filepath.Base(filePath)
	InfoMsgf(MsgImageRead, fileName, format, width, height, len(imageData))
	return d.AddImageFromData(imageData, fileName, format, width, height, config)
}

// generateSafeImageFileName generates a safe image file name.
// Converts file names with non-ASCII characters to safe ASCII file names for Microsoft Word compatibility.
func generateSafeImageFileName(imageID int, originalFileName string, format ImageFormat) string {
	// Get the file extension
	ext := filepath.Ext(originalFileName)
	if ext == "" {
		// If no extension, add one based on the format
		switch format {
		case ImageFormatPNG:
			ext = ".png"
		case ImageFormatJPEG:
			ext = ".jpeg"
		case ImageFormatGIF:
			ext = ".gif"
		default:
			ext = ".png"
		}
	}

	// Generate a safe file name using the image ID
	safeFileName := fmt.Sprintf("image%d%s", imageID, ext)
	return safeFileName
}

// AddImageFromData adds an image to the document from binary data.
func (d *Document) AddImageFromData(imageData []byte, fileName string, format ImageFormat, width, height int, config *ImageConfig) (*ImageInfo, error) {
	if d.documentRelationships == nil {
		d.documentRelationships = &Relationships{
			Xmlns:         "http://schemas.openxmlformats.org/package/2006/relationships",
			Relationships: []Relationship{},
		}
	}

	// Use the document-level image ID counter to ensure ID uniqueness
	imageID := d.nextImageID
	d.nextImageID++ // Increment counter

	// Generate a safe file name (avoid non-ASCII characters causing Word open errors)
	safeFileName := generateSafeImageFileName(imageID, fileName, format)

	// Generate relationship ID; note: rId1 is reserved for styles.xml, images start from rId2
	relationID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2)

	// Add image relationship using the safe file name
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, Relationship{
		ID:     relationID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
		Target: fmt.Sprintf("media/%s", safeFileName),
	})

	// Store image data using the safe file name
	if d.parts == nil {
		d.parts = make(map[string][]byte)
	}
	d.parts[fmt.Sprintf("word/media/%s", safeFileName)] = imageData

	// Update content types
	d.addImageContentType(format)

	// Create image info
	imageInfo := &ImageInfo{
		ID:         strconv.Itoa(imageID),
		RelationID: relationID,
		Format:     format,
		Width:      width,
		Height:     height,
		Data:       imageData,
		Config:     config,
	}

	// Create image paragraph and add to document
	paragraph := d.createImageParagraph(imageInfo)
	d.Body.AddElement(paragraph)

	return imageInfo, nil
}

// AddImageFromDataWithoutElement adds an image to the document from binary data without creating a paragraph element.
// This method is used by the template engine and other scenarios that manage image paragraphs themselves.
func (d *Document) AddImageFromDataWithoutElement(imageData []byte, fileName string, format ImageFormat, width, height int, config *ImageConfig) (*ImageInfo, error) {
	if d.documentRelationships == nil {
		d.documentRelationships = &Relationships{
			Xmlns:         "http://schemas.openxmlformats.org/package/2006/relationships",
			Relationships: []Relationship{},
		}
	}

	// Use the document-level image ID counter to ensure ID uniqueness
	imageID := d.nextImageID
	d.nextImageID++ // Increment counter

	// Generate a safe file name (avoid non-ASCII characters causing Word open errors)
	safeFileName := generateSafeImageFileName(imageID, fileName, format)

	// Generate relationship ID; note: rId1 is reserved for styles.xml, images start from rId2
	relationID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2)

	// Add image relationship using the safe file name
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, Relationship{
		ID:     relationID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
		Target: fmt.Sprintf("media/%s", safeFileName),
	})

	// Store image data using the safe file name
	if d.parts == nil {
		d.parts = make(map[string][]byte)
	}
	d.parts[fmt.Sprintf("word/media/%s", safeFileName)] = imageData

	// Update content types
	d.addImageContentType(format)

	// Create image info
	imageInfo := &ImageInfo{
		ID:         strconv.Itoa(imageID),
		RelationID: relationID,
		Format:     format,
		Width:      width,
		Height:     height,
		Data:       imageData,
		Config:     config,
	}

	// Note: this method does not create a paragraph element; the caller is responsible for managing that
	return imageInfo, nil
}

// createImageParagraph creates a paragraph containing an image.
func (d *Document) createImageParagraph(imageInfo *ImageInfo) *Paragraph {
	// Calculate image display size in EMU units
	displayWidth, displayHeight := d.calculateDisplaySize(imageInfo)

	// Get image description and title
	altText := "Image"
	title := "Image"
	if imageInfo.Config != nil {
		if imageInfo.Config.AltText != "" {
			altText = imageInfo.Config.AltText
		}
		if imageInfo.Config.Title != "" {
			title = imageInfo.Config.Title
		}
	}

	// Create Drawing element
	var drawing *DrawingElement

	// Check if this is a floating image
	if imageInfo.Config != nil &&
		(imageInfo.Config.Position == ImagePositionFloatLeft ||
			imageInfo.Config.Position == ImagePositionFloatRight) {
		// Create floating image
		drawing = d.createFloatingImageDrawing(imageInfo, displayWidth, displayHeight, altText, title)
	} else {
		// Create inline image
		drawing = d.createInlineImageDrawing(imageInfo, displayWidth, displayHeight, altText, title)
	}

	// Create paragraph containing the image
	paragraph := &Paragraph{
		Runs: []Run{
			{
				Drawing: drawing,
			},
		},
	}

	// Set paragraph alignment for inline images
	if imageInfo.Config != nil &&
		imageInfo.Config.Position == ImagePositionInline &&
		imageInfo.Config.Alignment != "" {
		paragraph.Properties = &ParagraphProperties{
			Justification: &Justification{Val: string(imageInfo.Config.Alignment)},
		}
	}

	return paragraph
}

// createInlineImageDrawing creates an inline image drawing element.
func (d *Document) createInlineImageDrawing(imageInfo *ImageInfo, displayWidth, displayHeight int64, altText, title string) *DrawingElement {
	return &DrawingElement{
		Inline: &InlineDrawing{
			DistT: "0",
			DistB: "0",
			DistL: "0",
			DistR: "0",
			Extent: &DrawingExtent{
				Cx: fmt.Sprintf("%d", displayWidth),
				Cy: fmt.Sprintf("%d", displayHeight),
			},
			DocPr: &DrawingDocPr{
				ID:    imageInfo.ID,
				Name:  fmt.Sprintf("Image %s", imageInfo.ID),
				Descr: altText,
				Title: title,
			},
			Graphic: d.createImageGraphic(imageInfo, displayWidth, displayHeight, altText, title),
		},
	}
}

// createFloatingImageDrawing creates a floating image drawing element.
func (d *Document) createFloatingImageDrawing(imageInfo *ImageInfo, displayWidth, displayHeight int64, altText, title string) *DrawingElement {
	config := imageInfo.Config

	// Calculate distances in EMU units
	distT := "0"
	distB := "0"
	distL := "0"
	distR := "0"

	if config.OffsetX > 0 {
		distL = fmt.Sprintf("%.0f", config.OffsetX*36000) // mm to EMU
		distR = fmt.Sprintf("%.0f", config.OffsetX*36000)
	}
	if config.OffsetY > 0 {
		distT = fmt.Sprintf("%.0f", config.OffsetY*36000)
		distB = fmt.Sprintf("%.0f", config.OffsetY*36000)
	}

	anchor := &AnchorDrawing{
		DistT:          distT,
		DistB:          distB,
		DistL:          distL,
		DistR:          distR,
		SimplePos:      "0",
		RelativeHeight: "251658240",
		BehindDoc:      "0",
		Locked:         "0",
		LayoutInCell:   "1",
		AllowOverlap:   "1",
		SimplePosition: &SimplePosition{
			X: "0",
			Y: "0",
		},
		Extent: &DrawingExtent{
			Cx: fmt.Sprintf("%d", displayWidth),
			Cy: fmt.Sprintf("%d", displayHeight),
		},
		EffectExtent: &EffectExtent{
			L: "0",
			T: "0",
			R: "0",
			B: "0",
		},
		DocPr: &DrawingDocPr{
			ID:    imageInfo.ID,
			Name:  fmt.Sprintf("Image %s", imageInfo.ID),
			Descr: altText,
			Title: title,
		},
		CNvGraphicFramePr: &CNvGraphicFramePr{
			GraphicFrameLocks: &GraphicFrameLocks{
				Xmlns:          "http://schemas.openxmlformats.org/drawingml/2006/main",
				NoChangeAspect: "1",
			},
		},
		Graphic: d.createImageGraphic(imageInfo, displayWidth, displayHeight, altText, title),
	}

	// Set position
	d.setFloatingImagePosition(anchor, config)

	// Set text wrapping
	d.setFloatingImageWrap(anchor, config)

	return &DrawingElement{
		Anchor: anchor,
	}
}

// setFloatingImagePosition sets the position of a floating image.
func (d *Document) setFloatingImagePosition(anchor *AnchorDrawing, config *ImageConfig) {
	var alignValue string
	switch config.Position {
	case ImagePositionFloatLeft:
		alignValue = "left"
	case ImagePositionFloatRight:
		alignValue = "right"
	default:
		alignValue = "center"
	}
	anchor.PositionH = &HorizontalPosition{
		RelativeFrom: "margin",
		Align:        &PosAlign{Value: alignValue},
	}

	// Set vertical position to top alignment
	anchor.PositionV = &VerticalPosition{
		RelativeFrom: "margin",
		Align: &PosAlign{
			Value: "top",
		},
	}

	// If there is an offset, use offset instead of alignment
	if config.OffsetX != 0 || config.OffsetY != 0 {
		if config.OffsetX != 0 {
			anchor.PositionH.Align = nil
			anchor.PositionH.PosOffset = &PosOffset{
				Value: fmt.Sprintf("%.0f", config.OffsetX*36000),
			}
		}
		if config.OffsetY != 0 {
			anchor.PositionV.Align = nil
			anchor.PositionV.PosOffset = &PosOffset{
				Value: fmt.Sprintf("%.0f", config.OffsetY*36000),
			}
		}
	}
}

// setFloatingImageWrap sets text wrapping for a floating image.
func (d *Document) setFloatingImageWrap(anchor *AnchorDrawing, config *ImageConfig) {
	// Calculate wrapping distances
	wrapDistL := "114300" // Default 3mm
	wrapDistR := "114300"
	wrapDistT := "0"
	wrapDistB := "0"

	if config.OffsetX > 0 {
		wrapDistL = fmt.Sprintf("%.0f", config.OffsetX*36000)
		wrapDistR = fmt.Sprintf("%.0f", config.OffsetX*36000)
	}
	if config.OffsetY > 0 {
		wrapDistT = fmt.Sprintf("%.0f", config.OffsetY*36000)
		wrapDistB = fmt.Sprintf("%.0f", config.OffsetY*36000)
	}

	// wrapTextForPosition determines wrap text value based on image position.
	wrapTextForPosition := func() string {
		switch config.Position {
		case ImagePositionFloatLeft:
			return "right"
		case ImagePositionFloatRight:
			return "left"
		default:
			return "bothSides"
		}
	}

	switch config.WrapText {
	case ImageWrapNone:
		anchor.WrapNone = &WrapNone{}
	case ImageWrapSquare:
		anchor.WrapSquare = &WrapSquare{
			WrapText: wrapTextForPosition(),
			DistT:    wrapDistT,
			DistB:    wrapDistB,
			DistL:    wrapDistL,
			DistR:    wrapDistR,
		}
	case ImageWrapTight:
		anchor.WrapTight = &WrapTight{
			WrapText:    wrapTextForPosition(),
			DistL:       wrapDistL,
			DistR:       wrapDistR,
			WrapPolygon: d.createDefaultWrapPolygon(),
		}
	case ImageWrapTopAndBottom:
		anchor.WrapTopAndBottom = &WrapTopAndBottom{
			DistT: wrapDistT,
			DistB: wrapDistB,
		}
	default:
		// Default to square wrapping
		anchor.WrapSquare = &WrapSquare{
			WrapText: wrapTextForPosition(),
			DistT:    wrapDistT,
			DistB:    wrapDistB,
			DistL:    wrapDistL,
			DistR:    wrapDistR,
		}
	}
}

// createDefaultWrapPolygon creates a default wrapping polygon.
// This method creates a rectangular wrapping path conforming to the OpenXML specification.
func (d *Document) createDefaultWrapPolygon() *WrapPolygon {
	return &WrapPolygon{
		Start: &PolygonStart{
			X: "0",
			Y: "0",
		},
		LineTo: []PolygonLineTo{
			{X: "0", Y: "21600"},
			{X: "21600", Y: "21600"},
			{X: "21600", Y: "0"},
			{X: "0", Y: "0"},
		},
	}
}

// createImageGraphic creates an image graphic element.
func (d *Document) createImageGraphic(imageInfo *ImageInfo, displayWidth, displayHeight int64, altText, title string) *DrawingGraphic {
	return &DrawingGraphic{
		Xmlns: "http://schemas.openxmlformats.org/drawingml/2006/main",
		GraphicData: &GraphicData{
			URI: "http://schemas.openxmlformats.org/drawingml/2006/picture",
			Pic: &PicElement{
				Xmlns: "http://schemas.openxmlformats.org/drawingml/2006/picture",
				NvPicPr: &NvPicPr{
					CNvPr: &CNvPr{
						ID:    imageInfo.ID,
						Name:  fmt.Sprintf("Image %s", imageInfo.ID),
						Descr: altText,
						Title: title,
					},
					CNvPicPr: &CNvPicPr{
						PicLocks: &PicLocks{
							NoChangeAspect: "1",
						},
					},
				},
				BlipFill: &BlipFill{
					Blip: &Blip{
						Embed: imageInfo.RelationID,
					},
					Stretch: &Stretch{
						FillRect: &FillRect{},
					},
				},
				SpPr: &SpPr{
					Xfrm: &Xfrm{
						Off: &Off{
							X: "0",
							Y: "0",
						},
						Ext: &Ext{
							Cx: fmt.Sprintf("%d", displayWidth),
							Cy: fmt.Sprintf("%d", displayHeight),
						},
					},
					PrstGeom: &PrstGeom{
						Prst:  "rect",
						AvLst: &AvLst{},
					},
				},
			},
		},
	}
}

// calculateDisplaySize calculates the image display size in EMU units.
func (d *Document) calculateDisplaySize(imageInfo *ImageInfo) (int64, int64) {
	config := imageInfo.Config
	originalWidth := int64(imageInfo.Width)
	originalHeight := int64(imageInfo.Height)

	// Default to original size (96 DPI)
	// 1 pixel = 9525 EMU (at 96 DPI)
	displayWidth := originalWidth * 9525
	displayHeight := originalHeight * 9525

	if config != nil && config.Size != nil {
		switch {
		case config.Size.Width > 0 && config.Size.Height > 0:
			// User specified explicit dimensions
			displayWidth = int64(config.Size.Width * 36000)   // mm to EMU
			displayHeight = int64(config.Size.Height * 36000) // mm to EMU
		case config.Size.Width > 0 && config.Size.KeepAspectRatio:
			// Only width specified, keep aspect ratio
			displayWidth = int64(config.Size.Width * 36000)
			ratio := float64(originalHeight) / float64(originalWidth)
			displayHeight = int64(float64(displayWidth) * ratio)
		case config.Size.Height > 0 && config.Size.KeepAspectRatio:
			// Only height specified, keep aspect ratio
			displayHeight = int64(config.Size.Height * 36000)
			ratio := float64(originalWidth) / float64(originalHeight)
			displayWidth = int64(float64(displayHeight) * ratio)
		}
	}

	return displayWidth, displayHeight
}

// detectImageFormat detects the image format from binary data.
func detectImageFormat(data []byte) (ImageFormat, error) {
	if len(data) < 3 {
		return "", fmt.Errorf("image data too short")
	}

	// Detect PNG
	if len(data) >= 8 && bytes.Equal(data[:8], []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A}) {
		return ImageFormatPNG, nil
	}

	// Detect JPEG
	if len(data) >= 3 && bytes.Equal(data[:3], []byte{0xFF, 0xD8, 0xFF}) {
		return ImageFormatJPEG, nil
	}

	// Detect GIF
	if len(data) >= 6 && (bytes.Equal(data[:6], []byte("GIF87a")) || bytes.Equal(data[:6], []byte("GIF89a"))) {
		return ImageFormatGIF, nil
	}

	return "", fmt.Errorf("unsupported image format")
}

// getImageDimensions gets the dimensions of an image.
func getImageDimensions(data []byte, format ImageFormat) (int, int, error) {
	reader := bytes.NewReader(data)

	var img image.Image
	var err error

	switch format {
	case ImageFormatPNG:
		img, err = png.Decode(reader)
	case ImageFormatJPEG:
		img, err = jpeg.Decode(reader)
	case ImageFormatGIF:
		img, err = gif.Decode(reader)
	default:
		return 0, 0, fmt.Errorf("unsupported image format: %s", format)
	}

	if err != nil {
		return 0, 0, fmt.Errorf("failed to decode image: %w", err)
	}

	bounds := img.Bounds()
	return bounds.Dx(), bounds.Dy(), nil
}

// addImageContentType adds the content type for an image format.
func (d *Document) addImageContentType(format ImageFormat) {
	if d.contentTypes == nil {
		d.contentTypes = &ContentTypes{
			Xmlns:     "http://schemas.openxmlformats.org/package/2006/content-types",
			Defaults:  []Default{},
			Overrides: []Override{},
		}
	}

	var extension, contentType string
	switch format {
	case ImageFormatPNG:
		extension = "png"
		contentType = "image/png"
	case ImageFormatJPEG:
		extension = "jpeg"
		contentType = "image/jpeg"
	case ImageFormatGIF:
		extension = "gif"
		contentType = "image/gif"
	default:
		return
	}

	// Check if the same default type already exists
	for _, def := range d.contentTypes.Defaults {
		if def.Extension == extension {
			return
		}
	}

	// Add default content type
	d.contentTypes.Defaults = append(d.contentTypes.Defaults, Default{
		Extension:   extension,
		ContentType: contentType,
	})
}

// ResizeImage resizes an image.
func (d *Document) ResizeImage(imageInfo *ImageInfo, size *ImageSize) error {
	if imageInfo == nil {
		return fmt.Errorf("image info cannot be nil")
	}

	if imageInfo.Config == nil {
		imageInfo.Config = &ImageConfig{}
	}

	imageInfo.Config.Size = size

	// If the image has already been added to the document, it needs to be regenerated
	// Note: this is a simplified implementation; real-world use may require a more complex update mechanism
	return nil
}

// SetImagePosition sets the position of an image.
func (d *Document) SetImagePosition(imageInfo *ImageInfo, position ImagePosition, offsetX, offsetY float64) error {
	if imageInfo == nil {
		return fmt.Errorf("image info cannot be nil")
	}

	if imageInfo.Config == nil {
		imageInfo.Config = &ImageConfig{}
	}

	imageInfo.Config.Position = position
	imageInfo.Config.OffsetX = offsetX
	imageInfo.Config.OffsetY = offsetY

	// If the position changes (from inline to float or vice versa), the Drawing element may need to be regenerated
	// Note: this is a simplified implementation; real-world use may require a more complex update mechanism
	return nil
}

// SetImageWrapText sets the text wrapping style for an image.
func (d *Document) SetImageWrapText(imageInfo *ImageInfo, wrapText ImageWrapText) error {
	if imageInfo == nil {
		return fmt.Errorf("image info cannot be nil")
	}

	if imageInfo.Config == nil {
		imageInfo.Config = &ImageConfig{}
	}

	imageInfo.Config.WrapText = wrapText
	return nil
}

// SetImageAltText sets the alt text for an image.
func (d *Document) SetImageAltText(imageInfo *ImageInfo, altText string) error {
	if imageInfo == nil {
		return fmt.Errorf("image info cannot be nil")
	}

	if imageInfo.Config == nil {
		imageInfo.Config = &ImageConfig{}
	}

	imageInfo.Config.AltText = altText
	return nil
}

// SetImageTitle sets the title for an image.
func (d *Document) SetImageTitle(imageInfo *ImageInfo, title string) error {
	if imageInfo == nil {
		return fmt.Errorf("image info cannot be nil")
	}

	if imageInfo.Config == nil {
		imageInfo.Config = &ImageConfig{}
	}

	imageInfo.Config.Title = title
	return nil
}

// AddCellImage adds an image to a table cell.
//
// This method adds an image to a table cell, supporting both file path and binary data.
// Since images require document-level resource relationship management, this method must be called on a Document.
//
// Parameters:
//   - table: target table
//   - row: row index (0-based)
//   - col: column index (0-based)
//   - config: cell image configuration
//
// Returns:
//   - *ImageInfo: information about the added image
//   - error: an error if the operation fails
//
// Example:
//
//	table, _ := doc.AddTable(&document.TableConfig{Rows: 2, Cols: 2, Width: 6000})
//	imageConfig := &document.CellImageConfig{
//		FilePath: "logo.png",
//		Width:    50, // 50mm width
//		KeepAspectRatio: true,
//	}
//	imageInfo, err := doc.AddCellImage(table, 0, 0, imageConfig)
func (d *Document) AddCellImage(table *Table, row, col int, config *CellImageConfig) (*ImageInfo, error) {
	if table == nil {
		return nil, fmt.Errorf("table cannot be nil")
	}

	cell, err := table.GetCell(row, col)
	if err != nil {
		return nil, err
	}

	var imageData []byte
	var format ImageFormat
	var width, height int

	// Get image from file or data
	switch {
	case config.FilePath != "":
		// Read image from file
		imageData, err = os.ReadFile(config.FilePath)
		if err != nil {
			Errorf("failed to read image file %s: %v", config.FilePath, err)
			return nil, fmt.Errorf("failed to read image file: %w", err)
		}

		// Detect image format
		format, err = detectImageFormat(imageData)
		if err != nil {
			Errorf("failed to detect image format %s: %v", config.FilePath, err)
			return nil, fmt.Errorf("failed to detect image format: %w", err)
		}

		// Get image dimensions
		width, height, err = getImageDimensions(imageData, format)
		if err != nil {
			Errorf("failed to get image dimensions %s: %v", config.FilePath, err)
			return nil, fmt.Errorf("failed to get image dimensions: %w", err)
		}
	case len(config.Data) > 0:
		// Use the provided binary data
		imageData = config.Data

		if config.Format == "" {
			// Detect image format
			format, err = detectImageFormat(imageData)
			if err != nil {
				return nil, fmt.Errorf("failed to detect image format: %w", err)
			}
		} else {
			format = config.Format
		}

		// Get image dimensions
		width, height, err = getImageDimensions(imageData, format)
		if err != nil {
			return nil, fmt.Errorf("failed to get image dimensions: %w", err)
		}
	default:
		return nil, fmt.Errorf("either image file path or binary data must be provided")
	}

	// Create image configuration
	imageConfig := &ImageConfig{
		Position:  ImagePositionInline,
		Alignment: AlignCenter,
		AltText:   config.AltText,
		Title:     config.Title,
	}

	if config.Width > 0 || config.Height > 0 {
		imageConfig.Size = &ImageSize{
			Width:           config.Width,
			Height:          config.Height,
			KeepAspectRatio: config.KeepAspectRatio,
		}
	}

	// Use the Document method to add image resources without adding to the document body
	fileName := "cell_image.png"
	if config.FilePath != "" {
		fileName = config.FilePath
	}

	imageInfo, err := d.AddImageFromDataWithoutElement(imageData, fileName, format, width, height, imageConfig)
	if err != nil {
		return nil, err
	}

	// Create a paragraph containing the image and add it to the cell
	paragraph := d.createImageParagraph(imageInfo)
	cell.Paragraphs = append(cell.Paragraphs, *paragraph)

	InfoMsgf(MsgImageAddedToCell, row, col, imageInfo.ID)
	return imageInfo, nil
}

// AddCellImageFromFile adds an image to a table cell from a file (convenience method).
//
// This method is a convenience wrapper around AddCellImage that adds an image directly from a file path.
//
// Parameters:
//   - table: target table
//   - row: row index (0-based)
//   - col: column index (0-based)
//   - filePath: image file path
//   - widthMM: image width in millimeters, 0 to use original size
//
// Returns:
//   - *ImageInfo: information about the added image
//   - error: an error if the operation fails
func (d *Document) AddCellImageFromFile(table *Table, row, col int, filePath string, widthMM float64) (*ImageInfo, error) {
	return d.AddCellImage(table, row, col, &CellImageConfig{
		FilePath:        filePath,
		Width:           widthMM,
		KeepAspectRatio: true,
	})
}

// AddCellImageFromData adds an image to a table cell from binary data (convenience method).
//
// This method is a convenience wrapper around AddCellImage that adds an image directly from binary data.
//
// Parameters:
//   - table: target table
//   - row: row index (0-based)
//   - col: column index (0-based)
//   - data: image binary data
//   - widthMM: image width in millimeters, 0 to use original size
//
// Returns:
//   - *ImageInfo: information about the added image
//   - error: an error if the operation fails
func (d *Document) AddCellImageFromData(table *Table, row, col int, data []byte, widthMM float64) (*ImageInfo, error) {
	return d.AddCellImage(table, row, col, &CellImageConfig{
		Data:            data,
		Width:           widthMM,
		KeepAspectRatio: true,
	})
}

// SetImageAlignment sets the alignment of an image.
//
// This method sets the alignment for inline images (ImagePositionInline).
// For floating images, use SetImagePosition instead.
//
// The alignment parameter specifies the alignment type and supports the following values:
//   - AlignLeft: left alignment
//   - AlignCenter: center alignment
//   - AlignRight: right alignment
//   - AlignJustify: justified alignment
//
// Example:
//
//	imageInfo, err := doc.AddImageFromFile("image.png", nil)
//	if err != nil {
//		return err
//	}
//	err = doc.SetImageAlignment(imageInfo, document.AlignCenter)
func (d *Document) SetImageAlignment(imageInfo *ImageInfo, alignment AlignmentType) error {
	if imageInfo == nil {
		return fmt.Errorf("image info cannot be nil")
	}

	if imageInfo.Config == nil {
		imageInfo.Config = &ImageConfig{}
	}

	// Update configuration
	imageInfo.Config.Alignment = alignment

	// Find the paragraph containing this image and update its alignment
	for _, element := range d.Body.Elements {
		if paragraph, ok := element.(*Paragraph); ok {
			// Check if the paragraph contains the specified image
			for _, run := range paragraph.Runs {
				if run.Drawing != nil && run.Drawing.Inline != nil {
					// Check if docPr ID matches
					if run.Drawing.Inline.DocPr != nil && run.Drawing.Inline.DocPr.ID == imageInfo.ID {
						// Update paragraph alignment
						if paragraph.Properties == nil {
							paragraph.Properties = &ParagraphProperties{}
						}
						paragraph.Properties.Justification = &Justification{Val: string(alignment)}
						return nil
					}
				}
			}
		}
	}

	// Paragraph containing the image not found: keep the config updated and return nil to support setting alignment before the image is inserted into a paragraph.
	DebugMsgf(MsgImageParagraphNotFound, imageInfo.ID)
	return nil
}

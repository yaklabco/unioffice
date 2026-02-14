package document

import (
	"errors"
	"fmt"
	"image"
	"math/rand"

	unioffice "github.com/yaklabco/unioffice/v2"
	"github.com/yaklabco/unioffice/v2/common"
	"github.com/yaklabco/unioffice/v2/common/tempstorage"
	"github.com/yaklabco/unioffice/v2/measurement"
	"github.com/yaklabco/unioffice/v2/schema/soo/dml"
	"github.com/yaklabco/unioffice/v2/schema/soo/dml/picture"
	"github.com/yaklabco/unioffice/v2/schema/soo/wml"
)

// AddImageSVG registers an SVG image and its PNG fallback with the document.
// Both images are added to the document's relationship list. The SVG content
// type (image/svg+xml) is registered in the content types.
// Returns the PNG ImageRef, SVG ImageRef, and any error.
func (d *Document) AddImageSVG(svgData []byte, pngFallback []byte, width, height int) (common.ImageRef, common.ImageRef, error) {
	var empty common.ImageRef
	if len(svgData) == 0 || len(pngFallback) == 0 {
		return empty, empty, errors.New("both SVG and PNG data are required")
	}
	if width <= 0 || height <= 0 {
		return empty, empty, errors.New("width and height must be positive")
	}

	// Register PNG image
	pngImage := common.Image{
		Size:   image.Point{X: width, Y: height},
		Format: "png",
		Data:   &pngFallback,
	}
	pngRef, err := d.AddImage(pngImage)
	if err != nil {
		return empty, empty, fmt.Errorf("adding PNG image: %w", err)
	}

	// Register SVG image
	svgImage := common.Image{
		Size:   image.Point{X: width, Y: height},
		Format: "svg",
		Data:   &svgData,
	}
	svgRef := common.MakeImageRef(svgImage, &d.DocBase, d._ead)
	if svgImage.Path != "" {
		if err := tempstorage.Add(svgImage.Path); err != nil {
			return empty, empty, err
		}
	}
	d.Images = append(d.Images, svgRef)
	target := fmt.Sprintf("media/image%d.%s", len(d.Images), svgImage.Format)
	rel := d._ead.AddRelationship(target, unioffice.ImageType)
	d.ContentTypes.EnsureDefault("svg", "image/svg+xml")
	svgRef.SetRelID(rel.X().IdAttr)
	svgRef.SetTarget(target)

	return pngRef, svgRef, nil
}

// AddDrawingInlineSVG adds an SVG image with PNG fallback as an inline drawing
// using mc:AlternateContent. Office versions supporting asvg render the SVG;
// older versions display the PNG fallback.
func (r Run) AddDrawingInlineSVG(svgImg, pngImg common.ImageRef) (InlineDrawing, error) {
	// Build the Choice drawing (PNG blip + SVG extension in ExtLst)
	choiceDrawing := wml.NewCT_Drawing()
	choiceInline := wml.NewWdInline()
	inlineDraw := InlineDrawing{r._gdedf, choiceInline}

	choiceInline.CNvGraphicFramePr = dml.NewCT_NonVisualGraphicFrameProperties()
	choiceDrawing.DrawingChoice = append(choiceDrawing.DrawingChoice, &wml.CT_DrawingChoice{Inline: choiceInline})

	choiceInline.Graphic = dml.NewGraphic()
	choiceInline.Graphic.GraphicData = dml.NewCT_GraphicalObjectData()
	choiceInline.Graphic.GraphicData.UriAttr = "http://schemas.openxmlformats.org/drawingml/2006/picture"

	choiceInline.DistTAttr = unioffice.Uint32(0)
	choiceInline.DistLAttr = unioffice.Uint32(0)
	choiceInline.DistBAttr = unioffice.Uint32(0)
	choiceInline.DistRAttr = unioffice.Uint32(0)
	choiceInline.Extent.CxAttr = int64(float64(pngImg.Size().X*measurement.Pixel72) / measurement.EMU)
	choiceInline.Extent.CyAttr = int64(float64(pngImg.Size().Y*measurement.Pixel72) / measurement.EMU)

	docPrID := 0x7FFFFFFF & rand.Uint32()
	choiceInline.DocPr.IdAttr = docPrID

	// Build the pic:pic element with blip + SVG extension
	choicePic := picture.NewPic()
	choicePic.NvPicPr.CNvPr.IdAttr = docPrID

	pngRelID := pngImg.RelID()
	if pngRelID == "" {
		return inlineDraw, errors.New("couldn't find reference to PNG image within document relations")
	}
	svgRelID := svgImg.RelID()
	if svgRelID == "" {
		return inlineDraw, errors.New("couldn't find reference to SVG image within document relations")
	}

	choiceInline.Graphic.GraphicData.Any = append(choiceInline.Graphic.GraphicData.Any, choicePic)

	// BlipFill with PNG as the main blip, SVG in ExtLst
	choicePic.BlipFill = dml.NewCT_BlipFillProperties()
	choicePic.BlipFill.Blip = dml.NewCT_Blip()
	choicePic.BlipFill.Blip.EmbedAttr = &pngRelID

	// Add SVG extension to the blip's ExtLst
	choicePic.BlipFill.Blip.ExtLst = dml.NewCT_OfficeArtExtensionList()
	svgExt := dml.NewCT_OfficeArtExtension()
	svgExt.UriAttr = dml.SVGBlipURI
	svgExt.Any = append(svgExt.Any, &dml.SVGBlip{EmbedAttr: svgRelID})
	choicePic.BlipFill.Blip.ExtLst.Ext = append(choicePic.BlipFill.Blip.ExtLst.Ext, svgExt)

	choicePic.BlipFill.FillModePropertiesChoice.Stretch = dml.NewCT_StretchInfoProperties()
	choicePic.BlipFill.FillModePropertiesChoice.Stretch.FillRect = dml.NewCT_RelativeRect()

	// Shape properties with dimensions
	choicePic.SpPr = dml.NewCT_ShapeProperties()
	choicePic.SpPr.Xfrm = dml.NewCT_Transform2D()
	choicePic.SpPr.Xfrm.Off = dml.NewCT_Point2D()
	choicePic.SpPr.Xfrm.Off.XAttr.ST_CoordinateUnqualified = unioffice.Int64(0)
	choicePic.SpPr.Xfrm.Off.YAttr.ST_CoordinateUnqualified = unioffice.Int64(0)
	choicePic.SpPr.Xfrm.Ext = dml.NewCT_PositiveSize2D()
	choicePic.SpPr.Xfrm.Ext.CxAttr = int64(pngImg.Size().X * measurement.Point)
	choicePic.SpPr.Xfrm.Ext.CyAttr = int64(pngImg.Size().Y * measurement.Point)
	choicePic.SpPr.GeometryChoice.PrstGeom = dml.NewCT_PresetGeometry2D()
	choicePic.SpPr.GeometryChoice.PrstGeom.PrstAttr = dml.ST_ShapeTypeRect

	// Build the Fallback drawing (PNG only, no SVG extension)
	fallbackDrawing := wml.NewCT_Drawing()
	fallbackInline := wml.NewWdInline()
	fallbackInline.CNvGraphicFramePr = dml.NewCT_NonVisualGraphicFrameProperties()
	fallbackDrawing.DrawingChoice = append(fallbackDrawing.DrawingChoice, &wml.CT_DrawingChoice{Inline: fallbackInline})

	fallbackInline.Graphic = dml.NewGraphic()
	fallbackInline.Graphic.GraphicData = dml.NewCT_GraphicalObjectData()
	fallbackInline.Graphic.GraphicData.UriAttr = "http://schemas.openxmlformats.org/drawingml/2006/picture"
	fallbackInline.DistTAttr = unioffice.Uint32(0)
	fallbackInline.DistLAttr = unioffice.Uint32(0)
	fallbackInline.DistBAttr = unioffice.Uint32(0)
	fallbackInline.DistRAttr = unioffice.Uint32(0)
	fallbackInline.Extent.CxAttr = choiceInline.Extent.CxAttr
	fallbackInline.Extent.CyAttr = choiceInline.Extent.CyAttr
	fallbackInline.DocPr.IdAttr = 0x7FFFFFFF & rand.Uint32()

	fallbackPic := picture.NewPic()
	fallbackPic.NvPicPr.CNvPr.IdAttr = fallbackInline.DocPr.IdAttr
	fallbackInline.Graphic.GraphicData.Any = append(fallbackInline.Graphic.GraphicData.Any, fallbackPic)
	fallbackPic.BlipFill = dml.NewCT_BlipFillProperties()
	fallbackPic.BlipFill.Blip = dml.NewCT_Blip()
	fallbackPic.BlipFill.Blip.EmbedAttr = &pngRelID
	fallbackPic.BlipFill.FillModePropertiesChoice.Stretch = dml.NewCT_StretchInfoProperties()
	fallbackPic.BlipFill.FillModePropertiesChoice.Stretch.FillRect = dml.NewCT_RelativeRect()
	fallbackPic.SpPr = dml.NewCT_ShapeProperties()
	fallbackPic.SpPr.Xfrm = dml.NewCT_Transform2D()
	fallbackPic.SpPr.Xfrm.Off = dml.NewCT_Point2D()
	fallbackPic.SpPr.Xfrm.Off.XAttr.ST_CoordinateUnqualified = unioffice.Int64(0)
	fallbackPic.SpPr.Xfrm.Off.YAttr.ST_CoordinateUnqualified = unioffice.Int64(0)
	fallbackPic.SpPr.Xfrm.Ext = dml.NewCT_PositiveSize2D()
	fallbackPic.SpPr.Xfrm.Ext.CxAttr = choicePic.SpPr.Xfrm.Ext.CxAttr
	fallbackPic.SpPr.Xfrm.Ext.CyAttr = choicePic.SpPr.Xfrm.Ext.CyAttr
	fallbackPic.SpPr.GeometryChoice.PrstGeom = dml.NewCT_PresetGeometry2D()
	fallbackPic.SpPr.GeometryChoice.PrstGeom.PrstAttr = dml.ST_ShapeTypeRect

	// Build mc:Choice with the Choice drawing
	choice := wml.NewAC_ChoiceRun()
	choice.SetRequires("asvg")
	choice.Drawing = choiceDrawing

	// Build AlternateContentSVGRun
	acRun := &wml.AlternateContentSVGRun{
		Choice: choice,
		Fallback: &wml.FallbackDrawing{
			Drawing: fallbackDrawing,
		},
	}

	// Append to run's Extra slice
	r._bbdb.Extra = append(r._bbdb.Extra, acRun)

	return inlineDraw, nil
}

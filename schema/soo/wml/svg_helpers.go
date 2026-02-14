package wml

import "encoding/xml"

// SetRequires sets the mc:Choice Requires attribute value.
// This is needed because the _egdddc field on AC_ChoiceRun is private.
func (a *AC_ChoiceRun) SetRequires(requires string) {
	a._egdddc = requires
}

// FallbackDrawing wraps a CT_Drawing for use as mc:Fallback content.
type FallbackDrawing struct {
	Drawing *CT_Drawing
}

// MarshalXML marshals the FallbackDrawing as <mc:Fallback><w:drawing>...</w:drawing></mc:Fallback>.
func (f *FallbackDrawing) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	fb := xml.StartElement{Name: xml.Name{Local: "mc:Fallback"}}
	if err := e.EncodeToken(fb); err != nil {
		return err
	}
	if f.Drawing != nil {
		dStart := xml.StartElement{Name: xml.Name{Local: "w:drawing"}}
		if err := e.EncodeElement(f.Drawing, dStart); err != nil {
			return err
		}
	}
	return e.EncodeToken(xml.EndElement{Name: fb.Name})
}

// UnmarshalXML unmarshals the FallbackDrawing from XML.
func (f *FallbackDrawing) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	return d.Skip()
}

// AlternateContentSVGRun is an SVG-aware version of AlternateContentRun
// that includes the asvg namespace declaration required for SVG embedding.
type AlternateContentSVGRun struct {
	Choice   *AC_ChoiceRun
	Fallback *FallbackDrawing
}

// acElementName is the fully-qualified mc:AlternateContent element name
// matching what CT_R.MarshalXML uses for the standard AlternateContentRun.
var acElementName = xml.Name{
	Space: "http://schemas.openxmlformats.org/markup-compatibility/2006",
	Local: "mc:AlternateContent",
}

// MarshalXML marshals AlternateContentSVGRun as mc:AlternateContent with
// all required namespace declarations including asvg.
func (a *AlternateContentSVGRun) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// When called from CT_R.Extra's else branch, start has an empty name.
	// Provide the correct mc:AlternateContent name.
	name := start.Name
	if name.Local == "" {
		name = acElementName
	}

	acStart := xml.StartElement{
		Name: name,
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xmlns:wpg"}, Value: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"},
			{Name: xml.Name{Local: "xmlns:mc"}, Value: "http://schemas.openxmlformats.org/markup-compatibility/2006"},
			{Name: xml.Name{Local: "xmlns:w"}, Value: "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
			{Name: xml.Name{Local: "xmlns:wp"}, Value: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"},
			{Name: xml.Name{Local: "xmlns:wp14"}, Value: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"},
			{Name: xml.Name{Local: "xmlns:a"}, Value: "http://schemas.openxmlformats.org/drawingml/2006/main"},
			{Name: xml.Name{Local: "xmlns:pic"}, Value: "http://schemas.openxmlformats.org/drawingml/2006/picture"},
			{Name: xml.Name{Local: "xmlns:r"}, Value: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"},
			{Name: xml.Name{Local: "xmlns:wps"}, Value: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"},
			{Name: xml.Name{Local: "xmlns:v"}, Value: "urn:schemas-microsoft-com:vml"},
			{Name: xml.Name{Local: "xmlns:w14"}, Value: "http://schemas.microsoft.com/office/word/2010/wordml"},
			{Name: xml.Name{Local: "xmlns:o"}, Value: "urn:schemas-microsoft-com:office:office"},
			{Name: xml.Name{Local: "xmlns:w10"}, Value: "urn:schemas-microsoft-com:office:word"},
			{Name: xml.Name{Local: "xmlns:asvg"}, Value: "http://schemas.microsoft.com/office/drawing/2016/SVG/main"},
			{Name: xml.Name{Local: "mc:Ignorable"}, Value: "wp14 w14 w10 asvg"},
		},
	}
	if err := e.EncodeToken(acStart); err != nil {
		return err
	}

	if a.Choice != nil {
		// Use EncodeElement which writes the start tag, content, and end tag.
		choiceStart := xml.StartElement{
			Name: xml.Name{Local: "mc:Choice"},
			Attr: []xml.Attr{
				{Name: xml.Name{Local: "Requires"}, Value: "asvg"},
			},
		}
		if err := e.EncodeElement(a.Choice, choiceStart); err != nil {
			return err
		}
	}

	if a.Fallback != nil {
		if err := a.Fallback.MarshalXML(e, xml.StartElement{}); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: acStart.Name})
}

// UnmarshalXML unmarshals the AlternateContentSVGRun from XML.
func (a *AlternateContentSVGRun) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	return d.Skip()
}

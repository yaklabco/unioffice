package dml

import "encoding/xml"

// SVGBlipURI is the OOXML extension URI for SVG blip support.
const SVGBlipURI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"

// SVGBlipNS is the namespace for the asvg:svgBlip element.
const SVGBlipNS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"

// SVGBlip represents an asvg:svgBlip element referencing an SVG image
// inside a CT_Blip's ExtLst.
type SVGBlip struct {
	EmbedAttr string // relationship ID for the SVG
}

// MarshalXML marshals the SVGBlip as <asvg:svgBlip xmlns:asvg="..." r:embed="..."/>.
func (s *SVGBlip) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	el := xml.StartElement{
		Name: xml.Name{Local: "asvg:svgBlip"},
		Attr: []xml.Attr{
			{Name: xml.Name{Local: "xmlns:asvg"}, Value: SVGBlipNS},
			{Name: xml.Name{Local: "r:embed"}, Value: s.EmbedAttr},
		},
	}
	if err := e.EncodeToken(el); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: el.Name})
}

// UnmarshalXML unmarshals the SVGBlip from XML.
func (s *SVGBlip) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "embed" {
			s.EmbedAttr = attr.Value
		}
	}
	return d.Skip()
}

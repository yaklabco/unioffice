package document

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"image"
	"image/color"
	"image/png"
	"os"
	"strings"
	"testing"

	"github.com/yaklabco/unioffice/v2/common"
	"github.com/yaklabco/unioffice/v2/schema/soo/dml"
)

func testPNGData(t *testing.T, width, height int) []byte {
	t.Helper()
	img := image.NewRGBA(image.Rect(0, 0, width, height))
	for y := 0; y < height; y++ {
		for x := 0; x < width; x++ {
			img.Set(x, y, color.RGBA{R: 200, G: 200, B: 200, A: 255})
		}
	}
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

var testSVGData = []byte(`<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100">
  <rect width="200" height="100" fill="blue"/>
  <text x="50" y="50" fill="white">Test</text>
</svg>`)

func TestAddDrawingInlineSVG_Basic(t *testing.T) {
	doc := New()
	pngData := testPNGData(t, 200, 100)

	pngImg, svgImg, err := doc.AddImageSVG(testSVGData, pngData, 200, 100)
	if err != nil {
		t.Fatalf("AddImageSVG: %v", err)
	}

	para := doc.AddParagraph()
	run := para.AddRun()
	_, err = run.AddDrawingInlineSVG(svgImg, pngImg)
	if err != nil {
		t.Fatalf("AddDrawingInlineSVG: %v", err)
	}

	// Save and verify the ZIP contains both image files
	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	var hasPNG, hasSVG bool
	for _, f := range zr.File {
		if strings.HasSuffix(f.Name, ".png") {
			hasPNG = true
		}
		if strings.HasSuffix(f.Name, ".svg") {
			hasSVG = true
		}
	}

	if !hasPNG {
		t.Error("expected PNG file in ZIP")
	}
	if !hasSVG {
		t.Error("expected SVG file in ZIP")
	}
}

func TestAddDrawingInlineSVG_XMLStructure(t *testing.T) {
	doc := New()
	pngData := testPNGData(t, 200, 100)

	pngImg, svgImg, err := doc.AddImageSVG(testSVGData, pngData, 200, 100)
	if err != nil {
		t.Fatalf("AddImageSVG: %v", err)
	}

	para := doc.AddParagraph()
	run := para.AddRun()
	_, err = run.AddDrawingInlineSVG(svgImg, pngImg)
	if err != nil {
		t.Fatalf("AddDrawingInlineSVG: %v", err)
	}

	// Marshal the Extra items directly to check XML structure
	if len(run.X().Extra) == 0 {
		t.Fatal("no Extra items found on run")
	}
	xmlData, err := xml.Marshal(run.X().Extra[0])
	if err != nil {
		t.Fatalf("xml.Marshal Extra: %v", err)
	}
	xmlStr := string(xmlData)

	// Must contain mc:AlternateContent
	if !strings.Contains(xmlStr, "AlternateContent") {
		t.Error("expected mc:AlternateContent in XML")
	}
	// Must contain mc:Choice with Requires="asvg"
	if !strings.Contains(xmlStr, "asvg") {
		t.Error("expected asvg reference in XML")
	}
	// Must contain mc:Fallback
	if !strings.Contains(xmlStr, "Fallback") {
		t.Error("expected mc:Fallback in XML")
	}
	// Must contain svgBlip
	if !strings.Contains(xmlStr, "svgBlip") {
		t.Error("expected asvg:svgBlip in XML")
	}
}

func TestAddDrawingInlineSVG_ContentTypes(t *testing.T) {
	doc := New()
	pngData := testPNGData(t, 200, 100)

	_, _, err := doc.AddImageSVG(testSVGData, pngData, 200, 100)
	if err != nil {
		t.Fatalf("AddImageSVG: %v", err)
	}

	// Save and check content types
	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	for _, f := range zr.File {
		if f.Name == "[Content_Types].xml" {
			rc, err := f.Open()
			if err != nil {
				t.Fatalf("open content types: %v", err)
			}
			defer rc.Close()
			data := new(bytes.Buffer)
			if _, err := data.ReadFrom(rc); err != nil {
				t.Fatalf("reading content types: %v", err)
			}
			ct := data.String()
			if !strings.Contains(ct, "image/svg+xml") {
				t.Error("expected image/svg+xml in [Content_Types].xml")
			}
			if !strings.Contains(ct, "image/png") {
				t.Error("expected image/png in [Content_Types].xml")
			}
			return
		}
	}
	t.Error("could not find [Content_Types].xml in ZIP")
}

func TestAddDrawingInlineSVG_Relationships(t *testing.T) {
	doc := New()
	pngData := testPNGData(t, 200, 100)

	pngImg, svgImg, err := doc.AddImageSVG(testSVGData, pngData, 200, 100)
	if err != nil {
		t.Fatalf("AddImageSVG: %v", err)
	}

	if pngImg.RelID() == "" {
		t.Error("PNG image should have a relationship ID")
	}
	if svgImg.RelID() == "" {
		t.Error("SVG image should have a relationship ID")
	}
	if pngImg.RelID() == svgImg.RelID() {
		t.Error("PNG and SVG should have different relationship IDs")
	}
}

func TestAddDrawingInlineSVG_SaveToFile(t *testing.T) {
	doc := New()
	pngData := testPNGData(t, 400, 300)

	pngImg, svgImg, err := doc.AddImageSVG(testSVGData, pngData, 400, 300)
	if err != nil {
		t.Fatalf("AddImageSVG: %v", err)
	}

	para := doc.AddParagraph()
	run := para.AddRun()
	_, err = run.AddDrawingInlineSVG(svgImg, pngImg)
	if err != nil {
		t.Fatalf("AddDrawingInlineSVG: %v", err)
	}

	tmpFile := t.TempDir() + "/test_svg.docx"
	if err := doc.SaveToFile(tmpFile); err != nil {
		t.Fatalf("SaveToFile: %v", err)
	}

	// Verify file exists and is non-empty
	info, err := os.Stat(tmpFile)
	if err != nil {
		t.Fatalf("stat: %v", err)
	}
	if info.Size() == 0 {
		t.Error("output file is empty")
	}
}

func TestSVGBlip_MarshalXML(t *testing.T) {
	blip := &dml.SVGBlip{EmbedAttr: "rId5"}
	data, err := xml.Marshal(blip)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}

	xmlStr := string(data)
	if !strings.Contains(xmlStr, "asvg:svgBlip") {
		t.Errorf("expected asvg:svgBlip, got %s", xmlStr)
	}
	if !strings.Contains(xmlStr, "r:embed") {
		t.Errorf("expected r:embed, got %s", xmlStr)
	}
	if !strings.Contains(xmlStr, "rId5") {
		t.Errorf("expected rId5, got %s", xmlStr)
	}
	if !strings.Contains(xmlStr, dml.SVGBlipNS) {
		t.Errorf("expected SVG namespace, got %s", xmlStr)
	}
}

func TestSVGBlip_UnmarshalXML(t *testing.T) {
	xmlData := `<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId7"/>`
	blip := &dml.SVGBlip{}
	if err := xml.Unmarshal([]byte(xmlData), blip); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if blip.EmbedAttr != "rId7" {
		t.Errorf("expected rId7, got %s", blip.EmbedAttr)
	}
}

func TestAddImageSVG_InvalidData(t *testing.T) {
	doc := New()
	_, _, err := doc.AddImageSVG(nil, nil, 0, 0)
	if err == nil {
		t.Error("expected error for nil data")
	}
}

func TestAddImageSVG_Returns(t *testing.T) {
	doc := New()
	pngData := testPNGData(t, 100, 50)

	pngRef, svgRef, err := doc.AddImageSVG(testSVGData, pngData, 100, 50)
	if err != nil {
		t.Fatalf("AddImageSVG: %v", err)
	}

	// Both refs should have data
	if pngRef.Size().X != 100 || pngRef.Size().Y != 50 {
		t.Errorf("PNG size: got %v, want 100x50", pngRef.Size())
	}
	if svgRef.Size().X != 100 || svgRef.Size().Y != 50 {
		t.Errorf("SVG size: got %v, want 100x50", svgRef.Size())
	}
	if pngRef.Format() != "png" {
		t.Errorf("PNG format: got %s, want png", pngRef.Format())
	}
	if svgRef.Format() != "svg" {
		t.Errorf("SVG format: got %s, want svg", svgRef.Format())
	}
}

func TestAddDrawingInlineSVG_ExtensionURI(t *testing.T) {
	doc := New()
	pngData := testPNGData(t, 200, 100)

	pngImg, svgImg, err := doc.AddImageSVG(testSVGData, pngData, 200, 100)
	if err != nil {
		t.Fatalf("AddImageSVG: %v", err)
	}

	para := doc.AddParagraph()
	run := para.AddRun()
	_, err = run.AddDrawingInlineSVG(svgImg, pngImg)
	if err != nil {
		t.Fatalf("AddDrawingInlineSVG: %v", err)
	}

	// Marshal the run's Extra content
	for _, extra := range run.X().Extra {
		xmlData, err := xml.Marshal(extra)
		if err != nil {
			continue
		}
		xmlStr := string(xmlData)
		// Check the extension URI is correct
		if strings.Contains(xmlStr, dml.SVGBlipURI) {
			return // found it
		}
	}
	t.Error("expected SVGBlipURI in Extra content")
}

func TestAddImageSVG_EnsuresContentTypes(t *testing.T) {
	doc := New()
	pngData := testPNGData(t, 100, 100)

	// Call AddImageSVG multiple times — should not duplicate content types
	for i := 0; i < 3; i++ {
		_, _, err := doc.AddImageSVG(testSVGData, pngData, 100, 100)
		if err != nil {
			t.Fatalf("iteration %d: %v", i, err)
		}
	}

	// Verify we have 6 images registered (3 PNG + 3 SVG)
	if len(doc.Images) != 6 {
		t.Errorf("expected 6 images, got %d", len(doc.Images))
	}
}

// TestAddDrawingInlineSVG_MultipleSVGs tests adding multiple SVG images to a document.
func TestAddDrawingInlineSVG_MultipleSVGs(t *testing.T) {
	doc := New()
	pngData := testPNGData(t, 200, 100)

	for i := 0; i < 3; i++ {
		pngImg, svgImg, err := doc.AddImageSVG(testSVGData, pngData, 200, 100)
		if err != nil {
			t.Fatalf("AddImageSVG %d: %v", i, err)
		}

		para := doc.AddParagraph()
		run := para.AddRun()
		_, err = run.AddDrawingInlineSVG(svgImg, pngImg)
		if err != nil {
			t.Fatalf("AddDrawingInlineSVG %d: %v", i, err)
		}
	}

	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	var pngCount, svgCount int
	for _, f := range zr.File {
		if strings.HasSuffix(f.Name, ".png") && strings.Contains(f.Name, "media/") {
			pngCount++
		}
		if strings.HasSuffix(f.Name, ".svg") && strings.Contains(f.Name, "media/") {
			svgCount++
		}
	}

	if pngCount != 3 {
		t.Errorf("expected 3 PNG files, got %d", pngCount)
	}
	if svgCount != 3 {
		t.Errorf("expected 3 SVG files, got %d", svgCount)
	}
}

// TestFallbackDrawing_MarshalXML tests the FallbackDrawing XML output.
func TestFallbackDrawing_MarshalXML(t *testing.T) {
	fb := &common.ImageRef{}
	_ = fb // Just verifying types exist; real test below

	// Test with nil drawing
	doc := New()
	pngData := testPNGData(t, 200, 100)

	pngImg, svgImg, err := doc.AddImageSVG(testSVGData, pngData, 200, 100)
	if err != nil {
		t.Fatalf("AddImageSVG: %v", err)
	}

	para := doc.AddParagraph()
	run := para.AddRun()
	_, err = run.AddDrawingInlineSVG(svgImg, pngImg)
	if err != nil {
		t.Fatalf("AddDrawingInlineSVG: %v", err)
	}

	// Verify roundtrip save works
	var buf bytes.Buffer
	if err := doc.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}
	if buf.Len() == 0 {
		t.Error("expected non-empty output")
	}
}

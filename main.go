package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io/ioutil"
	"log"
	"os"
)

// DocumentEntry represents an element in the document. It can be either a figure (SVG with legend)
// or a text paragraph.
type DocumentEntry struct {
	// Type can be "figure" or "text"
	Type string

	// For a figure entry, provide the path to the SVG file and the legend text.
	SvgPath string
	Legend  string

	// For a text entry, provide the text.
	Text string
}

// addFile writes a file entry to the ZIP archive.
func addFile(zipWriter *zip.Writer, name, content string) error {
	writer, err := zipWriter.Create(name)
	if err != nil {
		return err
	}
	_, err = writer.Write([]byte(content))
	return err
}

// addDir creates a directory entry in the ZIP archive.
func addDir(zipWriter *zip.Writer, name string) error {
	header := &zip.FileHeader{
		Name:   name,
		Method: zip.Store,
	}
	header.SetMode(0755 | os.ModeDir)
	_, err := zipWriter.CreateHeader(header)
	return err
}

// xmlEscape escapes special XML characters.
func xmlEscape(s string) string {
	var buf bytes.Buffer
	xml.EscapeText(&buf, []byte(s))
	return buf.String()
}

// GenerateDocxFromEntries builds a DOCX file (ZIP archive) from a list of document entries.
// Figure entries will output an image (SVG) and its legend; text entries output a text paragraph.
func GenerateDocxFromEntries(entries []DocumentEntry, outputPath string) error {
	// First, count how many figure entries there are and read their SVG files.
	figureCount := 0
	figureData := make([][]byte, 0)
	for _, entry := range entries {
		if entry.Type == "figure" {
			figureCount++
			data, err := ioutil.ReadFile(entry.SvgPath)
			if err != nil {
				return fmt.Errorf("error reading SVG file %q: %v", entry.SvgPath, err)
			}
			figureData = append(figureData, data)
		}
	}

	// --- Calculate the full usable width in EMUs ---
	// We assume a page width of 11906 twips and left/right margins of 1440 twips.
	// Usable width = 11906 - 1440 - 1440 = 9026 twips.
	// One twip = 1/20 point, and one point ≈ 12700 EMUs.
	fullWidthEmu := 5737100 // ≈ 9026/20 * 12700

	// --- Determine scaled height for images ---
	// Assume an original image size of 3000000 EMUs (width) x 2000000 EMUs (height).
	scaleFactor := float64(fullWidthEmu) / 3000000.0
	newHeightEmu := int(2000000 * scaleFactor)

	// --- Build the document body ---
	// We'll iterate over the entries in order. For each "figure" entry, we increment a figure index,
	// add a paragraph with the inline drawing, and then a paragraph with the legend.
	// For "text" entries, we simply add a paragraph with the text.
	var docBody bytes.Buffer
	figureIndex := 0

	for _, entry := range entries {
		if entry.Type == "figure" {
			figureIndex++
			// Paragraph with the image drawing.
			docBody.WriteString(fmt.Sprintf(`
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="%d" cy="%d"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:docPr id="%d" name="Picture %d"/>
            <wp:cNvGraphicFramePr>
              <a:graphicFrameLocks noChangeAspect="1"/>
            </wp:cNvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr>
                    <pic:cNvPr id="0" name="Picture %d"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="rId%d"/>
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="%d" cy="%d"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                      <a:avLst/>
                    </a:prstGeom>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>`, fullWidthEmu, newHeightEmu, figureIndex, figureIndex, figureIndex, figureIndex, fullWidthEmu, newHeightEmu))
			// Paragraph with the legend text.
			docBody.WriteString(fmt.Sprintf(`
    <w:p>
      <w:r>
        <w:t>%s</w:t>
      </w:r>
    </w:p>`, xmlEscape(entry.Legend)))
		} else if entry.Type == "text" {
			// Paragraph with the text.
			docBody.WriteString(fmt.Sprintf(`
    <w:p>
      <w:r>
        <w:t>%s</w:t>
      </w:r>
    </w:p>`, xmlEscape(entry.Text)))
		}
	}

	// Wrap the body in the complete document structure.
	documentXML := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>%s
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`, docBody.String())

	// --- Build the document relationships file (word/_rels/document.xml.rels) ---
	// Only figure entries need a relationship entry.
	var docRels bytes.Buffer
	docRels.WriteString(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	for i := 1; i <= figureCount; i++ {
		docRels.WriteString(fmt.Sprintf(`
    <Relationship Id="rId%d" 
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" 
        Target="media/image%d.svg"/>`, i, i))
	}
	docRels.WriteString("\n</Relationships>")

	// --- Build the [Content_Types].xml ---
	// Declare the main document and each image override.
	var contentTypes bytes.Buffer
	contentTypes.WriteString(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="xml" ContentType="application/xml"/>
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`)
	for i := 1; i <= figureCount; i++ {
		contentTypes.WriteString(fmt.Sprintf(`
    <Override PartName="/word/media/image%d.svg" ContentType="image/svg+xml"/>`, i))
	}
	contentTypes.WriteString("\n</Types>")

	// --- Build the root relationships file (_rels/.rels) ---
	relsRoot := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" 
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" 
        Target="word/document.xml"/>
</Relationships>`

	// --- Create the DOCX (ZIP archive) ---
	outFile, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("error creating output file: %v", err)
	}
	defer outFile.Close()

	zipWriter := zip.NewWriter(outFile)
	defer zipWriter.Close()

	// Add explicit directory entries.
	if err := addDir(zipWriter, "_rels/"); err != nil {
		return fmt.Errorf("error adding _rels/ directory: %v", err)
	}
	if err := addDir(zipWriter, "word/"); err != nil {
		return fmt.Errorf("error adding word/ directory: %v", err)
	}
	if err := addDir(zipWriter, "word/_rels/"); err != nil {
		return fmt.Errorf("error adding word/_rels/ directory: %v", err)
	}
	if err := addDir(zipWriter, "word/media/"); err != nil {
		return fmt.Errorf("error adding word/media/ directory: %v", err)
	}

	// Add XML parts.
	if err := addFile(zipWriter, "[Content_Types].xml", contentTypes.String()); err != nil {
		return fmt.Errorf("error adding [Content_Types].xml: %v", err)
	}
	if err := addFile(zipWriter, "_rels/.rels", relsRoot); err != nil {
		return fmt.Errorf("error adding _rels/.rels: %v", err)
	}
	if err := addFile(zipWriter, "word/document.xml", documentXML); err != nil {
		return fmt.Errorf("error adding word/document.xml: %v", err)
	}
	if err := addFile(zipWriter, "word/_rels/document.xml.rels", docRels.String()); err != nil {
		return fmt.Errorf("error adding word/_rels/document.xml.rels: %v", err)
	}

	// Add each figure's SVG file to the media folder.
	figureIndex = 0
	for _, entry := range entries {
		if entry.Type == "figure" {
			figureIndex++
			imgName := fmt.Sprintf("word/media/image%d.svg", figureIndex)
			writer, err := zipWriter.Create(imgName)
			if err != nil {
				return fmt.Errorf("error creating image entry %q in ZIP: %v", imgName, err)
			}
			if _, err := writer.Write(figureData[figureIndex-1]); err != nil {
				return fmt.Errorf("error writing SVG data for %q: %v", imgName, err)
			}
		}
	}

	return nil
}

func main() {
	// Define a list of document entries.
	// Entries of Type "figure" must specify an SVG file path and a legend.
	// Entries of Type "text" specify the text to output.
	entries := []DocumentEntry{
		{Type: "figure", SvgPath: "input1.svg", Legend: "Figure 1: First SVG image."},
		{Type: "text", Text: "This is a paragraph of descriptive text following the first figure."},
		{Type: "figure", SvgPath: "input2.svg", Legend: "Figure 2: Second SVG image."},
		{Type: "text", Text: "Another text paragraph following the second figure."},
		// Add more entries as needed.
	}

	outputPath := "output.docx" // Output DOCX file path.

	if err := GenerateDocxFromEntries(entries, outputPath); err != nil {
		log.Fatalf("Failed to generate DOCX: %v", err)
	}
	log.Printf("DOCX generated successfully at %s", outputPath)
}

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
	// The name must end with a slash.
	header := &zip.FileHeader{
		Name:   name,
		Method: zip.Store, // no compression needed for directories
	}
	// Mark this as a directory.
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

// GenerateDocxFromSvg creates a DOCX file that includes a legend and embeds an SVG image.
// It builds the DOCX structure (a ZIP archive with XML parts) manually.
func GenerateDocxFromSvg(svgPath, legend, outputPath string) error {
	// Read the SVG file from disk.
	svgData, err := ioutil.ReadFile(svgPath)
	if err != nil {
		return fmt.Errorf("error reading SVG file: %v", err)
	}

	// Prepare the XML contents.
	contentTypes := `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="xml" ContentType="application/xml"/>
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/media/image1.svg" ContentType="image/svg+xml"/>
</Types>`

	relsRoot := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" 
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" 
        Target="word/document.xml"/>
</Relationships>`

	documentXML := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <!-- Paragraph with the legend text -->
    <w:p>
      <w:r>
        <w:t>%s</w:t>
      </w:r>
    </w:p>
    <!-- Paragraph with the embedded SVG image -->
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr>
                    <pic:cNvPr id="0" name="Picture 1"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="rId1"/>
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="3000000" cy="2000000"/>
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
    </w:p>
    <!-- Section properties -->
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`, xmlEscape(legend))

	documentRels := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" 
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" 
        Target="media/image1.svg"/>
</Relationships>`

	// Create the output DOCX file.
	outFile, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("error creating output file: %v", err)
	}
	defer outFile.Close()

	zipWriter := zip.NewWriter(outFile)
	// Ensure the ZIP archive is closed properly.
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

	// Add the XML parts to the ZIP archive.
	if err := addFile(zipWriter, "[Content_Types].xml", contentTypes); err != nil {
		return fmt.Errorf("error adding [Content_Types].xml: %v", err)
	}
	if err := addFile(zipWriter, "_rels/.rels", relsRoot); err != nil {
		return fmt.Errorf("error adding _rels/.rels: %v", err)
	}
	if err := addFile(zipWriter, "word/document.xml", documentXML); err != nil {
		return fmt.Errorf("error adding word/document.xml: %v", err)
	}
	if err := addFile(zipWriter, "word/_rels/document.xml.rels", documentRels); err != nil {
		return fmt.Errorf("error adding word/_rels/document.xml.rels: %v", err)
	}

	// Add the SVG image into the word/media folder.
	imgWriter, err := zipWriter.Create("word/media/image1.svg")
	if err != nil {
		return fmt.Errorf("error creating image entry in ZIP: %v", err)
	}
	if _, err := imgWriter.Write(svgData); err != nil {
		return fmt.Errorf("error writing SVG data: %v", err)
	}

	return nil
}

func main() {
	svgPath := "input.svg"                                      // Path to your SVG file.
	legend := "Figure 1: This is the legend for the SVG image." // Legend text.
	outputPath := "output.docx"                                 // Output DOCX file path.

	if err := GenerateDocxFromSvg(svgPath, legend, outputPath); err != nil {
		log.Fatalf("Failed to generate DOCX: %v", err)
	}
	log.Printf("DOCX generated successfully at %s", outputPath)
}

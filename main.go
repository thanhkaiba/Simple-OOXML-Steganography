package main

import (
	"archive/zip"
	"bufio"
	"bytes"
	"github.com/beevik/etree"
	"os"
	"time"
)

func main() {
	err := HideAndSave([]byte("test test"), "./test.docx", "./output.docx")
	if err != nil {
		panic(err)
	}
}

func Hide(secretData []byte, filename string) ([]byte, error) {
	zr, _ := zip.OpenReader("./test.docx")
	defer zr.Close()

	var b bytes.Buffer
	zwf := bufio.NewWriter(&b)
	zw := zip.NewWriter(zwf)
	defer zw.Close()

	err := HideData(secretData, zr, zw)
	if err != nil {
		return nil, err
	}

	return b.Bytes(), nil
}

func HideAndSave(secretData []byte, input string, output string) error {
	zr, _ := zip.OpenReader(input)
	defer zr.Close()

	zwf, _ := os.Create(output)
	defer zwf.Close()

	zw := zip.NewWriter(zwf)
	defer zw.Close()

	return HideData(secretData, zr, zw)
}

func HideData(secretData []byte, zr *zip.ReadCloser, zw *zip.Writer) error {
	for _, zipItem := range zr.File {
		if zipItem.Name == "_rels/.rels" {

			if err := createSecretTextFile(secretData, zipItem.Method, zw); err != nil {
				return err
			}

			f, err := zipItem.Open()
			if err != nil {
				return err
			}

			// store relationship of new file in _rels/.rels
			if relsWriter, err := zw.CreateHeader(&zipItem.FileHeader); err == nil {
				doc := etree.NewDocument()
				if _, err := doc.ReadFrom(f); err != nil {
					return err
				}
				result := doc.SelectElements("Relationships")
				if len(result) == 1 {
					Relationships := result[0]
					ele := etree.NewElement("Relationship")
					ele.CreateAttr("Id", "rId100")
					ele.CreateAttr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument/a")
					ele.CreateAttr("Target", "customData.xml")
					Relationships.AddChild(ele)
				}
				_, err := doc.WriteTo(relsWriter)
				if err != nil {
					return err
				}
			} else {
				return err
			}

		} else {
			if err := zw.Copy(zipItem); err != nil {
				return err
			}
		}

	}
	return nil
}

func createSecretTextFile(secretData []byte, method uint16, zw *zip.Writer) (err error) {
	// create customData.xml to store secret text
	header := &zip.FileHeader{
		Name:     "customData.xml",
		Method:   method, // deflate also works, but at a cost
		Modified: time.Now(),
	}

	if entryWriter, err := zw.CreateHeader(header); err == nil {
		_, err = entryWriter.Write(secretData)
	}
	return
}

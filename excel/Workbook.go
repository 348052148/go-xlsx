package excel
import (
	"fmt"
	"encoding/xml"
	"archive/zip"
	"io"
	"strings"
	"io/ioutil"
)
type Workbook struct {
	currentSheet string
	sheetXlsFileList map[string]io.ReadCloser
	stringsList []string

	sheetStyle *sheetStyle

	//xlsx内容
	workBookLabel *workBookLabel
	relationshipsLabel *relationshipsLabel
}

type workBookLabel struct {
	XMLName   xml.Name  `xml:"workbook"`
	Info      Info      `xml:"fileVersion"`
	SheetMetaLabel sheetMetaLabel `xml:"sheets"`
}

type Info struct {
	AppName string `xml:"appName,attr"`
}

type sheetMetaLabel struct {
	SheetInfoLabels []sheetInfoLabel `xml:"sheet"`
}

type sheetInfoLabel struct {
	SheetId string `xml:"sheetId,attr"`
	State   string `xml:"state,attr"`
	Name    string `xml:"name,attr"`
	RID		string `xml:"id,r,attr"`
}
//引用
type relationshipsLabel struct {
	XMLName   xml.Name  `xml:"Relationships"`
	RelationshipLabels 	[]relationshipLabel `xml:"Relationship"`
}
type relationshipLabel struct {
	RID 	string `xml:"Id,attr"`
	Target 	string `xml:"Target,attr"`
}


func (book *Workbook) Echo()  {
	fmt.Println(book.workBookLabel)
}

func (book *Workbook) NewWorkSheet() *workSheet {
	return new(workSheet)
}

//改变活跃sheet
func (book *Workbook) ChangeSheet(sheetName string) *workSheet {
	book.currentSheet = sheetName

	workSheet := new(workSheet)
	//搜索sheetName
	sheetIndex := ""
	for _, s := range book.workBookLabel.SheetMetaLabel.SheetInfoLabels {
		if s.Name == sheetName {
			sheetIndex = book.getSheetPathByrId(s.RID)
			break
		}
	}
	fmt.Println(sheetIndex)

	workSheet.newWorkSheet(book.sheetXlsFileList[sheetIndex], book)

	return workSheet
}
//初始化
func (book *Workbook) NewWorkBook(filePath string) *Workbook {
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return nil
	}
	book.sheetStyle = new(sheetStyle)
	book.gennerateWorkBookStruct(reader)
	return book
}

//根据sheet-rId获取 sheet名称
func (book *Workbook) getSheetPathByrId(rId string) string{
	for _,ships := range book.relationshipsLabel.RelationshipLabels {
		if(ships.RID == rId) {
			return book.workBookLabel.Info.AppName+"/"+ships.Target
		}
	}
	return ""
}

//生成workbook数据结构
func (book *Workbook) gennerateWorkBookStruct(reader *zip.ReadCloser) {
	book.sheetXlsFileList = make(map[string]io.ReadCloser, 0)
	for _, file := range reader.File {
		if file.FileHeader.Name == "xl/workbook.xml" {
			f, err := file.Open()
			if err != nil {

			}
			b, err := ioutil.ReadAll(f)
			if err != nil {
				fmt.Println("ERROR")
				return
			}
			workBookLabel := new(workBookLabel)
			xml.Unmarshal(b, workBookLabel)
			book.workBookLabel = workBookLabel
			//处理workbook
		} else if strings.Index(file.FileHeader.Name, "xl/worksheets/sheet") != -1 {
			f, err := file.Open()
			if err != nil {

			}
			//sheet存入
			book.sheetXlsFileList[file.FileHeader.Name] = f
		} else if file.FileHeader.Name == "xl/sharedStrings.xml" {
			f, err := file.Open()
			if err != nil {

			}
			// 处理共享字符集
			sharedStrings := new(sharedStrings)
			sharedStrings.newSharedString(f)
			book.stringsList = sharedStrings.getMapString()
		} else if (file.FileHeader.Name == "xl/_rels/workbook.xml.rels"){
			f, err := file.Open()
			if err != nil {

			}
			// relsShips := new(relationshipsLabel)
			rels, err := ioutil.ReadAll(f) 
			relationshipsLabel := new(relationshipsLabel)
			err = xml.Unmarshal(rels, relationshipsLabel)
			book.relationshipsLabel = relationshipsLabel
		} else if file.FileHeader.Name == "xl/styles.xml" {
			f, err := file.Open()
			if err != nil {

			}
			//处理样式
			book.sheetStyle.NewStyleSheet(f)
		} else {
			//fmt.Println(file.FileHeader.Name)
		}
	}
}
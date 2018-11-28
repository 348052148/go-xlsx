package excel

import (
	"fmt"
	"io"
	"io/ioutil"
	"encoding/xml"
	"strconv"
	"strings"
	"regexp"
	"time"
	"encoding/json"
)

type workSheet struct {
	workBook *Workbook
	Data [][]string
	strList []string

	//
	workSheetLabel *workSheetLabel
}

type workSheetLabel struct {
	XMLName      xml.Name  `xml:"worksheet"`
	SheetData    sheetDataLabel `xml:"sheetData"`
	MergeCells	 mergeCellsLabel `xml:"mergeCells"`
}

type mergeCellsLabel struct {
	List  []mergeCellLabel `xml:"mergeCell"`
}
type mergeCellLabel struct {
	Ref string `xml:"ref,attr"`
}

type sheetDataLabel struct {
	XMLName xml.Name `xml:"sheetData"`
	Rows    []rowLabel    `xml:"row"`
}

type rowLabel struct {
	XMLName xml.Name `xml:"row"`
	Span    string   `xml:"spans,attr"`
	R       string   `xml:"r,attr"`
	Cols    []colLabel    `xml:"c"`
}

type colLabel struct {
	XMLName xml.Name `xml:"c"`
	S       string   `xml:"s,attr"`
	T       string   `xml:"t,attr"`
	R       string   `xml:"r,attr"`
	F       string   `xml:"f"`
	V       string   `xml:"v"`
	IS      string   `xml:"is"`
}

func (sheet *workSheet) Echo()  {
	
	fmt.Println("WORKSHEER")
}
//格式化输出json
func (sheet *workSheet) FormatJson() []byte {
	data, err := json.Marshal(sheet.getSheetData())
	if err !=nil {
		panic("格式化Json失败了哦")
	}
	return data
}
//格式化输出数组
func (sheet *workSheet) FormatArray() [][]string {
	return sheet.getSheetData()
}
//获取sheet数据
func (sheet *workSheet) getSheetData() [][]string {

	for _, row := range sheet.workSheetLabel.SheetData.Rows {
		lens := strings.Split(row.Span, ":")
		i, err := strconv.Atoi(lens[1])
		if err != nil {

		}
		var tmp = make([]string, i)
		// var mcell string
		for _, col := range row.Cols {
			colIndex := getIndexByCol(col.R, row.R, i)
			//样式id
			//styleId := col.S

			//R暂不处理

			if col.V == "" {
				col.V = col.IS
			}
			
			col.V = sheet.hanldCellType(col)

			tmp[colIndex] = col.V //+":" + w.GetCellType(col)

		}
		sheet.Data = append(sheet.Data, tmp)
	}

	for _, c := range sheet.workSheetLabel.MergeCells.List {
		rangeColume(c.Ref,sheet)
	}

	return sheet.Data
}
//新建worksheet
func (sheet *workSheet) newWorkSheet(fileReader io.ReadCloser, workBook *Workbook) *workSheet {
	b, err := ioutil.ReadAll(fileReader)

	if err != nil {
		return nil
	}
	workSheetLabel := new(workSheetLabel)
	err = xml.Unmarshal(b, workSheetLabel)
	sheet.workSheetLabel = workSheetLabel
	sheet.workBook = workBook

	sheet.strList = workBook.stringsList
	return sheet
}
// 获取单元格类型
func (sheet *workSheet) getCellType(col colLabel) string {
	var cellType string = "NUMBER"
	switch col.T {
	case "b":
		cellType = "BOOL"
		break
	case "s":
		cellType = "SST"
		break
	case "e":
		cellType = "ERR"
		break
	case "str":
		cellType = "FORMULA"
		break
	case "inlineStr":
		cellType = "INLINESTR"
		break
	}

	i, _ :=strconv.Atoi(col.S)
	s := sheet.workBook.sheetStyle.GetStyleVal(i)

	sr := strings.Split(s, ";")

	s = sr[0]

	p,_ := regexp.MatchString("%", s)
	if p {
		cellType = "Percentage"
	}

	b, _ := regexp.MatchString("[ymdhs]", s)

	if b && strings.Index(s, "Red") == -1 {
		//处理获取是否是日期类型
		cellType = "DATE"
	}

	return cellType
}
//获取单元格索引
func getIndexByCol(colName string, s string, length int) int {
	Colume := make(map[string]int)
	k := ""
	for i := 0; i < length; i++ {
		k = ""
		d := i
		for true {
			if d/26 == 0 {
				k = string(d+65) + k
				break
			} else {
				k = string(d%26+65) + k
				d = d/26 - 1
			}
		}
		Colume[k] = i
	}

	colIndex := strings.Replace(colName, s, "", -1)

	return Colume[colIndex]
}

func (sheet *workSheet) hanldCellType(col colLabel) string {
	switch sheet.getCellType(col) {
	case "SST":
		i, err := strconv.Atoi(col.V)
		if err != nil {
			i = 0
		}
		col.V = sheet.strList[i]
		break
	case "BOOL":
		break
	case "INLINESTR":
		break
	case "ERR":
		break
	case "FORMULA":
		break
	case "Percentage":
		//$Value = sprintf('%.2f%%', round(100 * $Value, 2));
		t, _ := strconv.ParseFloat(col.V, 64)
		col.V = fmt.Sprintf("%.2f%%", 100 * t)
		break
	case "DATE":
		date := strings.Split(col.V, ".")
		days,_ := strconv.ParseInt(date[0],10,64)

		if days > 60 {
			days--
		}
		
		times := 0.0
		if(len(date) >=2 ){
			times,_ = strconv.ParseFloat("0." + date[1],64)
		}
		
		seconds :=  times * 86400

		local2, _ := time.LoadLocation("Local")
		
		tm2, _ := time.Parse("01/02/2006", "01/01/1900")

		seconds1 := tm2.In(local2).Unix() - 86400 + days * 86400 + int64(seconds)

		col.V = strconv.FormatInt(seconds1, 10)
		
		break
	case "NUMBER":
		
	}
	return col.V
}

//讲单元格区间的赋值
func rangeColume(s string, w *workSheet) {
	var darr [2][2]int64
	sarr := strings.Split(s, ":");
	darr[0] = getColume(sarr[0]);
	darr[1] = getColume(sarr[1]);
	// fmt.Println(s)
	// fmt.Println(darr)
	for i:=darr[0][0]-1; i < darr[1][0]; i++ {
		for j:=darr[0][1]-1; j< darr[1][1]; j++ {
			w.Data[i][j] = w.Data[darr[0][0]-1][darr[0][1]-1];
		}
	}
}
//反转字符串操作
func reverse(str string) string {
    rs := []rune(str)
    len := len(rs)
    var tt []rune

    tt = make([]rune, 0)
    for i := 0; i < len; i++ {
        tt = append(tt, rs[len-i-1])
    }
    return string(tt[0:])
}
// 获取以数字来标示的excel 行和列
func getColume(col string) [2]int64 {
	var  d [2]int64 
	r := regexp.MustCompile("\\d+").FindString(col)

	cs := regexp.MustCompile("[A-Z]+").FindString(col)
	var cv int64 = 0;
	for index, i :=  range reverse(cs) { 
		cv = cv + (int64(i) - 64)+ 26 * int64(index)
	}

	d[0],_ = strconv.ParseInt(r, 10, 64)
	d[1] = cv
	return d
}
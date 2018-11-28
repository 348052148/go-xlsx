package excel

import (
	"encoding/xml"
	"io"
	"io/ioutil"
)

type sheetStyle struct {
	StyleList map[string]string
	CellStyles map[int]string

	stylesLable *stylesLabel
}

type stylesLabel struct {
	XMLName   xml.Name  `xml:"styleSheet"`
	NumFmts   numFmtsLabel 	`xml:"numFmts"`
	CellXfs  cellXfsLabel	`xml:"cellXfs"`
}

type cellXfsLabel struct {
	Xfs 	[]xfLabel 	`xml:"xf"`
}

type xfLabel struct {
	NumFmtId string `xml:"numFmtId,attr"` 
}

type numFmtsLabel struct {
	List []numFmtLabel `xml:"numFmt"`
}

type numFmtLabel struct {
	NumFmtId string `xml:"numFmtId,attr"`
	FormatCode string `xml:"formatCode,attr"`
}


func (ss *sheetStyle) StyleInit() {
	ss.StyleList = make(map[string]string)
	ss.StyleList["0"] = "";
	ss.StyleList["1"] = "0";
	ss.StyleList["2"] = "0.00";
	ss.StyleList["3"] = "#,##0";
	ss.StyleList["4"] = "#,##0.00";

	ss.StyleList["9"] = "0%";
	ss.StyleList["10"] = "0.00%";
	ss.StyleList["11"] = "0.00E+00";
	ss.StyleList["12"] = "# ?/?";
	ss.StyleList["13"] = "# ??/??";
	ss.StyleList["14"] = "yyyy/m/d";
	ss.StyleList["15"] = "d-mmm-yy";
	ss.StyleList["16"] = "d-mmm";
	ss.StyleList["17"] = "mmm-yy";
	ss.StyleList["18"] = "h:mm AM/PM";
	ss.StyleList["19"] = "h:mm:ss AM/PM";
	ss.StyleList["20"] = "h:mm";
	ss.StyleList["21"] = "h:mm:ss";
	ss.StyleList["22"] = "yyyy/m/d h:mm";

	ss.StyleList["31"] = "yyyy年m月d日";
	ss.StyleList["32"] = "h时mmi分";
	ss.StyleList["33"] = "h时mmi分ss秒";

	ss.StyleList["37"] = "#,##0 ;(#,##0)";
	ss.StyleList["38"] = "#,##0 ;[Red](#,##0)";
	ss.StyleList["39"] = "#,##0.00;(#,##0.00)";
	ss.StyleList["40"] = "#,##0.00;[Red](#,##0.00)";

	ss.StyleList["44"] = "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)";
	ss.StyleList["45"] = "mm:ss";
	ss.StyleList["46"] = "[h]:mm:ss";
	ss.StyleList["47"] = "mm:ss.0";
	ss.StyleList["48"] = "##0.0E+0";
	ss.StyleList["49"] = "@";

	ss.StyleList["55"] = "AM/PM h时mmi分";
	ss.StyleList["56"] = "AM/PM h时mmi分ss秒";
	ss.StyleList["58"] = "m月d日";
	// CHT & CHS
	ss.StyleList["27"] = "yyyy年m月";
	ss.StyleList["30"] = "m/d/yy";
	ss.StyleList["36"] = "[$-404]e/m/d";
	ss.StyleList["50"] = "[$-404]e/m/d";
	ss.StyleList["57"] = "[$-404]e/m/d";
	// THA
	ss.StyleList["59"] = "t0";
	ss.StyleList["60"] = "t0.00";
	ss.StyleList["61"] = "t#,##0";
	ss.StyleList["62"] = "t#,##0.00";
	ss.StyleList["67"] = "t0%";
	ss.StyleList["68"] = "t0.00%";
	ss.StyleList["69"] = "t# ?/?";
	ss.StyleList["70"] = "t# ??/??";

}

func (ss *sheetStyle) GetStyleVal(fmtId int) string{
	return ss.StyleList[ss.CellStyles[fmtId]]
}

func (ss *sheetStyle) NewStyleSheet(reader io.ReadCloser) *sheetStyle {
	ss.StyleInit()
	styles := new(stylesLabel)
	rels, err := ioutil.ReadAll(reader) 
	err = xml.Unmarshal(rels, styles)
	if err!= nil {

	}
	//只选择自定义样式
	for _, numFmt := range styles.NumFmts.List {
		ss.StyleList[numFmt.NumFmtId] = numFmt.FormatCode
	}
	ss.CellStyles = make(map[int]string)
	//只选择id
	for i, xf := range styles.CellXfs.Xfs {
		ss.CellStyles[i] = xf.NumFmtId
	}

	return nil
}


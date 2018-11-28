package excel

import (
	"io"
	"encoding/xml"
	"io/ioutil"
)

type sharedStrings struct {
	strList [] string
	sharedStringLabel *sharedStringLabel
}


type sharedStringLabel struct {
	XMLName xml.Name `xml:"sst"`
	//	count   string   `xml:count,attr`
	SI   []siInfoLabel `xml:"si"`
}
type siInfoLabel struct {
	XMLName xml.Name `xml:"si"`
	T       string   `xml:"t"`
}

func (str *sharedStrings) newSharedString(FileReader io.ReadCloser) *sharedStrings {
	b, err := ioutil.ReadAll(FileReader)
	if err != nil {
		return nil
	}
	sharedStringLabel := new(sharedStringLabel)
	err = xml.Unmarshal(b, sharedStringLabel)
	str.sharedStringLabel = sharedStringLabel
	return str
}

func (str *sharedStrings) getMapString() []string {
	for _, t := range str.sharedStringLabel.SI {
		str.strList = append(str.strList, t.T)
	}
	return str.strList
}

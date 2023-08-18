package excelstructure

import (
	"github.com/xuri/excelize/v2"
)

var boolTrueValue = []string{
	"true",
	"True",
	"TRUE",
	"1",
	"是",
	"yes",
	"Yes",
	"YES",
	"y",
	"Y",
}

// Parser parser.
type Parser struct {
	FileName string
	// DataIndexOffset 数据索引偏移量,第一行可能为表头，则偏移量为1，如果有注释占用一行，则为2
	DataIndexOffset int
	// BoolTrueValues bool类型的true可选值 boolTrueValue
	BoolTrueValues []string
	// IsCheckEmpty 序列化到结构提的时候是否检测空值如果为空值则报错
	IsCheckEmpty bool
	// IsEmptyFunc 空值校验函数 可自定义空值
	IsEmptyFunc func(v string) bool
	// IsCoordinatesABS
	//    if true Cell.Coordinates is "A1"
	//    if false Cell.Coordinates is "$A$1"
	IsCoordinatesABS bool
	ExcelData        *Data
	// AllowFieldRepeat 允许表头字段重复
	AllowFieldRepeat bool
	currentSheetName string

	errsMap    map[string]error
	hasComment bool
	// fieldHeadRowIndex 表头行索引，第一行为表头，则索引为1
	fieldHeadRowIndex int
	excelFile         *excelize.File
}

// NewParser 传入文件名
// sheetIndex从1开始
//
//	ExcelTag excel tag:
//		ExcelField: map the excel head filed
//		ExcelDefault: if excel field is empty, use this default value
func NewParser(fileName string) *Parser {
	return &Parser{
		FileName:          fileName,
		DataIndexOffset:   1,
		fieldHeadRowIndex: 1,
		BoolTrueValues:    boolTrueValue,
		IsEmptyFunc: func(v string) bool {
			return v == ""
		},
		errsMap: make(map[string]error),
	}
}

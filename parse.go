package excelstructure

import (
	"fmt"

	sliceutil "github.com/booyangcc/utils/sliceutil"
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
	fileName string
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
	serializers       map[string]Serializer
}

// NewParser 传入文件名
// sheetIndex从1开始
//
//	ExcelTag excel tag:
//		ExcelField: map the excel head field
//		ExcelDefault: if excel field is empty, use this default value
func NewParser() *Parser {
	return &Parser{
		DataIndexOffset:   1,
		fieldHeadRowIndex: 1,
		BoolTrueValues:    boolTrueValue,
		IsEmptyFunc: func(v string) bool {
			return v == ""
		},
		errsMap: make(map[string]error),
	}
}

// RegisterSerializer 注册序列化器
func (p *Parser) RegisterSerializer(name string, serializer Serializer) error {
	if p.serializers == nil {
		p.serializers = make(map[string]Serializer)
	}

	if _, ok := p.serializers[name]; ok {
		return ErrorSerializerNameRepeat
	}

	if serializer.Marshal == nil || serializer.Unmarshal == nil {
		return ErrorSerializerHandlerEmpty
	}

	p.serializers[name] = serializer
	return nil
}

// Parse parse.
func (p *Parser) Parse(fileName string) (*Data, error) {
	p.fileName = fileName
	excelFile, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, NewError(fileName, "", "", err)
	}
	defer func() {
		if err1 := excelFile.Close(); err1 != nil {
			fmt.Println(err1.Error())
		}
	}()

	p.excelFile = excelFile
	if p.fieldHeadRowIndex < 1 {
		p.fieldHeadRowIndex = 1
	}
	if p.DataIndexOffset < 1 {
		p.DataIndexOffset = 1
	}

	excelData, err := p.getExcelFileData()
	if err != nil {
		return nil, NewError(fileName, "", "", err)
	}
	p.ExcelData = excelData

	return excelData, nil
}

func (p *Parser) getExcelFileData() (*Data, error) {
	fileName := p.fileName
	sheetMap := p.excelFile.GetSheetMap()
	if nil == sheetMap || len(sheetMap) == 0 {
		return nil, NewError(fileName, "", "", ErrorNoSheet)
	}

	sheetIndexData := make(map[int]*SheetData)
	sheetNameData := make(map[string]*SheetData)
	for sheetIndex, sheetName := range sheetMap {
		rows, err := p.excelFile.GetRows(sheetName)
		if err != nil {
			return nil, NewError(fileName, sheetName, fmt.Sprintf("sheet index %d", sheetIndex), err)
		}
		if len(rows) == 0 {
			sheetData := &SheetData{
				SheetName:       sheetName,
				FileName:        fileName,
				DataIndexOffset: p.DataIndexOffset,
			}
			sheetIndexData[sheetIndex] = sheetData
			sheetNameData[sheetName] = sheetData
			continue
		}
		// 输入数据为excel直观的行数 从1开始
		sheetFields := rows[p.fieldHeadRowIndex-1]

		fieldValid := make([]string, 0)
		repeatFieldName := ""
		for _, fieldName := range sheetFields {
			if sliceutil.InSlice(fieldName, fieldValid) {
				repeatFieldName = fieldName
				break
			}
			fieldValid = append(fieldValid, fieldName)
		}
		// 检查是否有相同字段
		if !p.AllowFieldRepeat && repeatFieldName != "" {
			return nil, NewError(fileName, sheetName, fmt.Sprintf("sheet index %d", sheetIndex), ErrorFieldRepeat)
		}

		parseRows := make(map[int]map[string]*Cell, 0)
		for index, row := range rows {
			excelIndex := index + 1
			if excelIndex <= p.DataIndexOffset {
				continue
			}
			p.getRow(excelIndex, row, sheetFields, parseRows)
		}

		sheetData := &SheetData{
			RowTotal:        len(rows),
			DataTotal:       len(rows) - p.DataIndexOffset,
			SheetName:       sheetName,
			FileName:        fileName,
			Rows:            parseRows,
			FieldKeys:       sheetFields,
			DataIndexOffset: p.DataIndexOffset,
		}

		sheetIndexData[sheetIndex] = sheetData
		sheetNameData[sheetName] = sheetData
	}

	excelData := &Data{
		FileName:       fileName,
		SheetIndexData: sheetIndexData,
		SheetNameData:  sheetNameData,
		SheetTotal:     p.excelFile.SheetCount,
		SheetList:      p.excelFile.GetSheetList(),
	}

	return excelData, nil
}

func (p *Parser) getRow(rowIndex int, rawRow, sheetFields []string, rowsData map[int]map[string]*Cell) {
	rowData := make(map[string]*Cell, 0)
	// excel起始行为1，所以这里要+1
	for colIndex, fieldName := range sheetFields {
		cell := &Cell{
			RowIndex: rowIndex,
			ColIndex: colIndex + 1,
			Key:      fieldName,
		}
		var c string
		var err error
		if p.IsCoordinatesABS {
			c, err = excelize.CoordinatesToCellName(colIndex+1, rowIndex, true)
		} else {
			c, err = excelize.CoordinatesToCellName(colIndex+1, rowIndex)
		}
		if err != nil {
			cell.ErrMsg = c
		}

		// 当前列索引大于当前行长度，说明当前行数据不足，当前列为空
		if colIndex > len(rawRow)-1 {
			cell.IsEmpty = true
		} else {
			cell.Value = rawRow[colIndex]
			cell.IsEmpty = p.IsEmptyFunc(rawRow[colIndex])
		}
		cell.Coordinates = c
		rowData[fieldName] = cell
	}

	rowsData[rowIndex] = rowData
}

package excel

import (
	"errors"
	"fmt"
	"reflect"
	"strconv"
	"strings"

	sliceutil "github.com/booyangcc/utils/sliceutil"
	"github.com/hashicorp/go-multierror"

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

// ErrElemNotStruct slice elem not struct
var ErrElemNotStruct = fmt.Errorf("slice elem must be struct")

// ExcelHeader excel header.
// ExcelDefault

// Cell cell.
type Cell struct {
	RowIndex    int
	ColIndex    int
	Value       string
	Key         string
	Coordinates string
	ErrMsg      string
	IsEmpty     bool
}

// SheetData sheet data.
type SheetData struct {
	// RowTotal  总行数
	RowTotal int
	// DataTotal 有效数据行数
	DataTotal int
	SheetName string
	// Rows  map[rowIndex]map[keyFiled]*Cell
	Rows      map[int]map[string]*Cell
	FieldKeys []string
	// DataIndexOffset 数据起始偏移量
	DataIndexOffset int
}

// GetCell get cell.
func (s *SheetData) GetCell(rowIndex int, fieldKey string, isCheckEmpty ...bool) (*Cell, error) {
	if s.DataTotal < rowIndex-s.DataIndexOffset {
		return nil, errors.New("row index out of range")
	}
	if rowIndex == 1 {
		return nil, errors.New("row index must > 1, because row 1 is header")
	}
	row := s.Rows[rowIndex]
	if !sliceutil.InSlice(fieldKey, s.FieldKeys) {
		return nil, fmt.Errorf("field key %s not in header", fieldKey)
	}
	cell, ok := row[fieldKey]
	if !ok {
		return nil, fmt.Errorf("field key %s not in header", fieldKey)
	}

	isCheck := false
	if len(isCheckEmpty) > 0 {
		isCheck = isCheckEmpty[0]
	}

	if isCheck && cell.IsEmpty {
		return nil, fmt.Errorf("row %d field %s is empty", rowIndex, fieldKey)
	}

	return cell, nil
}

// GetIntValue cong string 获取值
func (s *SheetData) GetIntValue(rowIndex int, fieldKey string, isCheckEmpty ...bool) (int, error) {
	v, err := s.GetCell(rowIndex, fieldKey, isCheckEmpty...)
	if err != nil {
		return 0, err
	}

	if v.Value == "" {
		return 0, nil
	}
	res, err := strconv.Atoi(v.Value)
	if err != nil {
		return 0, fmt.Errorf("cell %s field %s value %s is not int", v.Coordinates, fieldKey, v.Value)
	}
	return res, nil
}

// GetStringValue cong string 获取值
func (s *SheetData) GetStringValue(rowIndex int, fieldKey string, isCheckEmpty ...bool) (string, error) {
	v, err := s.GetCell(rowIndex, fieldKey, isCheckEmpty...)
	if err != nil {
		return "", err
	}

	return v.Value, nil
}

// GetIntValueWithMultiError cong string 获取值
func (s *SheetData) GetIntValueWithMultiError(
	rowIndex int, fieldKey string, errs error, isCheckEmpty ...bool,
) (int, error) {
	v, err := s.GetCell(rowIndex, fieldKey, isCheckEmpty...)
	if err != nil {
		errs = multierror.Append(errs, err)
		return 0, errs
	}

	if v.Value == "" {
		return 0, errs
	}

	res, err := strconv.Atoi(v.Value)
	if err != nil {
		errs = multierror.Append(errs,
			fmt.Errorf("cell %s field %s value %s is not int", v.Coordinates, fieldKey, v.Value))
		return 0, errs

	}

	return res, errs
}

// GetStringValueWithMultiError cong string 获取值
func (s *SheetData) GetStringValueWithMultiError(
	rowIndex int, fieldKey string, errs error, isCheckEmpty ...bool,
) (string, error) {
	v, err := s.GetCell(rowIndex, fieldKey, isCheckEmpty...)
	if err != nil {
		errs = multierror.Append(errs, err)
		return "", errs
	}

	return v.Value, errs
}

// Data excel data.
type Data struct {
	FileName       string
	SheetIndexData map[int]*SheetData
	SheetNameData  map[string]*SheetData
	SheetTotal     int
}

// Parser parser.
type Parser struct {
	FileName string
	// DataIndexOffset 数据索引偏移量,第一行可能为表头，则偏移量为1，如果有注释占用一行，则为2
	DataIndexOffset int
	// fieldHeadRowIndex 表头行索引，第一行为表头，则索引为1
	fieldHeadRowIndex int
	excelFile         *excelize.File
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
	errsMap          map[string]error
	// AllowFieldRepeat 允许表头字段重复
	AllowFieldRepeat bool
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

// Parse parse.
func (p *Parser) Parse() (*Data, error) {
	fileName := p.FileName
	excelFile, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, newExcelParseError(fileName, "root", "OpenReader err", err.Error())
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
		return nil, err
	}
	p.ExcelData = excelData
	return excelData, nil
}

// UnmarshalWithSheetIndex parse with sheet index. start with 1
func (p *Parser) UnmarshalWithSheetIndex(index int, output interface{}) error {
	excelData, err := p.Parse()
	if err != nil {
		return err
	}
	var errs error
	p.parseToStruct(index, excelData, output, &errs)
	return errs
}

// Unmarshal parse with sheet index 1
// output must be a pointer slice
// if the pointer field is pointer, and the value is empty ,the pointer field will be nil
func (p *Parser) Unmarshal(output interface{}) error {
	excelData, err := p.Parse()
	if err != nil {
		return err
	}
	var errs error
	p.parseToStruct(1, excelData, output, &errs)
	return errs
}

func (p *Parser) parseToStruct(sheetIndex int, excelData *Data, output interface{}, errs *error) {
	if len(excelData.SheetIndexData) < sheetIndex {
		*errs = multierror.Append(*errs, fmt.Errorf("sheet index %d out of range", sheetIndex))
		return
	}

	rv := reflect.ValueOf(output)
	if rv.Kind() == reflect.Ptr && rv.Elem().Kind() != reflect.Slice {
		*errs = multierror.Append(*errs, fmt.Errorf("output must be a array"))
		return
	}

	sheetData := excelData.SheetIndexData[sheetIndex]
	sliceType := rv.Elem().Type()
	sliceElemType := sliceType.Elem()

	var sliceElemStructType reflect.Type
	if sliceElemType.Kind() == reflect.Ptr {
		sliceElemStructType = sliceElemType.Elem()
		if sliceElemStructType.Kind() != reflect.Struct {
			*errs = multierror.Append(*errs, ErrElemNotStruct)
			return
		}
	} else if sliceElemType.Kind() == reflect.Struct {
		sliceElemStructType = sliceElemType
	} else {
		*errs = multierror.Append(*errs, ErrElemNotStruct)
		return
	}

	tagMap, err := parseFiledTagSetting(sliceElemStructType)
	if err != nil {
		*errs = multierror.Append(*errs, err)
		return
	}

	arr := reflect.MakeSlice(sliceType, 0, 0)
	for i := range sheetData.Rows {
		out := reflect.New(sliceElemStructType)

		if err := p.parseRowToStruct(i, sheetData, out, tagMap); err != nil {
			p.appendError(errs, err)
			continue
		}
		if sliceElemType.Kind() == reflect.Ptr {
			arr = reflect.Append(arr, out)
		}
		if sliceElemType.Kind() == reflect.Struct {
			arr = reflect.Append(arr, out.Elem())
		}

	}
	rv.Elem().Set(arr)
}

func (p *Parser) appendError(errs *error, err error) {
	if errs == nil {
		return
	}
	if err == nil {
		return
	}
	if p.errsMap == nil {
		p.errsMap = make(map[string]error)
	}

	if _, ok := p.errsMap[err.Error()]; ok {
		return
	}
	*errs = multierror.Append(*errs, err)
	p.errsMap[err.Error()] = err
}

// parse row to struct by tag setting
func (p *Parser) parseRowToStruct(
	rowIndex int, sheetData *SheetData, ve reflect.Value, tagMap map[string]TagSetting,
) (err error) {
	if ve.Kind() != reflect.Ptr {
		return errors.New("output must be a pointer")
	}
	if !ve.IsValid() {
		return errors.New("output is invalid")
	}
	vek := reflect.Indirect(ve)
	createStruct := vek.Type()
	for i := 0; i < createStruct.NumField(); i++ {
		field := createStruct.Field(i)
		tag := field.Tag
		et := tag.Get(TagName)
		columnName, df, skip := field.Name, "", false
		if len(et) > 0 {
			val, ok := tagMap[field.Name]
			if ok {
				columnName = val.Column
				df = val.Default
				skip = val.Skip
			}
		}
		if skip {
			continue
		}
		cell, err1 := sheetData.GetCell(rowIndex, columnName)
		if err1 != nil {
			return err1
		}

		err = p.filedSetAll(ve.Elem().FieldByName(field.Name), cell, df)
		if err != nil {
			return err
		}

	}
	return err
}

// nolint: gocyclo,gocritic
func (p *Parser) filedSetAll(field reflect.Value, cell *Cell, defaultValue string) error {
	value := cell.Value
	if p.IsCheckEmpty && cell.IsEmpty {
		return fmt.Errorf("%s field %s value is empty", cell.Coordinates, cell.Key)
	}
	if value == "" {
		value = defaultValue
	}
	var err error
	switch field.Kind() {
	case reflect.Ptr:
		valType := field.Type()
		valElemType := valType.Elem()
		if field.CanSet() {
			realVal := field
			if realVal.IsNil() {
				realVal = reflect.New(valElemType)
				if value != "" {
					err = p.filedSet(realVal.Elem(), value, cell)
					field.Set(realVal)
				} else {
					field.Set(reflect.Zero(valType))
				}
			}
		} else {
			err = fmt.Errorf("field %s can not set", cell.Key)
		}
	default:
		if field.CanSet() {
			err = p.filedSet(field, value, cell)
		} else {
			err = fmt.Errorf("field %s can not set", cell.Key)
		}
	}
	if err != nil {
		return err
	}
	return nil
}

func (p *Parser) filedSet(field reflect.Value, value string, cell *Cell) error {
	if !field.IsValid() {
		fmt.Println(field.Kind())
		return fmt.Errorf("field %s is invalid", cell.Key)
	}
	switch field.Kind() {
	case reflect.Int:
		intValue, err := strconv.Atoi(value)
		if err != nil {
			return err
		}
		if field.CanSet() {
			field.SetInt(int64(intValue))
		}

	case reflect.Uint:
		uintValue, err := strconv.ParseUint(value, 10, 32)
		if err != nil {
			return fmt.Errorf("field %s tag default value %s is not int, conv err  :%s",
				cell.Key, value, err.Error())
		}
		if field.CanSet() {
			field.SetUint(uintValue)
		}
	case reflect.String:
		if field.CanSet() {
			field.SetString(value)
		}
	case reflect.Bool:
		if field.CanSet() {
			field.SetBool(sliceutil.InSlice(value, p.BoolTrueValues))
		}
	case reflect.Float32:
		floatValue, err := strconv.ParseFloat(value, 32)
		if err != nil {
			return fmt.Errorf("field %s value %s is not float,conv err  :%s",
				cell.Key, value, err.Error())
		}
		if field.CanSet() {
			field.SetFloat(floatValue)
		}
	default:
		return fmt.Errorf("field %s type %s is not support", cell.Key, field.Kind().String())
	}
	return nil
}
func (p *Parser) getExcelFileData() (*Data, error) {
	fileName := p.FileName
	sheetMap := p.excelFile.GetSheetMap()
	if nil == sheetMap || len(sheetMap) == 0 {
		return nil, newExcelParseError(fileName, "root", "current file null data")
	}

	sheetIndexData := make(map[int]*SheetData)
	sheetNameData := make(map[string]*SheetData)
	for sheetIndex, sheetName := range sheetMap {
		rows, err := p.excelFile.GetRows(sheetName)
		if err != nil {
			return nil, newExcelParseError(fileName, sheetName, "get row err", err.Error())
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
			return nil, newExcelParseError(fileName, sheetName, fmt.Sprintf("field %s repeat", repeatFieldName))
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

type excelParseError struct {
	FileName  string
	SheetName string
	ErrMsg    string
}

func newExcelParseError(filName string, sheetName string, errMsg ...string) error {
	return &excelParseError{
		FileName:  filName,
		SheetName: sheetName,
		ErrMsg:    strings.Join(errMsg, ","),
	}
}

func (e *excelParseError) Error() string {
	return fmt.Sprintf(" %s, FileName: %s, SheetName: %s", e.ErrMsg, e.FileName, e.SheetName)
}

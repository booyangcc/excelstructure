package excelstructure

import (
	"fmt"
	"reflect"
	"strconv"

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

// Reader parser.
type Reader struct {
	FileName         string
	currentSheetName string
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

// NewReader 传入文件名
// sheetIndex从1开始
//
//	ExcelTag excel tag:
//		ExcelField: map the excel head filed
//		ExcelDefault: if excel field is empty, use this default value
func NewReader(fileName string) *Reader {
	return &Reader{
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
func (r *Reader) Parse() (*Data, error) {
	fileName := r.FileName
	excelFile, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, NewError(fileName, "", "", err)
	}
	defer func() {
		if err1 := excelFile.Close(); err1 != nil {
			fmt.Println(err1.Error())
		}
	}()

	r.excelFile = excelFile
	if r.fieldHeadRowIndex < 1 {
		r.fieldHeadRowIndex = 1
	}
	if r.DataIndexOffset < 1 {
		r.DataIndexOffset = 1
	}

	excelData, err := r.getExcelFileData()
	if err != nil {
		return nil, NewError(fileName, "", "", err)
	}
	r.ExcelData = excelData
	return excelData, nil
}

// UnmarshalWithSheetIndex parse with sheet index. start with 1
func (r *Reader) UnmarshalWithSheetIndex(index int, output interface{}) error {
	excelData, err := r.Parse()
	if err != nil {
		return err
	}
	var errs error
	r.parseToStruct(index, excelData, output, &errs)
	return errs
}

// Unmarshal parse with sheet index 1
// output must be a pointer slice
// if the pointer field is pointer, and the value is empty ,the pointer field will be nil
func (r *Reader) Unmarshal(output interface{}) error {
	excelData, err := r.Parse()
	if err != nil {
		return err
	}
	var errs error

	r.parseToStruct(1, excelData, output, &errs)
	return errs
}

func (r *Reader) parseToStruct(sheetIndex int, excelData *Data, output interface{}, errs *error) {
	if len(excelData.SheetIndexData) < sheetIndex {
		*errs = multierror.Append(*errs,
			NewError(r.FileName, "", fmt.Sprintf("sheetIndex %d", sheetIndex), ErrorSheetIndex))
		return
	}
	sheetData := excelData.SheetIndexData[sheetIndex]
	r.currentSheetName = sheetData.SheetName

	rv := reflect.ValueOf(output)
	if rv.Kind() == reflect.Ptr && rv.Elem().Kind() != reflect.Slice {
		*errs = multierror.Append(*errs,
			NewError(r.FileName, r.currentSheetName, "", ErrorInOutputType))
		return
	}

	sliceType := rv.Elem().Type()
	sliceElemType := sliceType.Elem()
	sliceElemStructType, err := getSliceElemType(r.FileName, r.currentSheetName, rv)
	if err != nil {
		*errs = multierror.Append(*errs, err)
		return
	}
	tagMap := parseFiledTagSetting(sliceElemStructType)

	arr := reflect.MakeSlice(sliceType, 0, 0)
	for i := range sheetData.Rows {
		out := reflect.New(sliceElemStructType)

		if err := r.parseRowToStruct(i, sheetData, out, tagMap); err != nil {
			r.appendError(errs, err)
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

func (r *Reader) appendError(errs *error, err error) {
	if errs == nil {
		return
	}
	if err == nil {
		return
	}
	if r.errsMap == nil {
		r.errsMap = make(map[string]error)
	}

	if _, ok := r.errsMap[err.Error()]; ok {
		return
	}
	*errs = multierror.Append(*errs, err)
	r.errsMap[err.Error()] = err
}

// parse row to struct by tag setting
func (r *Reader) parseRowToStruct(
	rowIndex int, sheetData *SheetData, ve reflect.Value, tagMap map[string]TagSetting,
) (err error) {
	if ve.Kind() != reflect.Ptr {
		return NewError(r.FileName, r.currentSheetName, fmt.Sprintf("row %d", rowIndex), ErrorTypePointer)
	}
	if !ve.IsValid() {
		return NewError(r.FileName, r.currentSheetName, fmt.Sprintf("row %d", rowIndex), ErrorFieldInvalid)
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

		err = r.filedSetAll(ve.Elem().FieldByName(field.Name), cell, df)
		if err != nil {
			return err
		}

	}
	return err
}

func (r *Reader) filedSetAll(field reflect.Value, cell *Cell, defaultValue string) error {
	value := cell.Value
	if r.IsCheckEmpty && cell.IsEmpty {
		return NewError(r.FileName, r.currentSheetName, cell.Coordinates, ErrorFieldValueEmpty)
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
					err = r.filedSet(realVal.Elem(), value, cell)
					field.Set(realVal)
				} else {
					field.Set(reflect.Zero(valType))
				}
			}
		} else {
			err = NewError(r.FileName, r.currentSheetName, cell.Coordinates, ErrorFieldNotSet)
		}
	default:
		if field.CanSet() {
			err = r.filedSet(field, value, cell)
		} else {
			err = NewError(r.FileName, r.currentSheetName, cell.Coordinates, ErrorFieldNotSet)
		}
	}
	if err != nil {
		return err
	}
	return nil
}

func (r *Reader) filedSet(field reflect.Value, value string, cell *Cell) error {
	if !field.IsValid() {
		fmt.Println(field.Kind())
		return NewError(r.FileName, r.currentSheetName, cell.Coordinates, ErrorFieldNotSet)
	}
	if !field.CanSet() {
		return NewError(r.FileName, r.currentSheetName, cell.Coordinates, ErrorFieldNotSet)
	}
	switch field.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		intValue, err := strconv.Atoi(value)
		if err != nil {
			return NewError(r.FileName, r.currentSheetName, cell.Coordinates, ErrorFieldNotMatch)
		}
		field.SetInt(int64(intValue))
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		uintValue, err := strconv.ParseUint(value, 10, 32)
		if err != nil {
			return NewError(r.FileName, r.currentSheetName, cell.Coordinates, ErrorFieldNotMatch)
		}
		field.SetUint(uintValue)
	case reflect.String:
		field.SetString(value)
	case reflect.Bool:
		field.SetBool(sliceutil.InSlice(value, r.BoolTrueValues))
	case reflect.Float32, reflect.Float64:
		floatValue, err := strconv.ParseFloat(value, 32)
		if err != nil {
			return NewError(r.FileName, r.currentSheetName, cell.Coordinates, ErrorFieldNotMatch)
		}
		field.SetFloat(floatValue)
	default:
		return NewError(r.FileName, r.currentSheetName, cell.Coordinates, ErrorFieldTypeNotSupport)
	}
	return nil
}
func (r *Reader) getExcelFileData() (*Data, error) {
	fileName := r.FileName
	sheetMap := r.excelFile.GetSheetMap()
	if nil == sheetMap || len(sheetMap) == 0 {
		return nil, NewError(fileName, "", "", ErrorNoSheet)
	}

	sheetIndexData := make(map[int]*SheetData)
	sheetNameData := make(map[string]*SheetData)
	for sheetIndex, sheetName := range sheetMap {
		rows, err := r.excelFile.GetRows(sheetName)
		if err != nil {
			return nil, NewError(fileName, sheetName, fmt.Sprintf("sheet index %d", sheetIndex), err)
		}
		// 输入数据为excel直观的行数 从1开始
		sheetFields := rows[r.fieldHeadRowIndex-1]

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
		if !r.AllowFieldRepeat && repeatFieldName != "" {
			return nil, NewError(fileName, sheetName, fmt.Sprintf("sheet index %d", sheetIndex), ErrorFieldRepeat)
		}

		parseRows := make(map[int]map[string]*Cell, 0)
		for index, row := range rows {
			excelIndex := index + 1
			if excelIndex <= r.DataIndexOffset {
				continue
			}
			r.getRow(excelIndex, row, sheetFields, parseRows)
		}

		sheetData := &SheetData{
			RowTotal:        len(rows),
			DataTotal:       len(rows) - r.DataIndexOffset,
			SheetName:       sheetName,
			FileName:        fileName,
			Rows:            parseRows,
			FieldKeys:       sheetFields,
			DataIndexOffset: r.DataIndexOffset,
		}

		sheetIndexData[sheetIndex] = sheetData
		sheetNameData[sheetName] = sheetData
	}

	excelData := &Data{
		FileName:       fileName,
		SheetIndexData: sheetIndexData,
		SheetNameData:  sheetNameData,
		SheetTotal:     r.excelFile.SheetCount,
	}

	return excelData, nil
}

func (r *Reader) getRow(rowIndex int, rawRow, sheetFields []string, rowsData map[int]map[string]*Cell) {
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
		if r.IsCoordinatesABS {
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
			cell.IsEmpty = r.IsEmptyFunc(rawRow[colIndex])
		}
		cell.Coordinates = c
		rowData[fieldName] = cell
	}

	rowsData[rowIndex] = rowData
}

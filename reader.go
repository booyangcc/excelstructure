package excelstructure

import (
	"fmt"
	"reflect"
	"strconv"

	sliceutil "github.com/booyangcc/utils/sliceutil"
	"github.com/hashicorp/go-multierror"

	"github.com/xuri/excelize/v2"
)

// Parse parse.
func (p *Parser) Parse() (*Data, error) {
	fileName := p.FileName
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

// ReadWithSheetName parse with sheet index. start with 1
func (p *Parser) ReadWithSheetName(name string, output interface{}) error {
	excelData, err := p.Parse()
	if err != nil {
		return err
	}

	return p.readToStruct(name, excelData, output)
}

// Parser parse with sheet index 1
// output must be a pointer slice
// if the pointer field is pointer, and the value is empty ,the pointer field will be nil
func (p *Parser) Read(output interface{}) error {
	excelData, err := p.Parse()
	if err != nil {
		return err
	}
	return p.readToStruct("", excelData, output)
}

func (p *Parser) readToStruct(sheetName string, excelData *Data, output interface{}) (errs error) {
	if sliceutil.InSlice(sheetName, excelData.SheetList) {
		errs = multierror.Append(errs,
			NewError(p.FileName, "", fmt.Sprintf("sheetName %s", sheetName), ErrorSheetName))
		return
	}

	if sheetName == "" {
		if len(excelData.SheetList) == 0 {
			return
		}
		sheetName = excelData.SheetList[0]
	}

	p.currentSheetName = p.excelFile.GetSheetList()[0]

	sheetData := excelData.SheetNameData[sheetName]

	rv := reflect.ValueOf(output)
	if rv.Kind() == reflect.Ptr && rv.Elem().Kind() != reflect.Slice {
		errs = multierror.Append(errs,
			NewError(p.FileName, p.currentSheetName, "", ErrorInOutputType))
		return
	}

	sliceType := rv.Elem().Type()
	sliceElemType := sliceType.Elem()
	sliceElemStructType, err := getSliceElemType(p.FileName, p.currentSheetName, rv)
	if err != nil {
		errs = multierror.Append(errs, err)
		return
	}
	tagMap := parseFiledTagSetting(sliceElemStructType)

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

	return
}

func (p *Parser) appendError(errs error, err error) {
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
	errs = multierror.Append(errs, err)
	p.errsMap[err.Error()] = err
}

// parse row to struct by tag setting
func (p *Parser) parseRowToStruct(
	rowIndex int, sheetData *SheetData, ve reflect.Value, tagMap map[string]TagSetting,
) (err error) {
	if ve.Kind() != reflect.Ptr {
		return NewError(p.FileName, p.currentSheetName, fmt.Sprintf("row %d", rowIndex), ErrorTypePointer)
	}
	if !ve.IsValid() {
		return NewError(p.FileName, p.currentSheetName, fmt.Sprintf("row %d", rowIndex), ErrorFieldInvalid)
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

func (p *Parser) filedSetAll(field reflect.Value, cell *Cell, defaultValue string) error {
	value := cell.Value
	if p.IsCheckEmpty && cell.IsEmpty {
		return NewError(p.FileName, p.currentSheetName, cell.Coordinates, ErrorFieldValueEmpty)
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
			err = NewError(p.FileName, p.currentSheetName, cell.Coordinates, ErrorFieldNotSet)
		}
	default:
		if field.CanSet() {
			err = p.filedSet(field, value, cell)
		} else {
			err = NewError(p.FileName, p.currentSheetName, cell.Coordinates, ErrorFieldNotSet)
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
		return NewError(p.FileName, p.currentSheetName, cell.Coordinates, ErrorFieldNotSet)
	}
	if !field.CanSet() {
		return NewError(p.FileName, p.currentSheetName, cell.Coordinates, ErrorFieldNotSet)
	}
	switch field.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		intValue, err := strconv.Atoi(value)
		if err != nil {
			return NewError(p.FileName, p.currentSheetName, cell.Coordinates, ErrorFieldNotMatch)
		}
		field.SetInt(int64(intValue))
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		uintValue, err := strconv.ParseUint(value, 10, 32)
		if err != nil {
			return NewError(p.FileName, p.currentSheetName, cell.Coordinates, ErrorFieldNotMatch)
		}
		field.SetUint(uintValue)
	case reflect.String:
		field.SetString(value)
	case reflect.Bool:
		field.SetBool(sliceutil.InSlice(value, p.BoolTrueValues))
	case reflect.Float32, reflect.Float64:
		floatValue, err := strconv.ParseFloat(value, 32)
		if err != nil {
			return NewError(p.FileName, p.currentSheetName, cell.Coordinates, ErrorFieldNotMatch)
		}
		field.SetFloat(floatValue)
	default:
		return NewError(p.FileName, p.currentSheetName, cell.Coordinates, ErrorFieldTypeNotSupport)
	}
	return nil
}
func (p *Parser) getExcelFileData() (*Data, error) {
	fileName := p.FileName
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

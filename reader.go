package excelstructure

import (
	"fmt"
	"reflect"
	"strconv"

	sliceutil "github.com/booyangcc/utils/sliceutil"
	"github.com/hashicorp/go-multierror"
)

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
	tagMap := parseFieldTagSetting(sliceElemStructType)

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
		columnName, df, skip, serializer := field.Name, "", false, "json"
		if len(et) > 0 {
			val, ok := tagMap[field.Name]
			if ok {
				columnName = val.Column
				df = val.Default
				skip = val.Skip
				serializer = val.Serializer
			}
		}
		if skip {
			continue
		}

		cell, err1 := sheetData.GetCell(rowIndex, columnName)
		if err1 != nil {
			return err1
		}

		if p.IsCheckEmpty && cell.IsEmpty {
			return NewError(p.FileName, p.currentSheetName, cell.Coordinates, ErrorFieldValueEmpty)
		}

		fieldValue := ve.Elem().FieldByName(field.Name)
		fieldType := field.Type
		if fieldType.Kind() == reflect.Ptr {
			fieldType = fieldType.Elem()
		}
		fieldTypeKind := uint(fieldType.Kind())
		if (uint(reflect.Invalid) < fieldTypeKind && fieldTypeKind < uint(reflect.Float64)) ||
			fieldTypeKind == uint(reflect.String) {
			err = p.fieldSetAll(fieldValue, cell, df)
			if err != nil {
				return err
			}
		} else {
			err = p.fieldUmarshal(fieldValue, cell, serializer)
			if err != nil {
				return err
			}
		}

	}
	return err
}

func (p *Parser) fieldUmarshal(field reflect.Value, cell *Cell, serializerName string) error {
	fieldType := field.Type()
	newField := reflect.New(fieldType)

	var serializer Serializer
	var ok bool
	if IsDefaultSerializer(serializerName) {
		serializer = DefaultSerializer
	} else {
		if serializer, ok = p.serializers[serializerName]; !ok {
			return NewError(p.FileName, p.currentSheetName, cell.Coordinates, ErrorSerializerNotExist)
		}
	}

	if len(cell.Value) > 0 {
		err := serializer.Unmarshal(cell.Value, newField.Interface())
		if err != nil {
			return NewError(p.FileName, p.currentSheetName, cell.Coordinates, err)
		}
	}

	if !field.CanSet() {
		return NewError(p.FileName, p.currentSheetName, cell.Coordinates, ErrorFieldNotSet)
	}

	field.Set(newField.Elem())
	return nil
}

func (p *Parser) fieldSetAll(field reflect.Value, cell *Cell, defaultValue string) error {
	value := cell.Value
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
					err = p.fieldSet(realVal.Elem(), value, cell)
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
			err = p.fieldSet(field, value, cell)
		} else {
			err = NewError(p.FileName, p.currentSheetName, cell.Coordinates, ErrorFieldNotSet)
		}
	}
	if err != nil {
		return err
	}
	return nil
}

func (p *Parser) fieldSet(field reflect.Value, value string, cell *Cell) error {
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

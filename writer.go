package excelstructure

import (
	"fmt"
	"reflect"

	"github.com/hashicorp/go-multierror"
	"github.com/xuri/excelize/v2"
)

// Write  写入单个sheet
// sheetName sheet名称，为空则为结构体元素的类型+s
// input必须是slice，slice的元素必须是struct
func (p *Parser) Write(fileName, sheetName string, input interface{}) error {
	return p.WriteWithMultiSheet(fileName, map[string]interface{}{
		"": input,
	})
}

// WriteWithSheetName  写入单个sheet
// input必须是slice，slice的元素必须是struct
func (p *Parser) WriteWithSheetName(fileName, sheetName string, input interface{}) error {
	return p.WriteWithMultiSheet(fileName, map[string]interface{}{
		sheetName: input,
	})
}

// WriteWithMultiSheet 写入多个结构体到多个sheet，key为sheetName，value为slice
func (p *Parser) WriteWithMultiSheet(fileName string, inputMap map[string]interface{}) error {
	excelFile := excelize.NewFile()
	p.fileName = fileName

	for sheetName, input := range inputMap {
		p.currentSheetName = sheetName
		// 返回的错误是多个错误的集合，已经是封装过的故直接返回
		err := p.writeToSheet(excelFile, input)
		if err != nil {
			return err
		}
	}
	sheetList := excelFile.GetSheetList()
	err := excelFile.DeleteSheet(sheetList[0])
	if err != nil {
		return NewError(p.fileName, "", "", err)
	}

	if err = excelFile.SaveAs(p.fileName); err != nil {
		return NewError(p.fileName, "", "", err)
	}

	return nil
}

func (p *Parser) writeToSheet(excelFile *excelize.File, input interface{}) (errs error) {
	rv := reflect.Indirect(reflect.ValueOf(input))
	if rv.Kind() != reflect.Slice {
		errs = multierror.Append(errs,
			NewError(p.fileName, p.currentSheetName, "", ErrorInOutputType))
		return
	}

	sliceElemStructType, err := getSliceElemType(p.fileName, p.currentSheetName, rv)
	if err != nil {
		errs = multierror.Append(errs, err)
		return
	}

	if len(p.currentSheetName) == 0 {
		p.currentSheetName = fmt.Sprintf("%ss", sliceElemStructType.Name())
	}

	_, err = excelFile.NewSheet(p.currentSheetName)
	if err != nil {
		errs = multierror.Append(errs, err)
		return
	}

	tagMap := parseFieldTagSetting(sliceElemStructType)
	err = p.writeHead(excelFile, tagMap, sliceElemStructType)
	if err != nil {
		errs = multierror.Append(errs, err)
		return
	}

	err = p.writeData(excelFile, tagMap, rv)
	if err != nil {
		errs = multierror.Append(errs, err)
		return
	}

	return
}

func (p *Parser) writeData(ef *excelize.File, tagMap map[string]TagSetting, rv reflect.Value) error {
	for i := 0; i < rv.Len(); i++ {
		elemValue := rv.Index(i)
		elemType := elemValue.Type()
		if elemType.Kind() == reflect.Ptr {
			elemType = elemType.Elem()
			elemValue = elemValue.Elem()
		}

		rowData := make([]interface{}, 0, elemType.NumField())
		for j := 0; j < elemType.NumField(); j++ {
			field := elemType.Field(j)
			fieldTagSetting, ok := tagMap[field.Name]
			if !ok {
				fieldTagSetting = TagSetting{
					Column:     field.Name,
					Serializer: JSONSerializerName,
				}
			}
			if fieldTagSetting.Column == "-" || fieldTagSetting.Skip {
				continue
			}

			elemValueField := elemValue.Field(j)
			realElemValue := elemValueField.Interface()

			fieldType := field.Type.Kind()
			if fieldType == reflect.Ptr {
				fieldType = field.Type.Elem().Kind()
			}
			if fieldType == reflect.Slice || fieldType == reflect.Map || fieldType == reflect.Struct ||
				fieldType == reflect.Interface {
				if len(fieldTagSetting.Serializer) == 0 {
					fieldTagSetting.Serializer = JSONSerializerName
				}

				var serializer Serializer
				var ok bool
				if IsDefaultSerializer(fieldTagSetting.Serializer) {
					serializer = DefaultSerializer
				} else {
					serializer, ok = p.serializers[fieldTagSetting.Serializer]
					if !ok {
						return NewError(p.fileName, p.currentSheetName, "", ErrorSerializerNotExist)
					}
				}

				v, err := serializer.Marshal(realElemValue)
				if err != nil {
					return NewError(p.fileName, p.currentSheetName, "", err)
				}
				rowData = append(rowData, v)
			} else {
				if elemValueField.Kind() == reflect.Ptr {
					realElemValue = elemValueField.Elem().Interface()
				}
				if elemValueField.IsZero() && fieldTagSetting.Default != "" {
					realElemValue = fieldTagSetting.Default
				}
				rowData = append(rowData, realElemValue)
			}

		}

		coords, err := excelize.CoordinatesToCellName(1, p.DataIndexOffset+i+1)
		if err != nil {
			return err
		}

		err = ef.SetSheetRow(p.currentSheetName, coords, &rowData)
		if err != nil {
			return err
		}
	}
	return nil
}

func (p *Parser) writeHead(ef *excelize.File, tagMap map[string]TagSetting, sliceElemType reflect.Type) error {
	heads := make([]string, 0)
	comments := make([]string, 0)
	hasComment := false
	for _, tag := range tagMap {
		if tag.Comment != "" {
			hasComment = true
		}
	}

	for i := 0; i < sliceElemType.NumField(); i++ {
		field := sliceElemType.Field(i)
		fieldTagSetting, ok := tagMap[field.Name]
		if !ok {
			fieldTagSetting = TagSetting{
				Column: field.Name,
			}
		}
		if fieldTagSetting.Column == "-" || fieldTagSetting.Skip {
			continue
		}
		heads = append(heads, fieldTagSetting.Column)
		if hasComment {
			comments = append(comments, fieldTagSetting.Comment)
		}
	}

	err := ef.SetSheetRow(p.currentSheetName, "A1", &heads)
	if err != nil {
		return err
	}
	if hasComment {
		err = ef.SetSheetRow(p.currentSheetName, "A2", &comments)
		if err != nil {
			return err
		}
		p.DataIndexOffset = 2
	}
	p.hasComment = hasComment
	return nil
}

func getSliceElemType(fileName, currentSheetName string, rv reflect.Value) (reflect.Type, error) {
	sliceType := rv.Type().Elem()
	sliceElemType := sliceType.Elem()

	if sliceElemType.Kind() == reflect.Ptr {
		sliceElemType = sliceElemType.Elem()
	}
	if sliceElemType.Kind() != reflect.Struct {
		return nil, NewError(fileName, currentSheetName, "", ErrorSliceElemType)
	}

	return sliceElemType, nil
}

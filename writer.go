package excelstructure

import (
	"fmt"
	"github.com/hashicorp/go-multierror"
	"github.com/xuri/excelize/v2"
	"reflect"
)

// Write struct to excel
func (p *Parser) Write(input interface{}) error {
	return p.writeToExcel(p.FileName, input)
}

func (p *Parser) writeToExcel(fileName string, input interface{}) (errs error) {
	rv := reflect.Indirect(reflect.ValueOf(input))
	if rv.Kind() != reflect.Slice {
		errs = multierror.Append(errs,
			NewError(p.FileName, p.currentSheetName, "", ErrorInOutputType))
		return
	}

	if rv.Len() == 0 {
		return
	}

	sliceElemStructType, err := getSliceElemType(p.FileName, p.currentSheetName, rv)
	if err != nil {
		errs = multierror.Append(errs, err)
		return
	}
	tagMap := parseFiledTagSetting(sliceElemStructType)

	excelFile := excelize.NewFile()

	sheetName := fmt.Sprintf("%ss", sliceElemStructType.Name())
	_, err = excelFile.NewSheet(sheetName)
	if err != nil {
		errs = multierror.Append(errs, err)
		return
	}
	p.currentSheetName = sheetName

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

	err = excelFile.DeleteSheet(excelFile.GetSheetList()[0])
	if err != nil {
		errs = multierror.Append(errs, err)
		return
	}

	if err = excelFile.SaveAs(fileName); err != nil {
		errs = multierror.Append(errs,
			NewError(p.FileName, p.currentSheetName, "", err))
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
		rowData := make([]interface{}, 0)
		fmt.Println(elemValue.String(), elemType.String(), elemType.NumField())
		for j := 0; j < elemType.NumField(); j++ {
			field := elemType.Field(i)
			filedTagSetting, ok := tagMap[field.Name]
			if !ok {
				filedTagSetting = TagSetting{
					Column: field.Name,
				}
			}
			if filedTagSetting.Column == "-" || filedTagSetting.Skip {
				continue
			}
			realElemValue := elemValue.Field(j).Interface()
			if elemValue.Field(j).Kind() == reflect.Ptr {
				realElemValue = elemValue.Field(j).Elem().Interface()
			}
			if elemValue.Field(j).IsZero() && filedTagSetting.Default != "" {
				realElemValue = filedTagSetting.Default
			}
			rowData = append(rowData, realElemValue)
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
		filedTagSetting, ok := tagMap[field.Name]
		if !ok {
			filedTagSetting = TagSetting{
				Column: field.Name,
			}
		}
		if filedTagSetting.Column == "-" || filedTagSetting.Skip {
			continue
		}
		heads = append(heads, filedTagSetting.Column)
		if hasComment {
			comments = append(comments, filedTagSetting.Comment)
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

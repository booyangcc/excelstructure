package excelstructure

import (
	"fmt"
	"github.com/hashicorp/go-multierror"
	"github.com/xuri/excelize/v2"
	"reflect"
)

// Writer write.
type Writer struct {
	FileName         string
	currentSheetName string
	errsMap          map[string]error
	// AllowFieldRepeat 允许表头字段重复
	AllowFieldRepeat bool
	hasComment       bool
	dataRowOffset    int
}

// NewWriter 传入文件名
func NewWriter(fileName string) *Writer {
	return &Writer{
		FileName:      fileName,
		errsMap:       make(map[string]error),
		dataRowOffset: 1,
	}
}

// Write struct to excel
func (w *Writer) Write(input interface{}) error {
	return w.writeToExcel(w.FileName, input)
}

func (w *Writer) writeToExcel(fileName string, input interface{}) (errs error) {
	rv := reflect.Indirect(reflect.ValueOf(input))
	if rv.Kind() != reflect.Slice {
		errs = multierror.Append(errs,
			NewError(w.FileName, w.currentSheetName, "", ErrorInOutputType))
		return
	}

	if rv.Len() == 0 {
		return
	}

	sliceElemStructType, err := getSliceElemType(w.FileName, w.currentSheetName, rv)
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
	w.currentSheetName = sheetName

	err = w.writeHead(excelFile, tagMap, sliceElemStructType)
	if err != nil {
		errs = multierror.Append(errs, err)
		return
	}
	err = w.writeData(excelFile, tagMap, rv)
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
			NewError(w.FileName, w.currentSheetName, "", err))
	}
	return
}

func (w *Writer) writeData(ef *excelize.File, tagMap map[string]TagSetting, rv reflect.Value) error {
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

		coords, err := excelize.CoordinatesToCellName(1, w.dataRowOffset+i+1)
		if err != nil {
			return err
		}
		err = ef.SetSheetRow(w.currentSheetName, coords, &rowData)
		if err != nil {
			return err
		}
	}
	return nil
}

func (w *Writer) writeHead(ef *excelize.File, tagMap map[string]TagSetting, sliceElemType reflect.Type) error {
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

	err := ef.SetSheetRow(w.currentSheetName, "A1", &heads)
	if err != nil {
		return err
	}
	if hasComment {
		err = ef.SetSheetRow(w.currentSheetName, "A2", &comments)
		if err != nil {
			return err
		}
		w.dataRowOffset = 2
	}
	w.hasComment = hasComment
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

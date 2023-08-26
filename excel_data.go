package excelstructure

import (
	"fmt"
	"strconv"

	sliceutil "github.com/booyangcc/utils/sliceutil"
	"github.com/hashicorp/go-multierror"
)

// Cell excel cell info
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
	// RowTotal  row total.
	RowTotal int
	// DataTotal data total.
	DataTotal int
	SheetName string
	FileName  string
	// Rows  map[rowIndex]map[keyField]*Cell
	Rows      map[int]map[string]*Cell
	FieldKeys []string
	// DataIndexOffset data index offset.
	DataIndexOffset int
}

// Data excel data.
type Data struct {
	FileName       string
	SheetIndexData map[int]*SheetData
	SheetNameData  map[string]*SheetData
	SheetTotal     int
	SheetList      []string
}

// GetCell get cell.
func (s *SheetData) GetCell(rowIndex int, fieldKey string, isCheckEmpty ...bool) (*Cell, error) {
	if s.DataTotal < rowIndex-s.DataIndexOffset {
		return nil, NewError(s.FileName, s.SheetName, fmt.Sprintf("rowIndex %d", rowIndex), ErrorDataRowOutOfRange)
	}
	if rowIndex == 1 {
		return nil, NewError(s.FileName, s.SheetName, fmt.Sprintf("rowIndex %d", rowIndex), ErrorRowIndexIsHeader)
	}
	row := s.Rows[rowIndex]
	if !sliceutil.InSlice(fieldKey, s.FieldKeys) {
		return nil, NewError(s.FileName, s.SheetName,
			fmt.Sprintf("rowIndex %d, fieldKey %s", rowIndex, fieldKey), ErrorFieldNotExist)
	}
	cell, ok := row[fieldKey]
	if !ok {
		return nil, NewError(s.FileName, s.SheetName,
			fmt.Sprintf("rowIndex %d, fieldKey %s", rowIndex, fieldKey), ErrorFieldNotExist)
	}

	isCheck := false
	if len(isCheckEmpty) > 0 {
		isCheck = isCheckEmpty[0]
	}

	if isCheck && cell.IsEmpty {
		return nil, NewError(s.FileName, s.SheetName,
			fmt.Sprintf("rowIndex %d, fieldKey %s", rowIndex, fieldKey), ErrorFieldValueEmpty)
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
		return 0, NewError(s.FileName, s.SheetName, v.Coordinates, ErrorFieldNotMatch)
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
		errs = multierror.Append(errs, NewError(s.FileName, s.SheetName, v.Coordinates, ErrorFieldNotMatch))
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

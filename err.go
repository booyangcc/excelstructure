package excelstructure

import (
	"errors"
	"fmt"
)

var (
	// ErrorDataRowOutOfRange data row out of range
	ErrorDataRowOutOfRange = errors.New("out of range")
	// ErrorRowIndexIsHeader row index is header
	ErrorRowIndexIsHeader = errors.New("row index must > 1, because row 1 is header")
	// ErrorSheetName sheet name invalid
	ErrorSheetName = errors.New("sheet name not exist")
	// ErrorInOutputType output type invalid
	ErrorInOutputType = errors.New("output type must be slice or struct pointer")
	// ErrorSliceElemType elem type must struct or struct pointer
	ErrorSliceElemType = errors.New("slice elem type must struct or struct pointer")
	// ErrorTypePointer output not pointer
	ErrorTypePointer = errors.New("output type must be pointer")
	// ErrorNoSheet no sheet
	ErrorNoSheet = errors.New("no sheet")

	// ErrorFieldNotExist Field invalid
	ErrorFieldNotExist = errors.New("field not exist in header")
	// ErrorFieldRepeat repeat field
	ErrorFieldRepeat = errors.New("field repeat")
	// ErrorFieldTypeNotSupport field type not support
	ErrorFieldTypeNotSupport = errors.New("field type not support")
	// ErrorFieldNotMatch Field type not match
	ErrorFieldNotMatch = errors.New("field tag default value or excel value not match struct field type")
	// ErrorFieldNotSet pointer field can not set
	ErrorFieldNotSet = errors.New("pointer field can not set")
	// ErrorFieldInvalid pointer field invalid
	ErrorFieldInvalid = errors.New("pointer or field invalid")
	// ErrorFieldValueEmpty  field empty
	ErrorFieldValueEmpty = errors.New("value is empty")
	// ErrorNoData no data
	ErrorNoData = errors.New("no data")

	// ErrorSerializerNameRepeat serializer name repeat
	ErrorSerializerNameRepeat = errors.New("serializer name repeat")
	// ErrorSerializerHandlerEmpty serializer handler empty
	ErrorSerializerHandlerEmpty = errors.New("serializer marshal or unmarshal handler empty")
	// ErrorSerializerNotExist serializer not exist
	ErrorSerializerNotExist = errors.New("serializer not exist")
)

// Error excel structure error
type Error struct {
	FileName    string
	SheetName   string
	Coordinates string
	Err         error
}

// NewError new error
func NewError(filName string, sheetName string, coordinates string, err error) error {
	return &Error{
		FileName:    filName,
		SheetName:   sheetName,
		Err:         err,
		Coordinates: coordinates,
	}
}

// Error error
func (e *Error) Error() string {
	return fmt.Sprintf("fileName: %s, SheetName: %s, Coordinates: %s , ErrMsg: %s",
		e.FileName, e.SheetName, e.Coordinates, e.Err.Error())
}

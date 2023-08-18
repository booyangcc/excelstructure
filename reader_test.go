package excelstructure

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

type Info struct {
	Name    string  `excel:"column:user_name;"`
	Phone   *string `excel:"column:phone;"`
	Age     string  `excel:"column:age;"`
	Man     bool    `excel:"column:man;default:true"`
	Address string  `excel:"column:address;skip"` // skip this field
}

func TestParser_UnmarshalWithSheetIndex(t *testing.T) {
	var info []*Info
	p := NewReader("./test_excel_file/test.xlsx")
	err := p.UnmarshalWithSheetIndex(2, &info)
	if err != nil {
		assert.Error(t, err)
	}
	require.Equal(t, 4, len(info))
	require.Equal(t, "booyang1", info[0].Name)
}

func TestParser_Unmarshal(t *testing.T) {
	var info []*Info
	p := NewReader("./test_excel_file/test.xlsx")
	err := p.Unmarshal(&info)
	if err != nil {
		assert.Error(t, err)
	}
	require.Equal(t, 4, len(info))
	require.Equal(t, "booyang", info[0].Name)
}

// excel data index offset is 2, the first two rows are not data,
// first row is title,second row is comment
func TestParser_UnmarshalWithComment(t *testing.T) {
	var info []*Info
	p := NewReader("./test_excel_file/test_with_comment.xlsx")
	p.DataIndexOffset = 2
	err := p.Unmarshal(&info)
	if err != nil {
		assert.Error(t, err)
	}
	require.Equal(t, 4, len(info))
	require.Equal(t, "booyang", info[0].Name)
}

func TestParser_UnmarshalWithCheckEmpty(t *testing.T) {
	var info []*Info
	p := NewReader("./test_excel_file/test_check_empty.xlsx")
	p.IsCheckEmpty = true
	err := p.Unmarshal(&info)
	// one row is empty, so the error is not nil, "C3 field age value is empty"
	assert.Error(t, err)
	// length is 4, but one row is empty, so the length is 3
	require.Equal(t, 3, len(info))
	require.Equal(t, "booyang", info[0].Name)
}

func TestSheetData_GetIntValue1(t *testing.T) {
	p := NewReader("./test_excel_file/test.xlsx")
	data, err := p.Parse()
	if err != nil {
		fmt.Println(err)
	}
	s := data.SheetIndexData[1]
	row2, err := s.GetIntValue(2, "age")
	assert.NoError(t, err)
	require.Equal(t, 20, row2)

	row3, err := s.GetIntValueWithMultiError(3, "age", err)
	assert.NoError(t, err)
	require.Equal(t, 3, row3)
}

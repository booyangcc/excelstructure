package excelstructure

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func Test_GetIntValue1(t *testing.T) {
	p := NewParser("./test_excel_file/test.xlsx")
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

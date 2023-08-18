package excelstructure

import (
	"github.com/booyangcc/utils/convutil"
	"github.com/stretchr/testify/assert"
	"testing"
)

func TestWriter_Marshal(t *testing.T) {
	w := NewWriter("./test_excel_file/test_write.xlsx")
	infos := []*Info{
		{
			Name:    "booyang",
			Phone:   convutil.String("123456789"),
			Age:     "18",
			Man:     true,
			Address: "beijing",
		},
		{
			Name:    "booyang1",
			Phone:   convutil.String("123456789"),
			Age:     "14",
			Man:     false,
			Address: "shanghai",
		},
	}
	err := w.Marshal(infos)
	assert.NoError(t, err)
}

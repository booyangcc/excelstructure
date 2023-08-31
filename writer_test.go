package excelstructure

import (
	"testing"

	"github.com/booyangcc/utils/convutil"
	"github.com/stretchr/testify/assert"
)

func TestWriter_Write(t *testing.T) {
	w := NewParser()
	infos := []*Info{
		{
			Name:    "booyang",
			Phone:   convutil.String("123456789"),
			Age:     "18",
			Man:     true,
			Address: []string{"beijing", "shanghai"},
		},
		{
			Name:    "booyang1",
			Phone:   convutil.String("123456789"),
			Age:     "14",
			Man:     false,
			Address: []string{"beijing", "shanghai"},
		},
	}
	err := w.Write("./test_excel_file/test_write.xlsx", "infos", infos)
	assert.NoError(t, err)
}

func TestWriter_WriteMulti(t *testing.T) {
	w := NewParser()
	_ = w.RegisterSerializer("mySerializer", mySerializer)
	infos1 := []*Info{
		{
			Name:    "booyang_sheet1",
			Phone:   convutil.String("123456789"),
			Age:     "18",
			Man:     true,
			Address: []string{"beijing", "shanghai"},
		},
		{
			Name:    "booyang1_sheet1",
			Phone:   convutil.String("123456789"),
			Age:     "14",
			Man:     false,
			Address: []string{"beijing", "shanghai"},
		},
	}

	infos2 := []*Info{
		{
			Name:    "booyang_sheet2",
			Phone:   convutil.String("123456789"),
			Age:     "18",
			Man:     true,
			Address: []string{"beijing", "shanghai"},
		},
		{
			Name:    "booyang1_sheet2",
			Phone:   convutil.String("123456789"),
			Age:     "14",
			Man:     false,
			Address: []string{"beijing", "shanghai"},
		},
	}
	err := w.WriteWithMultiSheet("./test_excel_file/test_write_multi.xlsx", map[string]interface{}{
		"infos1": infos1,
		"infos2": infos2,
	})
	assert.NoError(t, err)
}

// type Detail struct {
// 	Height int    `json:"height"`
// 	Weight int    `json:"weight"`
// 	Nation string `json:"nation"`
// }

// type Person struct {
// 	Name    string   `excel:"column:user_name;comment:person name"`
// 	Age     int      `excel:"column:age;"`
// 	Man     bool     `excel:"column:man;default:true"`
// 	Address []string `excel:"column:address;serializer"`
// 	Details *Detail  `excel:"column:details;serializer:json"`
// }

func Test_WriterSerializer(t *testing.T) {
	persons := []*Person{
		{
			Name:    "booyang",
			Age:     18,
			Man:     true,
			Address: []string{"beijing", "shanghai"},
			Detail: Detail{
				Height: 180,
				Weight: 70,
				Nation: "China",
			},
		},
		{
			Name:    "bob",
			Age:     17,
			Man:     true,
			Address: []string{"Lundon", "New York"},
			Detail: Detail{
				Height: 181,
				Weight: 60,
				Nation: "Britain",
			},
		},
	}
	w := NewParser()
	err := w.Write("./test_excel_file/test_serializer_write.xlsx", "", persons)
	assert.NoError(t, err)
}

package excelstructure

import (
	"testing"

	"github.com/booyangcc/utils/convutil"
	"github.com/stretchr/testify/assert"
)

func TestWriter_Marshal(t *testing.T) {
	w := NewParser("./test_excel_file/test_write.xlsx")
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
	err := w.Write(infos)
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
	w := NewParser("./test_excel_file/test_serializer_write.xlsx")
	err := w.Write(persons)
	assert.NoError(t, err)
}

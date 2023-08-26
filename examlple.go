package excelstructure

import (
	"encoding/json"
	"fmt"

	"github.com/booyangcc/utils/convutil"
)

type Detail struct {
	Height int    `json:"height"`
	Weight int    `json:"weight"`
	Nation string `json:"nation"`
}

type Info struct {
	Name    string   `excel:"column:user_name;comment:person name"`
	Phone   *string  `excel:"column:phone;comment:phone number"`
	Age     string   `excel:"column:age;"`
	Man     bool     `excel:"column:man;default:true"`
	Address []string `excel:"column:address;serializer:mySerializer"`
	Detail  Detail   `excel:"column:details;serializer:mySerializer"` // u can use custom serializer,default json serializer
}

var (
	mySerializer = Serializer{
		Marshal: func(v interface{}) (string, error) {
			bs, err := json.Marshal(v)
			if err != nil {
				return "", err
			}
			return string(bs), nil
		},
		Unmarshal: func(s string, v interface{}) error {
			return json.Unmarshal([]byte(s), v)
		},
	}
)

func TestWriteRead() {
	infos := []*Info{
		{
			Name:    "booyang",
			Phone:   convutil.String("123456789"),
			Age:     "18",
			Man:     true,
			Address: []string{"beijing", "shanghai"},
			Detail: Detail{
				Height: 180,
				Weight: 70,
				Nation: "China",
			},
		},
		{
			Name:    "booyang1",
			Phone:   convutil.String("123456789"),
			Age:     "14",
			Man:     false,
			Address: []string{"guangzhou", "xian"},
			Detail: Detail{
				Height: 181,
				Weight: 60,
				Nation: "Britain",
			},
		},
	}
	p := NewParser("./test_excel_file/test_write.xlsx")
	// use custom serializer
	p.RegisterSerializer("mySerializer", mySerializer)
	err := p.Write(infos)
	if err != nil {
		fmt.Println(err)
	}

	// because the struct field has comment tag, so the comment has been written to row 2, so when read, data offset is 2
	p.DataIndexOffset = 2
	var newInfo []*Info
	err = p.Read(&newInfo)
	if err != nil {
		fmt.Println(err)
	}
	fmt.Printf("oldInfo: %+v, newInfo: %+v", infos, newInfo)
}

func TestParse() {
	p := NewParser("./test_excel_file/test_write.xlsx")
	// use custom serializer
	p.RegisterSerializer("mySerializer", mySerializer)
	// because the struct field has comment tag, so the comment has been written to row 2, so when read, data offset is 2
	p.DataIndexOffset = 2
	excelData, err := p.Parse()
	if err != nil {
		fmt.Println(err)
	}

	s := excelData.SheetNameData["Infos"]
	row3UserName, err := s.GetStringValue(3, "user_name")
	if err != nil {
		fmt.Println(err)
	}

	row3age, err := s.GetIntValue(3, "age")
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println(row3UserName, row3age)

	row4UserName, err := s.GetStringValue(4, "user_name")
	if err != nil {
		fmt.Println(err)
	}

	row4age, err := s.GetIntValue(4, "age")
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println(row4UserName, row4age)
}

[中文文档](https://github.com/booyangcc/excelstructure/blob/main/README_zh.md)

## Introduction

excelstructure is a tool for converting between Excel and Go struct. It can convert a struct to Excel and vice versa. It also supports getting the value of a single field from a row.

## Basic Usage

### Installation

`go get github.com/booyangcc/excelstructure`

### Simple Usage

```golang
infos := []*Info{
    {
        Name:    "booyang",
        Phone:   convutil.String("123456789"),
    },
}

p := NewParser()
// Because the struct fields have comment tags, the data starts from the third row when writing, so the data offset when reading is 2
p.DataIndexOffset = 2
// Write
err = p.Write("./test_excel_file/test_write.xlsx", "Infos", infos)
if err != nil {
    fmt.Println(err)
}

var newInfo []*Info
// Read
err = p.Read("./test_excel_file/test_write.xlsx", &newInfo)
if err != nil {
    fmt.Println(err)
}

```

### Complete Example

```golang
package main

import (
	"fmt"
	"github.com/booyangcc/utils/convutil"
    "github.com/booyangcc/excelstructure"
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
	Detail  Detail   `excel:"column:details;serializer:mySerializer"` // You can use custom serializer, default use json, here we use our custom serializer mySerializer
}

var (
	// Custom serializer
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

// TestWriteRead Usage 1, write and read to struct
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
	p := NewParser()
	// Register custom serializer
	err := p.RegisterSerializer("mySerializer", mySerializer)
	if err != nil {
		return
	}
	// Write to single sheet
	err = p.Write("./test_excel_file/test_write.xlsx", "Infos", infos)
	if err != nil {
		fmt.Println(err)
	}

	/*
		// Write to two sheets, Info1 and Info2, with the same data
		err = p.WriteWithMultiSheet("./test_excel_file/test_write_multi.xlsx", map[string]interface{}{
			"Info1": infos,
			"Info2": infos,
		})
		if err != nil {
			fmt.Println(err)
		}
	*/
	// Because the struct fields have comment tags, the data starts from the third row when writing, so the data offset when reading is 2
	p.DataIndexOffset = 2
	var newInfo []*Info

	err = p.Read("./test_excel_file/test_write.xlsx", &newInfo)
	if err != nil {
		fmt.Println(err)
	}
	/*
		// Read from multiple sheets, read Info1 and Info2 to newInfo1 and newInfo2
		var newInfo1 []*Info
		var newInfo2 []*Info
		err = p.ReadWithMultiSheet("./test_excel_file/test_write_multi.xlsx", map[string]interface{}{
			"Info1": &newInfo1,
			"Info2": &newInfo2,
		})
	*/
	if err != nil {
		fmt.Println(err)
	}

	fmt.Printf("oldInfo: %+v, newInfo: %+v", infos, newInfo)
}

// TestParse Usage 2, directly get the data of a column in a row
func TestParse() {
	p := NewParser()
	// Use custom serializer
	err := p.RegisterSerializer("mySerializer", mySerializer)
	if err != nil {
		return
	}
	// Because the struct fields have comment tags, the data starts from the third row when writing, so the data offset when reading is 2
	p.DataIndexOffset = 2
	excelData, err := p.Parse("./test_excel_file/test_write.xlsx")
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

```

***excel data*** output excel
![输出数据](./imgs/test_write.png)

Usage
### Tag Configuration
The tag name is `excel`. The tag field configuration is as follows:
> Name string `excel:"column:user_name;comment:person name;skip;default:boo;serializer:mySerializer"`
- column: the header to parse or write to Excel
- comment: if any field in the struct contains this configuration in the excel tag, the second row of the output Excel file will be a comment
- skip: indicates that the current field is skipped and not parsed or written to Excel
- default: if the field is zero-value, use the default value instead
- serializer: serialization and deserialization of structures, slices, interfaces, and other types, supporting customization, default is json serializer

### Parser Usage
Parser parameters:
- FileName: the file to read
- DataIndexOffset: the data index offset. If the first row is a header, the offset is 1. If there is a comment occupying a row, the offset is 2. The default value is 1.
- BoolTrueValues: the optional values for true boolean values. The default values are [true,True,TRUE,1,是,yes,Yes,YES,y,Y]. You can specify them manually.
- IsCheckEmpty: whether to check for empty values when serializing to a struct. If a value is empty, an error is thrown.
- IsEmptyFunc: the callback function to check for empty values. The default function is func(v string) bool {return v==""}
- IsCoordinatesABS: the type of cell coordinate value. If true, the coordinate is A1. If false, the coordinate is 1.
- ExcelData: the parsed data values
- AllowFieldRepeat: whether to allow duplicate fields. If true, the fields will be overwritten.


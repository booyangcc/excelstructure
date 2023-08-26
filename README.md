[中文文档](https://github.com/booyangcc/excelstructure/blob/main/README_zh.md)

## Introduction

excelstructure is a tool for converting between Excel and Go struct. It can convert a struct to Excel and vice versa. It also supports getting the value of a single field from a row.

## Basic Usage

### Installation

`go get github.com/booyangcc/excelstructure`

### Example

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
	Detail Detail   `excel:"column:details;serializer:mySerializer"` // u can use custom serializer,default json serializer
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
### Parse Struct to Excel
```golang
p := excelstructure.NewParser("./test_write.xlsx")
p.RegisterSerializer("mySerializer", mySerializer)
err := p.Write(infos)  
if err != nil {  
    fmt.Println(err)  
}  
```
### Parsing Excel to Struct
#### Method 1: Parse Directly to Struct
```golang
p := NewParser("./test_write.xlsx")  
p.RegisterSerializer("mySerializer", mySerializer)

// because the struct field has comment tag, so the comment has been written to row 2, so when read, data offset is 2
p.DataIndexOffset = 2
var newInfo []*Info
err = p.Read(&newInfo)
if err != nil {
        fmt.Println(err)
}
```
#### Method 2: Get the Value of a Single Field in a Row
```golang
p := NewParser("./test_write.xlsx")
// Since the struct being written has a comment field, the second row in the Excel file will be a comment, so the data index offset is 2
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
```






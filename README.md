## Introduction

Excel is a library to unmarshal excel data to struct, and alse suppoert get row data.

## Basic Usag

### Installation

`go get github.com/booyangcc/excelstructure`

### example

```golang
package main

import (
	"fmt"
	"github.com/booyangcc/utils/convutil"
    "github.com/booyangcc/excelstructure"
)

type Info struct {
	Name    string  `excel:"column:user_name;comment:person name"`
	Phone   *string `excel:"column:phone;comment:phone number"`
	Age     string  `excel:"column:age;"`
	Man     bool    `excel:"column:man;default:true"`
	Address string  `excel:"column:address;skip"` // skip this field
}

var (
	infos = []*Info{
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
)

func TestWriteRead() {
	p := excelstructure.NewParser("./test_write.xlsx")
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
	r := excelstructure.NewParser("./test_write.xlsx")
	// because the struct field has comment tag, so the comment has been written to row 2, so when read, data offset is 2
	r.DataIndexOffset = 2
	excelData, err := r.Parse()
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






## Introduction

Excel is a library to unmarshal excel data to struct, and alse suppoert get row data.

## Basic Usag

### Installation

`go get github.com/booyangcc/excel`

### example

Need parse excel data

| user_name | phone  | age | man |
| --------- | ------ | --- | --- |
| booyang   | 232323 | 20  | Y   |
| bob       | 22222  | 3   | Y   |
| tom       | 111111 | 18  | Y   |
| sandy     | 33333  | 18  | N   |

unmarshal struct

```golang
package main

import (
    "fmt"

    "github.com/booyangcc/excel"
)

type Info struct {
	Name    string  `excel:"column:user_name;"`
	Phone   *string `excel:"column:phone;"`
	Age     string  `excel:"column:age;"`
	Man     bool    `excel:"column:man;default:true"`
	Address string  `excel:"column:address;skip"` // skip this field, if no this setting, will throw error, filed addree not in excel header
}

func main() {
    // unmarshal excel to struct
    var info []*Info
    p := excel.NewParser("./test_excel_file/test.xlsx")
    err := p.Unmarshal(&info)
    if err != nil {
        fmt.Println(err)
        return
    }
    fmt.Fprintln("%+v", info)
    
    excelData = p.ExcelData.SheetIndexData[1]
    // get row data
    row2, err := s.GetIntValue(2, "age")
    if err != nil {
        fmt.Println(err)
    }
    // 20
    fmt.Fprintln("%+v", row2)

    row3, err := s.GetIntValueWithMultiError(3, "age", err)
    if err != nil {
        fmt.Println(err)
    }
    // 3
    fmt.Fprintln("%+v", row3)
}

```






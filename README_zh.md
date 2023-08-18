## 介绍

excelstructure是一个excel结构体互转的工具，可以将结构体转换为excel，也可以将excel转换为结构体，同时也支持单行字段值获取

## 基础使用

### 安装

`go get github.com/booyangcc/excelstructure`

### 样例

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
	Address string  `excel:"column:address;skip"` // skip 代表这个字段不解析也不写入excel
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

	// 因为写入的结构体有comment字段故第二行为comment，所以读取的时候偏移量为2
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
	// 因为写入的结构体有comment字段故第二行为comment，所以读取的时候偏移量为2
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
## 使用

### tag配置
tag字段设置,tag名称`excel`
> Name string `excel:"column:user_name;comment:person name;skip;default:boo"`
- column：解析或写入excel的head头
- comment：任意一结构体的字段exceltag 包含了这个配置，则输出excel的时候第二行为comment
- skip：标注当前字段跳过，不解析也不写入excel
- default：解析或设置如果字段为零值则使用default替换


### parser使用
parser参数
- FileName 读的文件
- DataIndexOffset 数据索引偏移量,第一行可能为表头，则偏移量为1，如果有注释占用一行，则为2。默认为1
- BoolTrueValues bool值为true的可选项，默认为[true,True,TRUE,1,是,yes,Yes,YES,y,Y]。使用时可以手动指定
- IsCheckEmpty 序列化到结构提的时候是否检测空值如果为空值则报错
- IsEmptyFunc 检测是否为空的回调函数，默认为 `func(v string) bool  {return v==""}`
- IsCoordinatesABS cell坐标值类型 ，ture返回坐标A1, false为$A$1
- ExcelData 解析出来的数据值
- AllowFieldRepeat 是否允许重复字段允许则覆盖
### 解析结构体到excel
```golang
p := excelstructure.NewParser("./test_write.xlsx")
err := p.Write(infos)  
if err != nil {  
    fmt.Println(err)  
}  
```
### 解析excel到结构体
#### 使用方式一，直接解析到结构体
```golang
p := NewParser("./test_write.xlsx")
// 因为写入的结构体有comment字段故第二行为comment，所以读取的时候偏移量为2
p.DataIndexOffset = 2
var newInfo []*Info
err = p.Read(&newInfo)
if err != nil {
        fmt.Println(err)
}
```
#### 使用方式二，获取单行某个字段的值
```golang
p := NewParser("./test_write.xlsx")  
// 因为写入的结构体有comment字段故第二行为comment，所以读取的时候偏移量为2
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






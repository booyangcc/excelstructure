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
	Detail  Detail   `excel:"column:details;serializer:mySerializer"` // 你可以使用自定义序列化，默认使用json，此处使用我们的自定义序列化mySerializer
}

var (
	// 自定的序列化
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

// TestWriteRead 使用方式1，解析写入到结构体
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
	// 注册自定义序列化
	err := p.RegisterSerializer("mySerializer", mySerializer)
	if err != nil {
		return
	}
	// 写入单个
	err = p.Write("./test_excel_file/test_write.xlsx", "Infos", infos)
	if err != nil {
		fmt.Println(err)
	}

	/*
		// 写入两个sheet， 分别为Info1和Info2，数据相同
		err = p.WriteWithMultiSheet("./test_excel_file/test_write_multi.xlsx", map[string]interface{}{
			"Info1": infos,
			"Info2": infos,
		})
		if err != nil {
			fmt.Println(err)
		}
	*/
	// 因为结构体字段有comment标签，所以写入的时候，数据从第三行开始，所以读取的时候，数据偏移量为2
	p.DataIndexOffset = 2
	var newInfo []*Info

	err = p.Read("./test_excel_file/test_write.xlsx", &newInfo)
	if err != nil {
		fmt.Println(err)
	}
	/*
		// 多sheet读取，读取sheet Info1和Info2到newInfo1和newInfo2
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

// TestParse 使用方式二，直接获取行的某一列数据
func TestParse() {
	p := NewParser()
	// 使用自定义序列化
	err := p.RegisterSerializer("mySerializer", mySerializer)
	if err != nil {
		return
	}
	// 因为结构体字段有comment标签，所以写入的时候，数据从第三行开始，所以读取的时候，数据偏移量为2
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

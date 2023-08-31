package excelstructure

import (
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func Test_ParseReadWithSheetIndex(t *testing.T) {
	var info []*Info
	p := NewParser()
	err := p.ReadWithSheetName("./test_excel_file/test.xlsx", "Sheet1", &info)
	if err != nil {
		assert.Error(t, err)
	}
	require.Equal(t, 4, len(info))
	require.Equal(t, "booyang1", info[0].Name)
}

func Test_ParseRead(t *testing.T) {
	var info []*Info
	p := NewParser()
	err := p.Read("./test_excel_file/test.xlsx", &info)
	if err != nil {
		assert.Error(t, err)
	}
	require.Equal(t, 4, len(info))
	require.Equal(t, "booyang", info[0].Name)
}

// excel data index offset is 2, the first two rows are not data,
// first row is title,second row is comment
func Test_ParseReadWithComment(t *testing.T) {
	var info []*Info
	p := NewParser()
	p.DataIndexOffset = 2
	err := p.Read("./test_excel_file/test_with_comment.xlsx", &info)
	if err != nil {
		assert.Error(t, err)
	}
	require.Equal(t, 4, len(info))
	require.Equal(t, "booyang", info[0].Name)
}

func Test_ParseReadWithCheckEmpty(t *testing.T) {
	var info []*Info
	p := NewParser()
	p.IsCheckEmpty = true
	err := p.Read("./test_excel_file/test_check_empty.xlsx", &info)
	// one row is empty, so the error is not nil, "C3 field age value is empty"
	assert.Error(t, err)
	// length is 4, but one row is empty, so the length is 3
	require.Equal(t, 3, len(info))
	require.Equal(t, "booyang", info[0].Name)
}

type Person struct {
	Name    string   `excel:"column:user_name;comment:person name"`
	Age     int      `excel:"column:age;"`
	Man     bool     `excel:"column:man;default:true"`
	Address []string `excel:"column:address;serializer"`
	Detail  Detail   `excel:"column:details;serializer:json"`
}

func Test_ParseReadWithSerializer(t *testing.T) {
	var persons []*Person
	p := NewParser()
	err := p.Read("./test_excel_file/test_serializer.xlsx", &persons)
	assert.NoError(t, err)
	require.Equal(t, 2, len(persons))
	require.Equal(t, "booyang1", persons[0].Name)
	// ["陕西省西安市雁塔区","陕西省延安市宝塔区"]
	require.Equal(t, 2, len(persons[0].Address))
	// {"height":180,"weight":70,"nation":"China"}
	require.Equal(t, "China", persons[0].Detail.Nation)
}

func Test_ParseReadWithMultiSheet(t *testing.T) {
	var infos1, infos2 []*Info
	p := NewParser()
	_ = p.RegisterSerializer("mySerializer", mySerializer)
	p.DataIndexOffset = 2
	err := p.ReadWithMultiSheet("./test_excel_file/test_write_multi.xlsx", map[string]interface{}{
		"infos1": &infos1,
		"infos2": &infos2,
	})
	assert.NoError(t, err)
	require.Equal(t, "booyang_sheet1", infos1[0].Name)
	require.Equal(t, "booyang_sheet2", infos2[0].Name)
}

package excel

import (
	"reflect"
	"strings"
)

const (
	// TagName tag name.
	TagName = "excel"
)

// TagSetting tag setting.
type TagSetting struct {
	// column
	Column string
	// 	Type string
	Type string
	// 	default value
	Default   string
	Comment   string
	OmitEmpty bool
	Skip      bool
}

func parseTagSetting(str, sep, kvSep string) map[string]string {
	settings := map[string]string{}
	names := strings.Split(str, sep)

	for i := 0; i < len(names); i++ {
		values := strings.Split(names[i], kvSep)
		k := strings.TrimSpace(strings.ToLower(values[0]))

		if len(values) >= 2 {
			settings[k] = values[1]
		} else if k != "" {
			settings[k] = k
		}
	}

	return settings
}

func parseFiledTagSetting(sliceElemType reflect.Type) (map[string]TagSetting, error) {
	tagFiledMap := make(map[string]TagSetting)

	for i := 0; i < sliceElemType.NumField(); i++ {
		field := sliceElemType.Field(i)
		tag := field.Tag.Get(TagName)
		if _, ok := tagFiledMap[field.Name]; ok {
			continue
		}
		kvm := parseTagSetting(tag, ";", ":")
		tagFiled := TagSetting{
			Column:    kvm["column"],
			Type:      kvm["type"],
			Default:   kvm["default"],
			Comment:   kvm["comment"],
			OmitEmpty: kvm["omitempty"] == "omitempty",
			Skip:      kvm["skip"] == "skip",
		}
		tagFiledMap[field.Name] = tagFiled
	}

	return tagFiledMap, nil
}

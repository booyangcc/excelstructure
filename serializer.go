package excelstructure

import "encoding/json"

const (
	// JSONSerializerName json serializer name
	JSONSerializerName = "json"
	// SerializerName serializer name
	SerializerName = "serializer"
)

// DefaultSerializer default serializer
var DefaultSerializer = Serializer{
	Marshal:   DefaultMarshal,
	Unmarshal: DefaultUnmarshal,
}

// MarshalFunc marshal func
type MarshalFunc func(v interface{}) (string, error)

// UnmarshalFunc unmarshal func
type UnmarshalFunc func(data string, v interface{}) error

type Serializer struct {
	Marshal   MarshalFunc
	Unmarshal UnmarshalFunc
}

// NewSerializer new serializer, to parse excel cell data to struct field
func NewSerializer(marshal MarshalFunc, unmarshal UnmarshalFunc) Serializer {
	return Serializer{
		Marshal:   marshal,
		Unmarshal: unmarshal,
	}
}

// DefaultMarshal default marshal, json marshal
func DefaultMarshal(v interface{}) (string, error) {
	vbs, err := json.Marshal(v)
	if err != nil {
		return "", err
	}
	return string(vbs), nil

}

// DefaultUnmarshal default unmarshal, json unmarshal
func DefaultUnmarshal(data string, v interface{}) error {
	err := json.Unmarshal([]byte(data), v)
	if err != nil {
		return err
	}
	return nil
}

// IsDefaultSerializer is default serializer
func IsDefaultSerializer(serializerName string) bool {
	return serializerName == JSONSerializerName || serializerName == SerializerName
}

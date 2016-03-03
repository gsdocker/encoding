package encoding

import (
	"fmt"
	"reflect"
	"testing"
)

type A struct{}

func TestReflect(t *testing.T) {
	var val []*A

	fmt.Printf("type :%v\n", reflect.ValueOf(val).Type().Elem())
}

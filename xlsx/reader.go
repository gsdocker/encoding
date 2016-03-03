package xlsx

import (
	"fmt"
	"reflect"
	"regexp"
	"strconv"
	"strings"

	"github.com/gsdocker/gserrors"
	"github.com/gsdocker/gslogger"
	x "github.com/tealeg/xlsx"
)

// UnmarshalF .
type UnmarshalF func(reflect.Value, string) error

// ErrUnmarshalField .
type ErrUnmarshalField struct {
	Key   string
	Type  reflect.Type
	Field reflect.StructField
}

func (e *ErrUnmarshalField) Error() string {
	return "json: cannot unmarshal object key " + strconv.Quote(e.Key) + " into unexported field " + e.Field.Name + " of type " + e.Type.String()
}

// ErrInvalidUnmarshal .
type ErrInvalidUnmarshal struct {
	Type reflect.Type
}

func (e *ErrInvalidUnmarshal) Error() string {
	if e.Type == nil {
		return "xlsx: Unmarshal(nil)"
	}

	if e.Type.Kind() != reflect.Ptr {
		return "xlsx: Unmarshal(non-pointer " + e.Type.String() + ")"
	}
	return "xlsx: Unmarshal(nil " + e.Type.String() + ")"
}

// RowReader row reader
type RowReader struct {
	gslogger.Log                           // mixin logger
	Sheet        string                    // sheet name
	nameMapping  map[string]string         // name mapping
	unmarshalers map[string]UnmarshalF     // unmarshal functions
	pattern      map[string]*regexp.Regexp // column pattern
	Split        string                    // split chars
	header       *x.Row                    // current row
	row          *x.Row                    // current row
	id           int                       // row id
}

func (reader *Reader) newRowReader(name string, header, row *x.Row, id int) *RowReader {
	return &RowReader{
		nameMapping:  reader.NameMapping,
		unmarshalers: reader.Unmarshalers,
		pattern:      reader.Pattern,
		Log:          reader.Log,
		Sheet:        name,
		header:       header,
		row:          row,
		Split:        ",",
	}
}

func (reader *RowReader) Read(val interface{}) (err error) {

	defer func() {
		if e := recover(); e != nil {
			err = gserrors.Newf(nil, "catch panic :%v", e)
		}
	}()

	rv := reflect.ValueOf(val)

	if rv.Kind() != reflect.Ptr || rv.IsNil() {
		return &ErrInvalidUnmarshal{reflect.TypeOf(val)}
	}

	if rv.Elem().IsNil() {
		rv = rv.Elem()
		rv.Set(reflect.New(rv.Type().Elem()))
	}

	if rv.Elem().Kind() != reflect.Struct {
		return &ErrInvalidUnmarshal{reflect.TypeOf(val)}
	}

	rv = reflect.Indirect(rv)

	for i, cell := range reader.row.Cells {
		colname := reader.header.Cells[i].Value
		key := fmt.Sprintf("%s.%s", reader.Sheet, colname)

		if name, ok := reader.nameMapping[key]; ok {
			colname = name
			key = fmt.Sprintf("%s.%s", reader.Sheet, name)
		}

		if reader.unmarshalers != nil {
			if f, ok := reader.unmarshalers[key]; ok {
				if err := f(reflect.Indirect(rv), cell.Value); err != nil {
					return gserrors.Newf(err, "can't conv cell[%s:%d] '%s'", colname, reader.id, cell.Value)
				}
				continue
			}
		}

		field := rv.FieldByName(colname)

		if !field.IsValid() {
			reader.W("can't unmarshal col(%s)", colname)
			continue
		}

		if reader.readBuiltinType(key, cell.Value, field) {
			continue
		}

	}

	return nil
}

func (reader *RowReader) readBuiltinType(colname string, val string, assign reflect.Value) bool {

	switch assign.Type().Kind() {
	case reflect.Bool:
		if val == "true" || val == "1" {
			assign.SetBool(true)
		} else {
			assign.SetBool(false)
		}

	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		v, err := strconv.ParseInt(val, 0, 64)

		if err != nil {
			gserrors.Panicf(err, "can't conv cell[%s:%d] '%s' to int", colname, reader.id, val)
		}

		assign.SetInt(v)
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:

		v, err := strconv.ParseUint(val, 0, 64)

		if err != nil {
			gserrors.Panicf(err, "can't conv cell[%s:%d] '%s' to uint", colname, reader.id, val)
		}

		assign.SetUint(v)

	case reflect.Float32, reflect.Float64:

		val, err := strconv.ParseFloat(val, 64)

		if err != nil {
			gserrors.Panicf(err, "can't conv cell[%s:%d] '%s' to float", colname, reader.id, val)
		}

		assign.SetFloat(val)

	case reflect.String:
		assign.SetString(val)
	case reflect.Array:
	case reflect.Slice:

		pattern, ok := reader.pattern[colname]

		if !ok {
			gserrors.Panicf(nil, "can't conv %s(%d), not found convert pattern", colname, reader.id)
		}

		subs := strings.Split(val, reader.Split)

		slice := reflect.MakeSlice(assign.Type(), 0, len(subs))

		subType := assign.Type().Elem()

		if subType.Kind() == reflect.Ptr {
			subType = subType.Elem()
		}

		for _, sub := range subs {
			matched := pattern.FindStringSubmatch(sub)

			if matched == nil {

				if sub != "" {
					gserrors.Panicf(nil, "can't conv cell[%s:%d] '%s'", colname, reader.id, val)
				}

				continue
			}

			subval := reflect.New(subType)

			for i, match := range matched[1:] {

				if match == "" {
					continue
				}

				name := fmt.Sprintf("%s.%s", colname, subType.Field(i).Name)
				reader.readBuiltinType(name, match, reflect.Indirect(subval).Field(i))
			}

			slice = reflect.Append(slice, subval)
		}

		assign.Set(slice)

	default:
		return false
	}

	return true
}

// Reader xlsx reader
type Reader struct {
	gslogger.Log                           // mixin log
	file         *x.File                   // xlsx file
	Pattern      map[string]*regexp.Regexp // subtype pattern
	Unmarshalers map[string]UnmarshalF     // unmarshal functions
	NameMapping  map[string]string         // name mapping
}

// NewReader create new xlsx file reader
func NewReader(filename string) (*Reader, error) {
	file, err := x.OpenFile(filename)

	if err != nil {
		return nil, gserrors.Newf(err, "create new xlsx reader error :%s", filename)
	}

	return &Reader{
		Log:  gslogger.Get("xlsx"),
		file: file,
	}, err
}

// Read read all rows
func (reader *Reader) Read(sheetName string) (rows []*RowReader) {

	var sheet *x.Sheet
	ok := false

	for _, sheet = range reader.file.Sheets {
		if sheet.Name == sheetName {
			ok = true
			break
		}
	}

	if !ok {
		return nil
	}

	if len(sheet.Rows) < 2 {
		return nil
	}

	header := sheet.Rows[0]

	rows = make([]*RowReader, len(sheet.Rows)-1)

	for i, row := range sheet.Rows[1:] {
		rows[i] = reader.newRowReader(sheetName, header, row, i)
	}

	return
}

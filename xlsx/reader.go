package xlsx

import (
	"reflect"
	"strconv"

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
	gslogger.Log                       // mixin logger
	NameMapping  map[string]string     // name mapping
	Unmarshals   map[string]UnmarshalF // unmarshal functions
	header       *x.Row                // current row
	row          *x.Row                // current row
	id           int                   // row id
}

func (reader *Reader) newRowReader(header, row *x.Row, id int) *RowReader {
	return &RowReader{
		Log:    reader.Log,
		header: header,
		row:    row,
	}
}

func (reader *RowReader) Read(val interface{}) (err error) {

	defer func() {
		if e := recover(); e != nil {
			err = gserrors.Newf(nil, "catch unknown err :%v", e)
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

	valtype := reflect.Indirect(rv).Type()

	rv = reflect.Indirect(rv)

	for i, cell := range reader.row.Cells {
		colname := reader.header.Cells[i].Value

		if name, ok := reader.NameMapping[colname]; ok {
			colname = name
		}

		if reader.Unmarshals != nil {
			if f, ok := reader.Unmarshals[colname]; ok {
				if err := f(reflect.Indirect(rv), cell.Value); err != nil {
					return gserrors.Newf(err, "can't conv cell[%s:%d] '%s'", colname, reader.id, cell.Value)
				}
				continue
			}
		}

		fieldType, ok := valtype.FieldByName(colname)

		if !ok {
			reader.W("can't unmarshal col(%s)", colname)
			continue
		}

		field := rv.FieldByName(colname)

		switch fieldType.Type.Kind() {
		case reflect.Bool:
			val := cell.Value

			if val == "true" || val == "1" {
				field.SetBool(true)
			} else {
				field.SetBool(false)
			}

		case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
			val, err := cell.Int64()

			if err != nil {
				return gserrors.Newf(err, "can't conv cell[%s:%d] '%s' to int", colname, reader.id, cell.Value)
			}

			field.SetInt(val)
		case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:

			val, err := cell.Int64()

			if err != nil {
				return gserrors.Newf(err, "can't conv cell[%s:%d] '%s' to uint", colname, reader.id, cell.Value)
			}

			field.SetUint(uint64(val))

		case reflect.Float32, reflect.Float64:

			val, err := cell.Float()

			if err != nil {
				return gserrors.Newf(err, "can't conv cell[%s:%d] '%s' to float", colname, reader.id, cell.Value)
			}

			field.SetFloat(val)

		case reflect.String:
			field.SetString(cell.Value)
		default:
			reader.W("can't unmarshal col(%s)", colname)
		}

	}

	return nil
}

// Reader xlsx reader
type Reader struct {
	gslogger.Log         // mixin log
	file         *x.File // xlsx file
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
		rows[i] = reader.newRowReader(header, row, i)
	}

	return
}

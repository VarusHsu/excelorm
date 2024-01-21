package excelorm

import (
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"reflect"
	"strconv"
	"time"
)

type Option func(*options)

// Build 生成excel文件
// example usage:
// * define a struct
//
//	type User struct {
//		 ID        int64      `excel_header:"id"`
//		 Name      string     `excel_header:"name"`
//		 Age       int        `excel_header:"age"`
//	}
//
// * implement SheetModel interface
//
//	func (u User) SheetName() string {
//		 return "user"
//	}
//
// * append data to excel file
//
//	var sheetModels []SheetModel{
//		 User{
//		   ID:   1,
//		   Name: "张三",
//		   Age:  18,
//		 },
//		 User{
//		   ID:   2,
//		   Name: "李四",
//		   Age:  20,
//		 },
//	}
//
// * build Excel file
// err := Build("user.xlsx", sheetModels, WithTimeFormatLayout("2006/01/02 15:04:05"))
//
//	if err != nil {
//		 log.Fatal(err)
//	}
//
// * Multi-sheets
// define more structs which implement SheetModel interface
// then construct any of their objects to append to sheetModels
// different sheetModel better have different sheet name to avoid confusion
// rows ordered in Excel file is the same as sheetModels
func Build(fileName string, sheetModels []SheetModel, opts ...Option) error {
	// default options
	options := &options{
		timeFormatLayout: "2006-01-02 15:04:05",
		floatPrecision:   2,
		floatFmt:         'f',
		ifNullValue:      "",
	}

	// apply options
	for _, opt := range opts {
		opt(options)
	}

	f := excelize.NewFile()
	sheetLinesCount := make(map[string]int)
	for _, sheetModel := range sheetModels {
		sheetName := sheetModel.SheetName()
		if sheetName == "" {
			return errors.New("sheetModel must have a sheet name")
		}

		modelKind := reflect.TypeOf(sheetModel).Kind()
		switch modelKind {
		case reflect.Struct:
			l := sheetLinesCount[sheetName]
			err := appendRow(f, sheetModel, l, options)
			if err != nil {
				return err
			}
			sheetLinesCount[sheetName]++
			if l == 0 { // first line is header, so counter increase again
				sheetLinesCount[sheetName]++
			}
		default:
			return errors.New("sheetModel must be struct")
		}
	}
	f.DeleteSheet("Sheet1")
	return f.SaveAs(fileName)
}

type SheetModel interface {
	SheetName() string
}

type options struct {
	timeFormatLayout string // time.Time, *time.Time 的格式化版图
	floatPrecision   int    // 小数保留多少位
	floatFmt         byte   // 小数的格式，默认为'f',详细见 strconv.FormatFloat 的注释
	ifNullValue      string // null pointer		空值的默认显示
}

func WithTimeFormatLayout(layout string) Option {
	return func(options *options) {
		options.timeFormatLayout = layout
	}
}

func WithFloatPrecision(precision int) Option {
	return func(options *options) {
		options.floatPrecision = precision
	}
}

func WithFloatFmt(fmt byte) Option {
	return func(options *options) {
		options.floatFmt = fmt
	}
}

func WithIfNullValue(value string) Option {
	return func(options *options) {
		options.ifNullValue = value
	}
}

func appendRow(f *excelize.File, sheetModel SheetModel, line int, options *options) error {
	sheetName := sheetModel.SheetName()
	// find if sheetName exists
	sheetIndex := f.GetSheetIndex(sheetName)
	if sheetIndex == 0 {
		sheetIndex = f.NewSheet(sheetName) // create sheet
	}
	modelType := reflect.TypeOf(sheetModel)
	line++         // index start from 0 but excel start from 1
	if line == 1 { // set header
		for i := 0; i < modelType.NumField(); i++ {
			field := modelType.Field(i)
			header := field.Tag.Get("excel_header")
			if header == "" { // if no excel_header tag, use field name as header
				header = field.Name
			}
			cellName, err := coordinatesToCellName(i+1, 1)
			if err != nil {
				return err
			}
			f.SetCellValue(sheetName, cellName, header) // set header
		}
		line++ // set data first line
	}
	for i := 0; i < modelType.NumField(); i++ {
		field := modelType.Field(i)
		cellName, err := coordinatesToCellName(i+1, line)
		if err != nil {
			return err
		}

		fieldValue := reflect.ValueOf(sheetModel).Field(i) // get field value
		fieldKind := field.Type.Kind()                     // get field kind
	unAddrTo:
		switch fieldKind {
		case reflect.Pointer: // if field is pointer, get its value
			canAddr := fieldValue.Elem().CanAddr() // check if can get its value
			if !canAddr {
				f.SetCellValue(sheetName, cellName, options.ifNullValue) // null pointer
			} else {
				fieldValue = reflect.Indirect(fieldValue) // get value of pointer point to
				fieldKind = fieldValue.Kind()             // get kind of pointer point to
				goto unAddrTo                             // jump to unAddrTo, because now field is not pointer
			}
		case reflect.Struct, reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64,
			reflect.String, reflect.Bool, reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64,
			reflect.Float32, reflect.Float64:
			value := fieldValue.Interface() // get field value (type interface{})
			switch value.(type) {           // type assertion
			case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64, bool, string:
				f.SetCellValue(sheetName, cellName, value) // set cell value
			case float32: // convert float32 to string using options
				f.SetCellValue(sheetName,
					cellName,
					strconv.FormatFloat(
						float64(value.(float32)),
						options.floatFmt,
						options.floatPrecision,
						32,
					),
				)
			case float64: // convert float64 to string using options
				f.SetCellValue(sheetName,
					cellName,
					strconv.FormatFloat(
						value.(float64),
						options.floatFmt,
						options.floatPrecision,
						64,
					),
				)
			case time.Time: // convert time.Time to string using options
				f.SetCellValue(sheetName, cellName, value.(time.Time).Format(options.timeFormatLayout))
			default:
				return fmt.Errorf("unsupported type %T", value)
			}

		case reflect.Map, reflect.Slice, reflect.Array, reflect.Chan, reflect.Func, reflect.Interface,
			reflect.Invalid, reflect.UnsafePointer, reflect.Complex64, reflect.Complex128, reflect.Uintptr:
			return fmt.Errorf("unsupported type %s", fieldKind)
		}
	}
	return nil
}

// next code is copied and modified from https://github.com/360EntSecGroup-Skylar/excelize

// coordinatesToCellName converts [X, Y] coordinates to alpha-numeric cell
// name or returns an error.
// egs:
//
//	excelize.coordinatesToCellName(1, 1) // returns "A1", nil
func coordinatesToCellName(col, row int) (string, error) {
	const totalRows = 1048576
	if col < 1 || row < 1 {
		return "", fmt.Errorf("invalid cell reference [%d, %d]", col, row)
	}
	if row > totalRows {
		return "", errors.New("row number exceeds maximum limit")
	}
	sign := ""
	colName, err := columnNumberToName(col)
	return sign + colName + sign + strconv.Itoa(row), err
}

// columnNumberToName provides a function to convert the integer to Excel
// sheet column title.
func columnNumberToName(num int) (string, error) {
	const (
		minColumns = 1
		maxColumns = 16384
	)
	if num < minColumns || num > maxColumns {
		return "", fmt.Errorf("the column number must be greater than or equal to %d and less than or equal to %d", minColumns, maxColumns)
	}
	var col string
	for num > 0 {
		col = string(rune((num-1)%26+65)) + col
		num = (num - 1) / 26
	}
	return col, nil
}

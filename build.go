package excelorm

import (
	"bytes"
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"reflect"
	"strconv"
	"time"
)

type Option func(*options)

// WriteExcelSaveAs 生成excel文件并保存到本地
// example usage:
//	//define a struct
//	type Foo struct {
//		ID        int64      `excel_header:"id"`
//		Name      string     `excel_header:"name"`
//		CreatedAt time.Time  `excel_header:"created_at"`
//		DeletedAt *time.Time `excel_header:"deleted_at"`
//	}
//	// implement SheetModel interface
//	func (u Foo) SheetName() string {
//		return "foo sheet name"
//	}
//	//append data to excel file
//	bar1DeletedAt := time.Date(2024, 1, 3, 15, 4, 5, 0, time.Local)
//	sheetModels := []excelorm.SheetModel{
//		Foo{
//			ID:   1,
//			Name: "Bar1",
//			CreatedAt: time.Date(2024, 1, 2, 15, 4, 5, 0, time.Local),
//			DeletedAt: &bar1DeletedAt,
//		},
//		Foo{
//			ID:   2,
//			Name: "Bar2",
//			CreatedAt: time.Date(2024, 1, 2, 15, 4, 5, 0, time.Local),
//		},
//	}
//	//build Excel file
// 	if err := excelorm.WriteExcelSaveAs("foo.xlsx", sheetModels,
//		excelorm.WithTimeFormatLayout("2006/01/02 15:04:05"),
//		excelorm.WithIfNullValue("-"),
//	); err != nil {
//		 log.Fatal(err)
//	}
//  // After that code execute, you will get `foo.xlsx` file with named `foo sheet name`,
//	// It's content like next:
//	+-------------------------------------------------------+
//	| id | name |          created_at |          deleted_at |
//	+-------------------------------------------------------+
//	|  1 | Bar1 | 2024/01/02 15:04:05 | 2024/01/03 15:04:05 |
//	|  2 | Bar2 | 2024/01/02 15:04:05 |                   - |
//	+-------------------------------------------------------+
//	// Multi-sheets
//	// define more structs which implement SheetModel interface
//	// then construct any of their objects to append to sheetModels
//	// different sheetModel better have different sheet name to avoid confusion
//	// rows ordered in Excel file is the same as sheetModels
func WriteExcelSaveAs(fileName string, sheetModels []SheetModel, opts ...Option) error {
	time.Date(2024, 1, 2, 15, 4, 5, 0, time.Local)
	if fileName == "" {
		return errors.New("fileName can not be empty")
	}
	f, err := write(sheetModels, opts...)
	if err != nil {
		return err
	}
	return f.SaveAs(fileName)
}

func write(sheetModels []SheetModel, opts ...Option) (*excelize.File, error) {
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
		if sheetModel == nil {
			return nil, errors.New("nil reference row append is not allowed")
		}
		sheetName := sheetModel.SheetName()
		if sheetName == "" {
			return nil, errors.New("sheetModel must have a sheet name")
		}

		modelKind := reflect.TypeOf(sheetModel).Kind()
		switch modelKind {
		case reflect.Struct:
			l := sheetLinesCount[sheetName]
			err := appendRow(f, sheetModel, l, options)
			if err != nil {
				return nil, err
			}
			sheetLinesCount[sheetName]++
			if l == 0 && !options.headless { // first line is header, so counter increase again
				sheetLinesCount[sheetName]++
			}
		default:
			return nil, errors.New("sheetModel must be struct")
		}
	}
	err := setNoDataSheetHeaders(f, options)
	if err != nil {
		return nil, err
	}
	// delete default sheet
	var containsModelSheetNameEqSheet1 bool
	for _, sheetModel := range sheetModels {
		if sheetModel.SheetName() == "Sheet1" {
			containsModelSheetNameEqSheet1 = true
			break
		}
	}
	for _, sheetModel := range options.sheetHeaders {
		if sheetModel.SheetName() == "Sheet1" {
			containsModelSheetNameEqSheet1 = true
			break
		}
	}
	if !containsModelSheetNameEqSheet1 {
		f.DeleteSheet("Sheet1")
	}
	return f, nil
}

func setNoDataSheetHeaders(f *excelize.File, options *options) error {
	models := options.sheetHeaders
	if len(models) == 0 {
		return nil
	}
	for _, model := range models {
		sheetName := model.SheetName()
		idx := f.GetSheetIndex(sheetName)
		if idx != 0 {
			// sheet exists, continue
			continue
		}
		f.NewSheet(sheetName)

		// check if sheetModel is pointer
		if reflect.TypeOf(model).Kind() == reflect.Ptr {
			if reflect.ValueOf(model).Elem().CanAddr() { // check if sheetModel is nil
				// replace to sheetModel's reference value
				// if type(model) is SheetModel, then *model is still SheetModel
				model = reflect.Indirect(reflect.ValueOf(model)).Interface().(SheetModel)
			} else {
				return errors.New("nil reference row append is not allowed")
			}
		}

		modelType := reflect.TypeOf(model)
		for i := 0; i < modelType.NumField(); i++ {
			field := modelType.Field(i)
			header := field.Tag.Get("excel_header")
			if header == "" { // if no excel_header tag, use field name as header
				header = field.Name
			} else if header == "-" {
				continue // skip this field if header is "-"
			}
			cellName, err := coordinatesToCellName(i+1, 1)
			if err != nil {
				return err
			}
			f.SetCellValue(sheetName, cellName, header) // set header
		}
	}
	return nil
}

// WriteExcelAsBytesBuffer generate excel and save as excelize.File
func WriteExcelAsBytesBuffer(sheetModels []SheetModel, opts ...Option) (*bytes.Buffer, error) {
	buffer := new(bytes.Buffer)
	f, err := write(sheetModels, opts...)
	if err != nil {
		return nil, err
	}
	err = f.Write(buffer)
	if err != nil {
		return nil, err
	}
	return buffer, nil
}

type SheetModel interface {
	SheetName() string
}

type options struct {
	timeFormatLayout string       // time.Time, *time.Time 的格式化版图
	floatPrecision   int          // 小数保留多少位
	floatFmt         byte         // 小数的格式，默认为'f',详细见 strconv.FormatFloat 的注释
	ifNullValue      string       // null pointer		空值的默认显示
	sheetHeaders     []SheetModel // 当没有数据时，表头的默认显示
	trueValue        *string      // bool类型的true显示值
	falseValue       *string      // bool类型的false显示值
	integerAsString  bool         // int类型的字段是否以字符串形式显示(避免excel自动转为科学计数法)
	headless         bool         // 是否显示表头
}

// WithTimeFormatLayout 时间类型的格式化版图
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

// WithIfNullValue 当数据为nil时展示内容
func WithIfNullValue(value string) Option {
	return func(options *options) {
		options.ifNullValue = value
	}
}

// WithSheetHeaders 当没有数据时，默认也要展示表头
func WithSheetHeaders(headers ...SheetModel) Option {
	return func(options *options) {
		options.sheetHeaders = headers
	}
}

// WithBoolValueAs 当字段类型为bool时，true和false的展示内容
func WithBoolValueAs(trueValue, falseValue string) Option {
	return func(options *options) {
		options.trueValue = &trueValue
		options.falseValue = &falseValue
	}
}

// WithIntegerAsString int类型的字段是否以字符串形式显示(避免excel自动转为科学计数法)
func WithIntegerAsString() Option {
	return func(options *options) {
		options.integerAsString = true
	}
}

// WithHeadless 不显示表头
func WithHeadless() Option {
	return func(options *options) {
		options.headless = true
	}
}

func appendRow(f *excelize.File, sheetModel SheetModel, line int, options *options) error {
	sheetName := sheetModel.SheetName()
	// find if sheetName exists
	sheetIndex := f.GetSheetIndex(sheetName)
	if sheetIndex == 0 {
		f.NewSheet(sheetName) // create sheet
	}
	// check if sheetModel is pointer
	if reflect.TypeOf(sheetModel).Kind() == reflect.Ptr {
		if reflect.ValueOf(sheetModel).Elem().CanAddr() { // check if sheetModel is nil
			// replace to sheetModel's reference value
			// if type(sheetModel) is SheetModel, then *sheetModel is still SheetModel
			sheetModel = reflect.Indirect(reflect.ValueOf(sheetModel)).Interface().(SheetModel)
		} else {
			return errors.New("nil reference row append is not allowed")
		}
	}

	modelType := reflect.TypeOf(sheetModel)
	line++                              // index start from 0 but excel start from 1
	if line == 1 && !options.headless { // set header
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
			case int, int8, int16, int32, int64:
				if options.integerAsString {
					f.SetCellValue(sheetName, cellName, strconv.FormatInt(fieldValue.Int(), 10)) // set int cell value
				} else {
					f.SetCellValue(sheetName, cellName, value)
				}
			case uint, uint8, uint16, uint32, uint64:
				if options.integerAsString {
					f.SetCellValue(sheetName, cellName, strconv.FormatUint(fieldValue.Uint(), 10)) // set uint cell value
				} else {
					f.SetCellValue(sheetName, cellName, value)
				}
			case string:
				f.SetCellValue(sheetName, cellName, value) // set string cell value
			case bool: // convert bool to string using options
				b := value.(bool)
				if options.trueValue != nil && b { // if trueValue is set and value is true
					f.SetCellValue(sheetName, cellName, *options.trueValue)
				} else if options.falseValue != nil && !b { // if falseValue is set and value is false
					f.SetCellValue(sheetName, cellName, *options.falseValue)
				} else { // using default
					f.SetCellValue(sheetName, cellName, value)
				}
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
	colName, err := columnNumberToName(col)
	return colName + strconv.Itoa(row), err
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

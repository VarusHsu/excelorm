// Copyright (c) 2026 Varus Hsu
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of
// this software and associated documentation files (the "Software"), to deal in
// the Software without restriction, including without limitation the rights to
// use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
// the Software, and to permit persons to whom the Software is furnished to do so,
// subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
// FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
// COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
// IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
// CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

package excelorm

import (
	"bytes"
	"errors"
	"fmt"
	"reflect"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

type Option func(*options)

// WriteExcelSaveAs creates an Excel file and saves it to local storage.
// Example usage:
//
//	// Define a struct.
//	type Foo struct {
//		ID        int64      `excel_header:"id"`
//		Name      string     `excel_header:"name"`
//		CreatedAt time.Time  `excel_header:"created_at"`
//		DeletedAt *time.Time `excel_header:"deleted_at"`
//	}
//	// Implement the SheetModel interface.
//	func (u Foo) SheetName() string {
//		return "foo sheet name"
//	}
//	// Append data to the Excel file.
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
//	// Build and save the Excel file.
//	if err := excelorm.WriteExcelSaveAs("foo.xlsx", sheetModels,
//		excelorm.WithTimeFormatLayout("2006/01/02 15:04:05"),
//		excelorm.WithIfNullValue("-"),
//	); err != nil {
//		 log.Fatal(err)
//	}
//	// After this code runs, you will get a `foo.xlsx` file with a sheet named `foo sheet name`.
//	// Its content looks like this:
//	+-------------------------------------------------------+
//	| id | name |          created_at |          deleted_at |
//	+-------------------------------------------------------+
//	|  1 | Bar1 | 2024/01/02 15:04:05 | 2024/01/03 15:04:05 |
//	|  2 | Bar2 | 2024/01/02 15:04:05 |                   - |
//	+-------------------------------------------------------+
//	// Multiple sheets:
//	// Define more structs that implement the SheetModel interface,
//	// then append their objects to sheetModels.
//	// Different sheet models should use different sheet names to avoid confusion.
//	// Row order in the Excel file matches the order in sheetModels.
func WriteExcelSaveAs(fileName string, sheetModels []SheetModel, opts ...Option) error {
	if fileName == "" {
		return errors.New("fileName can not be empty")
	}
	f, err := write(sheetModels, opts...)
	if err != nil {
		return err
	}
	saveErr := f.SaveAs(fileName)
	closeErr := f.Close()
	return errors.Join(saveErr, closeErr)
}

func write(sheetModels []SheetModel, opts ...Option) (*excelize.File, error) {
	// Default options.
	options := &options{
		timeFormatLayout: "2006-01-02 15:04:05",
		floatPrecision:   2,
		floatFmt:         'f',
	}

	// Apply options.
	for _, opt := range opts {
		opt(options)
	}

	f := excelize.NewFile()
	cleanup := true
	defer func() {
		if cleanup {
			_ = f.Close()
		}
	}()
	swMap := make(map[string]*excelize.StreamWriter)
	lineNumMap := make(map[string]int)
	for _, sheetModel := range sheetModels {
		if _, err := modelStructValue(sheetModel); err != nil {
			return nil, err
		}
		sheetName := sheetModel.SheetName()
		if sheetName == "" {
			return nil, errors.New("sheetModel must have a sheet name")
		}

		line := lineNumMap[sheetName]
		sw, ok := swMap[sheetName]
		if !ok {
			var err error
			if _, err = f.NewSheet(sheetName); err != nil {
				return nil, err
			}
			if sw, err = f.NewStreamWriter(sheetName); err != nil {
				return nil, err
			}
			swMap[sheetName] = sw
		}
		lineNumMap[sheetName]++
		if err := appendRow(sw, sheetModel, line, options); err != nil {
			return nil, err
		}
	}
	err := setNoDataSheetHeaders(f, options)
	if err != nil {
		return nil, err
	}
	for _, sw := range swMap {
		if err = sw.Flush(); err != nil {
			return nil, err
		}
	}

	// Delete the default sheet.
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
		err := f.DeleteSheet("Sheet1")
		if err != nil {
			return nil, err
		}
	}
	cleanup = false
	return f, nil
}

func setNoDataSheetHeaders(f *excelize.File, options *options) error {
	models := options.sheetHeaders
	if len(models) == 0 {
		return nil
	}
	for _, model := range models {
		modelValue, err := modelStructValue(model)
		if err != nil {
			return err
		}
		sheetName := model.SheetName()
		idx, err := f.GetSheetIndex(sheetName)
		if err != nil {
			return err
		}
		if idx != -1 {
			// sheet exists, continue
			continue
		}
		if _, err = f.NewSheet(sheetName); err != nil {
			return err
		}

		modelType := modelValue.Type()
		for i := 0; i < modelType.NumField(); i++ {
			field := modelType.Field(i)
			header := field.Tag.Get("excel_header")
			if header == "" { // If no excel_header tag is set, use the field name.
				header = field.Name
			} else if header == "-" {
				continue // Skip this field when header is "-".
			}
			cellName, err := coordinatesToCellName(i+1, 1)
			if err != nil {
				return err
			}
			if err = f.SetCellValue(sheetName, cellName, header); err != nil { // Set header.
				return err
			}
		}
	}
	return nil
}

// WriteExcelAsBytesBuffer creates an Excel file and writes it to a bytes.Buffer.
// Usage is the same as WriteExcelSaveAs.
func WriteExcelAsBytesBuffer(sheetModels []SheetModel, opts ...Option) (*bytes.Buffer, error) {
	buffer := new(bytes.Buffer)
	f, err := write(sheetModels, opts...)
	if err != nil {
		return nil, err
	}
	writeErr := f.Write(buffer)
	closeErr := f.Close()
	if err = errors.Join(writeErr, closeErr); err != nil {
		return nil, err
	}
	return buffer, nil
}

type SheetModel interface {
	SheetName() string
}

type options struct {
	timeFormatLayout string       // Formatting layout for time.Time and *time.Time.
	floatPrecision   int          // Number of decimal places to keep.
	floatFmt         byte         // Float format, defaults to 'f'. See strconv.FormatFloat for details.
	ifNullValue      string       // Default display value for nil pointers.
	sheetHeaders     []SheetModel // Default headers to show when there is no data.
	trueValue        *string      // Display value for bool true.
	falseValue       *string      // Display value for bool false.
	integerAsString  bool         // Display integer fields as strings (avoid Excel scientific notation).
	headless         bool         // Whether to hide the header row.
}

// WithTimeFormatLayout sets the formatting layout used for time values.
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

// WithIfNullValue sets the display value when data is nil.
func WithIfNullValue(value string) Option {
	return func(options *options) {
		options.ifNullValue = value
	}
}

// WithSheetHeaders sets default headers to display even when there is no data.
func WithSheetHeaders(headers ...SheetModel) Option {
	return func(options *options) {
		options.sheetHeaders = headers
	}
}

// WithBoolValueAs sets display values for boolean true and false fields.
func WithBoolValueAs(trueValue, falseValue string) Option {
	return func(options *options) {
		options.trueValue = &trueValue
		options.falseValue = &falseValue
	}
}

// WithIntegerAsString displays integer fields as strings (to avoid Excel scientific notation).
func WithIntegerAsString() Option {
	return func(options *options) {
		options.integerAsString = true
	}
}

// WithHeadless disables writing the header row.
func WithHeadless() Option {
	return func(options *options) {
		options.headless = true
	}
}

func appendRow(sw *excelize.StreamWriter, sheetModel SheetModel, line int, options *options) error {
	modelValue, err := modelStructValue(sheetModel)
	if err != nil {
		return err
	}

	modelType := modelValue.Type()
	line++                              // Index starts from 0, but Excel starts from 1.
	if line == 1 && !options.headless { // Set header.
		var values []any
		for i := 0; i < modelType.NumField(); i++ {
			field := modelType.Field(i)
			header := field.Tag.Get("excelorm")
			if header == "" { // Deprecated
				header = field.Tag.Get("excel_header")
			}
			if header == "" { // If no excel_header tag is set, use the field name.
				header = field.Name
			}
			values = append(values, header)
		}
		cellName, err := coordinatesToCellName(1, line)
		if err != nil {
			return err
		}
		if err = sw.SetRow(cellName, values); err != nil { // Set header.
			return err
		}
	}

	if !options.headless {
		line++
	}

	var values = make([]any, 0, modelType.NumField())
	for i := 0; i < modelType.NumField(); i++ {
		field := modelType.Field(i)
		fieldValue := modelValue.Field(i) // Get field value.
		fieldKind := field.Type.Kind()    // Get field kind.
	unAddrTo:
		switch fieldKind {
		case reflect.Pointer: // If the field is a pointer, resolve its value.
			canAddr := fieldValue.Elem().CanAddr() // Check whether the pointed value is accessible.
			if !canAddr {
				values = append(values, options.ifNullValue)
			} else {
				fieldValue = reflect.Indirect(fieldValue) // Get the pointed value.
				fieldKind = fieldValue.Kind()             // Get the pointed value kind.
				goto unAddrTo                             // Re-check now that the field is no longer a pointer.
			}
		case reflect.Struct, reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64,
			reflect.String, reflect.Bool, reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64,
			reflect.Float32, reflect.Float64:
			valueInterface := fieldValue.Interface() // Get the field value (type interface{}).
			switch value := valueInterface.(type) {  // Type assertion.
			case int, int8, int16, int32, int64:
				if options.integerAsString {
					values = append(values, strconv.FormatInt(fieldValue.Int(), 10)) // Set integer cell value.
				} else {
					values = append(values, value)
				}
			case uint, uint8, uint16, uint32, uint64:
				if options.integerAsString {
					values = append(values, strconv.FormatUint(fieldValue.Uint(), 10)) // Set unsigned integer cell value.
				} else {
					values = append(values, value)
				}
			case string:
				values = append(values, value) // Set string cell value.
			case bool: // Convert bool using options.
				if options.trueValue != nil && value { // trueValue is set and value is true.
					values = append(values, *options.trueValue)
				} else if options.falseValue != nil && !value { // falseValue is set and value is false.
					values = append(values, *options.falseValue)
				} else { // Use default bool output.
					values = append(values, value)

				}
			case float32: // Format float32 using options.
				values = append(values, strconv.FormatFloat(float64(value), options.floatFmt, options.floatPrecision, 32))
			case float64: // Format float64 using options.
				values = append(values, strconv.FormatFloat(value, options.floatFmt, options.floatPrecision, 32))
			case time.Time: // Format time.Time using options.
				values = append(values, value.Format(options.timeFormatLayout))
			default:
				return fmt.Errorf("unsupported type %T", value)
			}

		case reflect.Map, reflect.Slice, reflect.Array, reflect.Chan, reflect.Func, reflect.Interface,
			reflect.Invalid, reflect.UnsafePointer, reflect.Complex64, reflect.Complex128, reflect.Uintptr:
			return fmt.Errorf("unsupported type %s", fieldKind)
		}
	}
	cellName, err := coordinatesToCellName(1, line)
	if err != nil {
		return err
	}
	if err = sw.SetRow(cellName, values); err != nil {
		return err
	}
	return nil
}

func modelStructValue(model SheetModel) (reflect.Value, error) {
	if model == nil {
		return reflect.Value{}, errors.New("nil reference row append is not allowed")
	}
	v := reflect.ValueOf(model)
	if v.Kind() == reflect.Ptr {
		if v.IsNil() {
			return reflect.Value{}, errors.New("nil reference row append is not allowed")
		}
		v = v.Elem()
	}
	if v.Kind() != reflect.Struct {
		return reflect.Value{}, errors.New("sheetModel must be struct")
	}
	return v, nil
}

// The following code is copied and modified from https://github.com/360EntSecGroup-Skylar/excelize.

// coordinatesToCellName converts [X, Y] coordinates to alpha-numeric cell
// name, or returns an error.
// Example:
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

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
	"strings"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

type ReadOption func(*readOptions)

type readOptions struct {
	timeFormatLayout string
	strict           bool
	ifNullValue      *string
	trueValue        *string
	falseValue       *string
}

var timeType = reflect.TypeOf(time.Time{})

var readModelMetaCache sync.Map // map[reflect.Type]*readModelMeta

type readValueKind int

const (
	readValueUnsupported readValueKind = iota
	readValueString
	readValueBool
	readValueInt
	readValueUint
	readValueFloat
	readValueTime
)

type readFieldMeta struct {
	index   int
	value   readValueKind
	bits    int
	ptr     bool
	goType  reflect.Type
	elemTyp reflect.Type
}

type readModelMeta struct {
	sheetName     string
	headers       []string
	fieldByHeader map[string]readFieldMeta
	headerSet     map[string]struct{}
}

type readColumnBinding struct {
	colIndex int
	header   string
	field    readFieldMeta
}

// WithReadTimeFormatLayout sets the parsing layout used for time values.
func WithReadTimeFormatLayout(layout string) ReadOption {
	return func(options *readOptions) {
		options.timeFormatLayout = layout
	}
}

// WithReadStrictMode enables strict header validation.
func WithReadStrictMode() ReadOption {
	return func(options *readOptions) {
		options.strict = true
	}
}

// WithReadIfNullValue maps a specific cell value to nil pointer fields.
func WithReadIfNullValue(value string) ReadOption {
	return func(options *readOptions) {
		options.ifNullValue = &value
	}
}

// WithReadBoolValueAs configures custom true/false input values.
func WithReadBoolValueAs(trueValue, falseValue string) ReadOption {
	return func(options *readOptions) {
		options.trueValue = &trueValue
		options.falseValue = &falseValue
	}
}

// ReadExcelToModels reads one sheet into a model slice.
// out must be a pointer to a slice whose element type is struct or *struct implementing SheetModel.
func ReadExcelToModels(fileName string, out any, opts ...ReadOption) error {
	if fileName == "" {
		return errors.New("fileName can not be empty")
	}
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return err
	}
	readErr := readModelsFromExcelFile(f, out, opts...)
	closeErr := f.Close()
	return errors.Join(readErr, closeErr)
}

// ReadExcelToModelsFromBytesBuffer reads one sheet into a model slice from a buffer.
func ReadExcelToModelsFromBytesBuffer(buffer *bytes.Buffer, out any, opts ...ReadOption) error {
	if buffer == nil {
		return errors.New("buffer can not be nil")
	}
	f, err := excelize.OpenReader(bytes.NewReader(buffer.Bytes()))
	if err != nil {
		return err
	}
	readErr := readModelsFromExcelFile(f, out, opts...)
	closeErr := f.Close()
	return errors.Join(readErr, closeErr)
}

// ReadExcelToMaps reads one sheet into []map[header]cellValue.
func ReadExcelToMaps(fileName, sheetName string, opts ...ReadOption) ([]map[string]string, error) {
	if fileName == "" {
		return nil, errors.New("fileName can not be empty")
	}
	if sheetName == "" {
		return nil, errors.New("sheetName can not be empty")
	}
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, err
	}
	maps, readErr := readMapsFromExcelFile(f, sheetName, opts...)
	closeErr := f.Close()
	if err = errors.Join(readErr, closeErr); err != nil {
		return nil, err
	}
	return maps, nil
}

// ReadExcelToMapsFromBytesBuffer reads one sheet into []map[header]cellValue from a buffer.
func ReadExcelToMapsFromBytesBuffer(buffer *bytes.Buffer, sheetName string, opts ...ReadOption) ([]map[string]string, error) {
	if buffer == nil {
		return nil, errors.New("buffer can not be nil")
	}
	if sheetName == "" {
		return nil, errors.New("sheetName can not be empty")
	}
	f, err := excelize.OpenReader(bytes.NewReader(buffer.Bytes()))
	if err != nil {
		return nil, err
	}
	maps, readErr := readMapsFromExcelFile(f, sheetName, opts...)
	closeErr := f.Close()
	if err = errors.Join(readErr, closeErr); err != nil {
		return nil, err
	}
	return maps, nil
}

func readModelsFromExcelFile(f *excelize.File, out any, opts ...ReadOption) error {
	options := defaultReadOptions()
	for _, opt := range opts {
		opt(options)
	}

	sliceValue, elemType, isPointerElem, err := targetSliceMeta(out)
	if err != nil {
		return err
	}
	meta, err := getReadModelMeta(elemType)
	if err != nil {
		return err
	}
	rows, err := f.GetRows(meta.sheetName)
	if err != nil {
		return err
	}
	if len(rows) == 0 {
		sliceValue.Set(reflect.MakeSlice(sliceValue.Type(), 0, 0))
		return nil
	}

	headers := rows[0]
	if options.strict {
		if err = validateHeaderRow(headers); err != nil {
			return err
		}
		if err = validateStrictModelHeaders(headers, meta); err != nil {
			return err
		}
	}

	bindings := buildReadColumnBindings(headers, meta)
	if len(bindings) == 0 {
		sliceValue.Set(reflect.MakeSlice(sliceValue.Type(), 0, 0))
		return nil
	}

	result := reflect.MakeSlice(sliceValue.Type(), 0, len(rows)-1)
	for rowIndex := 1; rowIndex < len(rows); rowIndex++ {
		row := rows[rowIndex]
		if isEmptyRow(row) {
			continue
		}

		modelValue := reflect.New(elemType).Elem()
		for _, binding := range bindings {
			raw := safeCellValue(row, binding.colIndex)
			if err = assignStringByMeta(modelValue.Field(binding.field.index), raw, binding.field, options); err != nil {
				return fmt.Errorf("sheet %q row %d col %d (%s): %w", meta.sheetName, rowIndex+1, binding.colIndex+1, binding.header, err)
			}
		}

		if isPointerElem {
			rowValue := reflect.New(elemType)
			rowValue.Elem().Set(modelValue)
			result = reflect.Append(result, rowValue)
		} else {
			result = reflect.Append(result, modelValue)
		}
	}

	sliceValue.Set(result)
	return nil
}

func readMapsFromExcelFile(f *excelize.File, sheetName string, opts ...ReadOption) ([]map[string]string, error) {
	options := defaultReadOptions()
	for _, opt := range opts {
		opt(options)
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, err
	}
	if len(rows) == 0 {
		return []map[string]string{}, nil
	}

	headers := rows[0]
	if options.strict {
		if err = validateHeaderRow(headers); err != nil {
			return nil, err
		}
	}

	maps := make([]map[string]string, 0, len(rows)-1)
	for rowIndex := 1; rowIndex < len(rows); rowIndex++ {
		row := rows[rowIndex]
		if isEmptyRow(row) {
			continue
		}
		item := make(map[string]string, len(headers))
		for colIndex, header := range headers {
			if header == "" {
				continue
			}
			item[header] = safeCellValue(row, colIndex)
		}
		maps = append(maps, item)
	}
	return maps, nil
}

func defaultReadOptions() *readOptions {
	return &readOptions{timeFormatLayout: "2006-01-02 15:04:05"}
}

func targetSliceMeta(out any) (reflect.Value, reflect.Type, bool, error) {
	if out == nil {
		return reflect.Value{}, nil, false, errors.New("out can not be nil")
	}
	outValue := reflect.ValueOf(out)
	if outValue.Kind() != reflect.Ptr || outValue.IsNil() {
		return reflect.Value{}, nil, false, errors.New("out must be a non-nil pointer to slice")
	}
	sliceValue := outValue.Elem()
	if sliceValue.Kind() != reflect.Slice {
		return reflect.Value{}, nil, false, errors.New("out must be a pointer to slice")
	}

	elemType := sliceValue.Type().Elem()
	isPointerElem := false
	if elemType.Kind() == reflect.Ptr {
		isPointerElem = true
		elemType = elemType.Elem()
	}
	if elemType.Kind() != reflect.Struct {
		return reflect.Value{}, nil, false, errors.New("slice element must be struct or *struct")
	}
	return sliceValue, elemType, isPointerElem, nil
}

func getReadModelMeta(modelType reflect.Type) (*readModelMeta, error) {
	if cached, ok := readModelMetaCache.Load(modelType); ok {
		return cached.(*readModelMeta), nil
	}

	sheetName, err := sheetNameFromType(modelType)
	if err != nil {
		return nil, err
	}

	meta := &readModelMeta{
		sheetName:     sheetName,
		headers:       make([]string, 0, modelType.NumField()),
		fieldByHeader: make(map[string]readFieldMeta, modelType.NumField()),
		headerSet:     make(map[string]struct{}, modelType.NumField()),
	}
	for i := 0; i < modelType.NumField(); i++ {
		field := modelType.Field(i)
		if field.PkgPath != "" { // skip unexported fields
			continue
		}
		header := field.Tag.Get("excelorm")
		if header == "" {
			header = field.Tag.Get("excel_header")
		}
		if header == "-" {
			continue
		}
		if header == "" {
			header = field.Name
		}
		if _, exists := meta.fieldByHeader[header]; exists {
			return nil, fmt.Errorf("duplicated model header %q", header)
		}
		fieldMeta := resolveReadFieldMeta(i, field.Type)
		meta.headers = append(meta.headers, header)
		meta.fieldByHeader[header] = fieldMeta
		meta.headerSet[header] = struct{}{}
	}

	stored, _ := readModelMetaCache.LoadOrStore(modelType, meta)
	return stored.(*readModelMeta), nil
}

func sheetNameFromType(modelType reflect.Type) (string, error) {
	if modelType.Kind() != reflect.Struct {
		return "", errors.New("sheetModel must be struct")
	}
	if model, ok := reflect.New(modelType).Interface().(SheetModel); ok {
		sheetName := model.SheetName()
		if sheetName == "" {
			return "", errors.New("sheetModel must have a sheet name")
		}
		return sheetName, nil
	}
	return "", errors.New("slice element must implement SheetModel")
}

func resolveReadFieldMeta(index int, typ reflect.Type) readFieldMeta {
	meta := readFieldMeta{index: index, goType: typ, elemTyp: typ}
	if typ.Kind() == reflect.Pointer {
		meta.ptr = true
		meta.elemTyp = typ.Elem()
	}
	switch meta.elemTyp.Kind() {
	case reflect.String:
		meta.value = readValueString
	case reflect.Bool:
		meta.value = readValueBool
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		meta.value = readValueInt
		meta.bits = meta.elemTyp.Bits()
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		meta.value = readValueUint
		meta.bits = meta.elemTyp.Bits()
	case reflect.Float32, reflect.Float64:
		meta.value = readValueFloat
		meta.bits = meta.elemTyp.Bits()
	case reflect.Struct:
		if meta.elemTyp == timeType {
			meta.value = readValueTime
			break
		}
		meta.value = readValueUnsupported
	default:
		meta.value = readValueUnsupported
	}
	return meta
}

func validateStrictModelHeaders(headers []string, meta *readModelMeta) error {
	headerSet := make(map[string]struct{}, len(headers))
	for _, header := range headers {
		headerSet[header] = struct{}{}
		if _, ok := meta.fieldByHeader[header]; !ok {
			return fmt.Errorf("strict mode: header %q has no matching field", header)
		}
	}
	for _, header := range meta.headers {
		if _, ok := headerSet[header]; !ok {
			return fmt.Errorf("strict mode: model field header %q is missing in sheet", header)
		}
	}
	return nil
}

func buildReadColumnBindings(headers []string, meta *readModelMeta) []readColumnBinding {
	bindings := make([]readColumnBinding, 0, len(headers))
	for colIndex, header := range headers {
		if header == "" {
			continue
		}
		fieldMeta, ok := meta.fieldByHeader[header]
		if !ok {
			continue
		}
		bindings = append(bindings, readColumnBinding{colIndex: colIndex, header: header, field: fieldMeta})
	}
	return bindings
}

func validateHeaderRow(headers []string) error {
	if len(headers) == 0 {
		return errors.New("strict mode: header row is empty")
	}
	seen := make(map[string]int, len(headers))
	for i, header := range headers {
		if header == "" {
			return fmt.Errorf("strict mode: empty header at column %d", i+1)
		}
		if firstIndex, exists := seen[header]; exists {
			return fmt.Errorf("strict mode: duplicated header %q at columns %d and %d", header, firstIndex+1, i+1)
		}
		seen[header] = i
	}
	return nil
}

func assignStringByMeta(fieldValue reflect.Value, raw string, fieldMeta readFieldMeta, options *readOptions) error {
	if fieldMeta.ptr {
		if raw == "" || (options.ifNullValue != nil && raw == *options.ifNullValue) {
			fieldValue.Set(reflect.Zero(fieldMeta.goType))
			return nil
		}
		target := reflect.New(fieldMeta.elemTyp)
		if err := assignPrimitiveByMeta(target.Elem(), raw, fieldMeta, options); err != nil {
			return err
		}
		fieldValue.Set(target)
		return nil
	}
	return assignPrimitiveByMeta(fieldValue, raw, fieldMeta, options)
}

func assignPrimitiveByMeta(fieldValue reflect.Value, raw string, fieldMeta readFieldMeta, options *readOptions) error {
	trimmed := strings.TrimSpace(raw)
	if trimmed == "" {
		fieldValue.Set(reflect.Zero(fieldMeta.elemTyp))
		return nil
	}

	switch fieldMeta.value {
	case readValueString:
		// Preserve original string without trimming, as users may want to keep leading/trailing spaces
		fieldValue.SetString(raw)
		return nil
	case readValueBool:
		value, err := parseBool(trimmed, options)
		if err != nil {
			return err
		}
		fieldValue.SetBool(value)
		return nil
	case readValueInt:
		value, err := strconv.ParseInt(trimmed, 10, fieldMeta.bits)
		if err != nil {
			return err
		}
		fieldValue.SetInt(value)
		return nil
	case readValueUint:
		value, err := strconv.ParseUint(trimmed, 10, fieldMeta.bits)
		if err != nil {
			return err
		}
		fieldValue.SetUint(value)
		return nil
	case readValueFloat:
		value, err := strconv.ParseFloat(trimmed, fieldMeta.bits)
		if err != nil {
			return err
		}
		fieldValue.SetFloat(value)
		return nil
	case readValueTime:
		parsedAt, err := time.Parse(options.timeFormatLayout, trimmed)
		if err != nil {
			return err
		}
		fieldValue.Set(reflect.ValueOf(parsedAt))
		return nil
	default:
		return fmt.Errorf("unsupported type %s", fieldMeta.elemTyp)
	}
}

func parseBool(value string, options *readOptions) (bool, error) {
	if options.trueValue != nil && value == *options.trueValue {
		return true, nil
	}
	if options.falseValue != nil && value == *options.falseValue {
		return false, nil
	}
	return strconv.ParseBool(strings.ToLower(value))
}

func safeCellValue(row []string, index int) string {
	if index >= len(row) {
		return ""
	}
	return row[index]
}

func isEmptyRow(row []string) bool {
	for _, cell := range row {
		if strings.TrimSpace(cell) != "" {
			return false
		}
	}
	return true
}

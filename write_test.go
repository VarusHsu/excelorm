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
	"path/filepath"
	"testing"
	"time"

	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

type Sheet1 struct {
	Col1  string     `excel_header:"string"`
	Col2  int        `excel_header:"int"`
	Col3  float64    `excel_header:"float"`
	Col4  bool       `excel_header:"bool"`
	Col5  time.Time  `excel_header:"time"`
	Col6  *string    `excel_header:"string pointer"`
	Col7  *int       `excel_header:"int pointer"`
	Col8  *float64   `excel_header:"float pointer"`
	Col9  *bool      `excel_header:"bool pointer"`
	Col10 *time.Time `excel_header:"time pointer"`
}

func (Sheet1) SheetName() string {
	return "sheet_one"
}

type Sheet2 struct {
	Col1  string     `excel_header:"string"`
	Col2  int        `excel_header:"int"`
	Col3  float64    `excel_header:"float"`
	Col4  bool       `excel_header:"bool"`
	Col5  time.Time  `excel_header:"time"`
	Col6  *string    `excel_header:"string pointer"`
	Col7  *int       `excel_header:"int pointer"`
	Col8  *float64   `excel_header:"float pointer"`
	Col9  *bool      `excel_header:"bool pointer"`
	Col10 *time.Time `excel_header:"time pointer"`
	Col11 float32    `excel_header:"float32"`
}

func (Sheet2) SheetName() string {
	return "sheet_two"
}

type Sheet3 struct {
	Col1 string `excel_header:"string"`
}

func (Sheet3) SheetName() string {
	return ""
}

type Sheet4 int

func (Sheet4) SheetName() string {
	return "sheet4"
}

type Sheet5 struct {
	Col1 string
}

func (Sheet5) SheetName() string {
	return "sheet5"
}

type Sheet6 struct {
	Col1 map[string]string `excel_header:"map"`
}

func (Sheet6) SheetName() string {
	return "sheet6"
}

type subStruct struct {
	Field string `excel_header:"field"`
}
type Sheet7 struct {
	SubStruct subStruct `excel_header:"subStruct"`
}

func (Sheet7) SheetName() string {
	return "sheet7"
}

type SheetHeaderSkip struct {
	Keep string `excel_header:"keep"`
	Drop string `excel_header:"-"`
}

func (SheetHeaderSkip) SheetName() string {
	return "header_skip"
}

type SheetDefault struct {
	Col1 string `excel_header:"col1"`
}

func (SheetDefault) SheetName() string {
	return "Sheet1"
}

type SheetUint struct {
	U uint `excel_header:"u"`
}

func (SheetUint) SheetName() string {
	return "sheet_uint"
}

type SheetLegacyTag struct {
	Legacy string `excelorm:"legacy_header"`
}

func (SheetLegacyTag) SheetName() string {
	return "sheet_legacy"
}

type NilHeader struct{}

func (*NilHeader) SheetName() string {
	return "nil_header"
}

type PointerReceiverSheet struct {
	Col1 string `excel_header:"col1"`
}

func (*PointerReceiverSheet) SheetName() string {
	return "pointer_receiver_sheet"
}

func baseModels(boolValue bool) []SheetModel {
	now := time.Date(2026, 1, 2, 3, 4, 5, 0, time.UTC)
	str := "string_value"
	return []SheetModel{
		Sheet1{
			Col1:  "string",
			Col2:  1,
			Col3:  1.1,
			Col4:  true,
			Col5:  now,
			Col6:  nil,
			Col7:  nil,
			Col8:  nil,
			Col9:  nil,
			Col10: nil,
		},
		Sheet2{
			Col1:  "string",
			Col2:  1,
			Col3:  1.1,
			Col4:  boolValue,
			Col5:  now,
			Col6:  &str,
			Col7:  nil,
			Col8:  nil,
			Col9:  nil,
			Col10: nil,
		},
	}
}

func TestWriteExcelSaveAs_Success(t *testing.T) {
	output := filepath.Join(t.TempDir(), "success.xlsx")
	require.NoError(t, WriteExcelSaveAs(output, baseModels(false)))

	t.Run("pointer struct", func(t *testing.T) {
		output := filepath.Join(t.TempDir(), "pointer_success.xlsx")
		now := time.Date(2026, 1, 2, 3, 4, 5, 0, time.UTC)
		models := []SheetModel{
			&Sheet1{Col1: "string", Col2: 1, Col3: 1.1, Col4: true, Col5: now},
		}
		require.NoError(t, WriteExcelSaveAs(output, models))

		file, err := excelize.OpenFile(output)
		require.NoError(t, err)
		t.Cleanup(func() {
			require.NoError(t, file.Close())
		})

		header, err := file.GetCellValue("sheet_one", "A1")
		require.NoError(t, err)
		require.Equal(t, "string", header)

		value, err := file.GetCellValue("sheet_one", "A2")
		require.NoError(t, err)
		require.Equal(t, "string", value)
	})
}

func TestWriteExcelSaveAs_Errors(t *testing.T) {
	testCases := []struct {
		name    string
		file    string
		models  []SheetModel
		errText string
	}{
		{
			name:    "empty file name",
			file:    "",
			models:  baseModels(false),
			errText: "fileName can not be empty",
		},
		{
			name:    "empty sheet name",
			file:    "invalid.xlsx",
			models:  []SheetModel{Sheet3{Col1: "string"}},
			errText: "sheetModel must have a sheet name",
		},
		{
			name:    "non struct sheet model",
			file:    "invalid.xlsx",
			models:  []SheetModel{Sheet4(1)},
			errText: "sheetModel must be struct",
		},
		{
			name: "unsupported map type",
			file: "invalid.xlsx",
			models: []SheetModel{
				Sheet6{Col1: map[string]string{"key": "value"}},
			},
			errText: "unsupported type map",
		},
		{
			name: "unsupported nested struct",
			file: "invalid.xlsx",
			models: []SheetModel{
				Sheet7{SubStruct: subStruct{Field: "field"}},
			},
			errText: "unsupported type excelorm.subStruct",
		},
		{
			name:    "nil row",
			file:    "invalid.xlsx",
			models:  []SheetModel{nil},
			errText: "nil reference row append is not allowed",
		},
		{
			name: "nil typed pointer row",
			file: "invalid.xlsx",
			models: func() []SheetModel {
				var model *Sheet1
				return []SheetModel{model}
			}(),
			errText: "nil reference row append is not allowed",
		},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			fileName := tc.file
			if fileName != "" {
				fileName = filepath.Join(t.TempDir(), fileName)
			}
			err := WriteExcelSaveAs(fileName, tc.models)
			require.EqualError(t, err, tc.errText)
		})
	}
}

func TestWriteExcelSaveAs_Options(t *testing.T) {
	testCases := []struct {
		name string
		opts []Option
	}{
		{name: "time format", opts: []Option{WithTimeFormatLayout("2006/01/02 15:04:05")}},
		{name: "null value", opts: []Option{WithIfNullValue("-")}},
		{name: "float precision", opts: []Option{WithFloatPrecision(10)}},
		{name: "float format", opts: []Option{WithFloatFmt('e')}},
		{name: "bool as words", opts: []Option{WithBoolValueAs("yes", "no")}},
		{name: "bool as digits", opts: []Option{WithBoolValueAs("1", "0")}},
		{name: "bool as literals", opts: []Option{WithBoolValueAs("true", "false")}},
		{name: "headless", opts: []Option{WithHeadless()}},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			output := filepath.Join(t.TempDir(), tc.name+".xlsx")
			require.NoError(t, WriteExcelSaveAs(output, baseModels(false), tc.opts...))
		})
	}
}

func TestWriteExcelAsBytesBuffer(t *testing.T) {
	t.Run("success", func(t *testing.T) {
		buffer, err := WriteExcelAsBytesBuffer(baseModels(false), WithIfNullValue("-"))
		require.NoError(t, err)
		require.NotNil(t, buffer)
		require.NotZero(t, buffer.Len())
	})

	t.Run("pointer struct", func(t *testing.T) {
		now := time.Date(2026, 1, 2, 3, 4, 5, 0, time.UTC)
		buffer, err := WriteExcelAsBytesBuffer([]SheetModel{
			&Sheet1{Col1: "string", Col2: 1, Col3: 1.1, Col4: true, Col5: now},
		})
		require.NoError(t, err)
		require.NotNil(t, buffer)
		require.NotZero(t, buffer.Len())

		file, err := excelize.OpenReader(bytes.NewReader(buffer.Bytes()))
		require.NoError(t, err)
		t.Cleanup(func() {
			require.NoError(t, file.Close())
		})

		header, err := file.GetCellValue("sheet_one", "A1")
		require.NoError(t, err)
		require.Equal(t, "string", header)

		value, err := file.GetCellValue("sheet_one", "A2")
		require.NoError(t, err)
		require.Equal(t, "string", value)
	})

	t.Run("pointer receiver sheet", func(t *testing.T) {
		buffer, err := WriteExcelAsBytesBuffer([]SheetModel{&PointerReceiverSheet{Col1: "v"}})
		require.NoError(t, err)
		require.NotNil(t, buffer)
		require.NotZero(t, buffer.Len())

		file, err := excelize.OpenReader(bytes.NewReader(buffer.Bytes()))
		require.NoError(t, err)
		t.Cleanup(func() {
			require.NoError(t, file.Close())
		})

		header, err := file.GetCellValue("pointer_receiver_sheet", "A1")
		require.NoError(t, err)
		require.Equal(t, "col1", header)

		value, err := file.GetCellValue("pointer_receiver_sheet", "A2")
		require.NoError(t, err)
		require.Equal(t, "v", value)
	})

	t.Run("nil row", func(t *testing.T) {
		_, err := WriteExcelAsBytesBuffer([]SheetModel{nil})
		require.EqualError(t, err, "nil reference row append is not allowed")
	})

	t.Run("nil typed pointer row", func(t *testing.T) {
		var model *Sheet1
		_, err := WriteExcelAsBytesBuffer([]SheetModel{model})
		require.EqualError(t, err, "nil reference row append is not allowed")
	})
}

func TestWithSheetHeaders_NoData(t *testing.T) {
	buffer, err := WriteExcelAsBytesBuffer(nil, WithSheetHeaders(Sheet1{}))
	require.NoError(t, err)

	file, err := excelize.OpenReader(bytes.NewReader(buffer.Bytes()))
	require.NoError(t, err)
	t.Cleanup(func() {
		require.NoError(t, file.Close())
	})

	value, err := file.GetCellValue("sheet_one", "A1")
	require.NoError(t, err)
	require.Equal(t, "string", value)
}

func TestWithHeadless_DoesNotWriteHeader(t *testing.T) {
	buffer, err := WriteExcelAsBytesBuffer(baseModels(true), WithHeadless())
	require.NoError(t, err)

	file, err := excelize.OpenReader(bytes.NewReader(buffer.Bytes()))
	require.NoError(t, err)
	t.Cleanup(func() {
		require.NoError(t, file.Close())
	})

	firstCell, err := file.GetCellValue("sheet_one", "A1")
	require.NoError(t, err)
	require.Equal(t, "string", firstCell)
}

func TestWithIntegerAsString_WritesStringCell(t *testing.T) {
	defaultBuffer, err := WriteExcelAsBytesBuffer(baseModels(false))
	require.NoError(t, err)

	defaultFile, err := excelize.OpenReader(bytes.NewReader(defaultBuffer.Bytes()))
	require.NoError(t, err)
	t.Cleanup(func() {
		require.NoError(t, defaultFile.Close())
	})

	defaultType, err := defaultFile.GetCellType("sheet_one", "B2")
	require.NoError(t, err)
	require.Equal(t, excelize.CellTypeUnset, defaultType)

	stringBuffer, err := WriteExcelAsBytesBuffer(baseModels(false), WithIntegerAsString())
	require.NoError(t, err)

	stringFile, err := excelize.OpenReader(bytes.NewReader(stringBuffer.Bytes()))
	require.NoError(t, err)
	t.Cleanup(func() {
		require.NoError(t, stringFile.Close())
	})

	stringType, err := stringFile.GetCellType("sheet_one", "B2")
	require.NoError(t, err)
	require.Contains(t, []excelize.CellType{excelize.CellTypeInlineString, excelize.CellTypeSharedString}, stringType)
}

func TestWithSheetHeaders_PointerAndSkipTags(t *testing.T) {
	buffer, err := WriteExcelAsBytesBuffer(nil, WithSheetHeaders(&SheetHeaderSkip{}))
	require.NoError(t, err)

	file, err := excelize.OpenReader(bytes.NewReader(buffer.Bytes()))
	require.NoError(t, err)
	t.Cleanup(func() {
		require.NoError(t, file.Close())
	})

	keep, err := file.GetCellValue("header_skip", "A1")
	require.NoError(t, err)
	require.Equal(t, "keep", keep)

	drop, err := file.GetCellValue("header_skip", "B1")
	require.NoError(t, err)
	require.Empty(t, drop)
}

func TestWithSheetHeaders_ExistingSheetKeepsHeader(t *testing.T) {
	buffer, err := WriteExcelAsBytesBuffer(baseModels(false), WithSheetHeaders(Sheet1{}))
	require.NoError(t, err)

	file, err := excelize.OpenReader(bytes.NewReader(buffer.Bytes()))
	require.NoError(t, err)
	t.Cleanup(func() {
		require.NoError(t, file.Close())
	})

	header, err := file.GetCellValue("sheet_one", "A1")
	require.NoError(t, err)
	require.Equal(t, "string", header)
}

func TestWriteExcel_KeepSheet1WhenRequested(t *testing.T) {
	buffer, err := WriteExcelAsBytesBuffer([]SheetModel{SheetDefault{Col1: "v"}})
	require.NoError(t, err)

	file, err := excelize.OpenReader(bytes.NewReader(buffer.Bytes()))
	require.NoError(t, err)
	t.Cleanup(func() {
		require.NoError(t, file.Close())
	})

	value, err := file.GetCellValue("Sheet1", "A1")
	require.NoError(t, err)
	require.Equal(t, "col1", value)
}

func TestCoordinatesToCellName(t *testing.T) {
	testCases := []struct {
		name    string
		col     int
		row     int
		cell    string
		errText string
	}{
		{name: "valid", col: 1, row: 1, cell: "A1"},
		{name: "invalid col", col: 0, row: 1, errText: "invalid cell reference [0, 1]"},
		{name: "invalid row", col: 1, row: 0, errText: "invalid cell reference [1, 0]"},
		{name: "row overflow", col: 1, row: 1048577, errText: "row number exceeds maximum limit"},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			cell, err := coordinatesToCellName(tc.col, tc.row)
			if tc.errText != "" {
				require.EqualError(t, err, tc.errText)
				return
			}
			require.NoError(t, err)
			require.Equal(t, tc.cell, cell)
		})
	}
}

func TestColumnNumberToName(t *testing.T) {
	testCases := []struct {
		name    string
		num     int
		nameOut string
		errText string
	}{
		{name: "valid A", num: 1, nameOut: "A"},
		{name: "valid AA", num: 27, nameOut: "AA"},
		{name: "too small", num: 0, errText: "the column number must be greater than or equal to 1 and less than or equal to 16384"},
		{name: "too large", num: 16385, errText: "the column number must be greater than or equal to 1 and less than or equal to 16384"},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			nameOut, err := columnNumberToName(tc.num)
			if tc.errText != "" {
				require.EqualError(t, err, tc.errText)
				return
			}
			require.NoError(t, err)
			require.Equal(t, tc.nameOut, nameOut)
		})
	}
}

func TestWithIntegerAsString_Uint(t *testing.T) {
	buffer, err := WriteExcelAsBytesBuffer([]SheetModel{SheetUint{U: 9}}, WithIntegerAsString())
	require.NoError(t, err)

	file, err := excelize.OpenReader(bytes.NewReader(buffer.Bytes()))
	require.NoError(t, err)
	t.Cleanup(func() {
		require.NoError(t, file.Close())
	})

	cellType, err := file.GetCellType("sheet_uint", "A2")
	require.NoError(t, err)
	require.Contains(t, []excelize.CellType{excelize.CellTypeInlineString, excelize.CellTypeSharedString}, cellType)
}

func TestDeprecatedExcelORMTagHeader(t *testing.T) {
	buffer, err := WriteExcelAsBytesBuffer([]SheetModel{SheetLegacyTag{Legacy: "v"}})
	require.NoError(t, err)

	file, err := excelize.OpenReader(bytes.NewReader(buffer.Bytes()))
	require.NoError(t, err)
	t.Cleanup(func() {
		require.NoError(t, file.Close())
	})

	header, err := file.GetCellValue("sheet_legacy", "A1")
	require.NoError(t, err)
	require.Equal(t, "legacy_header", header)
}

func TestAppendRow_PointerModel(t *testing.T) {
	f := excelize.NewFile()
	t.Cleanup(func() {
		require.NoError(t, f.Close())
	})
	_, err := f.NewSheet("pointer_sheet")
	require.NoError(t, err)

	sw, err := f.NewStreamWriter("pointer_sheet")
	require.NoError(t, err)

	now := time.Date(2026, 1, 2, 3, 4, 5, 0, time.UTC)
	model := &Sheet1{Col1: "string", Col2: 1, Col3: 1.1, Col4: true, Col5: now}

	err = appendRow(sw, model, 0, &options{timeFormatLayout: "2006-01-02 15:04:05", floatPrecision: 2, floatFmt: 'f'})
	require.NoError(t, err)
	require.NoError(t, sw.Flush())

	header, err := f.GetCellValue("pointer_sheet", "A1")
	require.NoError(t, err)
	require.Equal(t, "string", header)

	value, err := f.GetCellValue("pointer_sheet", "A2")
	require.NoError(t, err)
	require.Equal(t, "string", value)

	nullValue, err := f.GetCellValue("pointer_sheet", "F2")
	require.NoError(t, err)
	require.Empty(t, nullValue)
}

func TestAppendRow_NilPointerModel(t *testing.T) {
	f := excelize.NewFile()
	t.Cleanup(func() {
		require.NoError(t, f.Close())
	})
	_, err := f.NewSheet("pointer_sheet")
	require.NoError(t, err)

	sw, err := f.NewStreamWriter("pointer_sheet")
	require.NoError(t, err)

	var model *Sheet1
	err = appendRow(sw, model, 0, &options{timeFormatLayout: "2006-01-02 15:04:05", floatPrecision: 2, floatFmt: 'f'})
	require.EqualError(t, err, "nil reference row append is not allowed")
}

func TestWithSheetHeaders_NilPointer(t *testing.T) {
	var header *NilHeader
	_, err := WriteExcelAsBytesBuffer(nil, WithSheetHeaders(header))
	require.EqualError(t, err, "nil reference row append is not allowed")
}

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
	"path/filepath"
	"testing"
	"time"

	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

type ReadCustomSheet struct {
	Col1  string     `excel_header:"string"`
	Col4  bool       `excel_header:"bool"`
	Col5  time.Time  `excel_header:"time"`
	Col6  *string    `excel_header:"string pointer"`
	Col7  *int       `excel_header:"int pointer"`
	Col8  *float64   `excel_header:"float pointer"`
	Col9  *bool      `excel_header:"bool pointer"`
	Col10 *time.Time `excel_header:"time pointer"`
}

func (ReadCustomSheet) SheetName() string {
	return "sheet_one"
}

func TestReadExcelToModels(t *testing.T) {
	path := writeReadFixture(t)

	var rows []Sheet1
	err := ReadExcelToModels(path, &rows)
	require.NoError(t, err)
	require.Len(t, rows, 1)
	require.Equal(t, "string", rows[0].Col1)
	require.Equal(t, 1, rows[0].Col2)
	require.Equal(t, 1.1, rows[0].Col3)
	require.True(t, rows[0].Col4)
	require.Equal(t, time.Date(2026, 1, 2, 3, 4, 5, 0, time.UTC), rows[0].Col5)
	require.Nil(t, rows[0].Col6)
}

func TestReadExcelToModels_PointerSlice(t *testing.T) {
	path := writeReadFixture(t)

	var rows []*Sheet1
	err := ReadExcelToModels(path, &rows)
	require.NoError(t, err)
	require.Len(t, rows, 1)
	require.NotNil(t, rows[0])
	require.Equal(t, "string", rows[0].Col1)
}

func TestReadExcelToModelsFromBytesBuffer(t *testing.T) {
	buffer, err := WriteExcelAsBytesBuffer(baseModels(false))
	require.NoError(t, err)

	var rows []Sheet2
	err = ReadExcelToModelsFromBytesBuffer(buffer, &rows)
	require.NoError(t, err)
	require.Len(t, rows, 1)
	require.Equal(t, "string", rows[0].Col1)
	require.Equal(t, "string_value", *rows[0].Col6)
}

func TestReadExcelToMaps(t *testing.T) {
	path := writeReadFixture(t)
	rows, err := ReadExcelToMaps(path, "sheet_one")
	require.NoError(t, err)
	require.Len(t, rows, 1)
	require.Equal(t, "string", rows[0]["string"])
	require.Equal(t, "1", rows[0]["int"])
}

func TestReadExcelToMapsFromBytesBuffer(t *testing.T) {
	buffer, err := WriteExcelAsBytesBuffer(baseModels(false))
	require.NoError(t, err)

	rows, err := ReadExcelToMapsFromBytesBuffer(buffer, "sheet_one")
	require.NoError(t, err)
	require.Len(t, rows, 1)
	require.Equal(t, "string", rows[0]["string"])
}

func TestReadExcelToModels_HeaderOnlySheet(t *testing.T) {
	path := writeCustomSheet(t, "sheet_one", [][]string{
		{"string", "int", "float", "bool", "time", "string pointer", "int pointer", "float pointer", "bool pointer", "time pointer"},
	})

	var rows []Sheet1
	err := ReadExcelToModels(path, &rows)
	require.NoError(t, err)
	require.Len(t, rows, 0)
}

func TestReadExcelToMaps_HeaderOnlySheet(t *testing.T) {
	path := writeCustomSheet(t, "sheet_one", [][]string{
		{"string", "int", "float", "bool", "time", "string pointer", "int pointer", "float pointer", "bool pointer", "time pointer"},
	})

	rows, err := ReadExcelToMaps(path, "sheet_one")
	require.NoError(t, err)
	require.Len(t, rows, 0)
}

func TestReadExcelToModels_StrictMode(t *testing.T) {
	t.Run("rejects unknown header", func(t *testing.T) {
		path := writeCustomSheet(t, "sheet_one", [][]string{
			{"string", "unknown"},
			{"v", "x"},
		})

		var rows []Sheet1
		err := ReadExcelToModels(path, &rows, WithReadStrictMode())
		require.EqualError(t, err, "strict mode: header \"unknown\" has no matching field")
	})

	t.Run("rejects missing model header", func(t *testing.T) {
		path := writeCustomSheet(t, "sheet_one", [][]string{
			{"string"},
			{"v"},
		})

		var rows []Sheet1
		err := ReadExcelToModels(path, &rows, WithReadStrictMode())
		require.Error(t, err)
		require.ErrorContains(t, err, "strict mode: model field header")
		require.ErrorContains(t, err, "is missing in sheet")
	})

	t.Run("rejects duplicated headers", func(t *testing.T) {
		path := writeCustomSheet(t, "sheet_one", [][]string{
			{"string", "string"},
			{"v1", "v2"},
		})

		var rows []Sheet1
		err := ReadExcelToModels(path, &rows, WithReadStrictMode())
		require.EqualError(t, err, "strict mode: duplicated header \"string\" at columns 1 and 2")
	})
}

func TestReadExcelToModels_CustomOptions(t *testing.T) {
	path := writeCustomSheet(t, "sheet_one", [][]string{
		{"string", "bool", "time", "string pointer", "int pointer", "float pointer", "bool pointer", "time pointer"},
		{"a", "yes", "2026/01/02 03:04:05", "-", "-", "-", "-", "-"},
	})

	var rows []ReadCustomSheet
	err := ReadExcelToModels(path, &rows,
		WithReadStrictMode(),
		WithReadTimeFormatLayout("2006/01/02 15:04:05"),
		WithReadBoolValueAs("yes", "no"),
		WithReadIfNullValue("-"),
	)
	require.NoError(t, err)
	require.Len(t, rows, 1)
	require.Equal(t, "a", rows[0].Col1)
	require.True(t, rows[0].Col4)
	require.Equal(t, time.Date(2026, 1, 2, 3, 4, 5, 0, time.UTC), rows[0].Col5)
	require.Nil(t, rows[0].Col6)
	require.Nil(t, rows[0].Col7)
	require.Nil(t, rows[0].Col8)
	require.Nil(t, rows[0].Col9)
	require.Nil(t, rows[0].Col10)
}

func TestReadExcelToModels_Errors(t *testing.T) {
	t.Run("invalid out", func(t *testing.T) {
		path := writeReadFixture(t)
		var rows []Sheet1
		err := ReadExcelToModels(path, rows)
		require.EqualError(t, err, "out must be a non-nil pointer to slice")
	})

	t.Run("invalid element", func(t *testing.T) {
		path := writeReadFixture(t)
		var rows []int
		err := ReadExcelToModels(path, &rows)
		require.EqualError(t, err, "slice element must be struct or *struct")
	})

	t.Run("sheet model missing method", func(t *testing.T) {
		path := writeReadFixture(t)
		var rows []struct{ A string }
		err := ReadExcelToModels(path, &rows)
		require.EqualError(t, err, "slice element must implement SheetModel")
	})

	t.Run("map api requires sheet name", func(t *testing.T) {
		path := writeReadFixture(t)
		_, err := ReadExcelToMaps(path, "")
		require.EqualError(t, err, "sheetName can not be empty")
	})
}

func writeReadFixture(t *testing.T) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "read_fixture.xlsx")
	err := WriteExcelSaveAs(path, baseModels(false))
	require.NoError(t, err)
	return path
}

func writeCustomSheet(t *testing.T, sheetName string, rows [][]string) string {
	t.Helper()
	f := excelize.NewFile()
	t.Cleanup(func() {
		require.NoError(t, f.Close())
	})

	_, err := f.NewSheet(sheetName)
	require.NoError(t, err)
	for rowIndex, row := range rows {
		values := make([]any, len(row))
		for i, value := range row {
			values[i] = value
		}
		cell, err := coordinatesToCellName(1, rowIndex+1)
		require.NoError(t, err)
		require.NoError(t, f.SetSheetRow(sheetName, cell, &values))
	}
	require.NoError(t, f.DeleteSheet("Sheet1"))

	path := filepath.Join(t.TempDir(), "custom.xlsx")
	require.NoError(t, f.SaveAs(path))
	return path
}

package excelorm

import (
	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"testing"
	"time"
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
	return "sheet1"
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
	return "sheet2"
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

func TestWriteExcel(t *testing.T) {
	sheet1 := Sheet1{
		Col1:  "string",
		Col2:  1,
		Col3:  1.1,
		Col4:  true,
		Col5:  time.Now(),
		Col6:  nil,
		Col7:  nil,
		Col8:  nil,
		Col9:  nil,
		Col10: nil,
	}
	var a = "string_value"
	sheet2 := Sheet2{
		Col1:  "string",
		Col2:  1,
		Col3:  1.1,
		Col4:  true,
		Col5:  time.Now(),
		Col6:  &a,
		Col7:  nil,
		Col8:  nil,
		Col9:  nil,
		Col10: nil,
	}
	var models []SheetModel
	models = append(models, sheet1, sheet1, sheet1, sheet1, sheet1, sheet2, sheet2, sheet2, sheet2, sheet2)

	err := WriteExcelSaveAs("test1.xlsx", models)
	if err != nil {
		t.Error(err)
	}

	sheet3 := Sheet3{
		Col1: "string",
	}

	models = append(models, sheet3)
	err = WriteExcelSaveAs("test2.xlsx", models)
	require.EqualError(t, err, "sheetModel must have a sheet name")

	sheet4 := Sheet4(1)
	models = make([]SheetModel, 0)
	models = append(models, sheet4)
	err = WriteExcelSaveAs("test3.xlsx", models)
	require.EqualError(t, err, "sheetModel must be struct")

	sheet5 := Sheet5{
		Col1: "string",
	}
	models = make([]SheetModel, 0)
	models = append(models, sheet5)
	err = WriteExcelSaveAs("test4.xlsx", models)

	sheet6 := Sheet6{
		Col1: map[string]string{
			"key": "value",
		},
	}
	models = make([]SheetModel, 0)
	models = append(models, sheet6)
	assert.NoErrorf(t, err, "")

	err = WriteExcelSaveAs("test5.xlsx", models)
	require.EqualError(t, err, "unsupported type map")

	sheet7 := Sheet7{
		SubStruct: subStruct{
			Field: "field",
		},
	}
	models = make([]SheetModel, 0)
	models = append(models, sheet7)
	err = WriteExcelSaveAs("test6.xlsx", models)
	assert.EqualError(t, err, "unsupported type excelorm.subStruct")
}

func TestWithTimeFormatLayout(t *testing.T) {
	sheet1 := Sheet1{
		Col1:  "string",
		Col2:  1,
		Col3:  1.1,
		Col4:  true,
		Col5:  time.Now(),
		Col6:  nil,
		Col7:  nil,
		Col8:  nil,
		Col9:  nil,
		Col10: nil,
	}
	var a = "string_value"
	sheet2 := Sheet2{
		Col1:  "string",
		Col2:  1,
		Col3:  1.1,
		Col4:  true,
		Col5:  time.Now(),
		Col6:  &a,
		Col7:  nil,
		Col8:  nil,
		Col9:  nil,
		Col10: nil,
	}
	var models []SheetModel
	models = append(models, sheet1, sheet1, sheet1, sheet1, sheet1, sheet2, sheet2, sheet2, sheet2, sheet2)

	err := WriteExcelSaveAs("test7.xlsx", models, WithTimeFormatLayout("2006/01/02 15:04:05"))
	if err != nil {
		t.Error(err)
	}
}

func TestWithIfNullValue(t *testing.T) {
	sheet1 := Sheet1{
		Col1:  "string",
		Col2:  1,
		Col3:  1.1,
		Col4:  true,
		Col5:  time.Now(),
		Col6:  nil,
		Col7:  nil,
		Col8:  nil,
		Col9:  nil,
		Col10: nil,
	}
	var a = "string_value"
	sheet2 := Sheet2{
		Col1:  "string",
		Col2:  1,
		Col3:  1.1,
		Col4:  true,
		Col5:  time.Now(),
		Col6:  &a,
		Col7:  nil,
		Col8:  nil,
		Col9:  nil,
		Col10: nil,
	}
	var models []SheetModel
	models = append(models, sheet1, sheet1, sheet1, sheet1, sheet1, sheet2, sheet2, sheet2, sheet2, sheet2)

	err := WriteExcelSaveAs("test8.xlsx", models, WithIfNullValue("-"))
	if err != nil {
		t.Error(err)
	}
}

func TestWithFloatPrecision(t *testing.T) {
	sheet1 := Sheet1{
		Col1:  "string",
		Col2:  1,
		Col3:  1.1,
		Col4:  true,
		Col5:  time.Now(),
		Col6:  nil,
		Col7:  nil,
		Col8:  nil,
		Col9:  nil,
		Col10: nil,
	}
	var a = "string_value"
	sheet2 := Sheet2{
		Col1:  "string",
		Col2:  1,
		Col3:  1.1,
		Col4:  true,
		Col5:  time.Now(),
		Col6:  &a,
		Col7:  nil,
		Col8:  nil,
		Col9:  nil,
		Col10: nil,
	}
	var models []SheetModel
	models = append(models, sheet1, sheet1, sheet1, sheet1, sheet1, sheet2, sheet2, sheet2, sheet2, sheet2)

	err := WriteExcelSaveAs("test9.xlsx", models, WithFloatPrecision(10))
	if err != nil {
		t.Error(err)
	}
}

func TestWithFloatFmt(t *testing.T) {
	sheet1 := Sheet1{
		Col1:  "string",
		Col2:  1,
		Col3:  1.1,
		Col4:  true,
		Col5:  time.Now(),
		Col6:  nil,
		Col7:  nil,
		Col8:  nil,
		Col9:  nil,
		Col10: nil,
	}
	var a = "string_value"
	sheet2 := Sheet2{
		Col1:  "string",
		Col2:  1,
		Col3:  1.1,
		Col4:  true,
		Col5:  time.Now(),
		Col6:  &a,
		Col7:  nil,
		Col8:  nil,
		Col9:  nil,
		Col10: nil,
	}
	var models []SheetModel
	models = append(models, sheet1, sheet1, sheet1, sheet1, sheet1, sheet2, sheet2, sheet2, sheet2, sheet2)

	err := WriteExcelSaveAs("test10.xlsx", models, WithFloatFmt('e'))
	if err != nil {
		t.Error(err)
	}
}

func TestWithBoolValueAs(t *testing.T) {
	sheet1 := Sheet1{
		Col1:  "string",
		Col2:  1,
		Col3:  1.1,
		Col4:  true,
		Col5:  time.Now(),
		Col6:  nil,
		Col7:  nil,
		Col8:  nil,
		Col9:  nil,
		Col10: nil,
	}
	var a = "string_value"
	sheet2 := Sheet2{
		Col1:  "string",
		Col2:  1,
		Col3:  1.1,
		Col4:  false,
		Col5:  time.Now(),
		Col6:  &a,
		Col7:  nil,
		Col8:  nil,
		Col9:  nil,
		Col10: nil,
	}
	var models []SheetModel
	models = append(models, sheet1, sheet1, sheet1, sheet1, sheet1, sheet2, sheet2, sheet2, sheet2, sheet2)

	err := WriteExcelSaveAs("test11.xlsx", models, WithBoolValueAs("是", "否"))
	if err != nil {
		t.Error(err)
	}
	err = WriteExcelSaveAs("test12.xlsx", models, WithBoolValueAs("1", "0"))
	if err != nil {
		t.Error(err)
	}
	err = WriteExcelSaveAs("test13.xlsx", models, WithBoolValueAs("true", "false"))
	if err != nil {
		t.Error(err)
	}
	err = WriteExcelSaveAs("test14.xlsx", models) // default
	if err != nil {
		t.Error(err)
	}
}

func TestWithHeadless(t *testing.T) {
	sheet1 := Sheet1{
		Col1:  "string",
		Col2:  1,
		Col3:  1.1,
		Col4:  true,
		Col5:  time.Now(),
		Col6:  nil,
		Col7:  nil,
		Col8:  nil,
		Col9:  nil,
		Col10: nil,
	}
	var a = "string_value"
	sheet2 := Sheet2{
		Col1:  "string",
		Col2:  1,
		Col3:  1.1,
		Col4:  false,
		Col5:  time.Now(),
		Col6:  &a,
		Col7:  nil,
		Col8:  nil,
		Col9:  nil,
		Col10: nil,
	}
	var models []SheetModel
	models = append(models, sheet1, sheet1, sheet1, sheet1, sheet1, sheet2, sheet2, sheet2, sheet2, sheet2)

	err := WriteExcelSaveAs("test15.xlsx", models, WithHeadless())
	if err != nil {
		t.Error(err)
	}
}

func TestAppendNilRow(t *testing.T) {
	var models []SheetModel
	models = append(models, nil)
	err := WriteExcelSaveAs("test16.xlsx", models)
	require.EqualError(t, err, "nil reference row append is not allowed")

}

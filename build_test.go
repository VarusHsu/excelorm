package excelorm

import (
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
}

func (Sheet2) SheetName() string {
	return "sheet2"
}

func TestBuild(t *testing.T) {
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

	err := Build("test.xlsx", models)
	if err != nil {
		t.Error(err)
	}
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

	err := Build("test.xlsx", models, WithTimeFormatLayout("2006/01/02 15:04:05"))
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

	err := Build("test.xlsx", models, WithIfNullValue("-"))
	if err != nil {
		t.Error(err)
	}
}

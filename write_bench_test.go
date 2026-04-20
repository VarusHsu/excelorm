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
	"fmt"
	"testing"
	"time"
)

func benchmarkModels(rows int, asPointer bool, multiSheet bool) []SheetModel {
	models := make([]SheetModel, 0, rows)
	now := time.Date(2026, 1, 2, 3, 4, 5, 0, time.UTC)
	for i := 0; i < rows; i++ {
		name := fmt.Sprintf("name_%d", i)
		id := i
		rate := float64(i) / 10
		ok := i%2 == 0
		at := now.Add(time.Duration(i) * time.Second)

		if multiSheet && i%2 == 1 {
			if asPointer {
				models = append(models, &Sheet2{
					Col1:  name,
					Col2:  id,
					Col3:  rate,
					Col4:  ok,
					Col5:  at,
					Col6:  &name,
					Col11: float32(rate),
				})
				continue
			}
			models = append(models, Sheet2{
				Col1:  name,
				Col2:  id,
				Col3:  rate,
				Col4:  ok,
				Col5:  at,
				Col6:  &name,
				Col11: float32(rate),
			})
			continue
		}

		if asPointer {
			models = append(models, &Sheet1{
				Col1: name,
				Col2: id,
				Col3: rate,
				Col4: ok,
				Col5: at,
				Col6: &name,
			})
			continue
		}
		models = append(models, Sheet1{
			Col1: name,
			Col2: id,
			Col3: rate,
			Col4: ok,
			Col5: at,
			Col6: &name,
		})
	}
	return models
}

func BenchmarkWriteExcelAsBytesBuffer(b *testing.B) {
	cases := []struct {
		name       string
		rows       int
		asPointer  bool
		multiSheet bool
		opts       []Option
	}{
		{name: "rows_200_struct", rows: 200},
		{name: "rows_1000_struct", rows: 1000},
		{name: "rows_1000_pointer", rows: 1000, asPointer: true},
		{name: "rows_1000_multi_sheet", rows: 1000, multiSheet: true},
		{name: "rows_1000_headless", rows: 1000, opts: []Option{WithHeadless()}},
		{name: "rows_1000_integer_as_string", rows: 1000, opts: []Option{WithIntegerAsString()}},
	}

	for _, tc := range cases {
		b.Run(tc.name, func(b *testing.B) {
			models := benchmarkModels(tc.rows, tc.asPointer, tc.multiSheet)
			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				buffer, err := WriteExcelAsBytesBuffer(models, tc.opts...)
				if err != nil {
					b.Fatalf("WriteExcelAsBytesBuffer failed: %v", err)
				}
				if buffer.Len() == 0 {
					b.Fatal("empty buffer")
				}
			}
		})
	}
}

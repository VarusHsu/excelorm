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

import "testing"

func BenchmarkReadExcelToModelsFromBytesBuffer(b *testing.B) {
	buffer, err := WriteExcelAsBytesBuffer(benchmarkModels(1000, false, false))
	if err != nil {
		b.Fatalf("prepare benchmark buffer failed: %v", err)
	}

	b.Run("rows_1000", func(b *testing.B) {
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			var rows []Sheet1
			if err := ReadExcelToModelsFromBytesBuffer(buffer, &rows); err != nil {
				b.Fatalf("ReadExcelToModelsFromBytesBuffer failed: %v", err)
			}
			if len(rows) == 0 {
				b.Fatal("empty rows")
			}
		}
	})

	b.Run("rows_1000_strict", func(b *testing.B) {
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			var rows []Sheet1
			if err := ReadExcelToModelsFromBytesBuffer(buffer, &rows, WithReadStrictMode()); err != nil {
				b.Fatalf("ReadExcelToModelsFromBytesBuffer strict failed: %v", err)
			}
			if len(rows) == 0 {
				b.Fatal("empty rows")
			}
		}
	})

	b.Run("rows_1000_pointer_slice", func(b *testing.B) {
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			var rows []*Sheet1
			if err := ReadExcelToModelsFromBytesBuffer(buffer, &rows); err != nil {
				b.Fatalf("ReadExcelToModelsFromBytesBuffer pointer slice failed: %v", err)
			}
			if len(rows) == 0 {
				b.Fatal("empty rows")
			}
		}
	})
}

func BenchmarkReadExcelToMapsFromBytesBuffer(b *testing.B) {
	buffer, err := WriteExcelAsBytesBuffer(benchmarkModels(1000, false, false))
	if err != nil {
		b.Fatalf("prepare benchmark buffer failed: %v", err)
	}

	b.Run("rows_1000", func(b *testing.B) {
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			rows, err := ReadExcelToMapsFromBytesBuffer(buffer, "sheet_one")
			if err != nil {
				b.Fatalf("ReadExcelToMapsFromBytesBuffer failed: %v", err)
			}
			if len(rows) == 0 {
				b.Fatal("empty rows")
			}
		}
	})

	b.Run("rows_1000_strict", func(b *testing.B) {
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			rows, err := ReadExcelToMapsFromBytesBuffer(buffer, "sheet_one", WithReadStrictMode())
			if err != nil {
				b.Fatalf("ReadExcelToMapsFromBytesBuffer strict failed: %v", err)
			}
			if len(rows) == 0 {
				b.Fatal("empty rows")
			}
		}
	})
}

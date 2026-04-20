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

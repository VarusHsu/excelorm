# excelorm
![Code Coverage](https://img.shields.io/codecov/c/github/varushsu/excelorm.svg)
![Build Status](https://github.com/varushsu/excelorm/actions/workflows/go.yml/badge.svg)
![GitHub Release](https://img.shields.io/github/v/release/varushsu/excelorm)
![GitHub](https://img.shields.io/github/license/varushsu/excelorm)
![Hits](https://hits.seeyoufarm.com/api/count/incr/badge.svg?url=https://github.com/varushsu/excelorm)

A lightweight and easy-to-use Excel file generation tool for Go.

## Installation
```shell
 go get -u 'github.com/varushsu/excelorm@latest'
```

## Quick Start
* Define a struct with `excel_header` tags and implement the `SheetName` method.
```go
type Foo struct {
    ID        int64      `excel_header:"id"`
    Name      string     `excel_header:"name"`
    CreatedAt time.Time  `excel_header:"created_at"`
    DeletedAt *time.Time `excel_header:"deleted_at"`
}
func (u Foo) SheetName() string {
    return "foo sheet name"
}
```

* Construct sample data.
```go
bar1DeletedAt := time.Date(2024, 1, 3, 15, 4, 5, 0, time.Local)
sheetModels := []excelorm.SheetModel{
    Foo{
        ID:        1,
        Name:      "Bar1",
        CreatedAt: time.Date(2024, 1, 2, 15, 4, 5, 0, time.Local),
        DeletedAt: &bar1DeletedAt,
    },
    Foo{
        ID:        2,
        Name:      "Bar2",
        CreatedAt: time.Date(2024, 1, 2, 15, 4, 5, 0, time.Local),
    },
}
```
* Write the data to an Excel file.
```go
if err := excelorm.WriteExcelSaveAs("foo.xlsx", sheetModels,
    excelorm.WithTimeFormatLayout("2006/01/02 15:04:05"),
    excelorm.WithIfNullValue("-"), 
); err != nil {
    log.Fatal(err)
}
```
* You should see the following output in the file:<br>

| id | name | created_at          | deleted_at          |
|----|------|---------------------|---------------------|
| 1  | Bar1 | 2024/01/02 15:04:05 | 2024/01/03 15:04:05 |
| 2  | Bar2 | 2024/01/02 15:04:05 | -                   |


[foo.xlsx](foo.xlsx)

* To support multiple sheets, define more structs that implement `SheetName`.

## Read APIs

### Read into models

Use `ReadExcelToModels` to parse one sheet into a typed slice. The target must be a pointer to slice (`*[]T` or `*[]*T`), and `T` must implement `SheetName`.

```go
var rows []Foo
if err := excelorm.ReadExcelToModels("foo.xlsx", &rows); err != nil {
    log.Fatal(err)
}
```

### Read into maps

Use `ReadExcelToMaps` to read a sheet as `[]map[string]string`.

```go
rows, err := excelorm.ReadExcelToMaps("foo.xlsx", "foo sheet name")
if err != nil {
    log.Fatal(err)
}
fmt.Println(rows[0]["id"])
```

### Strict mode

Strict mode validates header consistency:

- Rejects empty headers.
- Rejects duplicated headers.
- Rejects headers in Excel that do not map to model fields.
- Rejects model fields that are missing in Excel.

```go
var rows []Foo
if err := excelorm.ReadExcelToModels("foo.xlsx", &rows, excelorm.WithReadStrictMode()); err != nil {
    log.Fatal(err)
}
```

### Read options

- `WithReadTimeFormatLayout(layout)` for parsing `time.Time`.
- `WithReadBoolValueAs(trueValue, falseValue)` for custom bool values.
- `WithReadIfNullValue(value)` to map marker values to nil pointers.


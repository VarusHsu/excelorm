# excelorm
a easier use excel file create tool for golang

## install
```shell
 go get -u 'github.com/varushsu/excelorm@latest'
```

## Quick Start
* define a struct with excel_header tag and implement `SheetName` method
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

* construct some data
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
* write to excel file
```go
if err := excelorm.WriteExcelSaveAs("foo.xlsx", sheetModels,
    excelorm.WithTimeFormatLayout("2006/01/02 15:04:05"),
    excelorm.WithIfNullValue("-"), 
); err != nil {
    log.Fatal(err)
}
```
* you can see the result in the file<br>

| id | name | created_at          | deleted_at          |
|----|------|---------------------|---------------------|
| 1  | Bar1 | 2024/01/02 15:04:05 | 2024/01/03 15:04:05 |
| 2  | Bar2 | 2024/01/02 15:04:05 | -                   |


[foo.xlsx](foo.xlsx)

* support multi-sheets by define more structs

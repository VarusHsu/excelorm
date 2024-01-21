# excelorm
a easier use excel write tool for golang

## install
```shell
go get github.com/varushsu/excelorm
```

## Quick Start
* define a struct with excel_header tag and implement `SheetName` method
```go
type User struct {
Name     string    `excel_header:"姓名"`
Age      int       `excel_header:"年龄"`
Birthday time.Time `excel_header:"生日"`
Jobs     *string   `excel_header:"工作"`
}
func (u User) SheetName() string {
    return "用户信息"
}
```

* construct some data
```go
user1 := User{
    Name: "张三",
    Age: 18,
    Birthday: time.Now(),
    Jobs: nil,
}
user2 := User{
    Name: "李四",
    Age: 20,
    Birthday: time.Now(),
    Jobs: toPtr("程序员"),
}
sheetModels := make([]excelorm.SheetModel, 0)
sheetModels = append(sheetModels, user1, user2)
```
* write to excel file
```go
err := excelorm.WriteExcel("test.xlsx", sheetModels)
if err != nil {
    panic(err)
}
```
* you can see the result in the file<br>
[test.xlsx](test.xlsx)

* support multi sheet by define more struct
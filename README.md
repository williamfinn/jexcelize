# jexcelize

Module for parsing excelize sheets to json or maps. After creating json we can unmarshal to a model.

## Installation

```bash
go get github.com/williamfinn/jexcelize
```

## Usage

```go
// open the excel file using excelize
f, err := excelize.OpenFile("sample.xlsx")

if err != nil {
	fmt.Errorf("failed to open file %v", err)
}

// pass the file to the parser
parser := NewExcelParser(f)
// parse the specific sheet
j, err := parser.RowsToJson("users")
// here we print string(j), unmarshalling to model is more useful implementation
fmt.Println(string(j))
```

## License
[MIT](https://choosealicense.com/licenses/mit/)
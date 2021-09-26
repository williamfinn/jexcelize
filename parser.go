package jexcelize

import (
	"encoding/json"
	"errors"
	"github.com/xuri/excelize/v2"
)

type ExcelParser struct {
	file *excelize.File
}

func NewExcelParser(file *excelize.File) *ExcelParser {
	return &ExcelParser{file: file}
}

// RowsToString gets all rows in a given sheet to [][]string
func (p *ExcelParser) RowsToString(sheet string) ([][]string, error) {
	return p.file.GetRows(sheet)
}

// RowsToMap converts rows to []map[string]interface{}, this can then be used to convert to json
func (p *ExcelParser) RowsToMap(sheet string) ([]map[string]interface{}, error) {
	allRows, err := p.RowsToString(sheet)
	if err != nil {
		return nil, err
	}

	if len(allRows) < 2 {
		return nil, errors.New("there must be at least 2 rows in sheet")
	}

	header := allRows[0]
	rows := allRows[1:]

	length := len(header)

	var values []map[string]interface{}
	for _, row := range rows {
		dict := map[string]interface{}{}
		// matching up header values to cell values
		for i, val := range row {
			if i < length && val != "" {
				dict[header[i]] = val
			}
		}
		values = append(values, dict)
	}

	return values, nil
}

// RowsToJson converts the RowsToMap return value of []map[string]interface{} to []byte which can then be unmarshalled
func (p *ExcelParser) RowsToJson(sheet string) ([]byte, error) {
	dictRows, err := p.RowsToMap(sheet)
	if err != nil {
		return nil, err
	}
	j, err := json.Marshal(dictRows)
	if err != nil {
		return nil, err
	}
	return j, nil
}
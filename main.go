package jexcelize

import (
	"encoding/json"
	"errors"
	"github.com/360EntSecGroup-Skylar/excelize"
)

type ExcelParser struct {
	file *excelize.File
}

func NewExcelParser(file *excelize.File) *ExcelParser {
	return &ExcelParser{file: file}
}

// get all rows in a given sheet to [][]string
func (p *ExcelParser) RowsToString(sheet string) ([][]string, error) {
	rows, err := p.file.Rows(sheet)
	if err != nil {
		return [][]string{}, err
	}
	var columns [][]string
	for rows.Next() {
		columns = append(columns, rows.Columns())
	}
	return columns, nil
}

// converts rows to []map[string]interface{}, this can then be used to convert to json
func (p *ExcelParser) RowsToMap(sheet string) ([]map[string]interface{}, error) {
	allRows, err := p.RowsToString(sheet)
	if err != nil {
		return []map[string]interface{}{}, err
	}

	if len(allRows) < 2 {
		return []map[string]interface{}{}, errors.New("there must be at least 2 rows in sheet")
	}

	header := allRows[0]
	length := len(header)
	rows := allRows[1:]

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

// converts the RowsToMap return value of []map[string]interface{} to []byte which can then be unmarshalled
func (p *ExcelParser) RowsToJson(sheet string) ([]byte, error) {
	dictRows, err := p.RowsToMap(sheet)
	if err != nil {
		return []byte{}, err
	}
	j, err := json.Marshal(dictRows)
	if err != nil {
		return []byte{}, err
	}
	return j, nil
}
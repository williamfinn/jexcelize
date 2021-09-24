package jexcelize

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"testing"
)

func TestMapCreation(t *testing.T) {
	f, err := excelize.OpenFile("sample.xlsx")
	if err != nil {
		t.Errorf("failed to open file %v", err)
	}

	parser := NewExcelParser(f)
	m, err := parser.RowsToMap("users")
	fmt.Println(m)
}

func TestJsonCreation(t *testing.T) {
	f, err := excelize.OpenFile("sample.xlsx")
	if err != nil {
		t.Errorf("failed to open file %v", err)
	}

	parser := NewExcelParser(f)
	jsonBytes, err := parser.RowsToJson("users")
	fmt.Println(string(jsonBytes))
}
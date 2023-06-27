package pglmmyexcel

import (
	"fmt"
	"strconv"
	"strings"

	mylog "github.com/asatisa/pglmmylog"
	myutil "github.com/asatisa/pglmmyutil"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const version = "1.0.0.01" // Update column A, B for Customer Data

type ExcelRow struct {
	No    int    // Index no
	Data  string // Original string
	Udata string // Upper case string
}

func ReadExcel(filename string) bool {
	fmt.Print(filename)
	return true
}

func GetVersion() string {
	return version
}

//var asource1 []string //array of source1

var GlobalSource []ExcelRow

// Get maximum rows of excel file. // limit read from INI file. and set data to array
func GetExcelMaxRows_SetData(excel_filename string, excel_sheet_name string) int {
	f, err := excelize.OpenFile(excel_filename)
	if err != nil {
		fmt.Println(err)
		return -1
	}

	var current_row int = 0
	var axis = ""
	excel_read_max_rows, _ := strconv.Atoi(myutil.ReadINI("config", "excel_read_max_rows"))
	mylog.PrintInfo("Excel Compare Filename: " + excel_filename)
	mylog.PrintInfo("		Count all rows of excel")
	GlobalSource = nil
	readCol := "A"
	if strings.Contains(strings.ToUpper(excel_filename), "CUSTOMER.XLSX") {
		readCol = "B"
	} else {
		readCol = "A"
	}
	for i := 1; i <= excel_read_max_rows; i++ {
		current_row = i - 1
		axis = fmt.Sprintf(readCol+"%d", i)
		cellVal := f.GetCellValue(excel_sheet_name, axis)
		//asource1 = append(asource1, cellVal)
		var xrow ExcelRow
		xrow.No = i
		xrow.Data = cellVal
		xrow.Udata = strings.ToUpper(cellVal)

		if cellVal != "" {
			GlobalSource = append(GlobalSource, xrow)
		} else if cellVal == "" {
			mylog.PrintInfo("		Count rows = ", current_row)
			return current_row
		}
	}

	return 0
}

// Get maximum rows of excel file. // limit read from INI file.
func GetExcelMaxRows(excel_filename string, excel_sheet_name string) int {
	f, err := excelize.OpenFile(excel_filename)
	if err != nil {
		fmt.Println(err)
		return -1
	}

	var current_row int = 0
	var axis = ""
	excel_read_max_rows, _ := strconv.Atoi(myutil.ReadINI("config", "excel_read_max_rows"))
	mylog.PrintInfo("Excel Compare Filename: " + excel_filename)
	mylog.PrintInfo("		Count all rows of excel")
	for i := 1; i <= excel_read_max_rows; i++ {
		current_row = i - 1
		axis = fmt.Sprintf("A%d", i)
		cellVal := f.GetCellValue(excel_sheet_name, axis)
		if cellVal == "" {
			mylog.PrintInfo("		Count rows = ", current_row)
			return current_row
		}
	}

	return 0
}

// get excel value from row column axis.
func GetExcelValue(excel_filename string, axis string, excel_sheet_name string) string {
	f, err := excelize.OpenFile(excel_filename)
	if err != nil {
		fmt.Println(err)
		return "N_A"
	}

	cellVal := f.GetCellValue(excel_sheet_name, axis)
	return cellVal
}

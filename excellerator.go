package main 

import (
	"fmt"
	"os/exec"

	"github.com/xuri/excelize/v2"
	//"github.com/go-ole/go-ole"
	//"github.com/go-ole/go-ole/oleutil"
)

type ExcelInput struct {
	Cell  string `json:"cell"`
	Sheet string `json:"sheet"`
	Value interface{} `json:"value"`
}

type Config struct {
	Inputs  []ExcelInput  `json:"inputs"`
	Outputs []ExcelInput `json:"outputs"`
}

// RunExcel updates, solves and extracts results from Excel
// If outputs are configured, Excel input returns those values
func RunExcel(link string, config Config) ([]ExcelInput, error) {
	if (len(config.Inputs) > 0) {
		err := updateCells(link, config.Inputs)
		if err != nil {
			return nil, err
		}
	}

	if (len(config.Outputs) >0) {
		output, err := pullOutputs(link, config.Outputs)
		if err != nil {
			return nil, err
		}
		return output, nil
	}
	return nil, nil
}

func updateCells(filePath string, inputs []ExcelInput) error {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return err
	}
	for _, input := range inputs {
		f.SetCellValue(input.Sheet, input.Cell, input.Value)
	}
	if err := f.Save(); err != nil {
		return err
	}
	return nil
}

func runFormulasWithPowerShell(filePath string) error {
	cmd := exec.Command("powershell", "-Command", fmt.Sprintf(`$excel = New-Object -ComObject Excel.Application; $workbook = $excel.Workbooks.Open("%s"); $excel.Visible = $true; $workbook.Save(); $workbook.Close(); $excel.Quit()`, filePath))
	return cmd.Run()
}

func SolveExcelSheet(link string) error {
	/*
	// Initialize COM
    ole.CoInitialize(0)
    defer ole.CoUninitialize()

    // Connect to Excel
    excel, err := oleutil.CreateObject("Excel.Application")
    if err != nil {
        return fmt.Errorf("failed to create Excel object: %w", err)
    }

    excelDispatch := excel.MustQueryInterface(ole.IID_IDispatch)
    defer excelDispatch.Release()

    // Open the workbook
    workbooks := oleutil.MustGetProperty(excelDispatch, "Workbooks").ToIDispatch()
    defer workbooks.Release()
    workbook := oleutil.MustCallMethod(workbooks, "Open", link).ToIDispatch()
    defer workbook.Release()

    // Set calculation mode to automatic
	//oleutil.PutProperty(excelDispatch, "Calculation", -4105) // xlCalculationAutomatic

    // Refresh all data connections
    oleutil.MustCallMethod(workbook, "RefreshAll")

    // Recalculate workbook formulas
    oleutil.MustCallMethod(workbook, "Calculate")

    // Save and close the workbook
    oleutil.MustCallMethod(workbook, "Save")
    oleutil.MustCallMethod(workbook, "Close", false)

    // Quit Excel
    oleutil.MustCallMethod(excelDispatch, "Quit")
	*/

    return nil
}


func pullOutputs(filePath string, cells []ExcelInput) ([]ExcelInput, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	var outputs []ExcelInput
	for _, cell := range cells {

		value, err := f.CalcCellValue(cell.Sheet, cell.Cell)
		if err != nil {
			return nil, err
		}
		outputs = append(outputs, ExcelInput{Cell: cell.Cell, Sheet: cell.Sheet, Value: value})
	}
	return outputs, nil
}



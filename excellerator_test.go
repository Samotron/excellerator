package main 

import (
	"fmt"
	"math/rand/v2"
	"os"
	"path/filepath"
	"runtime"
	"testing"

)
/*
func TestUpdateCells(t *testing.T) {
	filePath := "test.xlsx"
	inputs := []ExcelInput{
		{Cell: "A1", Value: "Hello"},
		{Cell: "B1", Value: 123},
	}

	// Create a new Excel file for testing
	f := excelize.NewFile()
	if err := f.SaveAs(filePath); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}
	defer os.Remove(filePath)

	if err := updateCells(filePath, inputs); err != nil {
		t.Errorf("updateCells() error = %v", err)
	}

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		t.Fatalf("Failed to open test file: %v", err)
	}

	for _, input := range inputs {
		value, err := f.GetCellValue("Sheet1", input.Cell)
		if err != nil {
			t.Errorf("GetCellValue() error = %v", err)
		}
		if value != input.Value {
			t.Errorf("Expected cell %s to have value %v, but got %v", input.Cell, input.Value, value)
		}
	}
}

func TestRunFormulasWithPowerShell(t *testing.T) {
	// This test is a placeholder as running PowerShell commands in tests is not ideal
	// and may not work in all environments.
	t.Skip("Skipping test for runFormulasWithPowerShell due to environment dependency")
}
*/

func TestSolveExcelSheet(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("Can't run this test on windows")
	}
	input := rand.IntN(100)
	output := input + 5
	testFile := "./testSolve.xlsx"
	wd, _ := os.Getwd()
	testFile = filepath.Join(wd, testFile)

	inputs := make([]ExcelInput, 1)
	inputs[0] = ExcelInput{"A1", "Sheet1", input}
	err := updateCells(testFile, inputs)
	fmt.Println("Edit File")
	if err != nil {
		fmt.Println(err)
		t.Fail()
	}
	err = SolveExcelSheet(testFile)
	if err != nil {
		fmt.Println(err)
		t.Fail()
	}
	cells := make([]ExcelInput, 1)
	cells[0] = ExcelInput{"B1", "Sheet1", nil}
	res, err := pullOutputs(testFile, cells)
	fmt.Println(res)
	if res[0].Value != output {
		t.Errorf("Result was incorrect, want %d got %s", output, res[0].Value)
	}

}
/*
func TestPullOutputs(t *testing.T) {
	filePath := "test.xlsx"
	cells := []string{"A1", "B1"}
	expectedOutputs := []ExcelOutput{
		{Cell: "A1", Value: "Hello"},
		{Cell: "B1", Value: "123"},
	}

	// Create a new Excel file for testing
	f := excelize.NewFile()
	f.SetCellValue("Sheet1", "A1", "Hello")
	f.SetCellValue("Sheet1", "B1", 123)
	if err := f.SaveAs(filePath); err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}
	defer os.Remove(filePath)

	outputs, err := pullOutputs(filePath, cells)
	if err != nil {
		t.Errorf("pullOutputs() error = %v", err)
	}

	for i, output := range outputs {
		if output.Cell != expectedOutputs[i].Cell || output.Value != expectedOutputs[i].Value {
			t.Errorf("Expected output %v, but got %v", expectedOutputs[i], output)
		}
	}
}
*/


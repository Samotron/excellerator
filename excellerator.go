package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"io/ioutil"
	"os/exec"

	charmLog "github.com/charmbracelet/log"
	"github.com/xuri/excelize/v2"
)

type ExcelInput struct {
	Cell  string
	Value interface{}
}

type ExcelOutput struct {
	Cell  string
	Value interface{}
}

func updateCells(filePath string, inputs []ExcelInput) error {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return err
	}
	for _, input := range inputs {
		f.SetCellValue("Sheet1", input.Cell, input.Value)
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

func pullOutputs(filePath string, cells []string) ([]ExcelOutput, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	var outputs []ExcelOutput
	for _, cell := range cells {
		value, err := f.GetCellValue("Sheet1", cell)
		if err != nil {
			return nil, err
		}
		outputs = append(outputs, ExcelOutput{Cell: cell, Value: value})
	}
	return outputs, nil
}

type Config struct {
	Inputs  []ExcelInput  `json:"inputs"`
	Outputs []ExcelOutput `json:"outputs"`
}

func main() {
	excelFilePath := flag.String("excel", "", "Path to the Excel file")
	configFilePath := flag.String("config", "", "Path to the JSON config file")
	configJSON := flag.String("config-json", "", "JSON string of the config")
	flag.Parse()

	if *excelFilePath == "" {
		fmt.Println("Usage: excellerator -excel <path to excel file> [-config <path to config file> | -config-json <config json string>]")
		flag.PrintDefaults()
		charmLog.Fatal("Excel file path must be provided")
	}

	var config Config
	var configData []byte
	var err error

	if *configFilePath != "" {
		charmLog.Info("Reading config file", "path", *configFilePath)
		configData, err = ioutil.ReadFile(*configFilePath)
		if err != nil {
			charmLog.Fatalf("Failed to read config file: %v", err)
		}
	} else if *configJSON != "" {
		charmLog.Info("Using config JSON string")
		configData = []byte(*configJSON)
	}

	if configData != nil {
		charmLog.Info("Parsing config")
		if err := json.Unmarshal(configData, &config); err != nil {
			charmLog.Fatalf("Failed to parse config: %v", err)
		}

		charmLog.Info("Updating cells in Excel file", "path", *excelFilePath)
		if err := updateCells(*excelFilePath, config.Inputs); err != nil {
			charmLog.Fatalf("Failed to update cells: %v", err)
		}
	}

	charmLog.Info("Running formulas with PowerShell", "path", *excelFilePath)
	if err := runFormulasWithPowerShell(*excelFilePath); err != nil {
		charmLog.Fatalf("Failed to run formulas with PowerShell: %v", err)
	}

	if configData != nil {
		charmLog.Info("Pulling outputs from Excel file", "path", *excelFilePath)
		var outputCells []string
		for _, output := range config.Outputs {
			outputCells = append(outputCells, output.Cell)
		}
		outputs, err := pullOutputs(*excelFilePath, outputCells)
		if err != nil {
			charmLog.Fatalf("Failed to pull outputs: %v", err)
		}

		charmLog.Info("Marshalling output data to JSON")
		outputData, err := json.MarshalIndent(outputs, "", "  ")
		if err != nil {
			charmLog.Fatalf("Failed to marshal outputs: %v", err)
		}

		fmt.Println(string(outputData))
	}
}

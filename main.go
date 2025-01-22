package main 

import (
	"encoding/json"
	"fmt"
	"log"
	"os"

	"github.com/urfave/cli/v2"
)

func main() {
app := &cli.App{
		Name:  "Excellerator",
		Usage: "A CLI tool for automating MS Excel",
		Commands: []*cli.Command{
			{
				Name:  "generate",
				Usage: "Generate an empty template file",
				Action: func(c *cli.Context) error {
					// Create a sample template
					template := ExcelInput{
						Sheet: "Sheet1",
						Cell:  "A1",
						Value: "54",
					}
					config := Config{
						Inputs:  []ExcelInput{template},
						Outputs: []ExcelInput{template},
					}

					// Serialize the template to JSON
					jsonData, err := json.MarshalIndent(config, "", "    ")
					if err != nil {
						return fmt.Errorf("failed to generate template JSON: %w", err)
					}

					// Print the generated JSON template to the console
					fmt.Println(string(jsonData))
					return nil
				},
			},
			{
				Name:  "run",
				Usage: "Run the Excel automation with a config file and Excel file",
				Flags: []cli.Flag{
					&cli.StringFlag{
						Name:     "config",
						Usage:    "Path to the JSON config file",
						Required: true,
					},
					&cli.StringFlag{
						Name:     "excel",
						Usage:    "Path to the Excel file",
						Required: true,
					},
				},
				Action: func(c *cli.Context) error {
					configPath := c.String("config")
					excelPath := c.String("excel")

					// Read the JSON config file
					jsonData, err := os.ReadFile(configPath)
					if err != nil {
						return fmt.Errorf("failed to read input file: %w", err)
					}

					// Parse the JSON config file
					var config Config
					if err := json.Unmarshal(jsonData, &config); err != nil {
						return fmt.Errorf("failed to parse config file: %w", err)
					}

					// Process the Excel file with the given configuration
					result, err := RunExcel(excelPath, config)
					if err != nil {
						return fmt.Errorf("failed to process Excel file: %w", err)
					}

					// Output the result as JSON
					jsonString, err := json.MarshalIndent(result, "", "    ")
					if err != nil {
						return fmt.Errorf("failed to generate output JSON: %w", err)
					}
					fmt.Println("Processing Result:")
					fmt.Println(string(jsonString))
					return nil
				},
			},
		},
	}

	// Run the CLI application
	if err := app.Run(os.Args); err != nil {
		log.Fatal(err)
	}
}

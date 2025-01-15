//go:generate go run generate_schema.go

package main

import (
	"encoding/json"
	"log"
	"os"

	"github.com/invopop/jsonschema"
)

type ExcelInput struct {
	Cell  string
	Value interface{}
}

type ExcelOutput struct {
	Cell  string `json:"cell"`
	Value string `json:"value"`
}

type Config struct {
	Inputs  []ExcelInput  `json:"inputs"`
	Outputs []ExcelOutput `json:"outputs"`
}

func main() {
	reflector := &jsonschema.Reflector{
		DoNotReference: true,
	}
	schema := reflector.Reflect(&Config{})

	file, err := os.Create("../config_schema.json")
	if err != nil {
		log.Fatalf("Failed to create schema file: %v", err)
	}
	defer file.Close()

	encoder := json.NewEncoder(file)
	encoder.SetIndent("", "  ")
	if err := encoder.Encode(schema); err != nil {
		log.Fatalf("Failed to encode schema to file: %v", err)
	}

	log.Println("Schema generated successfully")
}

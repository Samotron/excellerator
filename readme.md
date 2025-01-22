# Excellerator

Automate Microsoft Excel files from go!

## Installing
```shell
go install github.com/samotron/excellerator@latest
```

## Usage
To generate a template input file
```shell
excellerator generate > template.json
```

The template file can be manually updated with the cells that inputs and outputs are required from

To automate an excel file, prints the updated cells to stdout
```shell
excellerator run --excel test.xlsx --config template.json
```

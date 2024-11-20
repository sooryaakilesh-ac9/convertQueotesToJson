package utils

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// Quote represents the structure for each quote in the JSON output
type Quote struct {
	ID       int64    `json:"id"`
	Text     string   `json:"text"`
	Author   string   `json:"author,omitempty"`
	Year     int      `json:"year,omitempty"`
	Context  string   `json:"context,omitempty"`
	Tags     []string `json:"tags"`
	Language string   `json:"lang"`
}

// Metadata represents additional metadata information
type Metadata struct {
	Version     string `json:"version"`
	LastUpdated string `json:"lastUpdated"`
	TotalQuotes int    `json:"totalQuotes"`
	URL         string `json:"url"`
	Schema      struct {
		Format   string `json:"format"`
		Encoding string `json:"encoding"`
		FileType string `json:"filetype"`
	} `json:"schema"`
}

// QuotesData holds the entire JSON structure with quotes and metadata
type QuotesData struct {
	Quotes   []Quote  `json:"quotes"`
}

// OpenExcelFile opens the Excel file
func OpenExcelFile(fileName string) (*excelize.File, error) {
	file, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, fmt.Errorf("failed to open Excel file %s: %w", fileName, err)
	}
	return file, nil
}

// ReadQuotesFromExcel processes the Excel file and outputs JSON with quotes and metadata
func ReadQuotesFromExcel(fileNameValue string) error {
	fileName := fileNameValue

	file, err := OpenExcelFile(fileName)
	if err != nil {
		log.Printf("Error opening Excel file: %v", err)
		return err
	}
	defer func() {
		if err := file.Close(); err != nil {
			log.Printf("Error closing the Excel file: %v", err)
		}
	}()

	return ReadExcelFile(file)
}

// ReadExcelFile reads data from the first sheet, processes it in batches, and outputs accumulated JSON
func ReadExcelFile(file *excelize.File) error {
	var accumulatedQuotes []Quote
	batchSize := 100 // Set your desired batch size

	// Get all sheet names
	sheets := file.GetSheetList()
	if len(sheets) == 0 {
		return fmt.Errorf("no sheets found in the Excel file")
	}

	// Access the first sheet
	sheetName := sheets[0]

	// Read all rows in the specified sheet
	rows, err := file.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("unable to load cells: %w", err)
	}

	// Process each row in batches
	var batch []Quote
	for i, row := range rows {
		if i == 0 {
			// Skip header row if present
			continue
		}
		if len(row) < 2 {
			log.Printf("Skipping row %d due to insufficient columns: %v", i, row)
			continue // Skip rows with insufficient columns
		}

		// Process tags by removing spaces and splitting by commas
		rawTags := strings.ReplaceAll(row[0], " ", "") // Remove spaces
		tags := strings.Split(rawTags, ",")            // Split by commas

		// Create a Quote struct with data from the row
		quote := Quote{
			ID:       int64(i), // Generate an ID
			Text:     row[1],   // Column 1 as the quote text
			Tags:     tags,     // Column 0 as tags
			Language: "en-US",  // Default language
		}

		// Add quote to the current batch
		batch = append(batch, quote)

		// If batch size is reached, add the batch to the accumulated list
		if len(batch) >= batchSize {
			accumulatedQuotes = append(accumulatedQuotes, batch...)
			batch = nil // Reset the batch
		}
	}

	// Add any remaining quotes from the last incomplete batch
	if len(batch) > 0 {
		accumulatedQuotes = append(accumulatedQuotes, batch...)
	}

	// Create metadata for the accumulated quotes
	metadata := Metadata{
		Version:     "1.0",
		LastUpdated: time.Now().Format(time.RFC3339),
		TotalQuotes: len(accumulatedQuotes),
		URL:         "path/to/file", // Set URL if available
	}
	metadata.Schema.Format = "JSON"
	metadata.Schema.Encoding = "UTF-8"
	metadata.Schema.FileType = "text"

	// Combine accumulated quotes and metadata into the final structure
	quotesData := QuotesData{
		Quotes:   accumulatedQuotes,
	}

	// Write the accumulated quotes to a JSON file
	if err := WriteJSONToFile("quotes.json", quotesData); err != nil {
		log.Printf("Error writing JSON to file: %v", err)
		return err
	}

	// converting metadata to json encoding
	jsonMetadata, err := json.MarshalIndent(metadata, "", " ")
	if err != nil {
		return fmt.Errorf("error marshalling metadata to JSON: %v", err)
	}

	// writing metadata json file
	if err := os.WriteFile("quotesMetadata.json", jsonMetadata, 0644); err != nil {
		return fmt.Errorf("error writing metadata.json %v", err)
	}

	fmt.Println("JSON data successfully written to quotes_output.json")
	return nil
}

// WriteJSONToFile saves the JSON data to a specified file
func WriteJSONToFile(filename string, data QuotesData) error {
	// Convert data to JSON format with indentation
	jsonData, err := json.MarshalIndent(data, "", "  ")
	if err != nil {
		return fmt.Errorf("error marshalling JSON: %w", err)
	}

	// Write JSON data to file
	err = os.WriteFile(filename, jsonData, 0644)
	if err != nil {
		return fmt.Errorf("error writing JSON to file: %w", err)
	}

	return nil
}

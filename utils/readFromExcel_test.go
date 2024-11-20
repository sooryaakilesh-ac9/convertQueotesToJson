package utils

import (
	"encoding/json"
	"os"
	"path/filepath"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
	"github.com/xuri/excelize/v2"
)

// TestOpenExcelFile tests the Excel file opening functionality
func TestOpenExcelFile(t *testing.T) {
	tests := []struct {
		name     string
		fileName string
		wantErr  bool
	}{
		{
			name:     "non_existent_file",
			fileName: "nonexistent.xlsx",
			wantErr:  true,
		},
		{
			name:     "invalid_extension",
			fileName: "test.txt",
			wantErr:  true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, err := OpenExcelFile(tt.fileName)
			if tt.wantErr {
				assert.Error(t, err)
			} else {
				assert.NoError(t, err)
			}
		})
	}
}

// createTestExcelFile creates a temporary Excel file for testing
func createTestExcelFile(t *testing.T) (*excelize.File, string) {
	f := excelize.NewFile()
	defer f.Close()

	// Add test data
	sheet1Name := "Sheet1"
	f.SetCellValue(sheet1Name, "A1", "Tags")
	f.SetCellValue(sheet1Name, "B1", "Quote")

	// Row 1: Normal case
	f.SetCellValue(sheet1Name, "A2", "inspiration,motivation")
	f.SetCellValue(sheet1Name, "B2", "Test quote 1")

	// Row 2: Empty tags
	f.SetCellValue(sheet1Name, "A3", "")
	f.SetCellValue(sheet1Name, "B3", "Test quote 2")

	// Row 3: Tags with spaces
	f.SetCellValue(sheet1Name, "A4", "wisdom, life, philosophy")
	f.SetCellValue(sheet1Name, "B4", "Test quote 3")

	// Create temporary file
	tmpDir := t.TempDir()
	tmpFile := filepath.Join(tmpDir, "test.xlsx")

	err := f.SaveAs(tmpFile)
	require.NoError(t, err)

	return f, tmpFile
}

// TestReadQuotesFromExcel tests the complete Excel to JSON conversion process
func TestReadQuotesFromExcel(t *testing.T) {
	_, tmpFile := createTestExcelFile(t)

	err := ReadQuotesFromExcel(tmpFile)
	assert.NoError(t, err)

	// Verify output files exist
	assert.FileExists(t, "quotes.json")
	assert.FileExists(t, "quotesMetadata.json")

	// Clean up
	os.Remove("quotes.json")
	os.Remove("quotesMetadata.json")
}

// TestReadExcelFile tests the Excel file reading and processing
func TestReadExcelFile(t *testing.T) {
	f, _ := createTestExcelFile(t)

	err := ReadExcelFile(f)
	assert.NoError(t, err)

	// Read and verify the generated JSON file
	data, err := os.ReadFile("quotes.json")
	require.NoError(t, err)

	var quotesData QuotesData
	err = json.Unmarshal(data, &quotesData)
	require.NoError(t, err)

	// Verify quotes content
	assert.Len(t, quotesData.Quotes, 3)

	// Verify first quote
	assert.Equal(t, "Test quote 1", quotesData.Quotes[0].Text)
	assert.Equal(t, []string{"inspiration", "motivation"}, quotesData.Quotes[0].Tags)

	// Verify quote with empty tags
	assert.Equal(t, "Test quote 2", quotesData.Quotes[1].Text)
	assert.Equal(t, []string{""}, quotesData.Quotes[1].Tags)

	// Verify quote with spaced tags
	assert.Equal(t, "Test quote 3", quotesData.Quotes[2].Text)
	assert.Equal(t, []string{"wisdom", "life", "philosophy"}, quotesData.Quotes[2].Tags)

	// Clean up
	os.Remove("quotes.json")
	os.Remove("quotesMetadata.json")
}

// TestWriteJSONToFile tests JSON file writing functionality
func TestWriteJSONToFile(t *testing.T) {
	tests := []struct {
		name      string
		filename  string
		data      QuotesData
		wantErr   bool
		setupFunc func()
		cleanup   func()
	}{
		{
			name:     "valid_write",
			filename: "test_quotes.json",
			data: QuotesData{
				Quotes: []Quote{
					{
						ID:       1,
						Text:     "Test quote",
						Tags:     []string{"test"},
						Language: "en-US",
					},
				},
			},
			wantErr: false,
			cleanup: func() {
				os.Remove("test_quotes.json")
			},
		},
		{
			name:     "invalid_permissions",
			filename: "/root/test_quotes.json", // Should fail due to permissions
			data: QuotesData{
				Quotes: []Quote{},
			},
			wantErr: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if tt.setupFunc != nil {
				tt.setupFunc()
			}

			err := WriteJSONToFile(tt.filename, tt.data)

			if tt.wantErr {
				assert.Error(t, err)
			} else {
				assert.NoError(t, err)

				// Verify file contents
				data, err := os.ReadFile(tt.filename)
				require.NoError(t, err)

				var quotesData QuotesData
				err = json.Unmarshal(data, &quotesData)
				require.NoError(t, err)

				assert.Equal(t, tt.data, quotesData)
			}

			if tt.cleanup != nil {
				tt.cleanup()
			}
		})
	}
}

// TestMetadataGeneration tests the metadata generation
func TestMetadataGeneration(t *testing.T) {
	f, _ := createTestExcelFile(t)

	err := ReadExcelFile(f)
	require.NoError(t, err)

	// Read and verify metadata file
	data, err := os.ReadFile("quotesMetadata.json")
	require.NoError(t, err)

	var metadata Metadata
	err = json.Unmarshal(data, &metadata)
	require.NoError(t, err)

	// Verify metadata fields
	assert.Equal(t, "1.0", metadata.Version)
	assert.Equal(t, 3, metadata.TotalQuotes)
	assert.Equal(t, "JSON", metadata.Schema.Format)
	assert.Equal(t, "UTF-8", metadata.Schema.Encoding)
	assert.Equal(t, "text", metadata.Schema.FileType)

	// Verify LastUpdated is a valid RFC3339 timestamp
	_, err = time.Parse(time.RFC3339, metadata.LastUpdated)
	assert.NoError(t, err)

	// Clean up
	os.Remove("quotes.json")
	os.Remove("quotesMetadata.json")
}

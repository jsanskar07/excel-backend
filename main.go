package main

import (
	"fmt"
	"log"
	"net/http"
	"sort"

	"github.com/xuri/excelize/v2"
)

func mergeExcelHandler(w http.ResponseWriter, r *http.Request) {
	// Enable CORS
	w.Header().Set("Access-Control-Allow-Origin", "*")
	w.Header().Set("Access-Control-Allow-Methods", "POST, OPTIONS")
	w.Header().Set("Access-Control-Allow-Headers", "Content-Type")

	if r.Method == http.MethodOptions {
		w.WriteHeader(http.StatusOK)
		return
	}

	if r.Method != http.MethodPost {
		http.Error(w, "Invalid request method", http.StatusMethodNotAllowed)
		return
	}

	// Parse multipart form
	err := r.ParseMultipartForm(50 << 20) // 50 MB max
	if err != nil {
		http.Error(w, "Failed to parse form", http.StatusBadRequest)
		return
	}

	files := r.MultipartForm.File["files"]
	if len(files) == 0 {
		http.Error(w, "No files uploaded", http.StatusBadRequest)
		return
	}

	// Sort files by filename (if filenames have date info)
	print("Files have been fetched")
	sort.Slice(files, func(i, j int) bool {
		return files[i].Filename < files[j].Filename
	})

	// Create new Excel file for merged data
	mergedFile := excelize.NewFile()
	mergedSheet := mergedFile.GetSheetName(0)
	currentRow := 1

	for _, fh := range files {
		f, err := fh.Open()
		if err != nil {
			log.Println("Error opening file:", err)
			continue
		}
		defer f.Close()

		excel, err := excelize.OpenReader(f)
		if err != nil {
			log.Println("Error reading Excel:", err)
			continue
		}

		sheets := excel.GetSheetList()
		for _, sheet := range sheets {
			rows, err := excel.GetRows(sheet)
			if err != nil {
				log.Println("Error getting rows:", err)
				continue
			}

			for i, row := range rows {
				// Skip header row for all files except first
				if i == 0 && currentRow != 1 {
					continue
				}
				for colIndex, cell := range row {
					cellName, _ := excelize.CoordinatesToCellName(colIndex+1, currentRow)
					mergedFile.SetCellValue(mergedSheet, cellName, cell)
				}
				currentRow++
			}
		}
	}

	// Send merged file as response
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", `attachment; filename="merged.xlsx"`)
	if err := mergedFile.Write(w); err != nil {
		log.Println("Error writing merged file:", err)
		http.Error(w, "Failed to write merged file", http.StatusInternalServerError)
		return
	}
	fmt.Println("Merged file sent successfully")
}

func main() {
	http.HandleFunc("/merge-excel", mergeExcelHandler)
	fmt.Println("Server running on :8080")
	log.Fatal(http.ListenAndServe(":8080", nil))
}

// package main

// import (
// 	"fmt"
// 	"os"

// 	"github.com/360EntSecGroup-Skylar/excelize"
// )

// func writeToExcel(data []string, fileName string, sheetName string) error {
// 	// Check if the file already exists
// 	if _, err := os.Stat(fileName); os.IsNotExist(err) {
// 		// Create a new Excel file
// 		file := excelize.NewFile()

// 		// Create a new sheet
// 		index := file.NewSheet(fileName)
// 	} else {
// 		file, err := excelize.OpenFile(fileName)
// 		if index := file.GetSheetIndex(sheetName); index == 0 {
// 			index := file.NewSheet(fileName)
// 		}
// 	}
// 	// Set the data in the sheet
// 	for row, rowData := range data {
// 		cell := excelize.ToAlphaString(col+1) + fmt.Sprint(row+1)
// 		file.SetCellValue(sheetName, cell, cellData)
// 	}

// 	// Save the Excel file
// 	if err := file.SaveAs(fileName); err != nil {
// 		return err
// 	}

// 	return nil
// }

package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"
	"unicode"

	"github.com/xuri/excelize/v2"
)

func main() {
	fileName := "conditDB.xlsx"
	sheets := []string{"parziale", "total"}
	var listOfRows [][]string
	var newColumns []string

	startTime := time.Now()

	file, err := OpenFile(fileName)
	if err != nil {
		return
	}
	defer CloseFile(file)

	for i := 0; i < len(sheets); i++ {
		listOfRows = [][]string{}
		newColumns = []string{}

		title := strings.ToUpper(sheets[i])
		fmt.Println("\n", title)

		fmt.Println("1/5 - Prendendo le colonne")
		rows := GetRows(fileName, sheets[i])
		oldColumns := GetColumns(rows)

		fmt.Println("2/5 - Impostando le colonne in camelCase")
		newColumns = setNewColumns(rows)

		fmt.Println("3/5 - Modificando i valori")
		for j := 0; j < len(newColumns); j++ {
			fmt.Printf("	%v/%v - %s\n", j+1, len(newColumns), oldColumns[j])

			values := setRightValuesAndToUpperCase(file, sheets[i], oldColumns[j])
			listOfRows = append(listOfRows, values)
		}

		fmt.Println("4/5 - Creando il file excel")
		createFile(sheets[i], newColumns, listOfRows)

		fmt.Println("5/5 - Creando il file csv")
		createCsvFile(sheets[i]+".xlsx", "Sheet1")
	}

	fmt.Println("FATTO!")
	fmt.Println("Tempo impiegato: ", time.Since(startTime))

}

func OpenFile(fileName string) (*excelize.File, error) {
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Println(err)
		return nil, err
	}

	return f, nil
}

func CloseFile(file *excelize.File) error {
	return file.Close()
}

// Get all the rows in the selected sheet.
func GetRows(fileName string, sheet string) [][]string {
	f, err := OpenFile(fileName)
	if err != nil {
		return nil
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println(err)
		return nil
	}

	defer CloseFile(f)

	return rows
}

func GetColumns(rows [][]string) []string {
	return rows[0]
}

// Remove new lines from string and set text format to camelCase
func setNewColumns(rows [][]string) []string {
	var columns []string
	for _, column := range GetColumns(rows) {
		columns = append(columns, toCamelCase(strings.Replace(column, "\n", "", -1)))
	}

	return columns
}

func toCamelCase(s string) string {
	var result string
	wasSpaceOrSlash := false

	for i, r := range s {

		if i == 0 {
			result += string(unicode.ToLower(r))
			continue
		}

		// If the preview value was a space or a slash then the next one must be in uppercase
		if wasSpaceOrSlash {
			result += string(unicode.ToUpper(r))
			wasSpaceOrSlash = false
			continue
		}

		if string(r) == " " || string(r) == "/" {
			wasSpaceOrSlash = true
			continue
		}

		// --- Aggiunte specifiche --- //
		if s == "Tipo ACT_OF_GOD" {
			result += string(r)
			continue
		}

		if s == "età veicolo" && string(r) == "à" {
			result += "a"
			continue
		}
		// --- Fine aggiunte specifiche --- //

		result += string(unicode.ToLower(r))
	}

	return result
}

func setRightValuesAndToUpperCase(f *excelize.File, sheet string, column string) []string {
	cols, err := f.Cols(sheet)
	if err != nil {
		fmt.Println(err)
		return nil
	}

	var values []string
	var value string
	for cols.Next() {
		col, err := cols.Rows(excelize.Options{RawCellValue: true})
		if err != nil {
			fmt.Println(err)
		}

		if col[0] == column {
			for index, rowCell := range col {
				if index > 0 {
					value = rowCell

					if column == "Tipo ACT_OF_GOD" && rowCell == "" {
						values = append(values, "OTHER")
						continue
					}

					if column == "Taxi" && rowCell == "all" {
						values = append(values, "NO")
						continue
					}

					if column == "Toyota Dealer Network" && rowCell == "all" {
						values = append(values, "SI")
						continue
					}

					if column == "VHL COMM" && rowCell == "all" {
						values = append(values, "NO")
						continue
					}

					if column == "Decurtazione" && rowCell == "" {
						values = append(values, "0")
						continue
					}

					isDate := rowCell != col[0] && (column == "Valido dal" || column == "Valido al" || column == "del")
					if isDate {
						values = append(values, formatDate(value))
						continue
					}

					needsUpperCase := column == "Garanzia" || column == "Tipo ACT_OF_GOD" || column == "Taxi" || column == "Toyota Dealer Network" || column == "Brand Lusso" || column == "Lexus" || column == "VHL COMM" || column == "val ass" || column == "Franchigia"
					isTotal := sheet == "total"
					if needsUpperCase || isTotal {
						values = append(values, strings.ToUpper(value))
						continue
					}

					values = append(values, value)
				}
			}

			continue
		}
	}

	return values
}

func formatDate(oldDate string) (newDate string) {
	YYYYMMDD := "2006-01-02"
	i, err := strconv.Atoi(oldDate)
	if err != nil {
		fmt.Println(err)
	}

	t := time.Date(1900, time.January, -1+i, 0, 0, 0, 0, time.UTC).Format(YYYYMMDD)
	return t
}

func createFile(sheet string, newColumns []string, listOfRows [][]string) {
	alphabet := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	err := f.SetSheetRow("Sheet1", "A1", &newColumns)
	if err != nil {
		fmt.Println(err)
	}

	for i := 0; i < len(newColumns); i++ {
		err = f.SetSheetCol("Sheet1", alphabet[i]+"2", &listOfRows[i])
		if err != nil {
			fmt.Println(err)
		}
	}

	if err := f.SaveAs(sheet + ".xlsx"); err != nil {
		fmt.Println(err)
	}
}

func createCsvFile(fileName string, sheet string) {
	rows := GetRows(fileName, sheet)

	csvFileName := strings.Trim(fileName, ".xlsx")
	csvFile, err := os.Create(csvFileName + ".csv")
	if err != nil {
		fmt.Println("Error while creating the CSV file")
	}

	csvWriter := csv.NewWriter(csvFile)

	for _, row := range rows {
		csvWriter.Write(row)
	}

	csvWriter.Flush()
	csvFile.Close()
}

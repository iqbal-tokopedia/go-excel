package main

import (
	"fmt"
	"math/rand"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func main() {
	type Data struct {
		ID          int32
		Name        string
		Address     string
		PhoneNumber string
		DoB         string
		City        string
		Job         string
		Age         int32
	}

	// this is data to be write to excel
	var datas []Data
	for i := int32(0); i < 1000; i++ {
		strI := strconv.Itoa(int(i))
		d := Data{
			ID:          i,
			Name:        "Name" + strI,
			Address:     "Address" + strI,
			PhoneNumber: "085" + strI,
			DoB:         "10-01-1975",
			City:        "Tasikmalaya",
			Job:         "Software Engineer",
			Age:         int32(rand.Intn(45-15) + 15),
		}

		datas = append(datas, d)
	}

	f := excelize.NewFile()

	// Create a new sheet.
	index := f.NewSheet("Sheet1")

	// Header
	f.SetCellValue("Sheet1", "A1", "ID")
	f.SetCellValue("Sheet1", "B1", "Name")
	f.SetCellValue("Sheet1", "C1", "Address")
	f.SetCellValue("Sheet1", "D1", "Phone Number")
	f.SetCellValue("Sheet1", "E1", "Date of Birth")
	f.SetCellValue("Sheet1", "F1", "City")
	f.SetCellValue("Sheet1", "G1", "Job")
	f.SetCellValue("Sheet1", "H1", "Age")

	//Body
	for k, v := range datas {
		strKey := strconv.Itoa(k + 2)
		f.SetCellValue("Sheet1", "A"+strKey, v.ID)
		f.SetCellValue("Sheet1", "B"+strKey, v.Name)
		f.SetCellValue("Sheet1", "C"+strKey, v.Address)
		f.SetCellValue("Sheet1", "D"+strKey, v.PhoneNumber)
		f.SetCellValue("Sheet1", "E"+strKey, v.DoB)
		f.SetCellValue("Sheet1", "F"+strKey, v.City)
		f.SetCellValue("Sheet1", "G"+strKey, v.Job)
		f.SetCellValue("Sheet1", "H"+strKey, v.Age)
	}

	// Set active sheet of the workbook.
	f.SetActiveSheet(index)

	// Save spreadsheet by the given path.
	if err := f.SaveAs("your-filename-goes-here.xlsx"); err != nil {
		fmt.Println(err)
	}
}

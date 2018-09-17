package main

import (
	"context"
	"database/sql"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	_ "github.com/denisenkom/go-mssqldb"
	"log"
	"strconv"
)



func main() {
	xlsx, err := excelize.OpenFile("./sampleWrite.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
		var server = "winedb.database.windows.net"
	var port = 1433
	var user = "guy"
	var password = "Wilks2003"
	var database = "winedb"
	connString := fmt.Sprintf("server=%s;user id=%s;password=%s;port=%d;database=%s;",
		server, user, password, port, database)
	db, err := sql.Open("sqlserver", connString)
	if err != nil {
		log.Fatal("Error creating connection pool: ", err.Error())
	}
	ctx := context.Background()
	err = db.PingContext(ctx)
	if err != nil {
		log.Fatal(err.Error())
	}
	fmt.Printf("Connected!\n")
	var total int
	row := db.QueryRow("SELECT COUNT(*) FROM wineTest")
	row.Scan(&total)
	fmt.Println(total)
	names := make([]string, total)
	appelation := make([]string, total)
	composition := make([]string, total)
	targetYear := make([]string, total)
	points := make([]string, total)
	dateDrunk := make([]string, total)
	grade:= make([]string, total)
	comment := make([]string, total)
	price := make([]string, total)
	yearBought := make([]string, total)
	kLSKU := make([]string, total)
	quantity := make([]string, total)

	var temp string
	var insert string
	for a := 6538; a < len(names)+6538; a++{
		insert = ("SELECT nameOfWine FROM wineTest WHERE id = "+strconv.Itoa(a)+";")
		data := db.QueryRow(insert)
		data.Scan(&temp)
		fmt.Println(temp)
		names[a-6538] = temp
		insert = ("SELECT price FROM wineTest WHERE id = "+strconv.Itoa(a)+";")
		data = db.QueryRow(insert)
		data.Scan(&temp)
		price[a-6538] = temp
		insert = ("SELECT appelation FROM wineTest WHERE id = "+strconv.Itoa(a)+";")
		data = db.QueryRow(insert)
		data.Scan(&temp)
		appelation[a-6538] = temp
		insert = ("SELECT composition FROM wineTest WHERE id = "+strconv.Itoa(a)+";")
		data = db.QueryRow(insert)
		data.Scan(&temp)
		composition[a-6538] = temp
		insert = ("SELECT targetYear FROM wineTest WHERE id = "+strconv.Itoa(a)+";")
		data = db.QueryRow(insert)
		data.Scan(&temp)
		targetYear[a-6538] = temp
		insert = ("SELECT points FROM wineTest WHERE id = "+strconv.Itoa(a)+";")
		data = db.QueryRow(insert)
		data.Scan(&temp)
		points[a-6538] = temp
		insert = ("SELECT dateDrunk FROM wineTest WHERE id = "+strconv.Itoa(a)+";")
		data = db.QueryRow(insert)
		data.Scan(&temp)
		dateDrunk[a-6538] = temp
		insert = ("SELECT grade FROM wineTest WHERE id = "+strconv.Itoa(a)+";")
		data = db.QueryRow(insert)
		data.Scan(&temp)
		grade[a-6538] = temp
		insert = ("SELECT comments FROM wineTest WHERE id = "+strconv.Itoa(a)+";")
		data = db.QueryRow(insert)
		data.Scan(&temp)
		comment[a-6538] = temp
		insert = ("SELECT yearBought FROM wineTest WHERE id = "+strconv.Itoa(a)+";")
		data = db.QueryRow(insert)
		data.Scan(&temp)
		yearBought[a-6538] = temp
		insert = ("SELECT KLSKU FROM wineTest WHERE id = "+strconv.Itoa(a)+";")
		data = db.QueryRow(insert)
		data.Scan(&temp)
		kLSKU[a-6538] = temp
		insert = ("SELECT quantity FROM wineTest WHERE id = "+strconv.Itoa(a)+";")
		data = db.QueryRow(insert)
		data.Scan(&temp)
		quantity[a-6538] = temp
	}
	var tempCon string = "";
	for b := 0; b < len(names); b++{
		tempCon = "A" + strconv.Itoa(b+1)
		xlsx.SetCellValue("Sheet1", tempCon , names[b])
		tempCon = "B" + strconv.Itoa(b+1)
		xlsx.SetCellValue("Sheet1", tempCon , appelation[b])
		tempCon = "C" + strconv.Itoa(b+1)
		xlsx.SetCellValue("Sheet1", tempCon , composition[b])
		tempCon = "D" + strconv.Itoa(b+1)
		xlsx.SetCellValue("Sheet1", tempCon , targetYear[b])
		tempCon = "E" + strconv.Itoa(b+1)
		xlsx.SetCellValue("Sheet1", tempCon , points[b])
		tempCon = "F" + strconv.Itoa(b+1)
		xlsx.SetCellValue("Sheet1", tempCon , dateDrunk[b])
		tempCon = "G" + strconv.Itoa(b+1)
		xlsx.SetCellValue("Sheet1", tempCon , grade[b])
		tempCon = "H" + strconv.Itoa(b+1)
		xlsx.SetCellValue("Sheet1", tempCon , comment[b])
		tempCon = "I" + strconv.Itoa(b+1)
		xlsx.SetCellValue("Sheet1", tempCon , price[b])
		tempCon = "J" + strconv.Itoa(b+1)
		xlsx.SetCellValue("Sheet1", tempCon , yearBought[b])
		tempCon = "K" + strconv.Itoa(b+1)
		xlsx.SetCellValue("Sheet1", tempCon , kLSKU[b])
		tempCon = "L" + strconv.Itoa(b+1)
		xlsx.SetCellValue("Sheet1", tempCon , quantity[b])
	}
	xlsx.Save()
}

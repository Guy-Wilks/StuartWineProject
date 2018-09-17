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


func FloatToString(input_num float64) string {
	// to convert a float number to a string
	return strconv.FormatFloat(input_num, 'f', 6, 64)
}

var server = "winedb.database.windows.net"
var port = 1433
var user = "guy"
var password = "Wilks2003"
//TODO configure the regular wine table like the test wine table
var database = "winedb"
var names [1220]string
var appelation [1220]string
var composition [1220]string
var targetYear [1220]string
var points [1220]string
var dateDrunk [1220]string
var grade [1220]string
var comment [1220]string
var price [1220]string
var yearBought[1220]string
var kLSKU [1220]string
var quantity [1220]string

func main() {
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
	xlsx, err := excelize.OpenFile("./wineEXcl.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	// Get value from cell by given worksheet name and axis.
	for i := 0; i < len(names); i++ {
		temp := "A" + strconv.Itoa(i+1);
		names[i] = xlsx.GetCellValue("Sheet1", temp)
		temp = "B" + strconv.Itoa(i+1);
		appelation[i] = xlsx.GetCellValue("Sheet1", temp)
		temp = "C" + strconv.Itoa(i+1);
		composition[i] = xlsx.GetCellValue("Sheet1", temp)
		temp = "D" + strconv.Itoa(i+1);
		targetYear[i] = xlsx.GetCellValue("Sheet1", temp)
		temp = "E" + strconv.Itoa(i+1);
		points[i] = xlsx.GetCellValue("Sheet1", temp)
		temp = "F" + strconv.Itoa(i+1);
		dateDrunk[i] = xlsx.GetCellValue("Sheet1", temp)
		temp = "G" + strconv.Itoa(i+1);
		grade[i] = xlsx.GetCellValue("Sheet1", temp)
		temp = "H" + strconv.Itoa(i+1);
		comment[i] = xlsx.GetCellValue("Sheet1", temp)
		temp = "I" + strconv.Itoa(i+1);
		price[i]= xlsx.GetCellValue("Sheet1", temp)
		temp = "J" + strconv.Itoa(i+1);
		yearBought[i] = xlsx.GetCellValue("Sheet1", temp)
		temp = "K" + strconv.Itoa(i+1);
		kLSKU[i] = xlsx.GetCellValue("Sheet1", temp)
		temp = "L" + strconv.Itoa(i+1);
		quantity[i] = xlsx.GetCellValue("Sheet1", temp)
	}
	for a := 0; a < len(names); a++{
		err = db.PingContext(ctx)
		if err != nil {
			log.Fatal(err.Error())
		}
		var insert string = "INSERT INTO wineTest (nameOfWine, price, appelation, composition, targetYear, points, dateDrunk, grade, comments, yearBought, KLSKU, quantity) VALUES ("+ "'" + names [a] + "'"+","+ "'"+price[a]+ "'"+","+ "'"+appelation[a]+ "'"+","+ "'"+composition[a]+ "'"+","+ "'"+targetYear[a]+ "'"+","+ "'"+points[a]+ "'"+","+ "'"+dateDrunk[a]+ "'"+","+ "'"+grade[a]+ "'"+","+ "'"+comment[a]+ "'"+","+ "'"+yearBought[a]+ "'"+","+ "'"+kLSKU[a]+ "'"+","+ "'"+quantity[a]+ "'"+");"
		//fmt.Println(insert)
		db.Exec(insert)

		//		fmt.Println("Name: "+ names[a] + ", appelation: " +appelation[a]+ ", composition: " + composition[a] + ", targetYear: "+targetYear[a]+", points: "+points[a]+", date Drunk: "+dateDrunk[a]+", grade: "+grade[a]+", comment: "+comment[a]+", price: " + FloatToString(price[a]) + ", year bought: " + yearBought[a]+", K&L SKU: "+kLSKU[a]+", quantity: "+ quantity[a])
	}

}

package main

import (
	"database/sql"
	"fmt"
	"time"

	"github.com/mattn/go-adodb"
	_ "github.com/mattn/go-adodb"
)

const (
	Provider = "Provider=WinCCOLEDBProvider.1;"
	DSN      = "Catalog=CC_OS_1__21_12_14_16_25_11R;"
	DS       = "Data Source=10.1.1.1\\WINCC"
	db_type  = "adodb_with_cursorlocation"
)

type Data struct {
	valueId   int
	timeStamp string
	realValue float64
	quality   string
	flags     string
}

func main() {
	fmt.Println("Coba Koneksi WinCC DB! 5")

	conn_string := `Provider=WinCCOLEDBProvider.1;Persist Security Info=False;User ID="";Data Source=10.1.1.1\WINCC;Catalog=CC_OS_1__21_12_14_16_25_11R;Mode=Read;Location="";Mode=Read;Extended Properties=""`

	sql.Register(db_type, &adodb.AdodbDriver{
		CursorLocation: 3,
	})

	db, err := sql.Open(db_type, conn_string)
	if err != nil {
		fmt.Println("open", err)
		delay()
		return
	}
	defer db.Close()

	err = db.Ping()
	if err != nil {
		fmt.Println("Ping error : ", err)
		delay()
		return
	}

	query := `TAG:R,'SystemArchive\1511_AT_1379A/MS.PV_Out#Value','2022-02-08 07:35:00.000', '2022-02-08 07:40:00.000'`
	row, err := db.Query(query)
	if err != nil {
		fmt.Println("Query Error : ", err)
		delay()
		return
	}
	fmt.Println("Query Executed Sucesfully!")
	defer row.Close()

	var result = []Data{}
	for row.Next() {
		var res Data
		if err := row.Scan(
			&res.valueId,
			&res.timeStamp,
			&res.realValue,
			&res.quality,
			&res.flags); err != nil {
			fmt.Println("Error Scanning", err)
			delay()
			return
		}
		result = append(result, res)
	}

	err = row.Err()
	if err != nil {
		fmt.Println("Error row : ", err)
		delay()
		return
	}

	for _, index := range result {
		fmt.Print(index.valueId, "\t")
		fmt.Print(index.timeStamp, "\t")
		fmt.Print(index.realValue, "\t")
		fmt.Print(index.quality, "\t")
		fmt.Print(index.flags, "\t")
	}

}

func delay() {
	duration := time.Duration(10) * time.Second
	time.Sleep(duration)
}

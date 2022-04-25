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
	DSN      = "Catalog=CC_OS_1__21_12_14_16_25_11R"
	DS       = "Data Source=10.1.1.1\\WINCC;"
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

	// Original(tested) Query :
	// `Provider=WinCCOLEDBProvider.1;Persist Security Info=False;User ID="";Data Source=10.1.1.1\WINCC;Catalog=CC_OS_1__21_12_14_16_25_11R;Mode=Read;Location="";Mode=Read;Extended Properties=""`

	conn_string := Provider + DS + DSN

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

	output := "SystemArchive\\1511_AT_1379A/MS.PV_Out#Value"
	loc := time.FixedZone("UTC+7", +7*60*60)
	startTime := time.Date(2022, 02, 8, 7, 35, 0, 0, loc)
	endTime := time.Date(2022, 02, 8, 7, 40, 0, 0, loc)

	// Tested Query :
	//`TAG:R,'SystemArchive\1511_AT_1379A/MS.PV_Out#Value','2022-02-08 07:35:00.000', '2022-02-08 07:40:00.000'`

	query := "TAG:R '" + output + "', '" + startTime.String() + "', '" + endTime.String() + "' "
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

	for _, content := range result {
		fmt.Print(content.valueId, "\t")
		fmt.Print(content.timeStamp, "\t")
		fmt.Print(content.realValue, "\t")
		fmt.Print(content.quality, "\t")
		fmt.Println(content.flags, "\t")
	}

}

func delay() {
	duration := time.Duration(10) * time.Second
	time.Sleep(duration)
}

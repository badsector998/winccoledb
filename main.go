package main

import (
	"database/sql"
	"fmt"
	"time"

	"github.com/mattn/go-adodb"
	_ "github.com/mattn/go-adodb"
)

const (
	//provider harus cek lagi di server PDK
	Provider = "Provider=WinCCOLEDBProvider.1;"
	DSN      = "Catalog=CC_OS_1__21_12_14_16_25_11R;"
	DS       = "Data Source=10.1.1.1\\WINCC"
	db_type  = "adodb"
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

	sql.Register("adodb_with_cursorlocation", &adodb.AdodbDriver{
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

	// command, err := oleutil.CreateObject("ADODB.Command")

	query := `TAG:R,'SystemArchive\1511_AT_1379A/MS.PV_Out#Value','2022-02-08 07:35:00.000', '2022-02-08 07:40:00.000'`
	row, err := db.Query(query)
	if err != nil {
		fmt.Println("Query Error : ", err)
		delay()
		return
	}
	fmt.Println("Query Executed Sucesfully!")
	defer row.Close()

	var result []Data
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

}

func delay() {
	duration := time.Duration(10) * time.Second
	time.Sleep(duration)
}

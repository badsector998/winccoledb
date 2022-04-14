package main

import (
	"database/sql"
	"fmt"
	"time"

	"github.com/go-ole/go-ole"
	_ "github.com/mattn/go-adodb"
)

const (
	//provider harus cek lagi di server PDK
	Provider = "Provider=WinCCOLEDBProvider.1;"
	DSN      = "Catalog=CC_OS_1__21_12_14_16_25_11R;"
	DS       = "Data Source=10.1.1.1\\WINCC"
	db_type  = "adodb"
)

func main() {
	fmt.Println("Coba Koneksi WinCC DB!")
	ole.CoInitialize(0)
	defer ole.CoUninitialize()
	conn_string := Provider + DSN + DS
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
	defer row.Close()
	for row.Next() {
		var (
			valueId   int
			timeStamp string
			realValue string
			quality   string
			flags     string
		)
		// var valueName string
		err = row.Scan(&valueId, &timeStamp, &realValue, &quality, &flags)
		if err != nil {
			fmt.Println("Row Error : ", err)
		}
		fmt.Println(valueId, timeStamp, realValue, quality, flags)
	}

	err = row.Err()
	if err != nil {
		fmt.Println("Error row : ", err)
		delay()
		return
	}

	duration := time.Duration(10) * time.Second
	time.Sleep(duration)
}

func delay() {
	duration := time.Duration(10) * time.Second
	time.Sleep(duration)
}
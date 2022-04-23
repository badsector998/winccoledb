package main

import (
	"fmt"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

const (
	//provider harus cek lagi di server PDK
	Provider = "Provider=WinCCOLEDBProvider.1;"
	DSN      = "Catalog=CC_OS_1__21_12_14_16_25_11R;"
	DS       = "Data Source=10.1.1.1\\WINCC"
	db_type  = "adodb"
	DDSN     = "Catalog=CC_OS_1__21_12_14_16_25_11R"
)

func main() {
	fmt.Println("OLE TEST PROGRAM")

	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	query := `TAG:R,'SystemArchive\1511_AT_1379A/MS.PV_Out#Value','2022-02-08 07:35:00.000', '2022-02-08 07:40:00.000'`
	conn_string := Provider + DSN + DS

	//to do 4 : Execute Query then immedietly stores into RecorsetObject

	//to do 1 : Create Connection
	conn_service, err := oleutil.CreateObject("ADODB.Connection")
	if err != nil {
		fmt.Println("Error Creating Connection Object", err)
		return
	}

	db, err := conn_service.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		fmt.Println("Error DB Connection", err)
		return
	}
	conn_service.Release()

	_, _ = oleutil.PutProperty(db, "CursorLocation", 3)

	//to do 2 : Open Connection
	_, err = oleutil.CallMethod(db, "Open", conn_string)
	if err != nil {
		fmt.Println("Error DB Connection on Ping", err)
		return
	}

	//to do 3 : Create Query Statement
	cmd, err := oleutil.CreateObject("ADODB.Command")
	if err != nil {
		fmt.Print("Error Creating Command Object", err)
		return
	}

	sq, err := cmd.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		fmt.Println("Error Creating Query Interface", err)
		return
	}
	cmd.Release()

	_, err = oleutil.PutProperty(sq, "ActiveConnection", db)
	if err != nil {
		fmt.Println("Error Allocating Active Connection", err)
		return
	}

	_, err = oleutil.PutProperty(sq, "CommandText", query)
	if err != nil {
		fmt.Println("Error Preparing Query", err)
		return
	}

	_, err = oleutil.PutProperty(sq, "CommandType", 1)
	if err != nil {
		fmt.Println("Error Preparing Command Type", err)
		return
	}

	_, err = oleutil.PutProperty(sq, "Prepared", true)
	if err != nil {
		fmt.Println("Error Prepared Statement", err)
		return
	}

	// a. Create ADODB.RecordSet object
	// b. Convert ADODB.RecordSet into *ole.VARIANT type
	result, err := oleutil.CallMethod(sq, "Execute")
	if err != nil {
		fmt.Println("Result error", err)
		return
	}

	fmt.Println(result)

}

// func delay() {
// 	duration := time.Duration(10) * time.Second
// 	time.Sleep(duration)
// }

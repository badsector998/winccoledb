package main

import (
	"fmt"
	"time"
)

func main() {
	output := "SystemArchive\\1511_AT_1379A/MS.PV_Out#Value"
	loc := time.FixedZone("UTC+7", +7*60*60)
	startTime := time.Date(2022, 02, 8, 7, 35, 0, 0, loc)
	endTime := time.Date(2022, 02, 8, 7, 40, 0, 0, loc)

	fmt.Println(output)
	fmt.Println(loc)
	fmt.Println(startTime)
	fmt.Println(endTime)
}

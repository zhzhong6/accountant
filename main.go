package main

import (
	"fmt"
	"time"
)

func main() {
 str :="2021-06-24 04:11:20"
s,e:=	excelDateToDate(str)
fmt.Println(s,e)
}

func excelDateToDate(excelDate string) (time.Time,error) {
	baghdad, err := time.LoadLocation("Local")
	fmt.Println(err)
	return   time.ParseInLocation("2006-01-02 15:04:05"    , excelDate, baghdad)

}
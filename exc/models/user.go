package models

import (
	"fmt"
	"reflect"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"

	"accountant/exc/common"
)

const (
	filePath = "F:/excel/user.xlsx"
	Sheet1   = "Sheet1"
	filePath2 = "F:/excel/user2.xlsx"
)

type User struct {
	UserId   int       `excel:"用户Id"`
	Name     string    `excel:"用户名"`
	Age      int       `excel:"age"`
	Account  float64   `excel:"账户余额"`
	Birthday time.Time `excel:"出生"`  // 时间格式"yyyy-mm-dd hh:mm:ss"
}

func ImportUser() ([]*User, error) {
	u := make([]*User, 0)
	err := common.ImportFileToList(filePath, Sheet1, &User{}, &u)
	if err != nil {
		return nil, err
	}
	u =append(u,u...)
	err = common.OutputSliceToFile(filePath2, Sheet1, &u)
	return nil, nil
}

func RefactorWrite(records []*User) {
	xlsx := excelize.NewFile()
	index := xlsx.NewSheet("Sheet1")

	for i, t := range records {
		d := reflect.TypeOf(t).Elem()
		for j := 0; j < d.NumField(); j++ {
			// 设置表头
			if i == 0 {
				column := strings.Split(d.Field(j).Tag.Get("xlsx"), "-")[0]
				name := strings.Split(d.Field(j).Tag.Get("xlsx"), "-")[1]
				xlsx.SetCellValue("Sheet1", fmt.Sprintf("%s%d", column, i+1), name)
			}
			// 设置内容
			column := strings.Split(d.Field(j).Tag.Get("xlsx"), "-")[0]
			switch d.Field(j).Type.String() {
			case "string":
				xlsx.SetCellValue("Sheet1", fmt.Sprintf("%s%d", column, i+2), reflect.ValueOf(t).Elem().Field(j).String())
			case "int32":
				xlsx.SetCellValue("Sheet1", fmt.Sprintf("%s%d", column, i+2), reflect.ValueOf(t).Elem().Field(j).Int())
			case "int64":
				xlsx.SetCellValue("Sheet1", fmt.Sprintf("%s%d", column, i+2), reflect.ValueOf(t).Elem().Field(j).Int())
			case "bool":
				xlsx.SetCellValue("Sheet1", fmt.Sprintf("%s%d", column, i+2), reflect.ValueOf(t).Elem().Field(j).Bool())
			case "float32":
				xlsx.SetCellValue("Sheet1", fmt.Sprintf("%s%d", column, i+2), reflect.ValueOf(t).Elem().Field(j).Float())
			case "float64":
				xlsx.SetCellValue("Sheet1", fmt.Sprintf("%s%d", column, i+2), reflect.ValueOf(t).Elem().Field(j).Float())
				//case "time":
				//	xlsx.SetCellValue("Sheet1", fmt.Sprintf("%s%d", column, i+2), reflect.ValueOf(t).Elem().Field(j).time())
			}
		}
	}

	xlsx.SetActiveSheet(index)
	// 保存到xlsx中
	err := xlsx.SaveAs(filePath)
	if err != nil {
		fmt.Println(err)
	}
}

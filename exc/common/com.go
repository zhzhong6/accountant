package common

import (
	"encoding/json"
	"errors"
	"fmt"
	"reflect"
	"strconv"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func TypeMap(b interface{}) (map[string]string, map[string]string) {
	tr := make(map[string]string)
	fr := make(map[string]string)
	d := reflect.TypeOf(b).Elem()
	for j := 0; j < d.NumField(); j++ {
		tr[d.Field(j).Tag.Get("excel")] = d.Field(j).Type.String()
		fr[d.Field(j).Tag.Get("excel")] = d.Field(j).Name
	}
	return tr, fr
}

func ImportFileToList(filePath, sheetName string, modelOne interface{}, result interface{}) error {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		fmt.Println(err)
		return err
	}

	// Get all the rows in the Sheet1.
	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println(err)
		return err
	}

	if len(rows) == 0 {
		return errors.New(" rows length == 0")
	}
	keys := make([]string, 0)
	for _, c := range rows[0] {
		keys = append(keys, c)
	}

	typeMap, fieldMap := TypeMap(modelOne)

	res := make([]map[string]interface{}, 0)
	for i := 1; i < len(rows); i++ {
		r1 := make(map[string]interface{})
		for j := 0; j < len(rows[i]); j++ {
			var val interface{}
			var err error
			val = rows[i][j]
			switch typeMap[keys[j]] {
			case "int", "int32", "int64":
				val, err = strconv.Atoi(val.(string))
			case "float32", "float64":
				val, err = strconv.ParseFloat(val.(string), 64)
			case "bool":
				val, err = strconv.ParseBool(val.(string))
			case "time.Time":
				val, err = stringToTime(val.(string))
			default:
			}

			if err != nil {
				return err
			}
			r1[fieldMap[keys[j]]] = val

		}
		res = append(res, r1)
	}
	rr, err := json.Marshal(res)
	if err != nil {
		return err
	}
	err = json.Unmarshal(rr, result)
	return nil
}

func OutputSliceToFile(filePath, sheetName string, records interface{}) error {
	var err error
	reflectValue := reflect.Indirect(reflect.ValueOf(records))
	if reflectValue.Kind() != reflect.Slice {
		return errors.New(" must be slice ")
	}
	reflectLen := reflectValue.Len()
	mList := reflectValue.Slice(0, reflectLen)

	xlsx := excelize.NewFile()
	index := xlsx.NewSheet(sheetName)
	for i := 0; i < reflectLen; i++ {
		m := mList.Index(i).Elem().Interface()
		t := reflect.TypeOf(m)
		v := reflect.ValueOf(m)
		for j := 0; j < t.NumField(); j++ {
			column :=  j+'A'

			switch t.Field(j).Type.String() {
			case "string":
				err = xlsx.SetCellValue("Sheet1", fmt.Sprintf("%c%d", column, i+2), v.Field(j).String())
			case "int", "int32", "int64":
				err = xlsx.SetCellValue("Sheet1", fmt.Sprintf("%c%d", column, i+2), v.Field(j).Int())
			case "bool":
				err = xlsx.SetCellValue("Sheet1", fmt.Sprintf("%c%d", column, i+2), v.Field(j).Bool())
			case "float32", "float64":
				err = xlsx.SetCellValue("Sheet1", fmt.Sprintf("%c%d", column, i+2), v.Field(j).Float())
			case "time.Time":
			 	str:= fmt.Sprintf("%s",v.Field(j).Interface())
				err = xlsx.SetCellValue("Sheet1", fmt.Sprintf("%c%d", column, i+2), str)
			}
		}
	}
	if err != nil {
		return err
	}
	xlsx.SetActiveSheet(index)
	err = xlsx.SaveAs(filePath)
	if err != nil {
		return err
	}
	return nil

}

func stringToTime(excelDate string) (time.Time, error) {
	baghdad, err := time.LoadLocation("Local")
	fmt.Println(err)
	return time.ParseInLocation("2006-01-02 15:04:05", excelDate, baghdad)

}

// todo 再看
func timeToString(t time.Time) (interface{}, error) {

	return nil, nil
}



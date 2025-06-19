package read

import (
	"fmt"
	"reflect"
	"strconv"
	"strings"
	"sync"

	"github.com/spf13/viper"
	"github.com/tealeg/xlsx/v3"
)

const COL_ATT_NO = "序号"
const COL_ATT_NAME = "姓名"
const COL_ATT_REST_DAY = "休假"
const COL_ATT_DATE = "日期"
const COL_ATT_TEMP_4 = "4小时临勤"
const COL_ATT_TEMP_8 = "8小时临勤"
const COL_ATT_TEMP_12 = "12小时临勤"

type FixedData struct {
	No      int
	Name    string
	RestDay []int
	Area    string
}

type TempSumData struct {
	No      int
	Date    int
	Temp_4  int
	Temp_8  int
	Temp_12 int
	Area    string
}

var fixedChan = make(chan FixedData)
var tempChan = make(chan TempSumData)
var finishChan = make(chan string)
var fixedMap = make(map[string][]FixedData)
var tempMap = make(map[string][]TempSumData)

//读取固定岗及临勤统计
func ReadData(path string) (map[string][]FixedData, map[string][]TempSumData, error) {
	file, err := xlsx.OpenFile(viper.GetString("input"))

	if err != nil {
		return fixedMap, tempMap, err
	}

	var wg sync.WaitGroup
	count := len(file.Sheets)
	wg.Add(count)

	go handleChan(&wg, count)

	for _, sheet := range file.Sheets {
		if sheet.Name == "大医院" || sheet.Name == "中医院" || sheet.Name == "慕田峪" {
			go readFixed(sheet)
		} else {
			go readTemp(sheet)
		}
	}

	wg.Wait()

	return fixedMap, tempMap, err
}

func readFixed(sheet *xlsx.Sheet) {

	headerMap := readHeader(sheet)

	if len(headerMap) == 0 {
		finishChan <- "read header failed "
		return
	}

	maxRow := sheet.MaxRow
	for i := 1; i < maxRow; i++ {
		row, err := sheet.Row(i)

		if err != nil {
			fmt.Printf("read fixed err named %s, rowNo %d, err: %v", sheet.Name, i, err)
			continue
		}

		fixed := FixedData{}
		visitRowFixed(row, &headerMap, &fixed)
		if len(fixed.Name) != 0 {
			fixed.Area = sheet.Name
			fixedChan <- fixed
		}
	}
	finishChan <- "read finish " + sheet.Name

}

func visitRowFixed(row *xlsx.Row, headerMap *map[int]string, fixed *FixedData) error {
	for i := range row.Sheet.MaxCol {
		str, err := row.GetCell(i).FormattedValue()
		if err != nil {
			return err
		}

		refType := reflect.TypeOf(*fixed)
		if refType.Kind() != reflect.Struct {
			panic("not struct")
		}

		if fieldObj, ok := refType.FieldByName((*headerMap)[i]); ok {
			if fieldObj.Type.Kind() == reflect.Int {
				val, _ := strconv.Atoi(str)
				reflect.ValueOf(fixed).Elem().FieldByName((*headerMap)[i]).SetInt(int64(val))
			}
			if fieldObj.Type.Kind() == reflect.String {
				reflect.ValueOf(fixed).Elem().FieldByName((*headerMap)[i]).SetString(str)
			}
			if fieldObj.Type.Kind() == reflect.Slice {

				arr := strings.Split(str, ",")
				
				arrInt := make([]int, len(arr))

				for j := range arr {
					valInt, e := strconv.Atoi(arr[j])
					if e != nil {
						continue
					}
					arrInt[j] = valInt
				}

				reflect.ValueOf(fixed).Elem().FieldByName((*headerMap)[i]).Set(reflect.ValueOf(arrInt))

			}
		}
	}
	return nil
}

func readTemp(sheet *xlsx.Sheet) {
	headerMap := readHeader(sheet)

	if len(headerMap) == 0 {
		finishChan <- "read header failed "
		return
	}

	maxRow := sheet.MaxRow
	for i := 1; i < maxRow; i++ {

		row, err := sheet.Row(i)

		if err != nil {
			fmt.Printf("read fixed err named %s, rowNo %d, err: %v", sheet.Name, i, err)
			continue
		}

		temp := TempSumData{}
		visitRowTemp(row, &headerMap, &temp)
		if temp.Date > 0 && temp.Date < 32 {
			temp.Area = sheet.Name
			tempChan <- temp
		}
	}
	finishChan <- "read finish " + sheet.Name
}
func visitRowTemp(row *xlsx.Row, headerMap *map[int]string, temp *TempSumData) error {
	for i := range row.Sheet.MaxCol {
		str, err := row.GetCell(i).FormattedValue()
		if err != nil {
			return err
		}

		val, _ := strconv.Atoi(str)
		refType := reflect.TypeOf(*temp)
		if refType.Kind() != reflect.Struct {
			panic("not struct")
		}

		if fieldObj, ok := refType.FieldByName((*headerMap)[i]); ok {
			if fieldObj.Type.Kind() == reflect.Int {
				reflect.ValueOf(temp).Elem().FieldByName((*headerMap)[i]).SetInt(int64(val))
			}
			if fieldObj.Type.Kind() == reflect.String {
				reflect.ValueOf(temp).Elem().FieldByName((*headerMap)[i]).SetString(str)
			}
		}
	}
	return nil
}

func readHeader(sheet *xlsx.Sheet) map[int]string {
	headerMap := make(map[int]string)
	row, err := sheet.Row(0)
	if err != nil {
		panic(err)
	}
	for i := range row.Sheet.MaxCol {
		str, err := row.GetCell(i).FormattedValue()
		if err != nil {
			panic(err)
		}
		switch str {
		case COL_ATT_NO:
			headerMap[i] = "No"
		case COL_ATT_NAME:
			headerMap[i] = "Name"
		case COL_ATT_REST_DAY:
			headerMap[i] = "RestDay"
		case COL_ATT_DATE:
			headerMap[i] = "Date"
		case COL_ATT_TEMP_4:
			headerMap[i] = "Temp_4"
		case COL_ATT_TEMP_8:
			headerMap[i] = "Temp_8"
		case COL_ATT_TEMP_12:
			headerMap[i] = "Temp_12"
		default:
			fmt.Printf("unknown col name %s", str)
		}
	}
	return headerMap
}

func handleChan(wg *sync.WaitGroup, count int) {
	for {
		select {
		case fixed := <-fixedChan:
			handleFixed(fixed)
		case temp := <-tempChan:
			handleTemp(temp)
		case s := <-finishChan:
			fmt.Print(s)
			wg.Done()
			count--

			if count == 0 {
				return
			}
		}
	}
}

func handleFixed(fixed FixedData) {
	if list, found := fixedMap[fixed.Area]; found {
		list = append(list, fixed)
		fixedMap[fixed.Area] = list
	} else {
		list = []FixedData{fixed}
		fixedMap[fixed.Area] = list
	}
}

func handleTemp(temp TempSumData) {
	if list, found := tempMap[temp.Area]; found {
		list = append(list, temp)
		tempMap[temp.Area] = list
	} else {
		list = []TempSumData{temp}
		tempMap[temp.Area] = list
	}
}

package read

import (
	"fmt"
	"log"
	"parking/src/types"
	"reflect"
	"slices"
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

// temp attendance
// const COL_TEMP_ATT_NO = "序号"
// const COL_TEMP_ATT_DATE = "日期"
const COL_TEMP_ATT_NAME_LIST = "姓名列表"
const COL_TEMP_ATT_COUNT = "数量"
const COL_TEMP_ALL_TYPE = "工时类型"

var fixedChan = make(chan types.FixedData)
var tempChan = make(chan types.TempSumData)
var tempAttChan = make(chan types.TempAttendance)
var finishChan = make(chan string)
var fixedMap = make(map[string][]types.FixedData)
var tempMap = make(map[string][]types.TempSumData)
var tempAttMap = make(map[string][]types.TempAttendance)

// temp staff map key:name, value:TempStaff
var tempStaffMap = make(map[string][]types.TempStaff)

// 读取固定岗及临勤统计
func ReadData(path string) (map[string][]types.FixedData, map[string][]types.TempSumData, map[string][]types.TempAttendance, error) {
	file, err := xlsx.OpenFile(viper.GetString("input"))

	if err != nil {
		return fixedMap, tempMap, tempAttMap, err
	}

	var wg sync.WaitGroup
	count := len(file.Sheets)
	wg.Add(count)

	go handleChan(&wg, count)
	fixedAreas := viper.GetStringSlice("fixed_area")
	tempAreas := viper.GetStringSlice("temp_area")
	tempAttendanceAreas := viper.GetStringSlice("temp_attendance_sheet")

	for _, sheet := range file.Sheets {
		if slices.Index(fixedAreas, sheet.Name) != -1 {
			go readFixed(sheet)
		} else if slices.Index(tempAreas, sheet.Name) != -1 {
			go readTemp(sheet)
		} else if slices.Index(tempAttendanceAreas, sheet.Name) != -1 {
			go readTempWithStaffName(sheet)
		} else {
			fmt.Printf("---ignore sheet named %s \n", sheet.Name)
		}
	}

	wg.Wait()

	return fixedMap, tempMap, tempAttMap, err
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

		fixed := types.FixedData{} 
		visitRowFixed(row, &headerMap, &fixed)
		if len(fixed.Name) != 0 {
			fixed.Area = sheet.Name
			fixedChan <- fixed
		}
	}
	finishChan <- "read finish " + sheet.Name

}

func visitRowFixed(row *xlsx.Row, headerMap *map[int]string, fixed *types.FixedData) error {
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

func readTempWithStaffName(sheet *xlsx.Sheet) {
	headerMap := readHeader(sheet)

	if len(headerMap) == 0 {
		finishChan <- "temp attendance read header failed "
		return
	}

	maxRow := sheet.MaxRow
	for i := 1; i < maxRow; i++ {

		row, err := sheet.Row(i)

		if err != nil {
			fmt.Printf("read temp attendance err named %s, rowNo %d, err: %v", sheet.Name, i, err)
			continue
		}

		tempAtt := types.TempAttendance{}
		visitRowTemp(row, &headerMap, &tempAtt)
		if tempAtt.Date > 0 && tempAtt.Date < 32 {
			if len(tempAtt.NameList) != 0 && strings.Contains(tempAtt.NameList, ";") {
				tempAtt.Names = strings.Split(tempAtt.NameList, ";")
			}
			tempAtt.Area = sheet.Name
			tempAttChan <- tempAtt
		}
	}
	finishChan <- "temp attendance read finish " + sheet.Name
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

		temp := types.TempSumData{}
		visitRowTemp(row, &headerMap, &temp)
		if temp.Date > 0 && temp.Date < 32 {
			temp.Area = sheet.Name
			tempChan <- temp
		}
	}
	finishChan <- "read finish " + sheet.Name
}

func visitRowTemp[T types.TempSumData | types.TempAttendance](row *xlsx.Row, headerMap *map[int]string, temp *T) error {
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
		// case COL_TEMP_ATT_NO:
		// 	headerMap[i] = "No"
		// case COL_TEMP_ATT_DATE:
		// 	headerMap[i] = "Date"
		case COL_TEMP_ATT_NAME_LIST:
			headerMap[i] = "NameList"
		case COL_TEMP_ATT_COUNT:
			headerMap[i] = "Count"
		case COL_TEMP_ALL_TYPE:
			headerMap[i] = "Kind"
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
		case tempAtt := <-tempAttChan:
			handleTempAtt(tempAtt)
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

func handleTempAtt(tempAtt types.TempAttendance) {
	if len(tempAtt.Names) != tempAtt.Count {
		log.Panicf("read tempAtt panic Area:%s, No:%d, NameList:%s, Count:%d", tempAtt.Area, tempAtt.No, tempAtt.NameList, tempAtt.Count)
	}

		if list, found := tempAttMap[tempAtt.Area]; found {
		list = append(list, tempAtt)
		tempAttMap[tempAtt.Area] = list
	} else {
		list = []types.TempAttendance{tempAtt}
		tempAttMap[tempAtt.Area] = list
	}
}

func handleFixed(fixed types.FixedData) {
	if list, found := fixedMap[fixed.Area]; found {
		list = append(list, fixed)
		fixedMap[fixed.Area] = list
	} else {
		list = []types.FixedData{fixed}
		fixedMap[fixed.Area] = list
	}
}

func handleTemp(temp types.TempSumData) {
	if list, found := tempMap[temp.Area]; found {
		list = append(list, temp)
		tempMap[temp.Area] = list
	} else {
		list = []types.TempSumData{temp}
		tempMap[temp.Area] = list
	}
}

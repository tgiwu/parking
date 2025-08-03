package transfer

import (
	"bufio"
	"fmt"
	"os"
	"parking/src/types"
	"slices"
	"strings"
	"time"

	"github.com/spf13/viper"
)

var namePool = []string{}

func initNamePool() {
	file, err := os.Open(viper.GetString("name_pool"))

	if err != nil {
		panic(err)
	}

	defer file.Close()

	scanner := bufio.NewScanner(file)

	for scanner.Scan() {
		text := scanner.Text()
		if len(text) == 0 {
			continue
		}

		names := strings.SplitSeq(text, ",")

		for name := range names {
			n := strings.TrimSpace(name)
			if len(n) > 0 {
				namePool = append(namePool, n)
			}
		}

	}
}

func FixedTransfer(fixed map[string][]types.FixedData) []types.PageData {

	initNamePool()

	var pdl []types.PageData

	for key, fdl := range fixed {
		pd := types.PageData{}
		for i, item := range fdl {

			if len(pd.Area) == 0 {
				pd = types.PageData{
					Area:   key,
					Title:  fmt.Sprintf("停车场人员%d年度%d月固定岗出勤统计表", viper.GetInt("year"), viper.GetInt("month")),
					PDType: types.PD_TYPE_SITE,
				}
			}

			pda := types.PDAttendance{
				Id:   i + 1,
				Name: item.Name,
				Data: make(map[int]int),
			}
			if len(item.RestDay) > 0 {

				for _, day := range item.RestDay {
					pda.Data[day] = 1
				}
			}
			pd.Attendances = append(pd.Attendances, pda)

			if len(pd.Attendances) == types.LINES_PER_PAGE {
				pdl = append(pdl, pd)
			}
		}

		if len(pd.Attendances) != 0 {

			lastIndex := pd.Attendances[len(pd.Attendances)-1].Id
			remain := types.LINES_PER_PAGE - len(pd.Attendances)
			for remain > 0 {
				lastIndex++
				pd.Attendances = append(pd.Attendances, types.PDAttendance{
					Id:   lastIndex,
					Name: "",
					Data: map[int]int{},
				})
				remain--
			}

			pdl = append(pdl, pd)
		}
	}

	return pdl
}

func TempTransfer(temp map[string][]types.TempSumData, nameIndex int) []types.PageData {
	var pdl []types.PageData

	for area, tl := range temp {
		var att8List []types.PDAttendance
		var att12List []types.PDAttendance
		var att4List []types.PDAttendance
		id := 0
		/*
			to avid id index error ,loop tsd for every single temp
			a1  8:3, 12:0
			a2  8:1, 12:3
			a3  8:5, 12:0

			if we process temp_8 and temp_12 in single loop, id will be create by lines, it may lead id of temp_12 less than temp_8
		*/
		for _, tsd := range tl {
			if tsd.Temp_8 > 0 {
				fmt.Printf("current %d \n", tsd.Temp_8)
				for i := 1; i <= tsd.Temp_8; i++ {
					var pda types.PDAttendance
					if len(att8List) < i {
						pda = types.PDAttendance{Data: make(map[int]int)}
						pda.Data[tsd.Date] = 8
						pda.Id = id + 1
						pda.Name = namePool[nameIndex]
						nameIndex++
						id++
						att8List = append(att8List, pda)
					} else {
						pda = att8List[i-1]
						pda.Data[tsd.Date] = 8
						att8List[i-1] = pda
					}
				}
			}

			if tsd.Temp_4 > 0 {
				for i := 1; i <= tsd.Temp_4; i++ {
					var pda types.PDAttendance
					if len(att4List) < i {
						pda = types.PDAttendance{Data: make(map[int]int)}
						pda.Data[tsd.Date] = 4
						pda.Id = id + 1
						pda.Name = namePool[nameIndex]
						nameIndex++
						id++
						att4List = append(att4List, pda)
					} else {
						pda = att4List[i-1]
						pda.Data[tsd.Date] = 4
						att4List[i-1] = pda
					}
				}
			}
		}

		for _, tsd := range tl {
			if tsd.Temp_12 > 0 {
				for i := 1; i <= tsd.Temp_12; i++ {
					var pda types.PDAttendance
					if len(att12List) < i {
						pda = types.PDAttendance{Data: make(map[int]int)}
						pda.Data[tsd.Date] = 12
						pda.Id = id + 1
						pda.Name = namePool[nameIndex]
						nameIndex++
						id++
						att12List = append(att12List, pda)
					} else {
						pda = att12List[i-1]
						pda.Data[tsd.Date] = 12
						att12List[i-1] = pda
					}
				}
			}
		}

		for _, tsd := range tl {
			if tsd.Temp_4 > 0 {
				for i := 1; i <= tsd.Temp_4; i++ {
					var pda types.PDAttendance
					if len(att4List) < i {
						pda = types.PDAttendance{Data: make(map[int]int)}
						pda.Data[tsd.Date] = 4
						pda.Id = id + 1
						pda.Name = namePool[nameIndex]
						nameIndex++
						id++
						att4List = append(att4List, pda)
					} else {
						pda = att4List[i-1]
						pda.Data[tsd.Date] = 4
						att4List[i-1] = pda
					}
				}
			}
		}
		pdl = slices.Concat(pdl, splitToPage(area, att8List, att12List, att4List))
	}

	return pdl
}

func splitToPage(area string, att8List, att12List, att4List []types.PDAttendance) []types.PageData {
	if len(att8List) == 0 && len(att12List) == 0 && len(att4List) == 0 {
		return []types.PageData{}
	}

	pdl := []types.PageData{}

	//处理8小时和12小时
	if len(att8List) != 0 || len(att12List) != 0 {

		list := slices.Concat(att8List, att12List)

		entry := types.PageData{Attendances: []types.PDAttendance{}}
		for _, item := range list {
			entry.Attendances = append(entry.Attendances, item)

			if len(entry.Attendances) == types.LINES_PER_PAGE {
				entry.Title = fmt.Sprintf("停车场人员%d年%d月临勤出勤统计表", viper.GetInt("year"), viper.GetInt("month"))
				entry.PDType = types.PD_TYPE_TEMP
				entry.Area = area
				pdl = append(pdl, entry)
				entry = types.PageData{Attendances: []types.PDAttendance{}}
			}
		}

		//不满整页补充满
		if len(entry.Attendances) != 0 {
			entry.Title = fmt.Sprintf("停车场人员%d年%d月临勤出勤统计表", viper.GetInt("year"), viper.GetInt("month"))
			entry.PDType = types.PD_TYPE_TEMP
			entry.Area = area
			lastIndex := entry.Attendances[len(entry.Attendances)-1].Id
			remain := types.LINES_PER_PAGE - len(entry.Attendances)

			for remain > 0 {
				lastIndex++
				entry.Attendances = append(entry.Attendances, types.PDAttendance{Id: lastIndex})
				remain--

			}
			pdl = append(pdl, entry)
		}
	}

	//处理4小时
	if len(att4List) != 0 {

		entry := types.PageData{Attendances: []types.PDAttendance{}}
		for _, item := range att4List {
			entry.Attendances = append(entry.Attendances, item)

			if len(entry.Attendances) == types.LINES_PER_PAGE {
				entry.Title = fmt.Sprintf("停车场人员%d年%d月临勤出勤统计表", viper.GetInt("year"), viper.GetInt("month"))
				entry.PDType = types.PD_TYPE_NIGHT
				entry.Area = area
				pdl = append(pdl, entry)
				entry = types.PageData{Attendances: []types.PDAttendance{}}
			}
		}

		if len(entry.Attendances) != 0 {
			entry.Title = fmt.Sprintf("停车场人员%d年%d月临勤出勤统计表", viper.GetInt("year"), viper.GetInt("month"))
			entry.PDType = types.PD_TYPE_NIGHT
			entry.Area = area
			lastIndex := entry.Attendances[len(entry.Attendances)-1].Id
			remain := types.LINES_PER_PAGE - len(entry.Attendances)

			for remain > 0 {
				lastIndex++
				entry.Attendances = append(entry.Attendances, types.PDAttendance{Id: lastIndex})
				remain--

			}

			pdl = append(pdl, entry)
		}
	}

	return pdl
}

func CreateBillData(fixedMap map[string][]types.FixedData, tempMap map[string][]types.TempSumData) types.BillData {

	contractStart := viper.GetString("contract_start")
	contractEnd := viper.GetString("contract_end")
	contractStartTime, err := time.ParseInLocation(time.DateOnly, contractStart, time.Local)
	if err != nil {
		panic(err)
	}
	contractEndTime, _ := time.ParseInLocation(time.DateOnly, contractEnd, time.Local)

	var fixedDataMap = make(map[string]int)
	var temp8DataMap = make(map[string]int)
	var temp12DataMap = make(map[string]int)
	var temp4DataMap = make(map[string]int)
	for key, list := range fixedMap {
		if len(list) > 0 {
			fixedDataMap[key] = len(list)
		}
	}

	for area, list := range tempMap {
		for _, tsd := range list {
			if tsd.Temp_8 > 0 {
				if sum, found := temp8DataMap[area]; found {
					temp8DataMap[area] = sum + tsd.Temp_8
				} else {
					temp8DataMap[area] = tsd.Temp_8
				}
			}

			if tsd.Temp_12 > 0 {
				if sum, found := temp12DataMap[area]; found {
					temp12DataMap[area] = sum + tsd.Temp_12
				} else {
					temp12DataMap[area] = tsd.Temp_12
				}
			}

			if tsd.Temp_4 > 0 {
				if sum, found := temp4DataMap[area]; found {
					temp4DataMap[area] = sum + tsd.Temp_4
				} else {
					temp4DataMap[area] = tsd.Temp_4
				}
			}
		}
	}

	var billData = types.BillData{
		Year:               viper.GetInt("year"),
		Month:              viper.GetInt("month"),
		ContractStartYear:  contractStartTime.Year(),
		ContractStartMonth: int(contractStartTime.Month()),
		ContractStartDay:   contractStartTime.Day(),
		ContractEndYear:    contractEndTime.Year(),
		ContractEndMonth:   int(contractEndTime.Month()),
		ContractEndDay:     contractEndTime.Day(),
		FixedBillData:      fixedDataMap,
		TempBill8Data:      temp8DataMap,
		TempBill12Data:     temp12DataMap,
		TempBill4Data:      temp4DataMap,
	}
	fmt.Printf("bill : %+v\n", billData)

	return billData
}

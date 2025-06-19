package transfer

import (
	"bufio"
	"parking/src/read"
	"parking/src/types"
	"fmt"
	"os"
	"slices"
	"strings"

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

func FixedTransfer(fixed map[string][]read.FixedData) []types.PageData {

	initNamePool()

	var pdl []types.PageData

	for key, fdl := range fixed {
		pd := types.PageData{}
		for i, item := range fdl {

			if len(pd.Area) == 0 {
				pd = types.PageData{
					Area:   key,
					Title:  fmt.Sprintf("停车场人员%d年度%d月固定岗出勤统计表", viper.GetInt("year"), viper.GetInt("month")),
					IsTemp: false,
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

func TempTransfer(temp map[string][]read.TempSumData) []types.PageData {
	var pdl []types.PageData

	for area, tl := range temp {
		var att8List []types.PDAttendance
		var att12List []types.PDAttendance
		var att4List []types.PDAttendance

		nameIndex := 0
		for _, tsd := range tl {
			if tsd.Temp_8 > 0 {
				for i := 1; i <= tsd.Temp_8; i++ {
					var pda types.PDAttendance
					if len(att8List) <= i {
						pda = types.PDAttendance{Data: make(map[int]int)}
						pda.Data[tsd.Date] = 8
						pda.Id = nameIndex + 1
						pda.Name = namePool[nameIndex]
						nameIndex++
						att8List = append(att8List, pda)
					} else {
						pda = att8List[i]
						pda.Data[tsd.Date] = 8
						att8List[i] = pda
					}
				}
			}

			if tsd.Temp_12 > 0 {
				for i := 1; i <= tsd.Temp_12; i++ {
					var pda types.PDAttendance
					if len(att12List) <= i {
						pda = types.PDAttendance{Data: make(map[int]int)}
						pda.Data[tsd.Date] = 12
						pda.Id = nameIndex + 1
						pda.Name = namePool[nameIndex]
						nameIndex++
						att12List = append(att12List, pda)
					} else {
						pda = att12List[i]
						pda.Data[tsd.Date] = 12
						att12List[i] = pda
					}
				}
			}

			if tsd.Temp_4 > 0 {
				for i := 1; i <= tsd.Temp_4; i++ {
					var pda types.PDAttendance
					if len(att4List) <= i {
						pda = types.PDAttendance{Data: make(map[int]int)}
						pda.Data[tsd.Date] = 4
						pda.Id = nameIndex + 1
						pda.Name = namePool[nameIndex]
						nameIndex++
						att4List = append(att4List, pda)
					} else {
						pda = att4List[i]
						pda.Data[tsd.Date] = 4
						att4List[i] = pda
					}
				}
			}
		}
		pdl = slices.Concat(pdl, splitToPage(area, att8List, att12List, att4List))
	}

	return pdl
}

func splitToPage(area string, att8List, att12List, att4List []types.PDAttendance) []types.PageData {
	if len(att8List) == 0 && len(att12List) == 0 {
		return []types.PageData{}
	}

	list := slices.Concat(att8List, att12List)

	pdl := []types.PageData{}
	entry := types.PageData{Attendances: []types.PDAttendance{}}
	for _, item := range list {
		entry.Attendances = append(entry.Attendances, item)

		if len(entry.Attendances) == types.LINES_PER_PAGE {
			entry.Title = fmt.Sprintf("停车场人员%d年%d月临勤出勤统计表", viper.GetInt("year"), viper.GetInt("month"))
			entry.IsTemp = true
			entry.Area = area
			pdl = append(pdl, entry)
			entry = types.PageData{Attendances: []types.PDAttendance{}}
		}
	}

	if len(entry.Attendances) != 0 {
		entry.Title = fmt.Sprintf("停车场人员%d年%d月临勤出勤统计表", viper.GetInt("year"), viper.GetInt("month"))
		entry.IsTemp = true
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

	entry = types.PageData{Attendances: []types.PDAttendance{}}
	for _, item := range att4List {
		entry.Attendances = append(entry.Attendances, item)

		if len(entry.Attendances) == types.LINES_PER_PAGE {
			entry.Title = fmt.Sprintf("停车场人员%d年%d月临勤出勤统计表", viper.GetInt("year"), viper.GetInt("month"))
			entry.IsTemp = true
			entry.Area = area
			pdl = append(pdl, entry)
			entry = types.PageData{Attendances: []types.PDAttendance{}}
		}
	}

	if len(entry.Attendances) != 0 {
		entry.Title = fmt.Sprintf("停车场人员%d年%d月临勤出勤统计表", viper.GetInt("year"), viper.GetInt("month"))
		entry.IsTemp = true
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

	return pdl
}

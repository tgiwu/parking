package word

import (
	"errors"
	"fmt"
	"log"
	"parking/src/types"
	"path"
	"regexp"
	"slices"
	"strconv"
	"strings"
	"time"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
	"github.com/spf13/viper"
)

// var TABLE_HEADER = []string{"序", "姓名", "出勤", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "签字"}
var TABLE_HEADER = []string{"序", "姓名", "出勤"}

const DOC_TABLE_PAGE_TITLE = "DOC_TABLE_PAGE_TITLE"
const DOC_TABLE_PAGE_SUBTITLE = "DOC_TABLE_PAGE_SUBTITLE"
const TABLE_HEADER_CHAR = "TABLE_HEADER_CHAR"
const TABLE_HEADER_NUMBER = "TABLE_HEADER_NUMBER"
const TABLE_CONTENT_NO = "TABLE_CONTENT_NO"
const TABLE_CONTENT_NAME = "TABLE_CONTENT_NAME"
const TABLE_CONTENT_ATT_SUM = "TABLE_CONTENT_ATT_SUM"
const TABLE_CONTENT_NORMAL = "TABLE_CONTENT_NORMAL"
const TABLE_CONTENT_NORMAL_8 = "TABLE_CONTENT_NORMAL_8"
const TABLE_CONTENT_NORMAL_12 = "TABLE_CONTENT_NORMAL_12"
const TABLE_CONTENT_SIGN = "TABLE_CONTENT_SIGN"
const DOC_PAGE_SIGN = "DOC_PAGE_SIGN"

const DOC_BILL_TITLE = "DOC_BILL_TITLE"
const DOC_BILL_SUBTITLE = "DOC_BILL_SUBTITLE"
const DOC_BILL_CONTENT = "DOC_BILL_CONTENT"
const DOC_BILL_SITE_TITLE = "DOC_BILL_SITE_TITLE"
const DOC_BILL_SIGN = "DOC_BILL_SIGN"

const (
	TABLE_COL_INDEX_NO = iota
	TABLE_COL_INDEX_NAME
	TABLE_COL_INDEX_ATT_SUM
)

type TableType int
const (
	TABLE_TYPE_CHECKIN_FIXED = iota
	TABLE_TYPE_CHECKIN_TEMP
)

var TABLE_COL_INDEX_SIGN int

// 费用合计
// var billsum = 0

// 编写单页
func createFixedDocxSingle(doc *document.Document, pd *types.PageData) {

	fixedTitle(doc, pd)
	fixedSubtitle(doc, pd)

	table, _ := doc.AddTable(&document.TableConfig{
		Rows:  29,
		Cols:  TABLE_COL_INDEX_SIGN + 1,
		Width: 15500,
	})

	fixedHeader(table)
	FixedData(table, pd)

	signerArea(doc)

}

func CreateTempDocx(list []types.PageData) {
	document.SetGlobalLevel(document.LogLevelError)

	doc := document.New()
	setUpStyle(doc)
	doc.SetPageSize(document.PageSizeA4)
	doc.SetPageOrientation(document.OrientationLandscape)
	doc.SetPageMargins(5.56, 15, 5.56, 15)

	days := calcLastDayInMonth(viper.GetInt("year"), viper.GetInt("month"))

	for i := 1; i <= days; i++ {
		TABLE_HEADER = append(TABLE_HEADER, fmt.Sprintf("%d", i))
	}
	TABLE_HEADER = append(TABLE_HEADER, "签字")

	TABLE_COL_INDEX_SIGN = days + 3
}

// 创建固定岗考勤表
func CreateFixedDocx(list []types.PageData) {
	document.SetGlobalLevel(document.LogLevelError)

	doc := document.New()
	setUpStyle(doc)
	doc.SetPageSize(document.PageSizeA4)
	doc.SetPageOrientation(document.OrientationLandscape)
	doc.SetPageMargins(5.56, 15, 5.56, 15)

	days := calcLastDayInMonth(viper.GetInt("year"), viper.GetInt("month"))

	for i := 1; i <= days; i++ {
		TABLE_HEADER = append(TABLE_HEADER, fmt.Sprintf("%d", i))
	}
	TABLE_HEADER = append(TABLE_HEADER, "签字")

	TABLE_COL_INDEX_SIGN = days + 3

	for _, pd := range list {
		createFixedDocxSingle(doc, &pd)
	}

	fileName := fmt.Sprintf("%d年%d月停车场考勤.docx", viper.GetInt32("year"), viper.GetInt32("month"))
	doc.Save(path.Join(viper.GetString("output"), fileName))

}

func CreateBillDocMulti(billDataMap *map[int]types.BillData) {
	for _, data := range *billDataMap {
		CreateBillDoc(&data)
	}
}

func CreateBillDoc(billData *types.BillData) {
	doc := document.New()
	setUpBillStyle(doc)

	billTitle(doc, billData)
	billDescription(doc, billData)

	tempData(doc, billData)
	fixedData(doc, billData)
	sumText(doc, billData)
	signture(doc, billData)

	var billTypeStr string
	typesString := viper.GetStringSlice("type_strings")
	if billData.BillDataType > len(typesString) {
		billTypeStr = ""
	} else {
		billTypeStr = typesString[billData.BillDataType]
	}

	fileName := fmt.Sprintf("%d年%d月停车场结算单%s.docx", viper.GetInt32("year"), viper.GetInt32("month"), billTypeStr)
	doc.Save(path.Join(viper.GetString("output"), fileName))
}

func tableConfig(tt TableType) types.TableConfig{
	tableConfig := types.TableConfig{}
	switch tt {
	case TABLE_TYPE_CHECKIN_FIXED:
	case TABLE_TYPE_CHECKIN_TEMP:
	default:
		
	}
	return tableConfig
}

func DigitalToChar(money int) string {
	var sliceUnit = []string{"仟", "佰", "拾", "亿", "仟", "佰", "拾", "万", "仟", "佰", "拾", "元", "角", "分"}
	var upperDigitUnit = map[string]string{
		"0": "零",
		"1": "壹",
		"2": "贰",
		"3": "叁",
		"4": "肆",
		"5": "伍",
		"6": "陆",
		"7": "柒",
		"8": "捌",
		"9": "玖",
	}

	strMoney := strconv.Itoa(money * 100)

	if len(strMoney) > len(sliceUnit) {
		panic(errors.New("too big"))
	}

	units := sliceUnit[len(sliceUnit)-len(strMoney):]
	amount := make([]string, len(units))
	for idx, num := range strMoney {
		amount[idx] = fmt.Sprintf("%s%s", upperDigitUnit[string(num)], units[idx])
	}

	str := strings.Join(amount, "")
	reg, _ := regexp.Compile(`零角零分$`)
	str = reg.ReplaceAllString(str, "整")

	reg, _ = regexp.Compile(`零角`)
	str = reg.ReplaceAllString(str, "零")

	reg, _ = regexp.Compile(`零[仟佰拾]`)
	str = reg.ReplaceAllString(str, "零")

	reg, _ = regexp.Compile(`零{2,}`)
	str = reg.ReplaceAllString(str, "零")

	reg, _ = regexp.Compile(`零亿`)
	str = reg.ReplaceAllString(str, "亿")

	reg, _ = regexp.Compile(`零万`)
	str = reg.ReplaceAllString(str, "万")

	reg, _ = regexp.Compile(`零元`)
	str = reg.ReplaceAllString(str, "元")

	reg, _ = regexp.Compile(`零*元`)
	str = reg.ReplaceAllString(str, "元")

	reg, _ = regexp.Compile(`亿零{0,3}万`)
	str = reg.ReplaceAllString(str, "^元")

	reg, _ = regexp.Compile(`零元`)
	str = reg.ReplaceAllString(str, "零")

	return str
}

// calculate late day in month
func calcLastDayInMonth(year int, month int) int {

	loc, _ := time.LoadLocation("Local")

	date := time.Date(year, time.Month(month), 1, 0, 0, 0, 0, loc).AddDate(0, 1, -1)
	day := date.Day()
	return day
}

// 固定岗和临勤有重叠区域，为区别临勤岗名称后加L，写入文档时去掉
func formatAreaIfNeed(area string) string {
	fixedAreas := viper.GetStringSlice("fixed_area")
	if slices.Contains(fixedAreas, area) {
		return area
	}

	targetAreas := viper.GetStringSlice("target_area")
	tempAreas := viper.GetStringSlice("temp_area")
	index := slices.Index(tempAreas, area)
	if index == -1 {
		log.Panicf("known area %s", area)
	}
	return targetAreas[index%len(targetAreas)]

}

func CreateCheckinDocx(tempAttendanceMap *map[string][]types.TempAttendance) {
	for area, tempAttendanceList := range *tempAttendanceMap {
		CreateCheckinDocxSingle(&tempAttendanceList, area)
	}
}

func CreateCheckinDocxSingle(tempAttendanceList *[]types.TempAttendance, area string) {
	doc := document.New()
	setUpCheckinDocStyle(doc)

	doc.SetPageSize(document.PageSizeA4)
	doc.SetPageOrientation(document.OrientationLandscape)
	signPages(doc, tempAttendanceList, area)

	infoPages(doc, tempAttendanceList, area)

	checkInPages(doc, tempAttendanceList, area)

	doc.Save(path.Join(viper.GetString("output"), fmt.Sprintf("临勤（外雇）人员增补签到表_%s.docx", area)))

}

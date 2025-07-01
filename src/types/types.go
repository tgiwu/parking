package types

const COL_ATT_NO = "序号"
const COL_ATT_NAME = "姓名"
const COL_ATT_REST_DAY = "休假"
const COL_ATT_DATE = "日期"
const COL_ATT_TEMP_4 = "4小时临勤"
const COL_ATT_TEMP_8 = "8小时临勤"
const COL_ATT_TEMP_12 = "12小时临勤"

const TEMP_8_DOC_TXT = "(8小时)"
const TEMP_12_DOC_TXT = "(12小时)"
const TEMP_NIGHT = "夜航"

const PD_TYPE_SITE = 0
const PD_TYPE_TEMP = 1
const PD_TYPE_NIGHT = 2

const LINES_PER_PAGE = 27

type PageData struct {
	Title       string
	Area        string
	PDType      int
	Attendances []PDAttendance
}

type PDAttendance struct {
	Id   int
	Name string
	Data map[int]int //固定岗标记休假，临勤标记工时
}

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

type BillData struct {
	Year  int
	Month int

	ContractStartYear  int
	ContractStartMonth int
	ContractStartDay   int

	ContractEndYear  int
	ContractEndMonth int
	ContractEndDay   int

	TempBill8Data  map[string]int
	TempBill12Data map[string]int
	TempBill4Data  map[string]int
	FixedBillData  map[string]int
}

type ParagraphSimple struct {
	Text  string
	Style string
}

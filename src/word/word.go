package word

import (
	"errors"
	"fmt"
	"math/rand"
	"parking/src/types"
	"path"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
	"github.com/ZeroHawkeye/wordZero/pkg/style"
	"github.com/spf13/viper"
)

var TABLE_HEADER = []string{"序", "姓名", "出勤", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "签字"}

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

const TABLE_COL_INDEX_NO = 0
const TABLE_COL_INDEX_NAME = 1
const TABLE_COL_INDEX_ATT_SUM = 2
const TABLE_COL_INDEX_SIGN = 34

// 费用合计
var billsum = 0

// 编写单页
func writeDocxSingle(doc *document.Document, pd *types.PageData) {

	title(doc, pd)
	subtitle(doc, pd)

	table := doc.AddTable(&document.TableConfig{
		Rows:  29,
		Cols:  35,
		Width: 15500,
	})

	header(table)
	data(table, pd)

	signerArea(doc)

}

// 文档创建方法
func CreateDocx(list []types.PageData) {
	document.SetGlobalLevel(document.LogLevelError)

	doc := document.New()
	setUpStyle(doc)
	doc.SetPageSize(document.PageSizeA4)
	doc.SetPageOrientation(document.OrientationLandscape)
	doc.SetPageMargins(5.56, 15, 5.56, 15)

	for _, pd := range list {
		writeDocxSingle(doc, &pd)
	}

	doc.Save(path.Join(viper.GetString("output"), viper.GetString("file")))

}

// 预制风格
func setUpStyle(doc *document.Document) {
	styleManager := doc.GetStyleManager()
	quickAPI := style.NewQuickStyleAPI(styleManager)

	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   DOC_TABLE_PAGE_TITLE,
		Name: DOC_TABLE_PAGE_TITLE,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "center",
			SpaceBefore: 0, // 段前间距（缇）
			SpaceAfter:  0, // 段后间距（缇）
			LineSpacing: 0, // 行间距（缇）
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "(中文正文)",
			FontSize:  14,
			FontColor: "000000",
			Bold:      false,
		},
	})

	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   DOC_TABLE_PAGE_SUBTITLE,
		Name: DOC_TABLE_PAGE_SUBTITLE,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "left",
			SpaceBefore: 0, // 段前间距（缇）
			SpaceAfter:  0, // 段后间距（缇）
			LineSpacing: 0, // 行间距（缇）
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "宋体(中文正文)",
			FontSize:  14,
			FontColor: "000000",
			Bold:      false,
		},
	})

	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   DOC_PAGE_SIGN,
		Name: DOC_PAGE_SIGN,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "left",
			SpaceBefore: 0, // 段前间距（缇）
			SpaceAfter:  0, // 段后间距（缇）
			LineSpacing: 0, // 行间距（缇）
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "宋体(中文正文)",
			FontSize:  14,
			FontColor: "000000",
			Bold:      false,
		},
	})

}

// 表格标题
func title(doc *document.Document, pd *types.PageData) {
	var t string
	if pd.PDType == types.PD_TYPE_NIGHT || pd.PDType == types.PD_TYPE_TEMP {
		t = "临勤"
	} else {
		t = "固定岗"
	}
	titlePara := doc.AddParagraph(fmt.Sprintf("停车场人员%d年度%d月%s出勤统计表", viper.GetInt("year"), viper.GetInt("month"), t))
	titlePara.SetStyle(DOC_TABLE_PAGE_TITLE)
}

// 副标题和区域
func subtitle(doc *document.Document, pd *types.PageData) {

	subtitlePara := doc.AddParagraph(fmt.Sprintf("用工单位名称：%s                                   负责区域：%s", viper.GetString("corporation_name"), formatAreaIfNeed(pd.Area)))
	subtitlePara.SetStyle(DOC_TABLE_PAGE_SUBTITLE)
}

// 设置单元格宽度
func cellWidth(table *document.Table, row, col int) {

	switch col {
	case TABLE_COL_INDEX_NO:
		//序
		cell, err := table.GetCell(row, col)
		if err == nil {
			cw := cell.Properties.TableCellW
			cw.W = "250"
		}
	case TABLE_COL_INDEX_NAME:
		//姓名
		cell, err := table.GetCell(row, col)
		if err == nil {
			cw := cell.Properties.TableCellW
			cw.W = "1300"
		}
	case TABLE_COL_INDEX_ATT_SUM:
		//出勤
		cell, err := table.GetCell(row, col)
		if err == nil {
			cw := cell.Properties.TableCellW
			cw.W = "900"
		}
	case TABLE_COL_INDEX_SIGN:
		//签字
		cell, err := table.GetCell(row, col)
		if err == nil {
			cw := cell.Properties.TableCellW
			cw.W = "1300"
		}
	default:
		//日期
		cell, err := table.GetCell(row, col)
		if err == nil {
			cw := cell.Properties.TableCellW
			cw.W = "100"
		}
	}

}

// 表头
func header(table *document.Table) {
	for i := range table.GetColumnCount() {
		cellWidth(table, 0, i)
		var textFormate document.TextFormat

		if i == TABLE_COL_INDEX_NO || i == TABLE_COL_INDEX_NAME ||
			i == TABLE_COL_INDEX_ATT_SUM || i == TABLE_COL_INDEX_SIGN {
			textFormate = document.TextFormat{
				Bold:       false,
				Italic:     false,
				FontSize:   12,
				FontFamily: "宋体 (中文正文)",
			}
		} else {
			textFormate = document.TextFormat{
				Bold:       false,
				Italic:     false,
				FontSize:   7,
				FontFamily: "宋体 (中文正文)",
			}
		}

		cellformat, _ := table.GetCellFormat(0, i)
		cellformat.HorizontalAlign = document.CellAlignCenter
		cellformat.VerticalAlign = document.CellVAlignCenter
		cellformat.TextFormat = &textFormate
		table.SetCellFormat(0, i, cellformat)
		table.SetCellText(0, i, TABLE_HEADER[i])
	}
}

// 填充数据
func data(table *document.Table, pd *types.PageData) {
	days := calcLastDayInMonth(viper.GetInt("year"), viper.GetInt("month"))
	for row, pda := range pd.Attendances {
		table.SetRowHeight(row+1, &document.RowHeightConfig{Height: 16, Rule: document.RowHeightExact})
	loop:
		for col := range table.GetColumnCount() - 1 {
			cellWidth(table, row, col)
			if col-2 > days {
				break
			}
			var textFormate document.TextFormat

			switch col {
			case TABLE_COL_INDEX_NO:
				textFormate = document.TextFormat{
					Bold:       false,
					Italic:     false,
					FontSize:   7,
					FontFamily: "Calibri (西文正文)",
				}
			case TABLE_COL_INDEX_NAME:
				textFormate = document.TextFormat{
					Bold:       false,
					Italic:     false,
					FontSize:   12,
					FontFamily: "宋体 (中文正文)",
				}
			case TABLE_COL_INDEX_ATT_SUM:
				textFormate = document.TextFormat{
					Bold:       false,
					Italic:     false,
					FontSize:   9,
					FontFamily: "Calibri (西文正文)",
				}
			case TABLE_COL_INDEX_SIGN:
				textFormate = document.TextFormat{
					Bold:       false,
					Italic:     false,
					FontSize:   9,
					FontFamily: "宋体 (中文正文)",
				}
			default:
				textFormate = document.TextFormat{
					Bold:       false,
					Italic:     false,
					FontSize:   8,
					FontFamily: "Arial",
				}
			}

			cellformat, _ := table.GetCellFormat(row+1, col)
			cellformat.HorizontalAlign = document.CellAlignCenter
			cellformat.VerticalAlign = document.CellVAlignCenter
			cellformat.TextFormat = &textFormate
			table.SetCellFormat(row+1, col, cellformat)
			switch col {
			case TABLE_COL_INDEX_NO:
				table.SetCellText(row+1, col, fmt.Sprint(pda.Id))
			case TABLE_COL_INDEX_NAME:
				if len(pda.Name) == 0 {
					break loop
				}
				table.SetCellText(row+1, col, pda.Name)
			case TABLE_COL_INDEX_ATT_SUM:
				if pd.PDType == types.PD_TYPE_NIGHT || pd.PDType == types.PD_TYPE_TEMP {
					table.SetCellText(row+1, col, fmt.Sprint(len(pda.Data)))
				} else {
					table.SetCellText(row+1, col, fmt.Sprint(calcLastDayInMonth(viper.GetInt("year"), viper.GetInt("month"))-len(pda.Data)))
				}
			default:
				if pd.PDType == types.PD_TYPE_NIGHT || pd.PDType == types.PD_TYPE_TEMP {
					if v, found := pda.Data[col-2]; found {
						table.SetCellText(row+1, col, fmt.Sprint(v))
					} else {
						table.SetCellText(row+1, col, "")
					}
				} else {
					if v, found := pda.Data[col-2]; found {
						if v == 1 {
							table.SetCellText(row+1, col, "")
						} else {
							table.SetCellText(row+1, col, "√")
						}
					} else {
						table.SetCellText(row+1, col, "√")
					}
				}
			}
		}
	}
	//备注区域
	table.MergeCellsHorizontal(28, 1, 34)
	table.SetCellText(28, 1, "备注：")
}

// 签字区域
func signerArea(doc *document.Document) {
	subtitlePara := doc.AddParagraph("负责人签字：                      区域负责人签字：                   部门负责人签字：")
	subtitlePara.SetStyle(DOC_TABLE_PAGE_SUBTITLE)
}

// 计算目标月份天数
func calcLastDayInMonth(year int, month int) int {

	loc, _ := time.LoadLocation("Local")

	date := time.Date(year, time.Month(month), 1, 0, 0, 0, 0, loc).AddDate(0, 1, -1)
	day := date.Day()
	return day
}

func CreateBillDoc(billData *types.BillData) {
	doc := document.New()
	setUpBillStyle(doc)

	billTitle(doc, billData)
	billDescription(doc, billData)

	tempData(doc, billData)
	fixedData(doc, billData)
	sumText(doc)
	signture(doc)

	doc.Save(path.Join(viper.GetString("output"), "bill.docx"))
}

func setUpBillStyle(doc *document.Document) {
	styleManager := doc.GetStyleManager()
	quickAPI := style.NewQuickStyleAPI(styleManager)
	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   DOC_BILL_TITLE,
		Name: DOC_BILL_TITLE,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "center",
			SpaceBefore: 0, // 段前间距（缇）
			SpaceAfter:  0, // 段后间距（缇）
			LineSpacing: 0, // 行间距（缇）
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "微软雅黑",
			FontSize:  22,
			FontColor: "000000",
			Bold:      false,
		},
	})
	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   DOC_BILL_SUBTITLE,
		Name: DOC_BILL_SUBTITLE,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "center",
			SpaceBefore: 0, // 段前间距（缇）
			SpaceAfter:  0, // 段后间距（缇）
			LineSpacing: 0, // 行间距（缇）
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "仿宋",
			FontSize:  18,
			FontColor: "000000",
			Bold:      false,
		},
	})
	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   DOC_BILL_CONTENT,
		Name: DOC_BILL_CONTENT,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "left",
			SpaceBefore: 0, // 段前间距（缇）
			SpaceAfter:  0, // 段后间距（缇）
			LineSpacing: 0, // 行间距（缇）
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "仿宋",
			FontSize:  16,
			FontColor: "000000",
			Bold:      false,
		},
	})
	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   DOC_BILL_SITE_TITLE,
		Name: DOC_BILL_SITE_TITLE,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "left",
			SpaceBefore: 0, // 段前间距（缇）
			SpaceAfter:  0, // 段后间距（缇）
			LineSpacing: 0, // 行间距（缇）
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "仿宋",
			FontSize:  16,
			FontColor: "000000",
			Bold:      true,
		},
	})
	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   DOC_BILL_SIGN,
		Name: DOC_BILL_SIGN,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "left",
			SpaceBefore: 0, // 段前间距（缇）
			SpaceAfter:  0, // 段后间距（缇）
			LineSpacing: 0, // 行间距（缇）
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "仿宋",
			FontSize:  16,
			FontColor: "000000",
			Bold:      false,
		},
	})
}

func billTitle(doc *document.Document, billData *types.BillData) {
	paraTitle := doc.AddParagraph("停车场服务项目费用结算单")
	paraTitle.SetStyle(DOC_BILL_TITLE)
	paraSubtitle := doc.AddParagraph(fmt.Sprintf("（%d年%d月）", billData.Year, billData.Month))
	paraSubtitle.SetStyle(DOC_BILL_SUBTITLE)
}

func billDescription(doc *document.Document, billData *types.BillData) {
	paraCorporation := doc.AddParagraph(viper.GetString("corporation_name") + ":")
	paraCorporation.SetStyle(DOC_BILL_CONTENT)
	lastDay := calcLastDayInMonth(billData.Year, billData.Month)
	txt := fmt.Sprintf("    自%d年%d月%d日起，确认由贵公司为%s提供停车场服务工作，服务期限为%d.%d.%d-%d.%d.%d。贵公司已按合同要求完成%d年%d月各项工作，%d月份服务时间自%d月%d日至%d月%d日，共计%d天，费用明细如下：",
		billData.ContractStartYear, billData.ContractStartMonth, billData.ContractStartDay,
		viper.GetString("first_party"),
		billData.ContractStartYear, billData.ContractStartMonth, billData.ContractStartDay,
		billData.ContractEndYear, billData.ContractEndMonth, billData.ContractEndDay,
		billData.Year, billData.Month, billData.Month, billData.Month, 1, billData.Month, lastDay, lastDay)
	paraDescription := doc.AddParagraph(txt)
	paraDescription.SetStyle(DOC_BILL_CONTENT)
}

func tempData(doc *document.Document, billData *types.BillData) {
	if len(billData.TempBill12Data) == 0 &&
		len(billData.TempBill8Data) == 0 &&
		len(billData.TempBill4Data) == 0 {
		return
	}
	sTitle := fmt.Sprintf("    临勤岗：（8小时%d元/人/天；12小时%d元/人/天）", viper.GetInt("temp_8_day"), viper.GetInt("temp_12_day"))
	paraTitle := doc.AddParagraph(sTitle)
	paraTitle.SetStyle(DOC_BILL_SITE_TITLE)

	paraList := []types.ParagraphSimple{}
	tempAreas := viper.GetStringSlice("temp_area")

	for _, area := range tempAreas {
		temp4Count := 0
		if v, found := billData.TempBill4Data[area]; found {
			temp4Count = v
		}

		if v, found := billData.TempBill8Data[area]; found && (v > 0 || temp4Count > 0) {
			count := v + temp4Count/2
			if temp4Count%2 != 0 && rand.Int()%2 != 0 {
				count++
			}
			if count > 0 {
				sum := count * viper.GetInt("temp_8_day")
				paraList = append(paraList, types.ParagraphSimple{
					Text: fmt.Sprintf("    %s（8小时）：%d人*%d元/人/天=%d元;",
						formatAreaIfNeed(area), count, viper.GetInt("temp_8_day"), sum),
					Style: DOC_BILL_CONTENT,
				})
				billsum += sum
			}

		}

		if v, found := billData.TempBill12Data[area]; found && v > 0 {
			sum := v * viper.GetInt("temp_12_day")
			paraList = append(paraList, types.ParagraphSimple{
				Text: fmt.Sprintf("    %s（12小时）：%d人*%d元/人/天=%d元;",
					formatAreaIfNeed(area), v, viper.GetInt("temp_12_day"), sum),
				Style: DOC_BILL_CONTENT,
			})
			billsum += sum
		}
	}

	if len(paraList) > 0 {
		for _, line := range paraList {
			para := doc.AddParagraph(line.Text)
			para.SetStyle(line.Style)
		}
	}

}

func fixedData(doc *document.Document, billData *types.BillData) {
	sTitle := fmt.Sprintf("    固定岗：（%d元/人/月）", viper.GetInt("fixed_pay"))
	paraTitle := doc.AddParagraph(sTitle)
	paraTitle.SetStyle(DOC_BILL_SITE_TITLE)

	paraList := []types.ParagraphSimple{}
	fixedAreas := viper.GetStringSlice("fixed_area")

	for _, area := range fixedAreas {
		if v, found := billData.FixedBillData[area]; found && v > 0 {
			sum := v * viper.GetInt("fixed_pay")
			paraList = append(paraList, types.ParagraphSimple{
				Text:  fmt.Sprintf("    %s：%d元/人/月*%d人=%d元；", area, viper.GetInt("fixed_pay"), v, sum),
				Style: DOC_BILL_CONTENT,
			})
			billsum += sum
		}
	}

	if len(paraList) > 0 {
		for _, line := range paraList {
			para := doc.AddParagraph(line.Text)
			para.SetStyle(line.Style)
		}
	}
}

func sumText(doc *document.Document) {

	para := doc.AddParagraph(fmt.Sprintf("费用合计金额为：%d元（%s）", billsum, digitalToChar(billsum)))
	para.SetStyle(DOC_BILL_CONTENT)
}

func signture(doc *document.Document) {
	para1 := doc.AddParagraph(viper.GetString("first_party"))
	para1.SetStyle(DOC_BILL_SIGN)
	para1sign := doc.AddParagraph("（签字/盖章）")
	para1sign.SetStyle(DOC_BILL_SIGN)

	para2 := doc.AddParagraph(viper.GetString("corporation_name"))
	para2.SetStyle(DOC_BILL_SIGN)
	para2sign := doc.AddParagraph("（签字/盖章）")
	para2sign.SetStyle(DOC_BILL_SIGN)
}

func digitalToChar(money int) string {
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

// 固定岗和临勤有重叠区域，为区别临勤岗名称后加L，写入文档时去掉
func formatAreaIfNeed(area string) string {
	s, _ := strings.CutSuffix(area, "L")
	return s
}

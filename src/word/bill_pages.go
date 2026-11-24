package word

import (
	"fmt"
	"parking/src/types"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
	"github.com/ZeroHawkeye/wordZero/pkg/style"
	"github.com/spf13/viper"
)

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
	var billTypeStr string
	typesString := viper.GetStringSlice("type_strings")
	if billData.BillDataType > len(typesString) {
		billTypeStr = ""
	} else {
		billTypeStr = typesString[billData.BillDataType]
	}
	paraTitle := doc.AddParagraph(fmt.Sprintf("停车场%s服务项目费用结算单", billTypeStr))
	paraTitle.SetStyle(DOC_BILL_TITLE)
	paraSubtitle := doc.AddParagraph(fmt.Sprintf("（%d年%d月）", billData.Year, billData.Month))
	paraSubtitle.SetStyle(DOC_BILL_SUBTITLE)
}

func billDescription(doc *document.Document, billData *types.BillData) {
	foreword := ""
	switch billData.BillDataType {
	case types.TYPE_BILL_MEDICAL:
		foreword = viper.GetString("corp_medical")
	case types.TYPE_BILL_SCENIC:
		foreword = viper.GetString("corp_scenic")
	default:
		foreword = viper.GetString("corporation_name")
	}
	paraCorporation := doc.AddParagraph(foreword + ":")
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
		// temp4Count := 0
		// if v, found := billData.TempBill4Data[area]; found {
		// 	temp4Count = v
		// }

		if v, found := billData.TempBill8Data[area]; found && (v > 0) {
			count := v
			// if temp4Count%2 != 0 && rand.Int()%2 != 0 {
			// 	count++
			// }
			if count > 0 {
				sum := count * viper.GetInt("temp_8_day")
				paraList = append(paraList, types.ParagraphSimple{
					Text: fmt.Sprintf("    %s（8小时）：%d人*%d元/人/天=%d元;",
						formatAreaIfNeed(area), count, viper.GetInt("temp_8_day"), sum),
					Style: DOC_BILL_CONTENT,
				})
				billData.BillSum += sum
			}

		}

		if v, found := billData.TempBill12Data[area]; found && v > 0 {
			sum := v * viper.GetInt("temp_12_day")
			paraList = append(paraList, types.ParagraphSimple{
				Text: fmt.Sprintf("    %s（12小时）：%d人*%d元/人/天=%d元;",
					formatAreaIfNeed(area), v, viper.GetInt("temp_12_day"), sum),
				Style: DOC_BILL_CONTENT,
			})
			billData.BillSum += sum
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
			billData.BillSum += sum
		}
	}

	if len(paraList) > 0 {
		for _, line := range paraList {
			para := doc.AddParagraph(line.Text)
			para.SetStyle(line.Style)
		}
	}
}

func sumText(doc *document.Document, billData *types.BillData) {

	para := doc.AddParagraph(fmt.Sprintf("费用合计金额为：%d元（%s）", billData.BillSum, DigitalToChar(billData.BillSum)))
	para.SetStyle(DOC_BILL_CONTENT)
}

func signture(doc *document.Document, billData *types.BillData) {
	para1 := doc.AddParagraph(viper.GetString("first_party"))
	para1.SetStyle(DOC_BILL_SIGN)
	para1sign := doc.AddParagraph("单位负责人（签字/盖章）")
	para1sign.SetStyle(DOC_BILL_SIGN)

	secondParty := ""
	switch billData.BillDataType {
	case types.TYPE_BILL_SCENIC:
		secondParty = viper.GetString("corp_scenic")
	case types.TYPE_BILL_MEDICAL:
		secondParty = viper.GetString("corp_medical")
	default:
		secondParty = viper.GetString("corporation_name")
	}

	para2 := doc.AddParagraph(secondParty)
	para2.SetStyle(DOC_BILL_SIGN)
	para2sign := doc.AddParagraph("单位负责人（签字/盖章）")
	para2sign.SetStyle(DOC_BILL_SIGN)
}

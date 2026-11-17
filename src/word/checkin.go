package word

import (
	"fmt"
	"maps"
	"parking/src/types"
	"path"
	"slices"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
	"github.com/ZeroHawkeye/wordZero/pkg/style"
	"github.com/spf13/viper"
)

const (
	DOC_CHECKIN_PAGE_TITLE         = "doc_checkin_page_title"
	DOC_CHECKIN_INFO_PAGE_SUBTITLE = "doc_checkin_info_page_subtitle"
	DOC_CHECKIN_SIGN_PAGE_CONTENT  = "doc_checkin_sign_page_content"
	DOC_CHECKIN_INFO_PAGE_CONTENT  = "doc_checkin_info_page_content"
	DOC_CHECKIN_PAGE_SIGN          = "doc_checkin_page_sign"
)

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

	doc.Save(path.Join(viper.GetString("output"), fmt.Sprintf("临勤（外雇）人员增补签到表_%s.docx", area)))

}

func infoPages(doc *document.Document, tempAttendanceList *[]types.TempAttendance, area string) {

}

func infoPageSingle(doc *document.Document, tempAttendance types.TempAttendance, area string) {

	doc.SetPageMargins(5.56, 15, 5.56, 15)
	doc.SetPageOrientation(document.OrientationLandscape)
	staffs := make(map[string]bool)

	for _, name := range tempAttendance.Names {
		staffs[name] = true
	}

	namesPaged := make([][]string, 0)
	pageCount := len(staffs) / 12

	if len(staffs)%12 != 0 {
		pageCount += 1
	}

	if pageCount == 1 {
		namesPaged = append(namesPaged, slices.Collect(maps.Keys(staffs)))
	} else {
		for i := 0; i < pageCount; i++ {
			if i == pageCount-1 {
				namesPaged = append(namesPaged, slices.Collect(maps.Keys(staffs))[i*12:])
			} else {
				namesPaged = append(namesPaged, slices.Collect(maps.Keys(staffs))[i*12:(i+1)*12])
			}

		}
	}

}

func signPages(doc *document.Document, tempAttendanceList *[]types.TempAttendance, area string) {

	doc.SetPageMargins(5.56, 15, 5.56, 15)
	dates := make(map[int]bool, 0)
	for _, tempAttendance := range *tempAttendanceList {
		dates[tempAttendance.Date] = true
	}

	for date := range dates {
		signPageSingle(doc, date, area)
	}
}

func signPageSingle(doc *document.Document, date int, area string) {

	doc.SetPageMargins(5.56, 15, 5.56, 15)

	doc.AddParagraph("北京怀柔雁栖湖置业服务有限公司").SetStyle(DOC_CHECKIN_PAGE_TITLE)
	doc.AddParagraph("临勤（外雇）人员增补签到表").SetStyle(DOC_CHECKIN_PAGE_TITLE)

	doc.AddParagraph("\n").SetStyle(DOC_CHECKIN_PAGE_TITLE)

	doc.AddParagraph(fmt.Sprintf("单位名称：%s", formatAreaIfNeed(area))).SetStyle(DOC_CHECKIN_SIGN_PAGE_CONTENT)
	doc.AddParagraph(fmt.Sprintf("工作区域：%s", area)).SetStyle(DOC_CHECKIN_SIGN_PAGE_CONTENT)
	doc.AddParagraph(fmt.Sprintf("时   间：20%d年%d月%d日", viper.GetInt("year"), viper.GetInt("month"), date)).SetStyle(DOC_CHECKIN_SIGN_PAGE_CONTENT)
	doc.AddParagraph("工作时段：8 时 00 分至 17 时 00 分").SetStyle(DOC_CHECKIN_SIGN_PAGE_CONTENT)
	doc.AddParagraph(fmt.Sprintf("工作内容：%s", viper.GetString("work_content"))).SetStyle(DOC_CHECKIN_SIGN_PAGE_CONTENT)
	doc.AddParagraph("人员签到：").SetStyle(DOC_CHECKIN_SIGN_PAGE_CONTENT)

	for i := 0; i < 4; i++ {
		doc.AddParagraph("                                      ").SetStyle(DOC_CHECKIN_PAGE_SIGN)

	}
	doc.AddPageBreak()

}

// 预制风格
func setUpCheckinDocStyle(doc *document.Document) {
	styleManager := doc.GetStyleManager()
	quickAPI := style.NewQuickStyleAPI(styleManager)

	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   DOC_CHECKIN_PAGE_TITLE,
		Name: DOC_CHECKIN_PAGE_TITLE,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "center",
			SpaceBefore: 0, // 段前间距（缇）
			SpaceAfter:  0, // 段后间距（缇）
			LineSpacing: 0, // 行间距（缇）
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "宋体(中文正文)",
			FontSize:  24,
			FontColor: "000000",
			Bold:      false,
		},
	})

	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   DOC_CHECKIN_INFO_PAGE_SUBTITLE,
		Name: DOC_CHECKIN_INFO_PAGE_SUBTITLE,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "left",
			SpaceBefore: 0, // 段前间距（缇）
			SpaceAfter:  0, // 段后间距（缇）
			LineSpacing: 0, // 行间距（缇）
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "宋体(中文正文)",
			FontSize:  16,
			FontColor: "000000",
			Bold:      false,
		},
	})

	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   DOC_CHECKIN_SIGN_PAGE_CONTENT,
		Name: DOC_CHECKIN_SIGN_PAGE_CONTENT,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "left",
			SpaceBefore: 0, // 段前间距（缇）
			SpaceAfter:  0, // 段后间距（缇）
			LineSpacing: 0, // 行间距（缇）
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "宋体(中文正文)",
			FontSize:  18,
			FontColor: "000000",
			Bold:      false,
		},
	})

	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   DOC_CHECKIN_INFO_PAGE_CONTENT,
		Name: DOC_CHECKIN_INFO_PAGE_CONTENT,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "left",
			SpaceBefore: 0, // 段前间距（缇）
			SpaceAfter:  0, // 段后间距（缇）
			LineSpacing: 0, // 行间距（缇）
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "宋体(中文正文)",
			FontSize:  12,
			FontColor: "000000",
			Bold:      false,
		},
	})

	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   DOC_CHECKIN_PAGE_SIGN,
		Name: DOC_CHECKIN_PAGE_SIGN,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "left",
			SpaceBefore: 0, // 段前间距（缇）
			SpaceAfter:  0, // 段后间距（缇）
			LineSpacing: 0, // 行间距（缇）
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "宋体(中文正文)",
			FontSize:  12,
			FontColor: "000000",
			Bold:      false,
		},
	})

}

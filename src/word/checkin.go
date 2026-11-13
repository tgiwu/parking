package word

import (
	"fmt"
	"parking/src/types"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
	"github.com/ZeroHawkeye/wordZero/pkg/style"
	"github.com/spf13/viper"
)

const (
	DOC_CHECKIN_PAGE_TITLE    = "doc_checkin_page_title"
	DOC_CHECKIN_INFO_PAGE_SUBTITLE = "doc_checkin_info_page_subtitle"
	DOC_CHECKIN_SIGN_PAGE_CONTENT  = "doc_checkin_sign_page_content"
	DOC_CHECKIN_INFO_PAGE_CONTENT  = "doc_checkin_info_page_content"
)

func CreateCheckinDocx(tempAttendanceMap *map[string][]types.TempAttendance) {

}

func CreateCheckinDocxSingle(tempAttendanceList *[]types.TempAttendance, area string) {
	doc := document.New()
	setUpCheckinDocStyle(doc)
	
	doc.SetPageSize(document.PageSizeA4)
	doc.SetPageOrientation(document.OrientationLandscape)
	signPage(doc, tempAttendanceList, area)


}

func signPage(doc *document.Document, tempAttendanceList *[]types.TempAttendance, area string) {

	doc.SetPageMargins(5.56, 15, 5.56, 15)
	dates := make(map[int]bool, 0)
	for _, tempAttendance := range *tempAttendanceList {
		dates[tempAttendance.Date] = true
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

}
package word

import (
	"fmt"
	"parking/src/types"
	"path"
	"time"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
	"github.com/ZeroHawkeye/wordZero/pkg/style"
	"github.com/spf13/viper"
)

var TABLE_HEADER = []string{"序", "姓名", "出勤", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "签字"}

const DOC_PAGE_TITLE = "DOC_TITLE"
const DOC_PAGE_SUBTITLE = "DOC_SUBTITLE"
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

const TABLE_COL_INDEX_NO = 0
const TABLE_COL_INDEX_NAME = 1
const TABLE_COL_INDEX_ATT_SUM = 2
const TABLE_COL_INDEX_SIGN = 34

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
	doc := document.New()
	setUpStyle(doc)
	doc.SetPageSize(document.PageSizeA4)
	doc.SetPageOrientation(document.OrientationLandscape)
	doc.SetPageMargins(5.56, 15, 5.56, 15)

	for _, pd := range list {
		writeDocxSingle(doc, &pd)
		// if i != len(list) - 1 {
		//add page break
		// }
	}

	doc.Save(path.Join(viper.GetString("output"), viper.GetString("file")))

}

// 预制风格
func setUpStyle(doc *document.Document) {
	styleManager := doc.GetStyleManager()
	quickAPI := style.NewQuickStyleAPI(styleManager)

	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   DOC_PAGE_TITLE,
		Name: DOC_PAGE_TITLE,
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
		ID:   DOC_PAGE_SUBTITLE,
		Name: DOC_PAGE_SUBTITLE,
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

	// quickAPI.CreateQuickStyle(style.QuickStyleConfig{
	// 	ID:   TABLE_HEADER_CHAR,
	// 	Name: TABLE_HEADER_CHAR,
	// 	Type: style.StyleTypeParagraph,
	// 	ParagraphConfig: &style.QuickParagraphConfig{
	// 		Alignment:   "center",
	// 		SpaceBefore: 0, // 段前间距（缇）
	// 		SpaceAfter:  0, // 段后间距（缇）
	// 		LineSpacing: 0, // 行间距（缇）
	// 	},
	// 	RunConfig: &style.QuickRunConfig{
	// 		FontName:  "宋体(中文正文)",
	// 		FontSize:  12,
	// 		FontColor: "000000",
	// 		Bold:      false,
	// 	},
	// })
	// quickAPI.CreateQuickStyle(style.QuickStyleConfig{
	// 	ID:   TABLE_HEADER_NUMBER,
	// 	Name: TABLE_HEADER_NUMBER,
	// 	Type: style.StyleTypeParagraph,
	// 	ParagraphConfig: &style.QuickParagraphConfig{
	// 		Alignment:   "center",
	// 		SpaceBefore: 0, // 段前间距（缇）
	// 		SpaceAfter:  0, // 段后间距（缇）
	// 		LineSpacing: 0, // 行间距（缇）
	// 	},
	// 	RunConfig: &style.QuickRunConfig{
	// 		FontName:  "Calibri (西文正文)",
	// 		FontSize:  7,
	// 		FontColor: "000000",
	// 		Bold:      false,
	// 	},
	// })

	// quickAPI.CreateQuickStyle(style.QuickStyleConfig{
	// 	ID:   TABLE_CONTENT_NO,
	// 	Name: TABLE_CONTENT_NO,
	// 	Type: style.StyleTypeParagraph,
	// 	ParagraphConfig: &style.QuickParagraphConfig{
	// 		Alignment:   "center",
	// 		SpaceBefore: 0, // 段前间距（缇）
	// 		SpaceAfter:  0, // 段后间距（缇）
	// 		LineSpacing: 0, // 行间距（缇）
	// 	},
	// 	RunConfig: &style.QuickRunConfig{
	// 		Bold:     false,
	// 		Italic:   false,
	// 		FontSize: 7,
	// 		FontName: "Calibri (西文正文)",
	// 	},
	// })

	// quickAPI.CreateQuickStyle(style.QuickStyleConfig{
	// 	ID:   TABLE_CONTENT_NAME,
	// 	Name: TABLE_CONTENT_NAME,
	// 	Type: style.StyleTypeParagraph,
	// 	ParagraphConfig: &style.QuickParagraphConfig{
	// 		Alignment:   "center",
	// 		SpaceBefore: 0, // 段前间距（缇）
	// 		SpaceAfter:  0, // 段后间距（缇）
	// 		LineSpacing: 0, // 行间距（缇）
	// 	},
	// 	RunConfig: &style.QuickRunConfig{
	// 		Bold:     false,
	// 		Italic:   false,
	// 		FontSize: 12,
	// 		FontName: "宋体 (中文正文)",
	// 	},
	// })

	// quickAPI.CreateQuickStyle(style.QuickStyleConfig{
	// 	ID:   TABLE_CONTENT_ATT_SUM,
	// 	Name: TABLE_CONTENT_ATT_SUM,
	// 	Type: style.StyleTypeParagraph,
	// 	ParagraphConfig: &style.QuickParagraphConfig{
	// 		Alignment:   "center",
	// 		SpaceBefore: 0, // 段前间距（缇）
	// 		SpaceAfter:  0, // 段后间距（缇）
	// 		LineSpacing: 0, // 行间距（缇）
	// 	},
	// 	RunConfig: &style.QuickRunConfig{
	// 		Bold:     false,
	// 		Italic:   false,
	// 		FontSize: 9,
	// 		FontName: "Calibri (西文正文)",
	// 	},
	// })

	// quickAPI.CreateQuickStyle(style.QuickStyleConfig{
	// 	ID:   TABLE_CONTENT_NORMAL,
	// 	Name: TABLE_CONTENT_NORMAL,
	// 	Type: style.StyleTypeParagraph,
	// 	ParagraphConfig: &style.QuickParagraphConfig{
	// 		Alignment:   "center",
	// 		SpaceBefore: 0, // 段前间距（缇）
	// 		SpaceAfter:  0, // 段后间距（缇）
	// 		LineSpacing: 0, // 行间距（缇）
	// 	},
	// 	RunConfig: &style.QuickRunConfig{
	// 		Bold:     false,
	// 		Italic:   false,
	// 		FontSize: 10,
	// 		FontName: "Arial",
	// 	},
	// })

	// quickAPI.CreateQuickStyle(style.QuickStyleConfig{
	// 	ID:   TABLE_CONTENT_NORMAL_8,
	// 	Name: TABLE_CONTENT_NORMAL_8,
	// 	Type: style.StyleTypeNumbering,
	// 	ParagraphConfig: &style.QuickParagraphConfig{
	// 		Alignment:   "center",
	// 		SpaceBefore: 0, // 段前间距（缇）
	// 		SpaceAfter:  0, // 段后间距（缇）
	// 		LineSpacing: 0, // 行间距（缇）
	// 	},
	// 	RunConfig: &style.QuickRunConfig{
	// 		Bold:     false,
	// 		Italic:   false,
	// 		FontSize: 10,
	// 		FontName: "Arial",
	// 	},
	// })

	// quickAPI.CreateQuickStyle(style.QuickStyleConfig{
	// 	ID:   TABLE_CONTENT_NORMAL_12,
	// 	Name: TABLE_CONTENT_NORMAL_12,
	// 	Type: style.StyleTypeNumbering,
	// 	ParagraphConfig: &style.QuickParagraphConfig{
	// 		Alignment:   "center",
	// 		SpaceBefore: 0, // 段前间距（缇）
	// 		SpaceAfter:  0, // 段后间距（缇）
	// 		LineSpacing: 0, // 行间距（缇）
	// 	},
	// 	RunConfig: &style.QuickRunConfig{
	// 		Bold:     false,
	// 		Italic:   false,
	// 		FontSize: 8,
	// 		FontName: "Arial",
	// 	},
	// })

	// quickAPI.CreateQuickStyle(style.QuickStyleConfig{
	// 	ID:   TABLE_CONTENT_SIGN,
	// 	Name: TABLE_CONTENT_SIGN,
	// 	Type: style.StyleTypeParagraph,
	// 	ParagraphConfig: &style.QuickParagraphConfig{
	// 		Alignment:   "center",
	// 		SpaceBefore: 0, // 段前间距（缇）
	// 		SpaceAfter:  0, // 段后间距（缇）
	// 		LineSpacing: 0, // 行间距（缇）
	// 	},
	// 	RunConfig: &style.QuickRunConfig{
	// 		Bold:     false,
	// 		Italic:   false,
	// 		FontSize: 9,
	// 		FontName: "宋体 (中文正文)",
	// 	},
	// })

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
	if pd.IsTemp {
		t = "临勤"
	} else {
		t = "固定岗"
	}
	titlePara := doc.AddParagraph(fmt.Sprintf("停车场人员%d年度%d月%s出勤统计表", viper.GetInt("year"), viper.GetInt("month"), t))
	titlePara.SetStyle(DOC_PAGE_TITLE)
}

// 副标题和区域
func subtitle(doc *document.Document, pd *types.PageData) {

	subtitlePara := doc.AddParagraph(fmt.Sprintf("用工单位名称：%s                                   负责区域：%s", viper.GetString("corporation_name"), pd.Area))
	subtitlePara.SetStyle(DOC_PAGE_SUBTITLE)
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

			switch {
			case col == TABLE_COL_INDEX_NO:
				textFormate = document.TextFormat{
					Bold:       false,
					Italic:     false,
					FontSize:   7,
					FontFamily: "Calibri (西文正文)",
				}
			case col == TABLE_COL_INDEX_NAME:
				textFormate = document.TextFormat{
					Bold:       false,
					Italic:     false,
					FontSize:   12,
					FontFamily: "宋体 (中文正文)",
				}
			case col == TABLE_COL_INDEX_ATT_SUM:
				textFormate = document.TextFormat{
					Bold:       false,
					Italic:     false,
					FontSize:   9,
					FontFamily: "Calibri (西文正文)",
				}
			case col == TABLE_COL_INDEX_SIGN:
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
					FontSize:   10,
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
				if pd.IsTemp {
					table.SetCellText(row+1, col, fmt.Sprint(len(pda.Data)))
				} else {
					table.SetCellText(row+1, col, fmt.Sprint(31-len(pda.Data)))
				}
			default:
				if pd.IsTemp {
					if v, found := pda.Data[col-2]; found {
						table.SetCellText(row+1, col, fmt.Sprint(v))
					} else {
						table.SetCellText(row+1, col, "")
					}
				} else {
					if v, found := pda.Data[col-2]; found {
						if pda.Data[v] == 1 {
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
	subtitlePara.SetStyle(DOC_PAGE_SUBTITLE)
}

// 计算目标月份天数
func calcLastDayInMonth(year int, month int) int {

	yearStr := fmt.Sprint(year)
	monthStr := fmt.Sprint(month)

	if month < 10 {
		monthStr = "0" + monthStr
	}

	timeLayout := "2025-01-01 15:00:00"

	loc, _ := time.LoadLocation("Local")

	theTime, _ := time.ParseInLocation(timeLayout, yearStr+"-"+monthStr+"-01 00:00:00", loc)

	newMonth := theTime.Month()

	day := time.Date(year, newMonth+1, -1, 0, 0, 0, 0, loc).Day()

	return day
}

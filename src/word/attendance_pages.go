package word

import (
	"fmt"
	"parking/src/types"
	"slices"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
	"github.com/ZeroHawkeye/wordZero/pkg/style"
	"github.com/spf13/viper"
)

func checkInPages(doc *document.Document, tempAttendanceList *[]types.TempAttendance, area string) {

}

func checkInPageSingle() {

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
func fixedTitle(doc *document.Document, pd *types.PageData) {
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
func fixedSubtitle(doc *document.Document, pd *types.PageData) {
	var corp string
	switch {
	case slices.Contains(viper.GetStringSlice("area_scenic"), pd.Area):
		corp = viper.GetString("corp_scenic")
	case slices.Contains(viper.GetStringSlice("area_medical"), pd.Area):
		corp = viper.GetString("corp_medical")
	default:
		corp = viper.GetString("corporation_name")
	}

	subtitlePara := doc.AddParagraph(fmt.Sprintf("用工单位名称：%s                                   负责区域：%s", corp, formatAreaIfNeed(pd.Area)))
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
			cw.W = "1350"
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
func fixedHeader(table *document.Table) {
	for i := range table.GetColumnCount() {
		cellWidth(table, 0, i)
		var textFormate document.TextFormat

		if i == TABLE_COL_INDEX_NO || i == TABLE_COL_INDEX_NAME ||
			i == TABLE_COL_INDEX_ATT_SUM || i == TABLE_COL_INDEX_SIGN {
			textFormate = document.TextFormat{
				Bold:       false,
				Italic:     false,
				FontSize:   11,
				FontFamily: "宋体 (中文正文)",
			}
		} else {
			textFormate = document.TextFormat{
				Bold:       false,
				Italic:     false,
				FontSize:   6,
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
func FixedData(table *document.Table, pd *types.PageData) {
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
				//mark work when type is night or temp
				// if pd.PDType == types.PD_TYPE_NIGHT || pd.PDType == types.PD_TYPE_TEMP {
				// 	if v, found := pda.Data[col-2]; found {
				// 		table.SetCellText(row+1, col, fmt.Sprint(v))
				// 	} else {
				// 		table.SetCellText(row+1, col, "")
				// 	}
				// } else {
					//mark everything when type is site
					if v, found := pda.Data[col-2]; found {
						if v == 1 {
							table.SetCellText(row+1, col, "\\")
						} else {
							table.SetCellText(row+1, col, "√")
						}
					} else {
						table.SetCellText(row+1, col, "√")
					}
				// }
			}
		}
	}
	//area backup
	table.MergeCellsHorizontal(28, 1, len(TABLE_HEADER)-1)

	// if strings.HasSuffix(pd.Area, "L") {
	// 	table.SetCellText(28, 1, "注明：出勤用“/”表示")
	// } else {
		table.SetCellText(28, 1, "备注：出勤用“√”表示；休假用“\\”表示。")
	// }
}

// area sign
func signerArea(doc *document.Document) {
	subtitlePara := doc.AddParagraph("用工单位负责人：                      物业分管负责人：                   部门经理：")
	subtitlePara.SetStyle(DOC_TABLE_PAGE_SUBTITLE)
}

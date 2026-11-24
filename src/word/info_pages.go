package word

import (
	"maps"
	"parking/src/types"
	"slices"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
)

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
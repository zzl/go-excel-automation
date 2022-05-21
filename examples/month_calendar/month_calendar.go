package main

import (
	"fmt"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-excel-automation/excel"
	"time"
)

func main() {
	ole.Initialize()
	defer ole.Uninitialize()

	defer com.NewScope().Leave()

	xApp, _ := excel.NewApplicationInstance(true)
	xApp.SetVisible(true)

	objWs := xApp.Workbooks().Add().Worksheets().Item(1)
	ws := ole.As[*excel.Worksheet](objWs)

	ws.Range("1:1").SetRowHeight(30)
	ws.Range("A1:G8").SetHorizontalAlignment(excel.XlHAlign.XlHAlignCenter)
	ws.Range("A2:A8").Font().SetColorIndex(3)

	today := time.Now()
	rng := ws.Range("A1:G1")
	rng.Merge()
	rng.Font().SetBold(true)
	rng.SetNumberFormatLocal("@")
	rng.SetValue2(today.Format("Jan 2006"))
	rng.Select()

	for n, name := range []string{"Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"} {
		excel.RangeFromVar(ws.Cells().Item(2, n+1)).SetValue2(name)
	}

	daysArr := ole.NewArray2D[ole.Variant](6, 7, true)
	dateFrom := today.AddDate(0, 0, -today.Day()+1)
	dateTo := dateFrom.AddDate(0, 1, -1)
	for r := 0; r < 6; r++ {
		for c := 0; c < 7; c++ {
			day := r*7 + c + 1 - int(dateFrom.Weekday())
			if day >= dateFrom.Day() && day <= dateTo.Day() {
				sDay := fmt.Sprintf("%2d", day)
				daysArr.SetAt2(r, c, ole.VarScoped(sDay))
			}
		}
	}
	ws.Range("A3:G8").SetValue2(daysArr)
}

package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/briandowns/spinner"
	"os"
	"strconv"
	"strings"
	"time"
)

//运行
func Run(name string, f func() error) {
	fmt.Println("数据" + name + "开始...")
	s := spinner.New(spinner.CharSets[36], 100*time.Millisecond)
	s.Start()
	err := f()
	if err != nil {
		s.Stop()
		fmt.Println("数据" + name + "失败...")
		fmt.Println(err.Error())
		os.Exit(1)
	}
	time.Sleep(time.Second * 1)
	s.Stop()
	fmt.Println("数据" + name + "成功...")
}

//判断文件是否存在
func FileExists(name string) bool {
	if _, err := os.Stat(name); err != nil {
		if os.IsNotExist(err) {
			return false
		}
	}
	return true
}

// 计算日期相差多少月
func SubMonth(t1, t2 time.Time) (month int) {
	y1 := t1.Year()
	y2 := t2.Year()
	m1 := int(t1.Month())
	m2 := int(t2.Month())
	d1 := t1.Day()
	d2 := t2.Day()

	yearInterval := y1 - y2
	if m1 < m2 || m1 == m2 && d1 < d2 {
		yearInterval--
	}
	// 获取月数差值
	monthInterval := (m1 + 12) - m2
	//if d1 < d2 {
	//	monthInterval--
	//}
	monthInterval %= 12
	month = yearInterval*12 + monthInterval + 1
	return
}

// excel日期字段格式化 yyyy-mm-dd
func ConvertToFormatDay(excelDaysString string) string {
	baseDiffDay := 38719
	curDiffDay := excelDaysString
	b, _ := strconv.Atoi(curDiffDay)
	realDiffDay := b - baseDiffDay
	realDiffSecond := realDiffDay * 24 * 3600
	baseOriginSecond := 1136185445
	resultTime := time.Unix(int64(baseOriginSecond+realDiffSecond), 0).Format("2006-01-02")
	return resultTime
}

//获取单元格的背景颜色
func GetCellBgColor(f *excelize.File, sheet, axix string) string {
	styleID, err := f.GetCellStyle(sheet, axix)
	if err != nil {
		return err.Error()
	}
	fillID := *f.Styles.CellXfs.Xf[styleID].FillID
	fgColor := f.Styles.Fills.Fill[fillID].PatternFill.FgColor
	if fgColor.Theme != nil {
		children := f.Theme.ThemeElements.ClrScheme.Children
		if *fgColor.Theme < 4 {
			dklt := map[int]string{
				0: children[1].SysClr.LastClr,
				1: children[0].SysClr.LastClr,
				2: *children[3].SrgbClr.Val,
				3: *children[2].SrgbClr.Val,
			}
			return strings.TrimPrefix(
				excelize.ThemeColor(dklt[*fgColor.Theme], fgColor.Tint), "FF")
		}
		srgbClr := *children[*fgColor.Theme].SrgbClr.Val
		return strings.TrimPrefix(excelize.ThemeColor(srgbClr, fgColor.Tint), "FF")
	}
	return strings.TrimPrefix(fgColor.RGB, "FF")
}

package main

import (
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"path/filepath"
	"strconv"
)

func Calendar() (err error) {
	// 创建工作簿
	f := excelize.NewFile()

	for k, v := range sheetNameList {
		err = NewMonthSheet(f, k, v)
		if err != nil {
			continue
		}
	}

	// 保存工作簿
	savePath := filepath.Join(dirPath, appConfig.CalendarName)
	//fmt.Println("导出文件路径为：" + savePath)
	if err = f.SaveAs(savePath); err != nil {
		return err
	}
	return nil
}

func NewMonthSheet(f *excelize.File, k int, sheetName string) (err error) {
	var (
		monthStyle, titleStyle, dataStyle, blankStyle int
		addr                                          string
		sheet                                         = "Sheet1"
		top                                           = excelize.Border{Type: "top", Style: 1, Color: "DADEE0"}
		left                                          = excelize.Border{Type: "left", Style: 1, Color: "DADEE0"}
		right                                         = excelize.Border{Type: "right", Style: 1, Color: "DADEE0"}
		bottom                                        = excelize.Border{Type: "bottom", Style: 1, Color: "DADEE0"}
	)

	if k > 0 {
		f.NewSheet(sheet)
	}

	//例子
	//excelData = map[int][]interface{}{
	//	1: {"九月 2020"},
	//	3: {"日", "一", "二", "三", "四", "五", "六"},
	//	4:  {"", 1, 2, 3, 4, 5, 6},
	//	5:  {"aaa","bbb","ccc","ddd","eeee","ffff","gggg"},
	//	6:  {7, 8, 9, 10, 11, 12, 13},
	//	7:  {"aaa","bbb","ccc","ddd","eeee","ffff","gggg"},
	//	8:  {14, 15, 16, 17, 18, 19, 20},
	//	9:  {"aaa","bbb","ccc","ddd","eeee","ffff","gggg"},
	//	10: {21, 22, 23, 24, 25, 26, 27},
	//	11:  {"aaa","bbb","ccc","ddd","eeee","ffff","gggg"},
	//	12: {28, 29, 30, 31, "", "", ""},
	//	13:  {"aaa","bbb","ccc","ddd","eeee","ffff","gggg"},
	//}

	// 按行赋值
	for r, row := range excelDataList[k] {
		if addr, err = excelize.JoinCellName("A", r); err != nil {
			return err
		}
		if err = f.SetSheetRow(sheet, addr, &row); err != nil {
			return err
		}
	}
	// 设置自定义行高
	for r, ht := range heightMapList[k] {
		if err = f.SetRowHeight(sheet, r, ht); err != nil {
			return err
		}
	}
	// 设置列宽
	if err = f.SetColWidth(sheet, "A", "G", 16.5); err != nil {
		return err
	}
	// 合并月份单元格
	if err = f.MergeCell(sheet, "A1", "C1"); err != nil {
		return err
	}
	// 设置月份单元格样式
	if monthStyle, err = f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "1f7f3b", Bold: true, Size: 22, Family: "Microsoft YaHei"},
	}); err != nil {
		return err
	}
	// 设置月份单元格字体
	if err = f.SetCellStyle(sheet, "A1", "C1", monthStyle); err != nil {
		return err
	}
	// 创建周一至周日标题行样式
	if titleStyle, err = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Color: "1f7f3b", Bold: true, Family: "Microsoft YaHei"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"E6F4EA"}, Pattern: 1},
		Alignment: &excelize.Alignment{Vertical: "center", Horizontal: "center"},
		Border:    []excelize.Border{{Type: "top", Style: 2, Color: "1f7f3b"}},
	}); err != nil {
		return err
	}
	// 设置周一至周日标题行样式
	if err = f.SetCellStyle(sheet, "A3", "G3", titleStyle); err != nil {
		return err
	}
	// 创建日期单元格样式
	if dataStyle, err = f.NewStyle(&excelize.Style{
		Border: []excelize.Border{top, left, right},
		Protection: &excelize.Protection{
			Locked: true,
		},
	}); err != nil {
		return err
	}
	// 设置日期单元格样式
	for _, r := range rowsDateList[k] {
		if err = f.SetCellStyle(sheet, "A"+strconv.Itoa(r),
			"G"+strconv.Itoa(r), dataStyle); err != nil {
			return err
		}
	}
	// 创建空白单元格样式
	if blankStyle, err = f.NewStyle(&excelize.Style{
		Border: []excelize.Border{left, right, bottom},
		Alignment: &excelize.Alignment{
			WrapText: true,
		},
		Protection: &excelize.Protection{
			Locked: true,
		},
	}); err != nil {
		return err
	}
	// 设置空白单元格样式
	for _, r := range rowsContentList[k] {
		if err = f.SetCellStyle(sheet, "A"+strconv.Itoa(r),
			"G"+strconv.Itoa(r), blankStyle); err != nil {
			return err
		}
	}
	// 隐藏工作表网格线
	if err = f.SetSheetViewOptions(sheet, 0,
		excelize.ShowGridLines(false)); err != nil {
		return err
	}
	// 重命名工作表
	f.SetSheetName(sheet, sheetName)

	return nil
}

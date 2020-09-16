package main

import (
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/briandowns/spinner"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
)

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

//查找文件目录
func FindPath() (path, filePath string, err error) {

	appPath, err := filepath.Abs(filepath.Dir(os.Args[0]));
	if err != nil {
		return "", "", err
	}
	path = appPath
	workPath, err := os.Getwd()
	if err != nil {
		return "", "", err
	}
	var filename = "1.xlsx"

	filePath = filepath.Join(workPath, filename)
	if !FileExists(filePath) {
		filePath = filepath.Join(appPath, filename)
		path = appPath
		if !FileExists(filePath) {
			return "", "", errors.New("文件不存在")
		}
	}
	return path, filePath, nil
}

func FileExists(name string) bool {
	if _, err := os.Stat(name); err != nil {
		if os.IsNotExist(err) {
			return false
		}
	}
	return true
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

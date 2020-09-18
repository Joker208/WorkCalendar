package main

import (
	"fmt"
	"strconv"
	"time"
)

var (
	excelDataList 					[]map[int][]interface{}
	heightMapList                 	[]map[int]float64
	rowsDateList, rowsContentList 	[][]int
	sheetNameList                 	[]string
	index                         	int
)

func Handle() (err error) {

	heightMapList = make([]map[int]float64, 0)
	rowsDateList = make([][]int, 0)
	rowsContentList = make([][]int, 0)
	sheetNameList = make([]string, 0)

	lastDate, err := time.Parse("2006-01-02", lastDay)
	if err != nil {
		return err
	}

	monthSum := SubMonth(lastDate, time.Now())
	excelDataList = make([]map[int][]interface{}, 0, monthSum)

	year, month, _ := time.Now().Date()
	monthInt := int(month)

	for m := 0; m < monthSum; m++ {
		if monthInt > 12 {
			year++
			monthInt = monthInt - 12
		}
		handleMonth(year, monthInt)
		monthInt++
	}

	return nil

}

func handleMonth(year, month int) {

	//记录日期列和内容列
	rowsDate := make([]int, 0)
	rowsContent := make([]int, 0)

	//定位月初
	monthStart := time.Date(year, time.Month(month), 1, 0, 0, 0, 0, time.UTC)
	//定位月末
	monthEnd := time.Date(year, time.Month(month+1), 1, 0, 0, 0, 0, time.UTC)
	monthEnd = monthEnd.Add(-24 * time.Hour)
	//定位月初星期
	week := monthStart.Weekday()
	//定位月末星期
	week1 := monthEnd.Weekday()

	yearStr := strconv.Itoa(year)
	monthStr := strconv.Itoa(int(month))

	sheetName := yearStr + "年" + monthStr + "月"
	//日历初始内容
	excelData := make(map[int][]interface{})
	excelData[1] = []interface{}{yearStr + " " + monthStr + "月"}
	excelData[3] = []interface{}{"一", "二", "三", "四", "五", "六", "日"}

	//数据从4行开始
	rowNum := 4

	//判断每周需要行数
	weekRowMap := make(map[int]int)
	w := int(week)
	maxRow := 1
	index = 0
	for d1 := 1; d1 <= monthEnd.Day(); d1++ {
		dateStr := fmt.Sprintf("%s-%02d-%02d", yearStr, month, d1)
		if dateNum[dateStr] > maxRow {
			maxRow = dateNum[dateStr]
		}
		if w%7 == 0 {
			weekRowMap[index] = maxRow
			index++
			maxRow = 1
		}
		w++
	}

	index = 0
	//补全日历开头
	for i := 0; i < (int(week)+6)%7; i++ {
		excelData[rowNum] = append(excelData[rowNum], "")

		for j := 1; j < weekRowMap[index]+1; j++ {
			excelData[rowNum+j] = append(excelData[rowNum+j], "")
		}
	}

	//补全数据
	w = int(week)
	for d := 1; d <= monthEnd.Day(); d++ {
		excelData[rowNum] = append(excelData[rowNum], d)

		dateStr := fmt.Sprintf("%s-%02d-%02d", yearStr, month, d)
		dayContent := excelContent[dateStr]
		dayLength := len(dayContent)
		for j := 0; j < weekRowMap[index]; j++ {
			content := ""
			if j < dayLength {
				content = dayContent[j]
			}
			excelData[rowNum+j+1] = append(excelData[rowNum+j+1], content)
		}
		if w++; w%7 == 1 {
			rowsDate = append(rowsDate, rowNum)
			for j := 0; j < weekRowMap[index]; j++ {
				rowsContent = append(rowsContent, rowNum+j+1)
			}

			rowNum = rowNum + weekRowMap[index] + 1
			index++
			//存在最后几周没有
			if index >= len(weekRowMap) {
				weekRowMap[index] = 1
			}
		}
	}
	if week1 != time.Sunday {
		rowsDate = append(rowsDate, rowNum)
		for j := 1; j < weekRowMap[index]+1; j++ {
			rowsContent = append(rowsContent, rowNum+j)
		}
	}

	//补全日历结尾
	for i := 0; i < (7-int(week1))%7; i++ {
		excelData[rowNum] = append(excelData[rowNum], "")

		for j := 1; j < weekRowMap[index]+1; j++ {
			excelData[rowNum+j] = append(excelData[rowNum+j], "")
		}
	}

	//设置行高
	heightMap := make(map[int]float64)
	heightMap[1] = 45
	heightMap[3] = 22
	for _, v := range rowsContent {
		heightMap[v] = 44
	}

	//存到数组里
	excelDataList = append(excelDataList, excelData)
	heightMapList = append(heightMapList, heightMap)
	rowsDateList = append(rowsDateList, rowsDate)
	rowsContentList = append(rowsContentList, rowsContent)
	sheetNameList = append(sheetNameList, sheetName)

	return
}

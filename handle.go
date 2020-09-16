package main

import (
	"fmt"
	"strconv"
	"time"
)

var (
	excelData         map[int][]interface{}
	weekRowList       map[int]int
	heightList        map[int]float64
	rowList, rowList1 []int
	sheetName         string
	index             int
)

func Handle() (err error) {

	excelData = make(map[int][]interface{})

	year, month, _ := time.Now().Date()
	//定位月初
	monthStart := time.Date(year, month, 1, 0, 0, 0, 0, time.UTC)
	//定位月末
	monthEnd := time.Date(year, month+1, 1, 0, 0, 0, 0, time.UTC)
	monthEnd = monthEnd.Add(-24 * time.Hour)
	//定位月初星期
	week := monthStart.Weekday()

	//例子
	//data = map[int][]interface{}{
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

	yearStr := strconv.Itoa(year)
	monthStr := strconv.Itoa(int(month))

	sheetName = yearStr + "年" + monthStr + "月"
	//日历初始内容
	excelData[1] = []interface{}{yearStr + " " + monthStr + "月"}
	excelData[3] = []interface{}{"日", "一", "二", "三", "四", "五", "六"}

	//数据从4行开始
	rowNum := 4

	//判断每周需要行数
	weekRowList = make(map[int]int)
	w := week
	maxRow := 1
	index = 0
	for d1 := 1; d1 <= monthEnd.Day(); d1++ {
		dateStr := fmt.Sprintf("%s-%02d-%02d", yearStr, month, d1)
		if dateNum[dateStr] > maxRow {
			maxRow = dateNum[dateStr]
		}
		if w++; w%7 == 0 {
			weekRowList[index] = maxRow
			index++
			maxRow = 1
		}
	}

	index = 0
	//补全日历开头
	for i := 0; i < int(week); i++ {
		excelData[rowNum] = append(excelData[rowNum], "")

		for j := 1; j < weekRowList[index]+1; j++ {
			excelData[rowNum+j] = append(excelData[rowNum+j], "")
		}
	}

	//补全数据

	for d := 1; d <= monthEnd.Day(); d++ {
		excelData[rowNum] = append(excelData[rowNum], d)

		dateStr := fmt.Sprintf("%s-%02d-%02d", yearStr, month, d)
		dayContent := excelContent[dateStr]
		dayLength := len(dayContent)
		for j := 0; j < weekRowList[index]; j++ {
			content := ""
			if j < dayLength {
				content = dayContent[j]
			}
			excelData[rowNum+j+1] = append(excelData[rowNum+j+1], content)
		}

		if week++; week%7 == 0 {
			rowList = append(rowList, rowNum)
			for j := 0; j < weekRowList[index]; j++ {
				rowList1 = append(rowList1, rowNum+j+1)
			}

			rowNum = rowNum + weekRowList[index] + 1
			index++
			//存在最后几周没有
			if index >= len(weekRowList) {
				weekRowList[index] = 1
			}
		}
	}
	rowList = append(rowList, rowNum)
	for j := 1; j < weekRowList[index]+1; j++ {
		rowList1 = append(rowList1, rowNum+j)
	}

	//补全日历结尾
	for i := 0; i < int(7-week%7); i++ {
		excelData[rowNum] = append(excelData[rowNum], "")

		for j := 1; j < weekRowList[index]+1; j++ {
			excelData[rowNum+j] = append(excelData[rowNum+j], "")
		}
	}

	//设置行高
	heightList = make(map[int]float64)
	heightList[1] = 45
	heightList[3] = 22
	for _, v := range rowList1 {
		heightList[v] = 44
	}

	//fmt.Println("weekRowList")
	//fmt.Println(weekRowList)
	//fmt.Println("rowList")
	//fmt.Println(rowList)
	//fmt.Println("rowList1")
	//fmt.Println(rowList1)

	return nil

}

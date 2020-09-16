package main

import (
	"errors"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

var (
	excelContent map[string][]string
	dateNum      map[string]int
	path         string
)

func Scan() (err error) {

	var popo, task = "", ""
	var filePath = ""
	path, filePath, err = FindPath()
	if err != nil {
		return err
	}

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return err
	}

	//colorStr:= getCellBgColor(f, "Sheet1", "E2")
	//colorStr1:= getCellBgColor(f, "Sheet1", "E22")
	//fmt.Println("colorStr:",colorStr)
	//fmt.Println("colorStr1:",colorStr1)

	//新建map存放数据
	//按日期分组
	excelContent = make(map[string][]string)
	dateNum = make(map[string]int)
	titleList := [18]string{"aaa","bbb", "ccc"}

	// 获取 Sheet1 上所有单元格
	rows, err := f.GetRows("Sheet1")

	if len(rows) < 4 {
		return errors.New("数据量太少，导出失败")
	}

	//fmt.Println(len(rows))

	for i := 1; i < len(rows); i++ {
		rowLen := len(rows[i])
		rowTemp := rows[i]
		content := ""

		//以日期为标准，合并内容

		if rowLen < 3 {
			break
		}

		//取消合并内容
		if rowTemp[0] != "" {
			popo = rowTemp[0]
		} else {
			rowTemp[0] = popo
		}

		if rowTemp[1] != "" {
			task = rowTemp[1]
		} else {
			rowTemp[1] = task
		}

		//以日期为准合并内容
		date := ConvertToFormatDay(rowTemp[2])

		if _, ok := excelContent[date]; ok {
			//存在相同日期
			for j := 0; j < rowLen; j++ {
				//空置和日期不记录
				//if rowTemp[j] != "" && j != 2 && j != 3 {

				//改为只记录任务和内容
				if j == 1 || j == 4 {
					content += titleList[j] + ":" + rowTemp[j] + ";"
				}
			}
			dateNum[date]++
		} else {
			//新建日期事件
			for j := 0; j < rowLen; j++ {
				//空置和日期不记录
				//if rowTemp[j] != "" && j != 2 && j != 3 {

				//改为只记录任务和内容
				if j == 1 || j == 4 {
					content += titleList[j] + ":" + rowTemp[j] + ";"
				}
			}
			excelContent[date] = make([]string, 0)
			dateNum[date] = 1
		}
		excelContent[date] = append(excelContent[date], content)
		//fmt.Println("excelContent")
		//fmt.Println(excelContent)
		//fmt.Println("dateNum")
		//fmt.Println(dateNum)
	}
	return nil
}

package main

import "fmt"

func main() {
	//读取数据
	Run("读取",Scan)
	//处理数据
	Run("处理",Handle)
	//导出数据
	Run("导出",Calendar)

	//暂停
	var str string
	fmt.Scanln(&str)
}
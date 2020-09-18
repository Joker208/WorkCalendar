package main

import (
	"fmt"
	"gopkg.in/ini.v1"
	"os"
	"path/filepath"
)

type Config struct {
	FileName     string
	SheetName    string
	TitleList    []string
	CalendarName string
}

var (
	dirPath   string
	appConfig *Config
)

func init() {
	//配置文件路径
	appPath, err := filepath.Abs(filepath.Dir(os.Args[0]));
	if err != nil {
		os.Exit(1)
	}
	dirPath = appPath
	workPath, err := os.Getwd()
	if err != nil {
		os.Exit(1)
	}
	var filename = "app.conf"
	filePath := filepath.Join(workPath, filename)
	if !FileExists(filePath) {
		filePath = filepath.Join(appPath, filename)
		dirPath = appPath
		if !FileExists(filePath) {
			os.Exit(1)
		}
	}

	//导入配置
	cfg, err := ini.Load(filePath)
	if err != nil {
		fmt.Printf("找不到配置文件: %v", err)
		os.Exit(1)
	}

	appConfig = new(Config)

	err = cfg.MapTo(appConfig)
	if err != nil {
		fmt.Printf("转化失败: %v", err)
		os.Exit(1)
	}
}

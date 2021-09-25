package main

import (
	"fmt"
	"gopkg.in/yaml.v3"
	"io/ioutil"
)

type Conf struct {
	SourceData	 string `yaml:"SourceData"`
	SourceSheets string `yaml:"SourceSheets"`
	ToSheets	 string	`yaml:"ToSheets"`
	SplitByForm  string	`yaml:"SplitByForm"`
}

func ReadYaml()  *Conf{
	// 读取配置文件
	var conf Conf
	configs,err := ioutil.ReadFile("./configs/conf.yaml")
	if err != nil{
		fmt.Println(err)
	}
	// 转为struct
	err = yaml.Unmarshal(configs,&conf)
	if err != nil{
		fmt.Println(err)
	}
	return &conf
}
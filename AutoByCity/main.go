package main

import (
	"fmt"
	"os"
	"strconv"
	"strings"

	el "github.com/xuri/excelize/v2"
)

func main() {
	// 读取配置文件
	conf := ReadYaml()

	// 转为数组
	SourceSheetArr := strings.Split(conf.SourceSheets, ",")
	ToSheetsArr := strings.Split(conf.ToSheets, ",")
	SplitByForm := strings.Split(conf.SplitByForm, ",")

	if len(SourceSheetArr) != len(ToSheetsArr) {
		fmt.Println("需要生成的sheet表，数量与上一个参数数量相同")
		return
	}

	// 打开主文件
	f, err := el.OpenFile(conf.SourceData)
	if err != nil {
		fmt.Println(err)
	}

	// 生成文件
	for k, sheet := range SourceSheetArr {
		// 根据sheet表获取数据
		rows, err := f.GetRows(sheet)
		if err != nil {
			fmt.Println(err)
			return
		}
		// 获取分割字段
		index, err := strconv.Atoi(SplitByForm[k])
		if err != nil {
			fmt.Println(err)
			return
		}

		// 将二维数组转为map
		sheetMap, sheetHeader := rowsToMap(rows, index)
		// 生成文件
		MapToExcel(sheetMap, sheetHeader, ToSheetsArr[k])
	}
}

// row 转为 map key地市 -> value [][]string
func rowsToMap(rows [][]string, index int) (sheetMap map[string][][]string, header []string) {
	sheetMap = make(map[string][][]string)
	header = make([]string, 0)
	for k, row := range rows {
		sheetArr := make([][]string, 0)
		if k == 0 {
			header = append(header, row...)
		} else {
			sheetArr = append(sheetArr, row)
			if _, ok := sheetMap[row[index]]; ok {
				sheetMap[row[index]] = append(sheetMap[row[index]], sheetArr...)
			} else {
				sheetMap[row[index]] = sheetArr
			}
		}
	}
	return
}

// MapToExcel 将转为excel
func MapToExcel(sheetMap map[string][][]string, header []string, sheet string) {
	for k, v := range sheetMap {
		var f1 *el.File
		flag := isExist("./resource/" + k + ".xlsx")
		if !flag {
			f1 = el.NewFile()
		} else {
			f1, _ = el.OpenFile("./resource/" + k + ".xlsx")
		}
		index := f1.NewSheet(sheet)
		f1.DeleteSheet("Sheet1")
		f1.SetActiveSheet(index)
		for k, v := range header {
			f1.SetCellValue(sheet, axis(1, k), v)
		}
		for i := 0; i < len(v); i++ {
			row := i + 2 // 从第二行开始
			for k := 0; k < len(v[i]); k++ {
				f1.SetCellValue(sheet, axis(row, k), v[i][k])
			}
		}
		if err := f1.SaveAs("./resource/" + k + ".xlsx"); err != nil {
			fmt.Println(err)
		}
	}
}

// axis 根据列获取对应的excel的位置
func axis(row, column int) string {
	if column <= 25 {
		return fmt.Sprintf("%s%d", string(rune(65+column)), row)
	} else if column > 25 {
		a := (column + 1) / (26)
		column -= a * 26
		code := string(rune(64 + a))
		return fmt.Sprintf("%s%s%d", code, string(rune(65+column)), row)
	}
	return ""
}

// isExist 判断文件是否存在
func isExist(path string) bool {
	_, err := os.Stat(path)
	if err != nil {
		if os.IsExist(err) {
			return true
		}
		if os.IsNotExist(err) {
			return false
		}
		fmt.Println(err)
		return false
	}
	return true
}

package main

import (
	"bufio"
	"errors"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
	"unicode"

	"github.com/xuri/excelize/v2"
)

func GetLastNumbericCellValueByGetRows(f *excelize.File, sheetName string) (value float64, err error) {
	rows, err := f.GetRows(sheetName, excelize.Options{RawCellValue: true})
	if err != nil {
		return 0, errors.Join(err, fmt.Errorf("%s %s获取行数据失败", f.Path, sheetName))
	}
	l := len(rows)

	for i := l - 1; i >= 0; i-- {
		for _, v := range rows[i] {
			f, err := strconv.ParseFloat(strings.TrimSpace(v), 64)
			if err == nil {
				return f, nil
			}
		}
	}
	return 0, fmt.Errorf("%s %s没有找到数字值", f.Path, sheetName)
}

func GetLastNumbericCellValueByRows(f *excelize.File, sheeName string) (value float64, err error) {
	rows, err := f.Rows(sheeName)
	if err != nil {
		return 0, err
	}
	defer rows.Close()
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			fmt.Println("读取行数据错误", err)
			continue
		}
		for i := len(row) - 1; i >= 0; i-- {
			f, err := strconv.ParseFloat(strings.TrimSpace(row[i]), 64)
			if err == nil {
				value = f
				break
			}

		}
	}
	return value, nil
}

// 判断字符串是否在2个汉字到4个汉字之间
func IsStringLengthBetween2And4ChineseChars(s string) bool {
	count := 0
	for _, r := range s {
		if !unicode.Is(unicode.Han, r) {
			return false
		}
		count++
	}
	return count >= 2 && count <= 4
}

func GetXlsxFiles(dir string) ([]string, error) {
	// 遍历当前目录下的所有文件
	entries, err := os.ReadDir(dir)
	if err != nil {
		fmt.Println("读取目录失败:", err)
	}

	// 筛选出 Excel 文件
	excelFiles := []string{}
	for _, entry := range entries {
		if !entry.IsDir() && filepath.Ext(entry.Name()) == ".xlsx" && !strings.HasPrefix(entry.Name(), "~$") {
			excelFiles = append(excelFiles, entry.Name())
		}
	}
	return excelFiles, nil
}

func GetPersonPerformance(fileName string) (result map[string]float64, err error) {
	// 打开 Excel 文件
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	result = make(map[string]float64, 0)

	for _, sheetName := range f.GetSheetList() {
		if !IsStringLengthBetween2And4ChineseChars(sheetName) {
			continue
		}

		value, err := GetLastNumbericCellValueByGetRows(f, sheetName)
		if err != nil {
			fmt.Printf("Error: %v\n", err)
			continue
		}
		result[sheetName] = math.Round(value*100) / 100

	}
	return

}

func WriteToExcel(fileName string, data [][]any) error {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	sw, err := f.NewStreamWriter("Sheet1")
	if err != nil {
		return err
	}
	border := []excelize.Border{
		{
			Type:  "left",
			Style: 1,
			Color: "000000",
		}, {
			Type:  "right",
			Style: 1,
			Color: "000000",
		}, {
			Type:  "top",
			Style: 1,
			Color: "000000",
		}, {
			Type:  "bottom",
			Style: 1,
			Color: "000000",
		},
	}
	alignment := &excelize.Alignment{Horizontal: "center", Vertical: "center"}
	headerStyle := &excelize.Style{Font: &excelize.Font{Size: 14, Bold: true}, Border: border, Alignment: alignment}
	styleID, err := f.NewStyle(headerStyle)
	if err != nil {
		return errors.Join(err, errors.New("创建样式失败"))
	}
	rowStyle := &excelize.Style{Font: &excelize.Font{Size: 12}, Border: border, Alignment: alignment}
	rowStyleID, err := f.NewStyle(rowStyle)
	if err != nil {
		return errors.Join(err, errors.New("创建样式失败"))
	}
	// 设置列宽
	err = sw.SetColWidth(1, 1, 10.0)
	if err != nil {
		fmt.Println(err)
	}
	err = sw.SetColWidth(2, 4, 20.0)
	if err != nil {
		fmt.Println(err)
	}
	err = sw.MergeCell("A1", "E1")
	if err != nil {
		fmt.Println("合并单元格失败")
	}

	sw.SetRow("A1", []any{excelize.Cell{Value: fmt.Sprintf("质检部%s工作量统计表", time.Now().Format("2006年1月")), StyleID: styleID}}, excelize.RowOpts{Height: 30.0})
	header := []any{"序号", "姓名", "班组", "工作量", "备注"}
	sw.SetRow("A2", genCellsWithStyle(header, styleID))
	for i, row := range data {
		cell, err := excelize.JoinCellName("A", i+3)
		if err != nil {
			fmt.Println("join err:", err)
			continue
		}
		rowWithStyle := genCellsWithStyle(row, rowStyleID)
		if err := sw.SetRow(cell, rowWithStyle); err != nil {
			fmt.Println("set row err:", err)
			continue
		}

	}
	// 写入合计值
	l := len(data)
	sumStyle := &excelize.Style{Font: &excelize.Font{Size: 14, Bold: true}, Alignment: alignment}
	sumStyleID, err := f.NewStyle(sumStyle)
	if err != nil {
		return errors.Join(err, errors.New("创建汇总样式失败"))
	}
	sumRow := append(genCellsWithStyle([]any{"", "", "合计"}, sumStyleID), excelize.Cell{Formula: fmt.Sprintf("=SUM(D3:D%d)", l+2), StyleID: sumStyleID}, excelize.Cell{Value: "", StyleID: sumStyleID})
	err = sw.SetRow(fmt.Sprintf("A%d", l+3), sumRow)
	if err != nil {
		fmt.Println("写入合计值失败:", err)
		return err
	}

	if err := sw.Flush(); err != nil {
		return err
	}
	return f.SaveAs(fileName)
}

// 生成带样式的行数据
func genCellsWithStyle(row []any, styleID int) []any {
	cells := make([]any, 0, len(row))
	for _, v := range row {
		cells = append(cells, excelize.Cell{Value: v, StyleID: styleID})
	}
	return cells
}

func mergeMaps(map1, map2 map[string]float64) {
	for key, value := range map2 {
		map1[key] += value
	}
}

func Sumlize() (err error) {
	date := time.Now().Format("2006-01")
	savePath := fmt.Sprintf("./质检部工作量统计表%s.xlsx", date)
	// 获取绝对路径
	absolutePath, err := filepath.Abs(savePath)
	if err != nil {
		fmt.Println("获取绝对路径失败:", err)
		return
	}
	namePath, err := filepath.Abs("./质检部花名册.xlsx")
	if err != nil {
		fmt.Println("获取绝对路径失败:", err)
		return
	}
	xlsxFiles, err := GetXlsxFiles("./")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Printf("正在读取%d个xlsx文件: %v\n", len(xlsxFiles), xlsxFiles)

	writeData := make(map[string]float64)
	for _, fileName := range xlsxFiles {
		m, err := GetPersonPerformance(fileName)
		if err != nil {
			fmt.Println(err)
			continue
		}
		mergeMaps(writeData, m)
	}

	f, err := excelize.OpenFile(namePath)
	if err != nil {
		fmt.Println("获取员工花名册文件失败:", namePath)
		return
	}
	sheet := f.GetSheetList()[0]
	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println("获取员工花名册数据失败:", namePath, sheet)
		return
	}
	sumlize := make([][]any, 0)
	idx := 1
	for _, row := range rows[1:] {
		name := row[1]
		team := row[2]
		if value, ok := writeData[name]; ok {
			if value == 0 {
				continue
			}
			sumlize = append(sumlize, []any{idx, name, team, value, ""})
			idx++
		}

	}

	err = WriteToExcel(absolutePath, sumlize)
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Printf("文件已保存在%s中\n", absolutePath)
	return nil
}

func main() {
	err := Sumlize()
	if err != nil {
		// 等待用户输入回车退出
		fmt.Println("执行失败,按回车退出:\n", err)
		reader := bufio.NewReader(os.Stdin)
		_, _ = reader.ReadString('\n')
	}
}

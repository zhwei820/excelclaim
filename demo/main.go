package main

import (
	"fmt"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/zhwei820/excelclaim/excel"
)

func main() {
	url := "./test.xlsx"
	xlsx := excelize.NewFile()
	headerMap := map[string]string{
		"序号":   "序号",
		"日期":   "日期",
		"出发地":  "出发地",
		"到达地":  "到达地",
		"公务事由": "公务事由",
		"金额":   "金额",
		"备注":   "备注",
	}
	headerKeys := []string{
		"序号", "日期", "出发地", "到达地", "公务事由", "金额", "备注",
	}
	data := []map[string]string{}
	item := map[string]string{
		"序号":   "1",
		"日期":   "2018.06.07",
		"出发地":  "公司",
		"到达地":  "宝安",
		"公务事由": "加班",
		"金额":   "80.00",
		"备注":   "",
	}
	t1 := time.Now()
	for ii := 0; ii < 500000; ii++ {
		data = append(data, item)
	}
	t2 := time.Now()
	fmt.Println("dt", t2.Sub(t1).Seconds())
	t1 = time.Now()

	sheet1(xlsx, data, headerMap, headerKeys)
	// sheet2(xlsx)
	t2 = time.Now()
	fmt.Println("dt2", t2.Sub(t1).Seconds())

	err := xlsx.SaveAs(url)
	if err != nil {
		fmt.Println(err)
		return
	}
}

// sheet1
// headerMap: [{key:xx, value:yy}]
// data: [{key1:val1, key2:val2}]
func sheet1(xlsx *excelize.File, data []map[string]string, headerMap map[string]string, headerKeys []string) {
	sheet := excel.NewSheet(xlsx, "sheet1", len(headerMap), 24)

	header := []string{}
	for _, v := range headerKeys {
		header = append(header, headerMap[v])
	}
	sheet.WriteRow(header...)

	for _, item := range data {
		tmp := []string{}
		for _, v := range headerKeys {
			tmp = append(tmp, item[v])
		}
		sheet.WriteRow(tmp...)
	}
}

func sheet2(xlsx *excelize.File) {
	sheet := excel.NewSheet(xlsx, "加班餐费", 6, 22)
	sheet.SetAllColsWidth(7, 14, 8, 11, 8, 12)

	sheet.WriteRow("加班餐费明细")
	sheet.WriteRow("月份", "2018年06月", "姓名", "wwww", "部门", "研发部")
	sheet.WriteRow("序号", "日期", "事由", "中餐/晚餐", "金额", "备注")
	sheet.WriteRow("1", "2018-06-01", "加班", "晚餐", "20", "21:05")
	sheet.WriteRow("2", "2018-06-02", "加班", "晚餐", "20", "21:05")
	sheet.WriteRow("3", "2018-06-01", "加班", "晚餐", "20", "21:05")
	sheet.WriteRow("4", "2018-06-02", "加班", "晚餐", "20", "21:05")
	sheet.WriteRow("", "", "", "金额合计", "80.00")
	sheet.Apply(excel.NewExcelStyle(10, -1, false))
}

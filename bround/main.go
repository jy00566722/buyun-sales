package bround

import (
	"context"
	"fmt"
	"time"

	excelize "github.com/xuri/excelize/v2"
)

type ProgressInfo struct { //传给前端的进度
	Num  int    `json:"num"`
	Text string `json:"text"`
}

func Main_go(inputFilePath string, outFilePath string, ctx context.Context) error {

	// 创建新的 Excel 文件
	f := excelize.NewFile()
	defer f.Close()

	now := time.Now()
	// 生成文件名
	//filename := fmt.Sprintf("销售统计_%s.xlsx", now.Format("2006-01-02"))
	//outputFilePath := filepath.Join(filepath.Dir(inputFilePath), filename)

	// 调用各个函数，传入 Excel 文件和工作表名
	sheet1Name := now.Format("01.02") + "销量"
	err := getOneDaySale(f, sheet1Name, inputFilePath, ctx)
	if err != nil {
		fmt.Println("sheet1Name:", err)
		return err
	}
	sheet2Name := now.Format("01.02") + "客户"
	err = getCustomerSale(f, sheet2Name, inputFilePath, ctx)
	if err != nil {
		fmt.Println("sheet2Name:", err)
		return err
	}
	sheet3Name := now.Format("01") + "月货号+客户"
	err = getStyleSale(f, sheet3Name, inputFilePath, ctx)
	if err != nil {
		fmt.Println("sheet3Name:", err)
		return err
	}
	sheet4Name := now.Format("01") + "月货号"
	err = CreateStyleReport(f, sheet4Name, inputFilePath, ctx)
	if err != nil {
		fmt.Println("sheet4Name:", err)
		return err
	}
	// 保存文件
	if err := f.SaveAs(outFilePath); err != nil {
		fmt.Println("保存 Excel 文件失败:", err)
		return err
	}
	return nil
}

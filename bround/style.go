package bround

import (
	"context"
	"fmt"
	"sort"
	"strconv"
	"time"

	"github.com/wailsapp/wails/v2/pkg/runtime"
	"github.com/xuri/excelize/v2"
)

type StyleSale struct {
	Date     time.Time
	StyleID  string
	Quantity int
}

type StyleReport struct {
	StyleID    string
	DailySales map[string]int
	TotalSales int
}

func CreateStyleReport(f *excelize.File, sheetName string, inputFilePath string, ctx context.Context) error {
	// 1. 读取 Excel 文件
	styleSales, err := readStyleSalesData(inputFilePath)
	if err != nil {
		//fmt.Println("Error reading Excel file:", err)
		return err
	}
	runtime.EventsEmit(ctx, "progress", ProgressInfo{Num: 90, Text: "统计 货号 销量:正在分析数据"})
	// 2. 处理销售数据
	styleReports, dateRange := analyzeStyleSales(styleSales)
	// 2.1. 按日期排序
	latestDateStr := dateRange[len(dateRange)-1].Format("2006-01-02")
	sortedReports := sortReportsByLatestDateSales(styleReports, latestDateStr)

	// 3. 生成报告
	err = createStyleExcelReport(f, sheetName, sortedReports, dateRange)
	if err != nil {
		//fmt.Println("Error generating Excel report:", err)
		return err
	}

	fmt.Println("Style sales report generated successfully.")
	return nil
}

func readStyleSalesData(filename string) ([]StyleSale, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, fmt.Errorf("error opening file: %w", err)
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("no sheets found in the Excel file")
	}

	rows, err := f.GetRows(sheets[0])
	if err != nil {
		return nil, fmt.Errorf("error reading rows: %w", err)
	}

	var styleSales []StyleSale
	for i, row := range rows {
		if i == 0 { // Skip header row
			continue
		}
		if len(row) < 9 {
			continue // Skip rows with insufficient data
		}

		date, err := time.Parse("1/2/06 15:04", row[0])
		if err != nil {
			return nil, fmt.Errorf("error parsing date in row %d: %w", i+1, err)
		}

		quantity, err := strconv.Atoi(row[8]) // 配货数量 is in the 9th column (index 8)
		if err != nil {
			return nil, fmt.Errorf("error parsing quantity in row %d: %w", i+1, err)
		}

		styleSales = append(styleSales, StyleSale{
			Date:     date,
			StyleID:  row[3], // 货号 is in the 4th column (index 3)
			Quantity: quantity,
		})
	}

	return styleSales, nil
}

func analyzeStyleSales(styleSales []StyleSale) ([]StyleReport, []time.Time) {
	styleMap := make(map[string]*StyleReport)
	dateSet := make(map[string]bool)
	var latestDate time.Time

	// Find the latest date
	for _, sale := range styleSales {
		if sale.Date.After(latestDate) {
			latestDate = sale.Date
		}
	}

	latestDateStr := latestDate.Format("2006-01-02")

	for _, sale := range styleSales {
		dateStr := sale.Date.Format("2006-01-02")
		dateSet[dateStr] = true

		if report, exists := styleMap[sale.StyleID]; exists {
			report.DailySales[dateStr] += sale.Quantity
			report.TotalSales += sale.Quantity
		} else {
			styleMap[sale.StyleID] = &StyleReport{
				StyleID:    sale.StyleID,
				DailySales: map[string]int{dateStr: sale.Quantity},
				TotalSales: sale.Quantity,
			}
		}
	}

	var reports []StyleReport
	for _, report := range styleMap {
		if latestSale, exists := report.DailySales[latestDateStr]; exists && latestSale >= 10 {
			reports = append(reports, *report)
		}
	}

	sort.Slice(reports, func(i, j int) bool {
		return reports[i].TotalSales > reports[j].TotalSales
	})

	var dateRange []time.Time
	for dateStr := range dateSet {
		date, _ := time.Parse("2006-01-02", dateStr)
		dateRange = append(dateRange, date)
	}
	sort.Slice(dateRange, func(i, j int) bool {
		return dateRange[i].Before(dateRange[j])
	})

	return reports, dateRange
}
func sortReportsByLatestDateSales(reports []StyleReport, latestDateStr string) []StyleReport {
	sort.Slice(reports, func(i, j int) bool {
		salesI, existsI := reports[i].DailySales[latestDateStr]
		salesJ, existsJ := reports[j].DailySales[latestDateStr]

		if !existsI {
			return false
		}
		if !existsJ {
			return true
		}
		return salesI > salesJ
	})
	return reports
}
func createStyleExcelReport(f *excelize.File, sheetName string, reports []StyleReport, dateRange []time.Time) error {
	// 创建新的工作表
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("创建工作表失败: %w", err)
	}
	f.SetActiveSheet(index)

	// Set headers
	headers := []string{"货号"}
	for _, date := range dateRange {
		headers = append(headers, date.Format("01/02"))
	}
	headers = append(headers, "总计")

	for col, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		f.SetCellValue(sheetName, cell, header)
	}

	// Write data
	for row, report := range reports {
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", row+2), report.StyleID)

		for col, date := range dateRange {
			dateStr := date.Format("2006-01-02")
			cell, _ := excelize.CoordinatesToCellName(col+2, row+2)
			if quantity, exists := report.DailySales[dateStr]; exists {
				f.SetCellValue(sheetName, cell, quantity)
			}
		}

		totalCell, _ := excelize.CoordinatesToCellName(len(headers), row+2)
		f.SetCellValue(sheetName, totalCell, report.TotalSales)
	}

	// Save file
	//if err := f.SaveAs("style_sales_report.xlsx"); err != nil {
	//	return err
	//}

	return nil
}

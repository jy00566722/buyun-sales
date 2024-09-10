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

type StyleSaleRecord struct {
	Date      time.Time
	ProductID string
	Customer  string
	Quantity  int
}

type StyleCustomerStat struct {
	ProductID  string
	Customer   string
	DailySales map[string]int
	TotalSales int
}

func getStyleSale(f *excelize.File, sheetName string, inputFilePath string, ctx context.Context) error {
	// 1. 读取 Excel 文件
	records, err := readStyleExcelFile(inputFilePath)
	if err != nil {
		//fmt.Println("Error reading Excel file:", err)
		return err
	}
	runtime.EventsEmit(ctx, "progress", ProgressInfo{Num: 70, Text: "统计 客户+货号 销量:正在分析数据"})
	// 2. 计算统计信息
	stats, startDate, endDate, err := calculateStyleStats(records)
	if err != nil {
		//fmt.Println("Error calculating statistics:", err)
		return err
	}

	// 3. 生成新的 Excel 文件
	err = generateStyleExcelReport(f, sheetName, stats, startDate, endDate)
	if err != nil {
		//fmt.Println("Error generating Excel report:", err)
		return err
	}

	fmt.Println("Style sales statistics report generated successfully.")
	return nil
}

func readStyleExcelFile(filename string) ([]StyleSaleRecord, error) {
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

	var records []StyleSaleRecord
	for i, row := range rows {
		if i == 0 { // Skip header row
			continue
		}
		if len(row) < 12 {
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

		records = append(records, StyleSaleRecord{
			Date:      date,
			ProductID: row[3], // 货号 is in the 4th column (index 3)
			Customer:  row[2], // 客户 is in the 3rd column (index 2)
			Quantity:  quantity,
		})
	}

	return records, nil
}
func calculateStyleStats(records []StyleSaleRecord) (map[string][]StyleCustomerStat, time.Time, time.Time, error) {
	if len(records) == 0 {
		return nil, time.Time{}, time.Time{}, fmt.Errorf("no records provided")
	}

	salesMap := make(map[string]map[string]map[string]int)
	var startDate, endDate time.Time

	for _, record := range records {
		dateStr := record.Date.Format("2006-01-02")
		if startDate.IsZero() || record.Date.Before(startDate) {
			startDate = record.Date
		}
		if record.Date.After(endDate) {
			endDate = record.Date
		}

		if _, exists := salesMap[record.ProductID]; !exists {
			salesMap[record.ProductID] = make(map[string]map[string]int)
		}
		if _, exists := salesMap[record.ProductID][record.Customer]; !exists {
			salesMap[record.ProductID][record.Customer] = make(map[string]int)
		}
		salesMap[record.ProductID][record.Customer][dateStr] += record.Quantity
	}

	lastDate := endDate.Format("2006-01-02")

	// 创建一个结构体来存储产品ID和最后一天的销量
	type productLastDaySales struct {
		productID string
		sales     int
	}

	// 计算每个产品最后一天的销量
	var products []productLastDaySales
	for productID, customers := range salesMap {
		lastDaySales := 0
		for _, dailySales := range customers {
			if sales, exists := dailySales[lastDate]; exists {
				lastDaySales += sales
			}
		}
		products = append(products, productLastDaySales{productID, lastDaySales})
	}

	// 按最后一天的销量降序排序产品
	sort.Slice(products, func(i, j int) bool {
		return products[i].sales > products[j].sales
	})

	// 创建排序后的 stats map
	stats := make(map[string][]StyleCustomerStat)

	// 按排序后的顺序处理产品
	for _, product := range products {
		productID := product.productID
		customers := salesMap[productID]

		if product.sales < 10 {
			continue // 跳过最后一天销量小于10的产品
		}

		var customerStats []StyleCustomerStat

		for customer, dailySales := range customers {
			totalSales := 0
			for _, quantity := range dailySales {
				totalSales += quantity
			}

			if totalSales >= 20 {
				customerStats = append(customerStats, StyleCustomerStat{
					ProductID:  productID,
					Customer:   customer,
					DailySales: dailySales,
					TotalSales: totalSales,
				})
			}
		}

		// 对每个货号的客户按总销量降序排序
		sort.Slice(customerStats, func(i, j int) bool {
			return customerStats[i].TotalSales > customerStats[j].TotalSales
		})

		stats[productID] = customerStats
	}

	return stats, startDate, endDate, nil
}

func generateStyleExcelReport(f *excelize.File, sheetName string, salesStats map[string][]StyleCustomerStat, startDate, endDate time.Time) error {
	// 创建新的工作表
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("创建工作表失败: %w", err)
	}
	f.SetActiveSheet(index)

	// Set titles
	titles := []string{"货号", "客户"}
	currentDate := startDate
	for currentDate.Before(endDate) || currentDate.Equal(endDate) {
		titles = append(titles, currentDate.Format("01/02"))
		currentDate = currentDate.AddDate(0, 0, 1)
	}
	titles = append(titles, "总计")

	for i, title := range titles {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(sheetName, cell, title)
	}

	// Write data
	row := 2
	for productID, customerStats := range salesStats {
		startRow := row
		for _, stat := range customerStats {
			f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), productID)
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", row), stat.Customer)

			col := 3
			currentDate := startDate
			for currentDate.Before(endDate) || currentDate.Equal(endDate) {
				dateStr := currentDate.Format("2006-01-02")
				if quantity, exists := stat.DailySales[dateStr]; exists {
					cell, _ := excelize.CoordinatesToCellName(col, row)
					f.SetCellValue(sheetName, cell, quantity)
				}
				col++
				currentDate = currentDate.AddDate(0, 0, 1)
			}

			totalCell, _ := excelize.CoordinatesToCellName(col, row)
			f.SetCellValue(sheetName, totalCell, stat.TotalSales)

			row++
		}
		endRow := row - 1

		// Merge cells for product ID
		if startRow != endRow {
			f.MergeCell(sheetName, fmt.Sprintf("A%d", startRow), fmt.Sprintf("A%d", endRow))
		}
	}

	// Save file
	//filename := fmt.Sprintf("款式销售统计_%s_%s.xlsx", startDate.Format("2006-01-02"), endDate.Format("2006-01-02"))
	//if err := f.SaveAs(filename); err != nil {
	//	return err
	//}

	return nil
}

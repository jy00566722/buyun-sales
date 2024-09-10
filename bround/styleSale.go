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

type ProductStats struct {
	ProductID     string
	LastDaySales  int
	CustomerStats []StyleCustomerStat
}

func calculateStyleStats(records []StyleSaleRecord) ([]ProductStats, time.Time, time.Time, error) {
	if len(records) == 0 {
		return nil, time.Time{}, time.Time{}, fmt.Errorf("no records provided")
	}

	salesMap := make(map[string]map[string]map[string]int)
	var startDate, endDate time.Time

	// ... (populate salesMap, startDate, endDate as before)
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

	var productStats []ProductStats

	for productID, customers := range salesMap {
		lastDaySales := 0
		var customerStats []StyleCustomerStat

		for customer, dailySales := range customers {
			totalSales := 0
			for dateStr, quantity := range dailySales {
				totalSales += quantity
				if dateStr == lastDate {
					lastDaySales += quantity
				}
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

		if lastDaySales >= 10 {
			// Sort customers by total sales in descending order
			sort.Slice(customerStats, func(i, j int) bool {
				return customerStats[i].TotalSales > customerStats[j].TotalSales
			})

			productStats = append(productStats, ProductStats{
				ProductID:     productID,
				LastDaySales:  lastDaySales,
				CustomerStats: customerStats,
			})
		}
	}

	// Sort products by last day sales in descending order
	sort.Slice(productStats, func(i, j int) bool {
		return productStats[i].LastDaySales > productStats[j].LastDaySales
	})

	return productStats, startDate, endDate, nil
}

func generateStyleExcelReport(f *excelize.File, sheetName string, productStats []ProductStats, startDate, endDate time.Time) error {
	// Create new sheet
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("failed to create sheet: %w", err)
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
	for _, product := range productStats {
		startRow := row
		for _, stat := range product.CustomerStats {
			f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), product.ProductID)
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

	return nil
}

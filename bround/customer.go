package bround

import (
	"context"
	"fmt"
	"sort"
	"strconv"
	"time"

	"github.com/wailsapp/wails/v2/pkg/runtime"
	excelize "github.com/xuri/excelize/v2"
)

type CustomerSaleRecord struct {
	Date      time.Time
	ProductID string
	Customer  string
	Quantity  int
}

type ProductCustomerStat struct {
	ProductID string
	Customer  string
	Quantity  int
}

func getCustomerSale(f *excelize.File, sheetName string, inputFilePath string, ctx context.Context) error {
	//runtime.EventsEmit(ctx, "progress", "统计 客户 销量:开始读取文件")
	// 1. 读取 Excel 文件
	records, err := readCustomerExcelFile(inputFilePath)
	if err != nil {
		//fmt.Println("Error reading Excel file:", err)
		return err
	}
	// 2. 找出最近的日期
	latestDate := findLatestDateCustom(records)
	runtime.EventsEmit(ctx, "progress", ProgressInfo{Num: 35, Text: "统计 客户 销量:正在分析数据"})
	// 3. 计算统计信息
	stats, err := calculateCustomerStats(records, latestDate)
	if err != nil {
		//fmt.Println("Error calculating statistics:", err)
		return err
	}
	runtime.EventsEmit(ctx, "progress", ProgressInfo{Num: 40, Text: "统计 客户 销量:正在写入数据"})
	// 4. 生成新的 Excel 文件
	err = generateCustomerExcelReport(f, sheetName, stats)
	if err != nil {
		//fmt.Println("Error generating Excel report:", err)
		return err
	}

	fmt.Println("Customer sales statistics report generated successfully.")
	return nil
}

func findLatestDateCustom(records []CustomerSaleRecord) time.Time {
	var latestDate time.Time
	for _, record := range records {
		if record.Date.After(latestDate) {
			latestDate = record.Date
		}
	}
	fmt.Println("最新日期:", latestDate.Format("2006-01-02"))
	return latestDate
}
func readCustomerExcelFile(filename string) ([]CustomerSaleRecord, error) {
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

	var records []CustomerSaleRecord
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

		records = append(records, CustomerSaleRecord{
			Date:      date,
			ProductID: row[3], // 货号 is in the 4th column (index 3)
			Customer:  row[2], // 客户 is in the 3rd column (index 2)
			Quantity:  quantity,
		})
	}

	return records, nil
}

func calculateCustomerStats(records []CustomerSaleRecord, latestDate time.Time) (map[string][]ProductCustomerStat, error) {
	if len(records) == 0 {
		return nil, fmt.Errorf("no records provided")
	}

	// Normalize the latest date to the start of the day
	latestDate = time.Date(latestDate.Year(), latestDate.Month(), latestDate.Day(), 0, 0, 0, 0, latestDate.Location())

	// Create a map to store sales data for each product and customer
	salesMap := make(map[string]map[string]int)

	// Populate the salesMap
	for _, record := range records {
		if record.Date.Format("2006-01-02") == latestDate.Format("2006-01-02") {
			if _, exists := salesMap[record.ProductID]; !exists {
				salesMap[record.ProductID] = make(map[string]int)
			}
			salesMap[record.ProductID][record.Customer] += record.Quantity
		}
	}

	// Create the final stats map
	stats := make(map[string][]ProductCustomerStat)

	for productID, customers := range salesMap {
		var productTotal int
		var customerStats []ProductCustomerStat

		for customer, quantity := range customers {
			customerStats = append(customerStats, ProductCustomerStat{
				ProductID: productID,
				Customer:  customer,
				Quantity:  quantity,
			})
			productTotal += quantity
		}

		// Only include products with total sales > 10
		if productTotal >= 10 {
			// Sort customer stats by quantity in descending order
			sort.Slice(customerStats, func(i, j int) bool {
				return customerStats[i].Quantity > customerStats[j].Quantity
			})
			stats[productID] = customerStats
		}
	}

	return stats, nil
}

func generateCustomerExcelReport(f *excelize.File, sheetName string, salesStats map[string][]ProductCustomerStat) error {
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("创建工作表失败: %w", err)
	}
	f.SetActiveSheet(index)

	// Set titles
	titles := []string{"货号", "客户", "数量"}
	for i, title := range titles {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(sheetName, cell, title)
	}

	// Calculate total sales for each product
	productTotals := make(map[string]int)
	for productID, customerStats := range salesStats {
		total := 0
		for _, stat := range customerStats {
			total += stat.Quantity
		}
		productTotals[productID] = total
	}

	// Create a slice of product IDs and sort by total sales
	var productIDs []string
	for productID := range salesStats {
		productIDs = append(productIDs, productID)
	}
	sort.Slice(productIDs, func(i, j int) bool {
		return productTotals[productIDs[i]] > productTotals[productIDs[j]]
	})

	// Write data
	row := 2
	for _, productID := range productIDs {
		customerStats := salesStats[productID]
		startRow := row
		for _, stat := range customerStats {
			f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), productID)
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", row), stat.Customer)
			f.SetCellValue(sheetName, fmt.Sprintf("C%d", row), stat.Quantity)
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

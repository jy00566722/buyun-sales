package bround

import (
	"context"
	"fmt"
	"math"
	"sort"
	"strconv"
	"time"

	"github.com/wailsapp/wails/v2/pkg/runtime"
	"github.com/xuri/excelize/v2"
)

type SaleRecord struct {
	Date      time.Time
	ProductID string
	Quantity  int
}

type ProductStat struct {
	ProductID     string
	DailySales    int
	WeeklySales   int
	WeeklyCompare int
}

func getOneDaySale(f *excelize.File, sheetName string, inputFilePath string, ctx context.Context) error {
	runtime.EventsEmit(ctx, "progress", ProgressInfo{Num: 10, Text: "统计日销量:开始读取文件"})
	// 1. 读取 Excel 文件
	records, err := readExcelFile(inputFilePath)
	if err != nil {
		//fmt.Println("Error reading Excel file:", err)
		return err
	}
	// 2. 找出最近的日期
	latestDate := findLatestDate(records)

	// 3. 计算统计信息
	runtime.EventsEmit(ctx, "progress", ProgressInfo{Num: 20, Text: "统计日销量:正在分析数据"})
	stats, err := calculateStats(records, latestDate)
	if err != nil {
		//fmt.Println("Error calculating statistics:", err)
		return err
	}
	// 4. 生成新的 Excel 文件
	err = generateExcelReport(f, sheetName, stats)
	if err != nil {
		//fmt.Println("Error generating Excel report:", err)
		return err
	}
	runtime.EventsEmit(ctx, "progress", ProgressInfo{Num: 25, Text: "统计日销量:正在分析数据"})
	//runtime.EventsEmit(ctx, "progress", "统计日销量:写入数据表完毕")
	fmt.Println("Sales statistics report generated successfully.")
	return nil
}

func readExcelFile(filename string) ([]SaleRecord, error) {
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

	var records []SaleRecord
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

		records = append(records, SaleRecord{
			Date:      date,
			ProductID: row[3], // 货号 is in the 4th column (index 3)
			Quantity:  quantity,
		})
	}

	return records, nil
}

func findLatestDate(records []SaleRecord) time.Time {
	var latestDate time.Time
	for _, record := range records {
		if record.Date.After(latestDate) {
			latestDate = record.Date
		}
	}
	fmt.Println("最新日期:", latestDate.Format("2006-01-02"))
	return latestDate
}
func calculateStats(records []SaleRecord, latestDate time.Time) ([]ProductStat, error) {
	if len(records) == 0 {
		return nil, fmt.Errorf("没有提供记录")
	}

	// 将最新日期调整为当天的结束时间
	latestDate = time.Date(latestDate.Year(), latestDate.Month(), latestDate.Day(), 23, 59, 59, 0, latestDate.Location())

	// 找到实际的最早日期
	earliestActualDate := records[0].Date
	for _, record := range records {
		if record.Date.Before(earliestActualDate) {
			earliestActualDate = record.Date
		}
	}

	// 计算实际天数
	daysDifference := latestDate.Sub(earliestActualDate).Hours() / 24

	// 向上取整，确保包括不完整的第一天
	daysDifference = math.Ceil(daysDifference)

	if daysDifference < 7 {
		return nil, fmt.Errorf("数据不足:需要至少7天的数据,实际数据范围为 %v 到 %v(%.0f天)",
			earliestActualDate.Format("2006-01-02"), latestDate.Format("2006-01-02"), daysDifference)
	}

	fmt.Printf("数据范围：从 %v 到 %v(%.0f天)\n",
		earliestActualDate.Format("2006-01-02"), latestDate.Format("2006-01-02"), daysDifference)

	// Create a map to store sales data for each product
	salesMap := make(map[string]map[string]int)

	// Populate the salesMap
	for _, record := range records {
		// Normalize the record date to the start of the day
		normalizedDate := time.Date(record.Date.Year(), record.Date.Month(), record.Date.Day(), 0, 0, 0, 0, record.Date.Location())
		dateStr := normalizedDate.Format("2006-01-02")
		if _, exists := salesMap[record.ProductID]; !exists {
			salesMap[record.ProductID] = make(map[string]int)
		}
		salesMap[record.ProductID][dateStr] += record.Quantity
	}

	var stats []ProductStat

	for productID, sales := range salesMap {
		latestDateStr := latestDate.Format("2006-01-02")
		dailySales := sales[latestDateStr]
		currentWeekSales := 0
		previousWeekSales := 0

		for i := 0; i < 8; i++ {
			date := latestDate.AddDate(0, 0, -i)
			dateStr := date.Format("2006-01-02")
			if i < 7 {
				currentWeekSales += sales[dateStr]
			}
			if i > 0 && i <= 7 {
				previousWeekSales += sales[dateStr]
			}
		}

		weeklyCompare := currentWeekSales - previousWeekSales
		//这里判断，如果dailySales  currentWeekSales  weeklyCompare 都为0，则直接跳过
		if dailySales == 0 && currentWeekSales == 0 && weeklyCompare == 0 {
			continue
		}
		stats = append(stats, ProductStat{
			ProductID:     productID,
			DailySales:    dailySales,
			WeeklySales:   currentWeekSales,
			WeeklyCompare: weeklyCompare,
		})
	}

	// Sort stats by daily sales in descending order
	sort.Slice(stats, func(i, j int) bool {
		return stats[i].DailySales > stats[j].DailySales
	})

	return stats, nil
}

func generateExcelReport(f *excelize.File, sheetName string, salesStats []ProductStat) error {
	// 创建新的工作表
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("创建工作表失败: %w", err)
	}
	f.SetActiveSheet(index)

	// 设置标题
	titles := []string{"货号", "当日销量", "7日销量", "七日销量对比"}
	for i, title := range titles {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(sheetName, cell, title)
	}

	// 写入数据
	for i, sales := range salesStats {
		row := i + 2
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), sales.ProductID)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", row), sales.DailySales)
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", row), sales.WeeklySales)
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", row), sales.WeeklyCompare)
	}

	return nil
}

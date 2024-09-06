package main

import (
	"context"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	e "ExcelAnalyzer/bround"

	"github.com/wailsapp/wails/v2/pkg/runtime"
)

// App struct
type App struct {
	ctx              context.Context
	analyzedFilePath string
}

// NewApp creates a new App application struct
func NewApp() *App {
	return &App{}
}

// startup is called when the app starts. The context is saved
// so we can call the runtime methods
func (a *App) startup(ctx context.Context) {
	a.ctx = ctx
}

func (a *App) shutdown(ctx context.Context) {
	fmt.Println("Shutting down...")
}

// Greet returns a greeting for the given name
func (a *App) Greet(name string) string {
	return fmt.Sprintf("Hello %s, It's show time!", name)
}

// AnalyzeExcel analyzes the selected Excel file
func (a *App) AnalyzeExcel(filePath string) error {
	fmt.Println("待分析文件:", filePath)
	// 生成新的文件名
	dir := filepath.Dir(filePath)
	fileName := filepath.Base(filePath)
	fileExt := filepath.Ext(fileName)
	fileNameWithoutExt := fileName[:len(fileName)-len(fileExt)]
	newFileName := fileNameWithoutExt + "_分析完成" + fileExt
	a.analyzedFilePath = filepath.Join(dir, newFileName)
	// 这里调用您现有的Excel分析代码
	err := e.Main_go(filePath, a.analyzedFilePath, a.ctx)
	if err != nil {
		//判断err是否以"数据不足"开头
		if strings.HasPrefix(err.Error(), "数据不足") {
			runtime.EventsEmit(a.ctx, "error", err.Error())
		}
		return err
	}
	fmt.Println("分析完成:", filePath)
	return nil
}

// SaveExcel 保存分析后的Excel文件
func (a *App) SaveExcel() error {
	if a.analyzedFilePath == "" {
		return fmt.Errorf("没有可用的分析结果文件")
	}

	// 打开保存文件对话框
	filePath, err := runtime.SaveFileDialog(a.ctx, runtime.SaveDialogOptions{
		Title: "保存分析结果",
		Filters: []runtime.FileFilter{
			{
				DisplayName: "Excel文件 (*.xlsx)",
				Pattern:     "*.xlsx",
			},
		},
		DefaultFilename: filepath.Base(a.analyzedFilePath),
	})

	if err != nil {
		return fmt.Errorf("打开保存对话框失败: %w", err)
	}

	if filePath == "" {
		return nil // 用户取消了保存操作
	}

	// 复制文件到用户选择的位置
	err = copyFile(a.analyzedFilePath, filePath)
	if err != nil {
		return fmt.Errorf("保存文件失败: %w", err)
	}

	fmt.Printf("文件已保存到: %s\n", filePath)
	return nil
}

// OpenFileDialog opens a file dialog and returns the selected file path
func (a *App) OpenFileDialog() (string, error) {
	filePath, err := runtime.OpenFileDialog(a.ctx, runtime.OpenDialogOptions{
		Title: "Select Excel File",
		Filters: []runtime.FileFilter{
			{
				DisplayName: "Excel Files (*.xlsx)",
				Pattern:     "*.xlsx",
			},
		},
	})
	if err != nil {
		return "", err
	}
	return filePath, nil
}

// copyFile 复制文件的辅助函数
func copyFile(src, dst string) error {
	sourceFile, err := os.Open(src)
	if err != nil {
		return err
	}
	defer sourceFile.Close()

	destFile, err := os.Create(dst)
	if err != nil {
		return err
	}
	defer destFile.Close()

	_, err = io.Copy(destFile, sourceFile)
	if err != nil {
		return err
	}

	return nil
}

package main

import (
	"fmt"
	"github.com/fogleman/gg"
	"github.com/nfnt/resize"
	"github.com/xuri/excelize/v2"
	"image/color"
	"os"
	"path/filepath"
	"strings"
	"sync"
)

// 并行处理像素
var wg sync.WaitGroup

// 像素转磅
func weightPixel(pixels float64) float64 {
	pointsPerInch := 72.0
	pixelsPerPoint := 0.125
	return pixels / pixelsPerPoint / pointsPerInch
}

// 像素转磅
func pixelsToPoints(pixels int) float64 {
	pointsPerInch := 72.0
	inchesPerPoint := 1 / 96.0 //不同显示有差异
	return float64(pixels) * inchesPerPoint * pointsPerInch
}
func imageToASCII(imagePath, outputExcel string, scaleFactor float64) {
	// 打开图片并转换为灰度模式
	img, err := gg.LoadImage(imagePath)
	if err != nil {
		fmt.Println("Error loading image:", err)
		return
	}
	// 等比例缩小图片
	// 计算新的尺寸
	width := int(float64(img.Bounds().Dx()) * scaleFactor)
	height := int(float64(img.Bounds().Dy()) * scaleFactor)
	resizedImg := resize.Resize(uint(width), uint(height), img, resize.Lanczos3)
	//resizedImg := img
	context := gg.NewContext(int(resizedImg.Bounds().Dx()), int(resizedImg.Bounds().Dy()))
	context.DrawImage(resizedImg, 0, 0)
	imgWidth, imgHeight := context.Image().Bounds().Dx(), context.Image().Bounds().Dy()
	// 创建 Excel 工作簿
	file := excelize.NewFile()
	sheetName := "Sheet1"
	index, ok := file.NewSheet(sheetName)
	if ok != nil {
		fmt.Println("Error creating sheet:", err)
	}
	file.SetActiveSheet(index)
	// ASCII 字符集，按亮度递增
	asciiChars := "@%#*+=-:."
	// 自定义权重，权重越大的字符在图像中占据的面积越大
	asciiWeights := []float64{13, 8, 5, 3, 2, 1, 0.5, 0.2, 0.1}
	// 并行处理像素
	// 设置所有列的宽度为36
	for i := 0; i < imgWidth; i++ {
		colName := columnIndexToExcelName(i)
		file.SetColWidth(sheetName, colName, colName, weightPixel(36))
	}
	for h := 0; h < imgHeight; h++ {
		// 设置行高
		rowHeight := pixelsToPoints(15)
		file.SetRowHeight(sheetName, h+1, rowHeight)
		wg.Add(1)
		go func(h int) {
			defer wg.Done()
			for w := 0; w < imgWidth; w++ {
				// 处理像素
				// 获取像素的灰度值
				pixelColor := context.Image().At(w, h)
				r, g, b, _ := pixelColor.RGBA()
				pixelValue, _, _ := color.RGBToYCbCr(uint8(r/256), uint8(g/256), uint8(b/256))
				// 将灰度值映射到 ASCII 字符集
				asciiIndex := int((float64(pixelValue) / 255) * float64(len(asciiChars)-1))
				asciiChar := string(asciiChars[asciiIndex])
				// 根据灰度值选择权重
				weight := asciiWeights[asciiIndex]
				// 创建字体颜色
				fontColor := fmt.Sprintf("%02X%02X%02X", int(r/256), int(g/256), int(b/256))
				// 将 ASCII 字符和字体颜色写入 Excel 单元格
				cell := fmt.Sprintf("%s%d", columnIndexToExcelName(w), h+1)
				result := "."
				if weight > 0 {
					result = strings.Repeat(asciiChar, int(weight))
				}
				file.SetCellValue(sheetName, cell, result)
				// 使用 style 进行后续操作
				// 创建一个样式，设置字体颜色
				style, _ := file.NewStyle(&excelize.Style{
					Font: &excelize.Font{
						Color: fontColor, // 字体颜色
					},
					Border: []excelize.Border{}, // 空切片表示无边框
				})
				file.SetCellStyle(sheetName, cell, cell, style)
			}
		}(h)

	}
	// 等待所有 goroutine 完成
	wg.Wait()
	// 保存 Excel 文件
	if err := file.SaveAs(outputExcel); err != nil {
		fmt.Println("Error saving Excel file:", err)
	} else {
		fmt.Println("Excel file saved")
	}
}
func main() {
	//⚠️上传的图片像素太大会运行很慢
	//像素宽大小为不超过26*26=676，excel超列
	// 输入和输出文件夹
	inputFolder := "/Users/mumu/GolandProjects/ascii-to-excel/image"
	outputFolder := "ascii_excel"
	//缩小图片像素大小，太大运行会很慢
	scaleFactor := float64(1.0 / 1.0)
	// 如果输出文件夹不存在，创建它
	if _, err := os.Stat(outputFolder); os.IsNotExist(err) {
		os.Mkdir(outputFolder, os.ModePerm)
	}
	// 遍历输入文件夹中的所有图片文件
	files, err := os.ReadDir(inputFolder)
	if err != nil {
		fmt.Println("Error reading input folder:", err)
		return
	}
	for _, file := range files {
		if file.IsDir() {
			continue
		}
		filename := file.Name()
		if hasImageExtension(filename) {
			imageFile := filepath.Join(inputFolder, filename)
			excelOutput := filepath.Join(outputFolder, strings.TrimSuffix(filename, filepath.Ext(filename))+"_ascii.xlsx")
			// 调用之前的函数将 ASCII 图片保存到 Excel 文件中
			imageToASCII(imageFile, excelOutput, scaleFactor)
		}
	}
	fmt.Println("处理完成！")
}
func hasImageExtension(filename string) bool {
	ext := strings.ToLower(filepath.Ext(filename))
	return ext == ".png" || ext == ".jpg" || ext == ".jpeg"
}

// 生成excel列名，最大676列
func columnIndexToExcelName(index int) string {
	const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	var result string
	if index == 0 {
		return "A"
	}
	for index > 0 {
		index--
		result = string(letters[index%26]) + result
		index /= 26
		if index == 0 {
			break
		}
	}
	return result
}

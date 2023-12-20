package main

import (
	"flag"
	"fmt"
	"github.com/fogleman/gg"
	"github.com/nfnt/resize"
	"github.com/xuri/excelize/v2"
	"image/color"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
)

// 并行处理像素
var wg sync.WaitGroup

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

// 像素转磅
func weightPixel(pixels float64) float64 {
	pointsPerInch := 72.0
	pixelsPerPoint := 0.125
	return pixels / pixelsPerPoint / pointsPerInch
}

// 像素转磅
func pixelsToPoints(pixels float64) float64 {
	pointsPerInch := 72.0
	inchesPerPoint := 1 / 96.0
	return pixels * inchesPerPoint * pointsPerInch
}

// 打开并调整图片大小
func openAndResizeImage(imagePath string, scaleFactor float64) (*gg.Context, error) {
	// 打开图片
	img, err := gg.LoadImage(imagePath)
	if err != nil {
		return nil, fmt.Errorf("error loading image: %v", err)
	}
	// 等比例缩小图片
	width := int(float64(img.Bounds().Dx()) * scaleFactor)
	height := int(float64(img.Bounds().Dy()) * scaleFactor)
	resizedImg := resize.Resize(uint(width), uint(height), img, resize.Lanczos3)
	context := gg.NewContext(int(resizedImg.Bounds().Dx()), int(resizedImg.Bounds().Dy()))
	context.DrawImage(resizedImg, 0, 0)
	return context, nil
}

// 创建excel
func createExcelFile() (*excelize.File, int, error) {
	file := excelize.NewFile()
	index, err := file.NewSheet(sheetName)
	if err != nil {
		return nil, 0, fmt.Errorf("error creating sheet: %v", err)
	}
	file.SetActiveSheet(index)
	return file, index, nil
}

// 设置单元格列宽
func setColumnWidths(file *excelize.File, imgWidth int) {
	for i := 0; i < imgWidth; i++ {
		colName := columnIndexToExcelName(i)
		file.SetColWidth(sheetName, colName, colName, weightPixel(width))
	}
}

// 设置单元格行高
func setRowHeights(file *excelize.File, imgHeight int) {
	for h := 0; h < imgHeight; h++ {
		rowHeight := pixelsToPoints(height)
		file.SetRowHeight(sheetName, h+1, rowHeight)
	}
}

func processPixel(context *gg.Context, file *excelize.File, h, w int) {
	//获取像素点颜色
	pixelColor := context.Image().At(w, h)
	r, g, b, _ := pixelColor.RGBA()
	//获取像素点亮度值
	pixelValue, _, _ := color.RGBToYCbCr(uint8(r/256), uint8(g/256), uint8(b/256))
	asciiIndex := int((float64(pixelValue) / 255) * float64(len(asciiChars)-1))
	asciiChar := string(asciiChars[asciiIndex])
	weight := asciiWeights[asciiIndex]
	fontColor := fmt.Sprintf("%02X%02X%02X", int(r/256), int(g/256), int(b/256))
	result := strings.Repeat(asciiChar, int(weight))
	cell := fmt.Sprintf("%s%d", columnIndexToExcelName(w), h+1)
	file.SetCellValue(sheetName, cell, result)
	style, _ := file.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Color: fontColor, //字体颜色
		},
		Border: []excelize.Border{}, //空切片无边框
	})
	file.SetCellStyle(sheetName, cell, cell, style)
}

func imageToASCII(imagePath, outputExcel string) {
	context, err := openAndResizeImage(imagePath, scaleFactor)
	if err != nil {
		fmt.Println(err)
		return
	}
	imgWidth, imgHeight := context.Image().Bounds().Dx(), context.Image().Bounds().Dy()
	file, _, err := createExcelFile()
	if err != nil {
		fmt.Println(err)
		return
	}
	setColumnWidths(file, imgWidth)
	setRowHeights(file, imgHeight)
	for h := 0; h < imgHeight; h++ {
		wg.Add(1)
		go func(h int) {
			defer wg.Done()
			for w := 0; w < imgWidth; w++ {
				processPixel(context, file, h, w)
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

// 解析 ASCII 权重字符串为 float64 数组
func parseWeights(weightsStr string) []float64 {
	var weights []float64
	// 按逗号分割字符串
	weightStrs := strings.Split(weightsStr, ",")
	for _, wStr := range weightStrs {
		// 解析为 float64
		w, err := strconv.ParseFloat(wStr, 64)
		if err != nil {
			fmt.Println("Error parsing ASCII weight:", err)
			continue
		}
		weights = append(weights, w)
	}
	return weights
}

var (
	// 输入和输出文件夹
	inputFolder  string
	outputFolder string
	// sheet名称
	sheetName string
	// 缩小倍数
	scaleFactor     float64
	asciiWeightsStr string
	// 自定义权重，权重越大的字符在图像中占据的面积越大
	asciiWeights []float64
	// ASCII 字符集，按亮度递增
	asciiChars string
	//单元格列宽,单位像素
	width float64
	//单元格行高，单位像素
	height float64
)

func init() {
	//定义命令行参数
	flag.StringVar(&asciiWeightsStr, "weights", "13,8,5,3,2,1,0.5,0.2,0.1", "ascii字符权重数组")
	flag.StringVar(&asciiChars, "asciiChars", "@%#*+=-:.", "ascii字符集")
	flag.StringVar(&inputFolder, "input", "/Users/mumu/GolandProjects/ascii-to-excel/image", "输入文件夹地址")
	flag.StringVar(&outputFolder, "output", "ascii_excel", "输出文件夹")
	flag.StringVar(&sheetName, "sheet", "Sheet1", "sheet名称")
	flag.Float64Var(&scaleFactor, "scale", 1.0/3, "重新调整后的图片大小")
	flag.Float64Var(&width, "width", 36, "单元格列宽")
	flag.Float64Var(&height, "height", 15, "单元格行高")
}

func main() {
	//⚠️上传的图片像素太大会运行很慢 像素宽大小为不超过26*26=676，excel超列
	// 解析命令行参数
	flag.Parse()
	// 将字符串解析为 float64 数组
	asciiWeights = parseWeights(asciiWeightsStr)
	// 验证输入文件夹是否存在
	if _, err := os.Stat(inputFolder); os.IsNotExist(err) {
		fmt.Println("输入文件夹不存在", err)
		return
	}
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
			imageToASCII(imageFile, excelOutput)
		}
	}
	fmt.Println("处理完成！")
}

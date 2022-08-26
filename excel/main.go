package main

import (
	"fmt"
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/widget"
	"github.com/flopp/go-findfont"
	"github.com/xuri/excelize/v2"
	"os"
	"strconv"
	"strings"
)

// 初始化 中文字体的支持
func init() {
	fontPaths := findfont.List()
	for _, path := range fontPaths {
		fmt.Println(path)
		//楷体:simkai.ttf
		//黑体:simhei.ttf
		if strings.Contains(path, "simhei.ttf") {
			os.Setenv("FYNE_FONT", path)
			break
		}
	}
	fmt.Println("=============")
}

/**
xlsx  只支持Sheet1表头的读取 和 复制
*/

func main() {
	MyApp := app.New()
	c := MyApp.NewWindow("鸡东一中分班小程序")

	xlsxs(c)
	c.Resize(fyne.NewSize(600, 600)) // fyne 窗口的大小
	c.ShowAndRun()                   // 运行窗口
}

func xlsxs(w fyne.Window) {

	// excelize 解析处理的 二维数组
	var str [][]string
	var str1 [][]string
	var stus [][]*stu
	//标签控件
	label1 := widget.NewLabel("")
	// 选择文件对话框
	openFile := dialog.NewFileOpen(func(readCloser fyne.URIReadCloser, err error) {
		if readCloser == nil {
			return
		}
		// 设置标签控件 等于打开的文件地址
		label1.SetText(readCloser.URI().Path())
		// 传入 文件地址 返回 二位数组数据
		str = readers(readCloser.URI().Path())
	}, w)
	// 设置 选择文件对话框只支持查看 .xlsx 文件
	openFile.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"})) //只展示go文件

	// 按钮  打开 选择文件对话框
	button1 := widget.NewButton("choose Xinjian students' list.xlsx", func() {
		openFile.Show()
	})

	//标签控件
	label11 := widget.NewLabel("")
	// 选择文件对话框
	openFile1 := dialog.NewFileOpen(func(readCloser fyne.URIReadCloser, err error) {
		if readCloser == nil {
			return
		}
		// 设置标签控件 等于打开的文件地址
		label11.SetText(readCloser.URI().Path())
		// 传入 文件地址 返回 二位数组数据
		str1 = readers(readCloser.URI().Path())
	}, w)
	// 设置 选择文件对话框只支持查看 .xlsx 文件
	openFile1.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"})) //只展示go文件

	// 按钮  打开 选择文件对话框
	button11 := widget.NewButton("choose Hongshaonian students' list.xlsx", func() {
		openFile1.Show()
	})

	//标签控件
	label111 := widget.NewLabel("")
	// 选择文件对话框
	/*
		openFile11 := dialog.NewFileOpen(func(readCloser fyne.URIReadCloser, err error) {
			if readCloser == nil {
				return
			}
			// 设置标签控件 等于打开的文件地址
			label111.SetText(readCloser.URI().Path())
			// 传入 文件地址 返回 二位数组数据
			str = readers(readCloser.URI().Path())
		}, w)
	*/

	// 设置 选择文件对话框只支持查看 .xlsx 文件
	//openFile11.SetFilter(storage.NewExtensionFileFilter([]string{".xlsx"})) //只展示go文件

	// 按钮  打开 选择文件对话框
	button111 := widget.NewButton("start dividing", func() {
		stus = run()
		label111.SetText("dividing success,please save the result")
	})

	// 控件
	label2 := widget.NewLabel("")
	// 保存文件的 对话框
	saveFile := dialog.NewFileSave(func(writeCloser fyne.URIWriteCloser, err error) {
		if writeCloser == nil {
			return
		}
		// 标签控件的title 设置
		label2.SetText(writeCloser.URI().Path())
		// 传入 二维数组 和 保存文件地址  写入xlsx
		//setxlsx(str, writeCloser.URI().Path())
		saveResult(stus, "一中新生分班名单(带分数)")
		saveResult1(stus, writeCloser.URI().Path())
	}, w)

	// 显示 保存文件对话框
	button2 := widget.NewButton("choose a empty xlsx to save the result of dividing", func() {
		saveFile.Show()
	})

	// 设置窗口  控件
	w.SetContent(container.NewVBox(container.NewHBox(label1, button1), container.NewHBox(label11, button11), container.NewHBox(label111, button111), container.NewHBox(label2, button2)))
}

// 读取xlsx
func readers(path string) [][]string {
	f, err := excelize.OpenFile(path) // 打开文件
	if err != nil {
		fmt.Println(err)
		return nil
	}

	// 获取 Sheet1 上所有单元格
	rows, err := f.GetRows("Sheet1")
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
		}
		fmt.Println()
	}
	return rows
}

// 设置 xlsx表格的value的
func setxlsx(str [][]string, path string) {
	f := excelize.NewFile()
	// 创建一个工作表
	index := f.NewSheet("Sheet1")

	// 单元格标头
	var CellType = []string{
		"A",
		"B",
		"C",
		"D",
		"E",
		"F",
		"G",
		"H",
		"I",
		"J",
		"K",
		"L",
		"M",
		"N",
		"O",
		"P",
		"Q",
		"R",
		"S",
		"T",
		"U",
		"V",
		"W",
		"X",
		"Y",
		"Z",
	}

	// 将二维数组转为map
	con := map[string]string{}
	for i := 0; i < len(str); i++ {
		for j := 0; j < len(str[i]); j++ {
			con[CellType[j]+strconv.Itoa(i+1)] = str[i][j]
		}
	}
	fmt.Println(con)
	// 设置单元格的值

	// 设置 单元格内容
	for k, v := range con {
		f.SetCellValue("Sheet1", k, v)
	}
	// 设置工作簿的默认工作表
	f.SetActiveSheet(index)
	// 根据指定路径保存文件
	if err := f.SaveAs(path); err != nil {
		fmt.Println(err)
	}
}

func saveResult(stus [][]*stu, path string) {
	f2 := excelize.NewFile()

	f2.NewSheet("class 1")
	f2.NewSheet("class 2")
	f2.NewSheet("class 3")
	f2.NewSheet("class 4")
	f2.NewSheet("class 5")
	f2.NewSheet("class 6")
	f2.DeleteSheet("Sheet1")
	// 写入标题
	titleSlice := []interface{}{"姓名", "性别", "小学学校", "数学", "语文", "总分"}
	//titleSlice := []interface{}{"姓名", "性别", "小学学校"}
	_ = f2.SetSheetRow("class 1", "A1", &titleSlice)
	data0 := []interface{}{}
	for i := 0; i < len(stus[0]); i++ {
		data0 = append(data0, []interface{}{(stus[0][i].name), (stus[0][i].sex), (stus[0][i].school),
			(stus[0][i].score_math), (stus[0][i].score_chinese), (stus[0][i].score)})
	}
	// 遍历写入数据
	for key, datum := range data0 {
		axis := fmt.Sprintf("A%d", key+2)
		// 利用断言，转换类型
		tmp, _ := datum.([]interface{})
		_ = f2.SetSheetRow("class 1", axis, &tmp)
	}

	// 写入标题
	_ = f2.SetSheetRow("class 2", "A1", &titleSlice)
	data1 := []interface{}{}
	for i := 0; i < len(stus[1]); i++ {
		data1 = append(data1, []interface{}{(stus[1][i].name), (stus[1][i].sex), (stus[1][i].school),
			(stus[1][i].score_math), (stus[1][i].score_chinese), (stus[1][i].score)})
	}
	// 遍历写入数据
	for key, datum := range data1 {
		axis := fmt.Sprintf("A%d", key+2)
		// 利用断言，转换类型
		tmp, _ := datum.([]interface{})
		_ = f2.SetSheetRow("class 2", axis, &tmp)
	}

	_ = f2.SetSheetRow("class 3", "A1", &titleSlice)
	data2 := []interface{}{}
	for i := 0; i < len(stus[2]); i++ {
		data2 = append(data2, []interface{}{(stus[2][i].name), (stus[2][i].sex), (stus[2][i].school),
			(stus[2][i].score_math), (stus[2][i].score_chinese), (stus[2][i].score)})
	}
	// 遍历写入数据
	for key, datum := range data2 {
		axis := fmt.Sprintf("A%d", key+2)
		// 利用断言，转换类型
		tmp, _ := datum.([]interface{})
		_ = f2.SetSheetRow("class 3", axis, &tmp)
	}

	_ = f2.SetSheetRow("class 4", "A1", &titleSlice)
	data3 := []interface{}{}
	for i := 0; i < len(stus[3]); i++ {
		data3 = append(data3, []interface{}{(stus[3][i].name), (stus[3][i].sex), (stus[3][i].school),
			(stus[3][i].score_math), (stus[3][i].score_chinese), (stus[3][i].score)})
	}
	// 遍历写入数据
	for key, datum := range data3 {
		axis := fmt.Sprintf("A%d", key+2)
		// 利用断言，转换类型
		tmp, _ := datum.([]interface{})
		_ = f2.SetSheetRow("class 4", axis, &tmp)
	}

	_ = f2.SetSheetRow("class 5", "A1", &titleSlice)
	data4 := []interface{}{}
	for i := 0; i < len(stus[4]); i++ {
		data4 = append(data4, []interface{}{(stus[4][i].name), (stus[4][i].sex), (stus[4][i].school),
			(stus[4][i].score_math), (stus[4][i].score_chinese), (stus[4][i].score)})
	}
	// 遍历写入数据
	for key, datum := range data4 {
		axis := fmt.Sprintf("A%d", key+2)
		// 利用断言，转换类型
		tmp, _ := datum.([]interface{})
		_ = f2.SetSheetRow("class 5", axis, &tmp)
	}

	_ = f2.SetSheetRow("class 6", "A1", &titleSlice)
	data5 := []interface{}{}
	for i := 0; i < len(stus[5]); i++ {
		data5 = append(data5, []interface{}{(stus[5][i].name), (stus[5][i].sex), (stus[5][i].school),
			(stus[5][i].score_math), (stus[5][i].score_chinese), (stus[5][i].score)})
	}
	// 遍历写入数据
	for key, datum := range data5 {
		axis := fmt.Sprintf("A%d", key+2)
		// 利用断言，转换类型
		tmp, _ := datum.([]interface{})
		_ = f2.SetSheetRow("class 6", axis, &tmp)
	}
	// 设置工作簿的默认工作表
	f2.SetActiveSheet(0)
	// 根据指定路径保存文件
	if err := f2.SaveAs(path); err != nil {
		fmt.Println(err)
	}
	fmt.Println("执行成功")
}
func saveResult1(stus [][]*stu, path string) {
	f2 := excelize.NewFile()

	f2.NewSheet("class 1")
	f2.NewSheet("class 2")
	f2.NewSheet("class 3")
	f2.NewSheet("class 4")
	f2.NewSheet("class 5")
	f2.NewSheet("class 6")
	f2.DeleteSheet("Sheet1")
	// 写入标题
	titleSlice := []interface{}{"姓名", "性别", "小学学校"}
	//titleSlice := []interface{}{"姓名", "性别", "小学学校"}
	_ = f2.SetSheetRow("class 1", "A1", &titleSlice)
	data0 := []interface{}{}
	for i := 0; i < len(stus[0]); i++ {
		data0 = append(data0, []interface{}{(stus[0][i].name), (stus[0][i].sex), (stus[0][i].school)})
	}
	// 遍历写入数据
	for key, datum := range data0 {
		axis := fmt.Sprintf("A%d", key+2)
		// 利用断言，转换类型
		tmp, _ := datum.([]interface{})
		_ = f2.SetSheetRow("class 1", axis, &tmp)
	}

	// 写入标题
	_ = f2.SetSheetRow("class 2", "A1", &titleSlice)
	data1 := []interface{}{}
	for i := 0; i < len(stus[1]); i++ {
		data1 = append(data1, []interface{}{(stus[1][i].name), (stus[1][i].sex), (stus[1][i].school)})
	}
	// 遍历写入数据
	for key, datum := range data1 {
		axis := fmt.Sprintf("A%d", key+2)
		// 利用断言，转换类型
		tmp, _ := datum.([]interface{})
		_ = f2.SetSheetRow("class 2", axis, &tmp)
	}

	_ = f2.SetSheetRow("class 3", "A1", &titleSlice)
	data2 := []interface{}{}
	for i := 0; i < len(stus[2]); i++ {
		data2 = append(data2, []interface{}{(stus[2][i].name), (stus[2][i].sex), (stus[2][i].school)})
	}
	// 遍历写入数据
	for key, datum := range data2 {
		axis := fmt.Sprintf("A%d", key+2)
		// 利用断言，转换类型
		tmp, _ := datum.([]interface{})
		_ = f2.SetSheetRow("class 3", axis, &tmp)
	}

	_ = f2.SetSheetRow("class 4", "A1", &titleSlice)
	data3 := []interface{}{}
	for i := 0; i < len(stus[3]); i++ {
		data3 = append(data3, []interface{}{(stus[3][i].name), (stus[3][i].sex), (stus[3][i].school)})
	}
	// 遍历写入数据
	for key, datum := range data3 {
		axis := fmt.Sprintf("A%d", key+2)
		// 利用断言，转换类型
		tmp, _ := datum.([]interface{})
		_ = f2.SetSheetRow("class 4", axis, &tmp)
	}

	_ = f2.SetSheetRow("class 5", "A1", &titleSlice)
	data4 := []interface{}{}
	for i := 0; i < len(stus[4]); i++ {
		data4 = append(data4, []interface{}{(stus[4][i].name), (stus[4][i].sex), (stus[4][i].school)})
	}
	// 遍历写入数据
	for key, datum := range data4 {
		axis := fmt.Sprintf("A%d", key+2)
		// 利用断言，转换类型
		tmp, _ := datum.([]interface{})
		_ = f2.SetSheetRow("class 5", axis, &tmp)
	}

	_ = f2.SetSheetRow("class 6", "A1", &titleSlice)
	data5 := []interface{}{}
	for i := 0; i < len(stus[5]); i++ {
		data5 = append(data5, []interface{}{(stus[5][i].name), (stus[5][i].sex), (stus[5][i].school)})
	}
	// 遍历写入数据
	for key, datum := range data5 {
		axis := fmt.Sprintf("A%d", key+2)
		// 利用断言，转换类型
		tmp, _ := datum.([]interface{})
		_ = f2.SetSheetRow("class 6", axis, &tmp)
	}
	// 设置工作簿的默认工作表
	f2.SetActiveSheet(0)
	// 根据指定路径保存文件
	if err := f2.SaveAs(path); err != nil {
		fmt.Println(err)
	}
	fmt.Println("执行成功")
}

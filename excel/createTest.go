package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"math/rand"
	"sort"
	"strconv"
	"time"
)

type stu struct {
	name     string
	score    int64
	id       int64
	isDelete bool
}

/*const num1 = 50
const num2 = 50
const num3 = 50
const num4 = 50
const num5 = 50
const num6 = 50*/
const classesCount = 6

func main() {
	//学生信息表
	students := make([]stu, 0, 400)
	//stui:i班最终学生分配
	stus := make([][]*stu, 6)
	for i := range stus {
		stus[i] = make([]*stu, 0, 50)
	}
	//特殊学生map
	smap := make(map[string]string)
	//读取输入特殊学生信息表格

	f1, err1 := excelize.OpenFile("special.xlsx")
	if err1 != nil {
		fmt.Println(err1)
		return
	}
	defer func() {
		if err := f1.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	rows1, err1 := f1.GetRows("Sheet1")
	for i, row := range rows1 {
		if i == 0 {
			continue
		}
		smap[row[0]] = row[1]
	}

	//读取输入全量学生信息表格
	f, err := excelize.OpenFile("student.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// 获取 Sheet1 上所有单元格
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	for i, row := range rows {
		if i == 0 {
			continue
		}

		sco, err := strconv.ParseInt(row[1], 10, 64)
		if err != nil {
			return
		}
		_, ok := smap[row[0]]
		if ok {
			stus[5] = append(stus[5], &stu{
				name:     row[0],
				score:    sco,
				id:       int64(i),
				isDelete: true,
			})
		} else {
			students = append(students, stu{
				name:     row[0],
				score:    sco,
				id:       int64(i),
				isDelete: false,
			})
		}
	}
	amount := len(students) + len(smap)
	//num[i]表示i班的待分配人数
	nums := []int64{0, 0, 0, 0, 0, 0}

	if amount%6 == 0 {
		for i := 0; i < 6; i++ {
			nums[i] = int64(amount / classesCount)
		}
	} else {
		for i := 0; i < 6; i++ {
			nums[i] = int64(amount / classesCount)
		}
		/*余下的学生放到后面处理，随机分给五个班
		//把余下的学生分给前五个班
		for i := 1; i <= amount%6; i++ {
			nums[i-1]++
		}
		*/
	}
	nums[5] -= int64(len(smap))
	//对students的学生分数进行排序，初步方案:先选出6班，其余5个班平均分配
	sort.SliceStable(students, func(i, j int) bool {
		return students[i].score > students[j].score
	})
	//先分六班的
	if nums[5] > 0 {
		step := int64(len(students)) / nums[5]
		start := int64(0)
		tail := step
		for tail <= int64(len(students)) {
			s := students[start:tail]
			rand.Seed(time.Now().Unix())
			index := rand.Intn(int(step))
			stus[5] = append(stus[5], &s[index])
			s[index].isDelete = true
			start = tail
			tail += step
			nums[5]--
		}
		if nums[5] != 0 {
			fmt.Println("error,nums5 not zero")
			return
		}
	}
	//fstu：其他五班剩余待分配学生
	fstu := make([]*stu, 0)
	for index, _ := range students {
		if !students[index].isDelete {
			fstu = append(fstu, &students[index])
		}
	}
	rand.Seed(time.Now().Unix())
	ClassNo := []int{0, 1, 2, 3, 4}
	start := 0
	step := 5
	tail := step
	//先处理可以整除的部分
	for tail <= len(fstu) {
		//把0-4下标洗牌
		rand.Shuffle(len(ClassNo), func(i, j int) {
			ClassNo[i], ClassNo[j] = ClassNo[j], ClassNo[i]
		})
		s := fstu[start:tail]
		for j := 0; j < 5; j++ {
			ver := ClassNo[j]
			/*if nums[ver] > 0 {
				nums[ver]--
			} else {
				fmt.Println("error,num<0")
				return
			}*/
			stus[ver] = append(stus[ver], s[j])
		}
		start = tail
		tail = start + step
	}

	//处理剩余不足5人的学生，如果有的话
	if tail-step < len(fstu) {
		rand.Shuffle(len(ClassNo), func(i, j int) {
			ClassNo[i], ClassNo[j] = ClassNo[j], ClassNo[i]
		})
		s := fstu[tail-step:]
		for j := 0; j < len(s); j++ {
			ver := ClassNo[j]
			stus[ver] = append(stus[ver], s[j])
		}
	}
	if len(stus[0])+len(stus[1])+len(stus[2])+len(stus[3])+len(stus[4])+len(stus[5]) != amount {
		fmt.Println("error,numbers loss")
	}

	f2 := excelize.NewFile()

	index1 := f2.NewSheet("class 1")
	f2.NewSheet("class 2")
	f2.NewSheet("class 3")
	f2.NewSheet("class 4")
	f2.NewSheet("class 5")
	f2.NewSheet("class 6")
	f2.DeleteSheet("Sheet1")
	// 写入标题
	titleSlice := []interface{}{"姓名", "分数"}
	_ = f2.SetSheetRow("class 1", "A1", &titleSlice)

	data0 := []interface{}{}
	for i := 0; i < len(stus[0]); i++ {
		data0 = append(data0, []interface{}{(stus[0][i].name), (stus[0][i].score)})
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
		data1 = append(data1, []interface{}{(stus[1][i].name), (stus[1][i].score)})
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
		data2 = append(data2, []interface{}{(stus[2][i].name), (stus[2][i].score)})
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
		data3 = append(data3, []interface{}{(stus[3][i].name), (stus[3][i].score)})
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
		data4 = append(data4, []interface{}{(stus[4][i].name), (stus[4][i].score)})
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
		data5 = append(data5, []interface{}{(stus[5][i].name), (stus[5][i].score)})
	}
	// 遍历写入数据
	for key, datum := range data5 {
		axis := fmt.Sprintf("A%d", key+2)
		// 利用断言，转换类型
		tmp, _ := datum.([]interface{})
		_ = f2.SetSheetRow("class 6", axis, &tmp)
	}
	// 设置工作簿的默认工作表
	f2.SetActiveSheet(index1)
	// 根据指定路径保存文件
	if err := f2.SaveAs("list.xlsx"); err != nil {
		fmt.Println(err)
	}
	fmt.Println("执行成功")
}

package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"math/rand"
	"sort"
	"strconv"
	"strings"
)

type stu struct {
	name          string
	score         float64
	score_math    float64
	score_chinese float64
	id            int64
	sex           string
	school        string
	isDelete      bool
}

/*const num1 = 50
const num2 = 50
const num3 = 50
const num4 = 50
const num5 = 50
const num6 = 50*/
const classesCount = 6

var req = -1

func run() [][]*stu {
	fmt.Printf("触发func倒计数：%d", req)
	req--
	//学生信息表
	students := make([]stu, 0, 400)
	//stui:i班最终学生分配
	stus := make([][]*stu, 6)
	for i := range stus {
		stus[i] = make([]*stu, 0, 50)
	}
	nums := []int64{37, 38, 38, 38, 38, 38}
	//特殊学生map
	smap := make(map[string]string)
	smap1 := make(map[string]string)
	smap2 := make(map[string]string)

	if req < 0 {
		//读取输入特殊学生信息表格1
		f1, err1 := excelize.OpenFile("special.xlsx")
		if err1 != nil {
			fmt.Println(err1)
			return nil
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
			smap[row[0]] = "_"
		}
		//读取输入特殊学生信息表格2
		f2, err1 := excelize.OpenFile("special1.xlsx")
		if err1 != nil {
			fmt.Println(err1)
			return nil
		}
		defer func() {
			if err := f2.Close(); err != nil {
				fmt.Println(err)
			}
		}()
		rows2, err1 := f2.GetRows("Sheet1")
		for i, row := range rows2 {
			if i == 0 {
				continue
			}
			smap1[row[0]] = "_"
		}
		//读取输入特殊学生信息表格3
		f3, err1 := excelize.OpenFile("special2.xlsx")
		if err1 != nil {
			fmt.Println(err1)
			return nil
		}
		defer func() {
			if err := f3.Close(); err != nil {
				fmt.Println(err)
			}
		}()
		rows3, err1 := f3.GetRows("Sheet1")
		for i, row := range rows3 {
			if i == 0 {
				continue
			}
			smap2[row[0]] = "_"
		}
		fmt.Println("execute special")
	}
	//读取输入全量学生信息表格;两个表
	f, err := excelize.OpenFile("hsn_student.xlsx")
	if err != nil {
		fmt.Println(err)
		return nil
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// 获取 Sheet1 上所有单元格
	rows, err := f.GetRows("五年级学生成绩")
	if err != nil {
		fmt.Println(err)
		return nil
	}
	for i, row := range rows {
		if i == 0 {
			continue
		}
		//sco, err := strconv.ParseInt(row[1], 10, 64)
		sco_math, err := strconv.ParseFloat(row[3], 64)
		sco_chinese, err := strconv.ParseFloat(row[4], 64)
		sco := sco_chinese + sco_math
		if err != nil {
			return nil
		}
		_, ok := smap[row[0]]
		if ok {
			stus[5] = append(stus[5], &stu{
				name:          row[0],
				sex:           row[1],
				score_math:    sco_math,
				score_chinese: sco_chinese,
				score:         sco,
				school:        "红少年",
				id:            int64(i),
				isDelete:      true,
			})
			nums[5]--
		} else {
			if _, ok = smap1[row[0]]; ok {
				stus[1] = append(stus[1], &stu{
					name:          row[0],
					sex:           row[1],
					score_math:    sco_math,
					score_chinese: sco_chinese,
					score:         sco,
					school:        "红少年",
					id:            int64(i),
					isDelete:      true,
				})
				nums[1]--
			} else {
				if _, ok = smap2[row[0]]; ok && !(strings.Compare(row[0], "金子涵") == 0) {
					stus[2] = append(stus[2], &stu{
						name:          row[0],
						sex:           row[1],
						score_math:    sco_math,
						score_chinese: sco_chinese,
						score:         sco,
						school:        "红少年",
						id:            int64(i),
						isDelete:      true,
					})
					nums[2]--
				} else {
					students = append(students, stu{
						name:          row[0],
						sex:           row[1],
						score_math:    sco_math,
						score_chinese: sco_chinese,
						score:         sco,
						school:        "红少年",
						id:            int64(i),
						isDelete:      false,
					})
				}
			}
		}
	}
	//--------读第二张新建的表
	d, err := excelize.OpenFile("xj_student.xlsx")
	if err != nil {
		fmt.Println(err)
		return nil
	}
	defer func() {
		if err := d.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// 获取 Sheet1 上所有单元格
	drows, err := d.GetRows("一中")
	if err != nil {
		fmt.Println(err)
		return nil
	}
	for i, row := range drows {
		if i == 0 {
			continue
		}
		//sco, err := strconv.ParseInt(row[1], 10, 64)
		sco_math, err := strconv.ParseFloat(row[2], 64)
		sco_chinese, err := strconv.ParseFloat(row[3], 64)
		sco := sco_chinese + sco_math
		if err != nil {
			return nil
		}
		if i == 9 {
			fmt.Println(i)
		}
		_, ok := smap[row[1]]
		if ok {
			stus[5] = append(stus[5], &stu{
				name:          row[1],
				sex:           row[6],
				score_math:    sco_math,
				score_chinese: sco_chinese,
				score:         sco,
				school:        "新建",
				id:            int64(i),
				isDelete:      true,
			})
			nums[5]--
		} else if _, ok = smap1[row[1]]; ok {
			stus[1] = append(stus[1], &stu{
				name:          row[1],
				sex:           row[6],
				score_math:    sco_math,
				score_chinese: sco_chinese,
				score:         sco,
				school:        "新建",
				id:            int64(i),
				isDelete:      true,
			})
			nums[1]--
		} else if _, ok = smap2[row[1]]; ok {
			stus[2] = append(stus[2], &stu{
				name:          row[1],
				sex:           row[6],
				score_math:    sco_math,
				score_chinese: sco_chinese,
				score:         sco,
				school:        "新建",
				id:            int64(i),
				isDelete:      true,
			})
			nums[2]--
		} else {
			str := "新建"
			if strings.Compare("孙健航", row[1]) == 0 {
				str = "外地入学"
			}
			students = append(students, stu{
				name:          row[1],
				sex:           row[6],
				score_math:    sco_math,
				score_chinese: sco_chinese,
				score:         sco,
				school:        str,
				id:            int64(i),
				isDelete:      false,
			})
		}

	}

	//amount := len(students) + len(smap)+len(smap1)
	//
	//if amount%6 == 0 {
	//	for i := 0; i < 6; i++ {
	//		nums[i] = int64(amount / classesCount)
	//	}
	//} else {
	//	for i := 0; i < 6; i++ {
	//		nums[i] = int64(amount / classesCount)
	//	}
	//	/*余下的学生放到后面处理，随机分给五个班
	//	//把余下的学生分给前五个班
	//	for i := 1; i <= amount%6; i++ {
	//		nums[i-1]++
	//	}
	//	*/
	//}
	//nums[5] -= int64(len(smap))
	//对students的学生分数进行排序，初步方案:先选出6班，其余5个班平均分配
	sort.SliceStable(students, func(i, j int) bool {
		return students[i].score > students[j].score
	})

	//先分六班的
	//需要给六班分一个197+的高分
	//topNo := []int{0, 1, 2, 3}
	//rand.Seed(time.Now().Unix())
	//for {
	//	rand.Shuffle(len(topNo), func(i, j int) {
	//		topNo[i], topNo[j] = topNo[j], topNo[i]
	//	})
	//	one := topNo[0]
	//	if strings.Compare(students[one].name, "杨津福") == 0 {
	//		continue
	//	} else {
	//		stus[5] = append(stus[5], &students[one])
	//		students[one].isDelete = true
	//		nums[5]--
	//		break
	//	}
	//}

	//把后50名随机选人分给六班
	//if nums[5] > 0 {
	//	step := int64(len(students)-150) / nums[5]
	//	start := int64(150)
	//	tail := start + step
	//	for tail <= int64(len(students)) {
	//		s := students[start:tail]
	//		rand.Seed(time.Now().Unix())
	//		index := rand.Intn(int(step))
	//		stus[5] = append(stus[5], &s[index])
	//		s[index].isDelete = true
	//		start = tail
	//		tail += step
	//		nums[5]--
	//	}
	//	//if nums[5] != 0 {
	//	//	fmt.Println("error,nums5 not zero")
	//	//	return
	//	//}
	//}

	//fstu：其他五班剩余待分配学生
	fstu := make([]*stu, 0)
	for index, _ := range students {
		if !students[index].isDelete {
			fstu = append(fstu, &students[index])
		}
	}

	//ClassNo := []int{0, 1, 2, 3, 4, 5}
	//ClassNo1 := []int{0, 1, 2, 3, 4, 5}
	//flag := false
	start := 0
	step := 6
	tail := step
	count := 6
	//先处理可以整除的部分
	for tail <= len(fstu) && count > 0 {
		ClassNo := make([]int, 0)
		for i := range nums {
			if nums[i] == 1 {
				count--
			}
			if nums[i] > 0 {
				ClassNo = append(ClassNo, i)
			}
		}
		rand.Shuffle(len(ClassNo), func(i, j int) {
			ClassNo[i], ClassNo[j] = ClassNo[j], ClassNo[i]
		})
		s := fstu[start:tail]
		for j := 0; j < len(ClassNo); j++ {
			ver := ClassNo[j]
			if nums[ver] > 0 {
				nums[ver]--
			} else {
				fmt.Println("error,num<0")
				return nil
			}
			stus[ver] = append(stus[ver], s[j])
		}
		step = count
		start = tail
		tail = start + step
		//if flag {
		//	rand.Shuffle(len(ClassNo1), func(i, j int) {
		//		ClassNo1[i], ClassNo1[j] = ClassNo1[j], ClassNo1[i]
		//	})
		//	s := fstu[start:tail]
		//	for j := 0; j < 6; j++ {
		//		ver := ClassNo1[j]
		//		if nums[ver] > 0 {
		//			nums[ver]--
		//		} else {
		//			fmt.Println("error,num<0")
		//			return nil
		//		}
		//		stus[ver] = append(stus[ver], s[j])
		//		nums[ver]--
		//	}
		//	if nums[5] == 0 {
		//		flag = false
		//		step = 5
		//	}
		//	start = tail
		//	tail = start + step
		//} else {
		//	//把0-4下标洗牌
		//	rand.Shuffle(len(ClassNo), func(i, j int) {
		//		ClassNo[i], ClassNo[j] = ClassNo[j], ClassNo[i]
		//	})
		//	s := fstu[start:tail]
		//	for j := 0; j < 5; j++ {
		//		ver := ClassNo[j]
		//		/*if nums[ver] > 0 {
		//			nums[ver]--
		//		} else {
		//			fmt.Println("error,num<0")
		//			return
		//		}*/
		//		stus[ver] = append(stus[ver], s[j])
		//	}
		//	start = tail
		//	tail = start + step
		//}

	}

	//处理剩余不足5人的学生，如果有的话
	if tail-step < len(fstu) {
		ClassNo := make([]int, 0)
		for i := range nums {
			if nums[i] == 1 {
				count--
			}
			if nums[i] > 0 {
				ClassNo = append(ClassNo, i)
			}
		}
		rand.Shuffle(len(ClassNo), func(i, j int) {
			ClassNo[i], ClassNo[j] = ClassNo[j], ClassNo[i]
		})
		s := fstu[tail-step:]
		k := 0
		for j := 0; j < len(ClassNo); j++ {
			ver := ClassNo[j]
			nums[ver]--
			stus[ver] = append(stus[ver], s[k])
			k++
		}
	}
	//if len(stus[0])+len(stus[1])+len(stus[2])+len(stus[3])+len(stus[4])+len(stus[5]) != amount {
	//	fmt.Println("error,numbers loss")
	//}
	//对六班学生成绩排序
	sort.SliceStable(stus[5], func(i, j int) bool {
		return stus[5][i].score > stus[5][j].score
	})
	sort.SliceStable(stus[1], func(i, j int) bool {
		return stus[1][i].score > stus[1][j].score
	})
	sort.SliceStable(stus[2], func(i, j int) bool {
		return stus[2][i].score > stus[2][j].score
	})

	return stus
}

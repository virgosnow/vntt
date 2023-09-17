package main

import (
	"bufio"
	"fmt"
	"github.com/xuri/excelize/v2"
	"golang.org/x/text/encoding/unicode"
	"golang.org/x/text/transform"
	"log"
	"os"
	"regexp"
	"strings"
	"time"
)

func main() {
	println()
	//err := CheckExcelCount("./kimihane/input/Kimihane Couples - Chinese Sheet 1.xlsx")
	tm, err := CreateTransMap("./kimihane/input/")
	if err != nil {
		log.Println(err)
	}
	err = InsertTransMap("./kimihane/input/Kimihane Couples - Chinese Sheet 1.xlsx", tm)
	if err != nil {
		log.Println(err)
	}
	//println("今朝の転入届だ。")
	//println(tm["今朝の転入届だ。"])
}

// CreateTransMap 记录txt文件中的原文和译文，创建一个map组
func CreateTransMap(dirPath string) (map[string]string, error) {
	// 合并map
	mapPlus := func(m1, m2 map[string]string) map[string]string {
		m := map[string]string{}
		for k, v := range m1 {
			m[k] = v
		}
		for k, v := range m2 {
			m[k] = v
		}
		return m
	}

	tm := map[string]string{}
	dirs, _ := os.ReadDir(dirPath)
	for _, f := range dirs {
		if f.IsDir() {
			m, err := CreateTransMap(dirPath + "/" + f.Name())
			if err != nil {
				return nil, err
			}
			tm = mapPlus(tm, m)
		} else {
			if !strings.HasSuffix(f.Name(), ".txt") {
				println("跳过", f.Name())
				continue
			}
			file, err := os.Open(dirPath + "/" + f.Name())
			if err != nil {
				return nil, err
			}
			// 将UTF16内容转换成utf-8
			utf8Reader := transform.NewReader(file, unicode.UTF16(unicode.LittleEndian, unicode.UseBOM).NewDecoder())
			scanner := bufio.NewScanner(utf8Reader)
			var jpText, chText, num string
			for scanner.Scan() {
				text := scanner.Text()
				if strings.HasPrefix(text, "○") {
					newNum := strings.Split(text, "○")[1]
					if num == "" {
						num = newNum
					}
					if num != newNum {
						tm[jpText] = chText
						num = newNum
					}
					jpText = strings.Join(strings.Split(text, "○")[2:], "○")
					// trim
					jpText = strings.Trim(jpText, "　")
					// 去除ruby
					// 例：{浅生文|あそうふみ} => 浅生文
					jpText = ReplaceRuby(jpText, "{", "|", "}")
				} else if strings.HasPrefix(text, "●") {
					chText = strings.Join(strings.Split(text, "●")[2:], "●")
					// trim
					chText = strings.Trim(chText, "　")
					// 去除ruby
					// 例：{浅生文|あそうふみ} => 浅生文
					chText = ReplaceRuby(chText, "{", "|", "}")
				}
			}
			tm[jpText] = chText
		}
	}

	return tm, nil
}

func InsertTransMap(filepath string, tm map[string]string) error {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return err
	}
	for _, sheet := range f.GetSheetList() {
		rows, err := f.GetRows(sheet)
		if err != nil {
			return err
		}
		if len(rows[0]) < 6 || rows[0][2] != "Original" || rows[0][5] != "Chinese" {
			//log.Println(rows[0][2])
			continue
		}
		for k, row := range rows {
			if k == 0 || row[2] == "" {
				continue
			}
			//if strings.HasPrefix(row[0], "Image") {
			//	continue
			//}
			jpText := row[2]
			// 去除ruby
			// 例：≪天／・≫≪使／・≫ => 天使
			jpText = ReplaceRuby(jpText, "≪", "／", "≫")
			chText := tm[jpText]
			if chText != "" {
				cell, _ := excelize.CoordinatesToCellName(6, k+1)
				//println(fmt.Sprintf("f.SetCellStr(%v, %v, %v)", sheet, cell, chText))
				if err = f.SetCellStr(sheet, cell, chText); err != nil {
					return err
				}
			} else {
				//println(fmt.Sprintf("○%s_%d○%s\n●%s_%d●\n", sheet, k, jpText, sheet, k))
				println("["+sheet+"] not found:", jpText)
				cell, _ := excelize.CoordinatesToCellName(7, k+1)
				if err = f.SetCellStr(sheet, cell, "要确认"); err != nil {
					return err
				}
				//break
			}
		}
	}
	_ = f.SaveAs(filepath + ".fixed.xlsx")
	return nil
}

// CheckExcelCount 计算xlsx文件中每个sheet页的Original列的字数
func CheckExcelCount(filepath string) error {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return err
	}
	var totalCountSum int
	for _, sheet := range f.GetSheetList() {
		rows, err := f.GetRows(sheet)
		if err != nil {
			return err
		}
		if rows[0][2] != "Original" {
			//log.Println(rows[0][2])
			continue
		}
		var totalCount int
		for k, row := range rows {
			if k == 0 {
				continue
			}
			line := row[2]
			count := len([]rune(line))
			//log.Println("[" + strconv.Itoa(count) + "]" + line)
			totalCount += count
			totalCountSum += count
		}
		time.Sleep(time.Millisecond)
		println(fmt.Sprintf("%s\t%d", sheet, totalCount))
		//break
	}
	println("total:", totalCountSum)
	return nil
}

// ReplaceRuby 删除日文脚本里的注音
//
//	例：ReplaceRuby("{浅生文|あそうふみ}", "{", "|", "}") => 浅生文
//	注：Ruby就是日文脚本中经常有的汉字注音，写在汉字头上的，我们只有标拼音的时候用，日文有点滥用了，甚至用汉字标记汉字
func ReplaceRuby(s string, prefix, split, suffix string, showResult ...bool) string {
	// 转义
	ss := func(split string) string {
		if split == "|" {
			return "\\|"
		}
		return split
	}
	re, err := regexp.Compile(prefix + ".*" + ss(split) + ".*" + suffix)
	if err != nil {
		panic(err)
	}
	result := re.ReplaceAllStringFunc(s, func(s string) string {
		var str []rune
		var start bool
		for _, r := range []rune(s) {
			if r == []rune(prefix)[0] || r == []rune(suffix)[0] {
				start = true
				continue
			}
			if r == []rune(split)[0] {
				start = false
				continue
			}
			if start {
				str = append(str, r)
			}
		}
		if showResult != nil && showResult[0] {
			println(s + " -> " + string(str))
		}
		return string(str)
	})
	return result
}

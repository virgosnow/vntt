package excel

import (
	"crypto/md5"
	"encoding/json"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"log"
	"net/http"
	"net/url"
	"regexp"
	"sort"
	"strings"
)

type Excel struct {
	Output *excelize.File
}

func NewExcel() *Excel {
	return &Excel{
		Output: excelize.NewFile(),
	}
}

func (e *Excel) Test() error {
	s1 := []int64{1, 2, 3, 4, 5, 6, 7, 8, 9}
	var s2 []int64
	s2 = append(s2, s1...)
	s2[0] = 0
	log.Println(s1)
	log.Println(s2)
	return nil
}

func (e *Excel) Main() error {
	return e.HssHFind()
	//return e.HssH()
}

func (e *Excel) Trans() error {
	q := "try to take over the world"
	from := "en"
	to := "zh"
	appid := "20191207000363829"
	salt := "hakui"
	secret := "hFYOSHmjB66a3Sbhj3eG"
	signBef := appid + q + salt + secret
	sign := fmt.Sprintf("%x", md5.Sum([]byte(signBef)))
	log.Println(sign)
	reqUrl := "https://fanyi-api.baidu.com/api/trans/vip/translate?q=" + url.QueryEscape(q) +
		"&from=" + from +
		"&to=" + to +
		"&appid=" + appid +
		"&salt=" + salt +
		"&sign=" + sign
	type trans struct {
		Src string `json:"src"`
		Dst string `json:"dst"`
	}
	ret := struct {
		From  string  `json:"from"`
		To    string  `json:"to"`
		Trans []trans `json:"trans_result"`
		Error int64   `json:"error_code"`
	}{}
	body, err := HttpGet(reqUrl)
	log.Println(err)
	err = json.Unmarshal(body, &ret)
	log.Println(err)
	log.Println(ret)
	return nil
}

func HttpGet(reqUrl string) ([]byte, error) {
	resp, err := http.Get(reqUrl)
	if err != nil {
		return nil, err
	}
	defer func() { _ = resp.Body.Close() }()
	return ioutil.ReadAll(resp.Body)
}

func (e *Excel) HssHFind() error {
	f, err := excelize.OpenFile("./tcpclient/excel/hssh.xlsx")
	f2, err := excelize.OpenFile("./tcpclient/excel/hssh-renpy-tl.xlsx")
	if err != nil {
		return err
	}
	// 自动换行
	styleID, err := e.Output.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{Vertical: "center", WrapText: true}})
	if err != nil {
		return err
	}
	// 获取工作表中指定单元格的值
	sheetList := f.GetSheetList()
	for _, sheet := range sheetList {
		if sheet == "HSSH.xlsx|10日目" {
			break
		}
		rows, err := f.GetRows(sheet)
		rows2, err := f2.GetRows(sheet)
		if err != nil {
			return err
		}
		f2Has := make(map[string]int)
		m1 := regexp.MustCompile(`[\x20-\x7e]+`)
		for index2, row2 := range rows2 {
			ori2SimText := m1.ReplaceAllString(row2[2], "")
			f2Has[ori2SimText] = index2
		}
		// 删除特殊
		delete(f2Has, "……")
		e.Output.NewSheet(sheet)
		stream, err := e.Output.NewStreamWriter(sheet)
		if err != nil {
			return err
		}
		_ = stream.SetColWidth(1, 2, 50)
		err = stream.SetRow("A1", []interface{}{
			excelize.Cell{Value: "旧版原文"},
			excelize.Cell{Value: "新版原文"},
			excelize.Cell{Value: "差分字符量"},
		})
		err = stream.SetRow("A2", []interface{}{
			excelize.Cell{Value: ""},
			excelize.Cell{Value: ""},
			excelize.Cell{Formula: "SUM(C3:C1000)"},
		})
		var index2 int
		for k := range rows {
			if k == 0 || index2 >= len(rows2) {
				continue
			}
			oriText := rows[k][2]
			ori2Text := rows2[index2][2]
			ori2SimText := m1.ReplaceAllString(ori2Text, "")
			cell, _ := excelize.CoordinatesToCellName(1, k+2)
			if oriText == ori2SimText {
				err = stream.SetRow(cell, []interface{}{
					excelize.Cell{StyleID: styleID, Value: oriText},
					excelize.Cell{StyleID: styleID, Value: ori2Text},
					excelize.Cell{StyleID: styleID, Value: 0},
				})
				index2++
			} else if f2Has[oriText] == 0 || (f2Has[oriText] != 0 && f2Has[oriText] < index2) {
				err = stream.SetRow(cell, []interface{}{
					excelize.Cell{StyleID: styleID, Value: oriText},
					excelize.Cell{StyleID: styleID, Value: ""},
					excelize.Cell{StyleID: styleID, Formula: "LEN(" + cell + ")"},
				})
			} else if f2Has[oriText] != 0 {
				sumOri2Text := ""
				for oriText != ori2SimText && index2+1 < len(rows2) {
					sumOri2Text += ori2Text
					index2++
					ori2Text = rows2[index2][2]
					ori2SimText = m1.ReplaceAllString(ori2Text, "")
				}
				sumOri2Text += ori2Text
				index2++
				err = stream.SetRow(cell, []interface{}{
					excelize.Cell{StyleID: styleID, Value: oriText},
					excelize.Cell{StyleID: styleID, Value: sumOri2Text},
					excelize.Cell{StyleID: styleID, Formula: "LEN(" + cell + ")"},
				})
			}
		}
		err = stream.Flush()
		if err != nil {
			return err
		}
	}
	e.Output.DeleteSheet("Sheet1")
	if err := e.Output.SaveAs("hssh新旧文本差异.xlsx"); err != nil {
		fmt.Println(err)
	}
	return nil
}

func (e *Excel) HssH() error {
	f, err := excelize.OpenFile("./tcpclient/excel/hssh.xlsx")
	if err != nil {
		return err
	}
	// 获取工作表中指定单元格的值
	sheetList := f.GetSheetList()
	// 记录翻译
	trans := make(map[string]string)
	for _, sheet := range sheetList {
		rows, err := f.GetRows(sheet)
		if err != nil {
			return err
		}
		for k, row := range rows {
			if k == 0 || row[3] == "" {
				continue
			}
			key := strings.ReplaceAll(row[2], "ベリアル", "")
			key = strings.ReplaceAll(key, "ベル", "")
			// 保存
			trans[key] = row[3]
			// 增加命中率
			appendKeys := strings.SplitAfter(key, "。")
			appendValues := strings.SplitAfter(row[3], "。")
			appendValues2 := strings.SplitAfter(row[3], "，")
			if len(appendKeys) == len(appendValues) {
				for k, key := range appendKeys {
					trans[key] = appendValues[k]
				}
			} else if len(appendKeys) == len(appendValues2) {
				for k, key := range appendKeys {
					trans[key] = appendValues2[k]
				}
			}
			// 再一次增加命中率
			appendWKeys := strings.SplitAfter(key, "…")
			appendWValues := strings.SplitAfter(row[3], "……")
			if len(appendWKeys) == len(appendWValues) {
				for k, key := range appendWKeys {
					trans[key] = appendWValues[k]
				}
			}
		}
	}
	if err = f.Close(); err != nil {
		return err
	}
	// 开始匹配
	f, err = excelize.OpenFile("./tcpclient/excel/hssh-renpy-tl.xlsx")
	if err != nil {
		return err
	}
	// 获取工作表中指定单元格的值
	sheetList = f.GetSheetList()
	var totalCount, unmatchCount int64
	for _, sheet := range sheetList {
		if sheet != "30.rpy" {
			continue
		}
		rows, err := f.GetRows(sheet)
		if err != nil {
			return err
		}
		for k, row := range rows {
			if k == 0 {
				continue
			}
			var logOn bool
			logOn = true
			//m1 := regexp.MustCompile(`\{.*?\}`)
			m1 := regexp.MustCompile(`[\x20-\x7e]+`)
			keyText := m1.ReplaceAllString(row[2], "【数】")
			if logOn {
				log.Println(row[2])
				log.Println(keyText)
			}
			//log.Println(row[2])
			simText := m1.ReplaceAllString(row[2], "")
			// match
			matchText := trans[simText]
			// 增加命中率
			if matchText == "" {
				keys := strings.SplitAfter(simText, "。")
				for _, k := range keys {
					if k == "" {
						continue
					}
					if t := trans[k]; t != "" {
						matchText += t
					} else {
						matchText += "【待翻译】"
					}
				}
				if strings.Contains(matchText, "【待翻译】") {
					matchText = ""
				}
			}
			if logOn {
				if strings.HasPrefix(keyText, "【数】") {
					matchText = "【数】" + matchText
				}
				log.Println("译文：", matchText)
				log.Println("-----------------")
			}
			if matchText == "" {
				unmatchCount++
			}
			if matchText == "" && !logOn {
				log.Println(row[2])
				log.Println(m1.ReplaceAllString(row[2], "【数】"))
				log.Println("译文：", matchText)
				log.Println("-----------------")
				unmatchCount++
			}
			totalCount++
			//simText = strings.TrimLeft(simText, `\n `)
			//log.Println(simText)
		}
	}
	log.Println(fmt.Sprintf("total:%d, unmatch:%d", totalCount, unmatchCount))
	return nil
}

func (e *Excel) Find() error {
	f, err := excelize.OpenFile("./tcpclient/excel/hssh.xlsx")
	if err != nil {
		return err
	}
	// 获取工作表中指定单元格的值
	sheetList := f.GetSheetList()
	e.Output.NewSheet("重复部分")
	//cell, err := f.GetCellValue("Actors.json", "C2")
	//if err != nil {
	//	println(err.Error())
	//	return
	//}
	//println(cell)
	// 获取 Sheet1 上所有单元格

	existsMap := make(map[string][]string)
	existsZhMap := make(map[string][]string)

	// 自动换行
	styleID, err := e.Output.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{
		Vertical: "center",
		WrapText: true,
	}})
	if err != nil {
		return err
	}

	// 重复
	for _, sheet := range sheetList {
		rows, err := f.GetRows(sheet)
		if err != nil {
			return err
		}
		for k, row := range rows {
			if k == 0 {
				continue
			}
			var sheetExists, zhExists bool
			for _, v := range existsMap[row[2]] {
				if v == sheet {
					sheetExists = true
				}
			}
			if !sheetExists {
				existsMap[row[2]] = append(existsMap[row[2]], sheet)
			}
			for _, v := range existsZhMap[row[2]] {
				if v == row[3] {
					zhExists = true
				}
			}
			if !zhExists {
				existsZhMap[row[2]] = append(existsZhMap[row[2]], row[3])
				sort.Slice(existsZhMap[row[2]], func(i, j int) bool {
					return existsZhMap[row[2]][i] > existsZhMap[row[2]][j]
				})
			}
		}
	}

	// 标记
	for _, sheet := range sheetList {
		rows, err := f.GetRows(sheet)
		if err != nil {
			return err
		}
		e.Output.NewSheet(sheet)
		stream, err := e.Output.NewStreamWriter(sheet)
		if err != nil {
			return err
		}
		_ = stream.SetColWidth(1, 2, 50)
		err = stream.SetRow("A1", []interface{}{
			excelize.Cell{Value: "原文"},
			excelize.Cell{Value: "译文"},
			excelize.Cell{Value: "出现页"},
		})
		for k, row := range rows {
			if k == 0 {
				continue
			}
			cell, _ := excelize.CoordinatesToCellName(1, k+2)
			err = stream.SetRow(cell, []interface{}{
				excelize.Cell{StyleID: styleID, Value: row[2]},
				excelize.Cell{StyleID: styleID, Value: row[3]},
				excelize.Cell{Value: strings.Join(existsMap[row[2]], ",")},
			})
			//print(fmt.Sprintf("%s\t%s\t%s", row[4], row[2], row[3]))
			//for _, colCell := range row {
			//	print(colCell, "\t")
			//}
			//println()
		}
		err = stream.Flush()
		if err != nil {
			return err
		}
		//break
	}

	// 标记
	stream, err := e.Output.NewStreamWriter("重复部分")
	if err != nil {
		return err
	}
	_ = stream.SetColWidth(1, 3, 50)
	err = stream.SetRow("A1", []interface{}{
		excelize.Cell{Value: "原文"},
		excelize.Cell{Value: "译文"},
		excelize.Cell{Value: "冲突译文"},
		excelize.Cell{Value: "出现页"},
	})
	k := 2
	for _, sheet := range sheetList {
		rows, err := f.GetRows(sheet)
		if err != nil {
			return err
		}
		for _, row := range rows {
			if len(existsMap[row[2]]) > 1 {
				zhs := existsZhMap[row[2]]
				if len(zhs) > 1 {
					if zhs[len(zhs)-1] == "" {
						zhs = zhs[:len(zhs)-1]
					}
				}
				cell, _ := excelize.CoordinatesToCellName(1, k)
				err = stream.SetRow(cell, []interface{}{
					excelize.Cell{StyleID: styleID, Value: row[2]},
					excelize.Cell{StyleID: styleID, Value: zhs[0]},
					excelize.Cell{StyleID: styleID, Value: strings.Join(zhs[1:], "\n")},
					excelize.Cell{Value: strings.Join(existsMap[row[2]], ",")},
				})
				delete(existsMap, row[2])
				k++
			}
			//print(fmt.Sprintf("%s\t%s\t%s", row[4], row[2], row[3]))
			//for _, colCell := range row {
			//	print(colCell, "\t")
			//}
			//println()
		}
	}
	err = stream.Flush()
	if err != nil {
		return err
	}
	//break
	e.Output.DeleteSheet("Sheet1")
	if err := e.Output.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
	return nil
}

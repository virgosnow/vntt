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
	return e.HssH()
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

func (e *Excel) HssH() error {
	f, err := excelize.OpenFile("./hssh.xlsx")
	if err != nil {
		return err
	}
	// 获取工作表中指定单元格的值
	sheetList := f.GetSheetList()
	// 记录翻译
	trans := make(map[string]string)
	allText := ""
	newSheet := func(sheet string) (*excelize.StreamWriter, error) {
		e.Output.NewSheet(sheet)
		stream, err := e.Output.NewStreamWriter(sheet)
		if err != nil {
			return nil, err
		}
		_ = stream.SetColWidth(1, 3, 50)
		err = stream.SetRow("A1", []interface{}{excelize.Cell{Value: "new"}, excelize.Cell{Value: "old"}, excelize.Cell{Value: "訳文"}, excelize.Cell{Value: "差分"}})
		err = stream.SetRow("A2", []interface{}{excelize.Cell{Value: ""}, excelize.Cell{Value: ""}, excelize.Cell{Value: ""}, excelize.Cell{Formula: "SUM(D3:D1000)"}})
		return stream, nil
	}
	styleID, _ := e.Output.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{Vertical: "center", WrapText: true}})
	for _, sheet := range sheetList {
		rows, err := f.GetRows(sheet)
		if err != nil {
			return err
		}
		for k, row := range rows {
			if k == 0 || row[3] == "" {
				continue
			}
			key := row[2]
			// 保存
			trans[key] = row[3]
			allText += key
			// 增加命中率
			powerSplit := func(reg string) {
				compile := regexp.MustCompile(reg)
				smallKeys := compile.Split(key, -1)
				for k, com := range compile.FindAllString(key, -1) {
					smallKeys[k] += com
				}
				smallValues := compile.Split(row[3], -1)
				for k, com := range compile.FindAllString(row[3], -1) {
					smallValues[k] += com
				}
				if len(smallKeys) == len(smallValues) {
					for k, key := range smallKeys {
						value := smallValues[k]
						if _, ok := trans[key]; !ok {
							trans[key] = value
						}
					}
				}
			}
			powerSplit(`[、，]+`)
			powerSplit(`[。]+`)
			powerSplit(`[。、，]+`)
			powerSplit(`[…]+`)
			powerSplit(`[！]+`)
			powerSplit(`[？]+`)
			powerSplit(`[！？。、，…]+`)
		}
	}
	if err = f.Close(); err != nil {
		return err
	}
	// 开始匹配
	f, err = excelize.OpenFile("./hssh-renpy-tl.xlsx")
	if err != nil {
		return err
	}
	// 获取工作表中指定单元格的值
	sheetList = f.GetSheetList()
	selectSheet := map[string]bool{
		"10.rpy|0日目": true, "11.rpy|1日目": true, "14.rpy|4日目": true, "15.rpy|5日目": true, "16.rpy|6日目": true,
	}
	var totalCount, unmatchedCount, tmpCount int64
	for _, sheet := range sheetList {
		if !selectSheet[sheet] {
			continue
		}
		rows, err := f.GetRows(sheet)
		stream, err := newSheet(sheet)
		if err != nil {
			return err
		}
		for k, row := range rows {
			if k == 0 {
				continue
			}
			m1 := regexp.MustCompile(`[\x20-\x7e]+`)
			//keyText := m1.ReplaceAllString(row[2], "【数】")
			rowText := strings.ReplaceAll(row[2], "[na]", "ベリアル")
			rowText = strings.ReplaceAll(rowText, "[na2]", "ベル")
			simText := m1.ReplaceAllString(rowText, "")
			// match
			matchText := trans[simText]
			//if strings.HasPrefix(keyText, "【数】") {
			//	matchText = "【数】" + matchText
			//}
			if simText == "" {
				matchText = rowText
			} else if matchText == "" && strings.Contains(allText, simText) {
				powerMatch := func(text, reg string) string {
					if text != "" {
						return text
					}
					mustCompile := regexp.MustCompile(reg)
					smallKeys := mustCompile.Split(simText, -1)
					for k, com := range mustCompile.FindAllString(simText, -1) {
						smallKeys[k] += com
					}
					for _, k := range smallKeys {
						if t, ok := trans[k]; ok {
							text += t
						} else {
							text = ""
							break
						}
					}
					return text
				}
				// 不断抢救
				matchText = powerMatch(matchText, `[、，]+`)
				matchText = powerMatch(matchText, `[。]+`)
				matchText = powerMatch(matchText, `[、，。]+`)
				matchText = powerMatch(matchText, `[！]+`)
				matchText = powerMatch(matchText, `[？]+`)
				matchText = powerMatch(matchText, `[…]+`)
				matchText = powerMatch(matchText, `[！？。、，…]+`)
				if matchText == "" {
					tmpCount++
					log.Println("[" + simText + "]")
				}
			}
			if matchText == "" {
				unmatchedCount++
			}
			var logOn bool
			//logOn = true
			if logOn {
				log.Println(rowText)
				log.Println(simText)
				if matchText == "" {
					log.Println("译文：", "unmatched")
					log.Println("-----------------")
				} else {
					log.Println("译文：", matchText)
					log.Println("-----------------")
				}
			}
			totalCount++
			// 输出成文件
			cell, _ := excelize.CoordinatesToCellName(1, k+2)
			cell2, _ := excelize.CoordinatesToCellName(3, k+2)
			formula := "=IF(" + cell2 + "=\"\",LEN(" + cell + "),0)"
			if matchText == "" {
				simText = ""
			} else {
				matchText = strings.ReplaceAll(matchText, "贝利艾尔", "[na]")
				matchText = strings.ReplaceAll(matchText, "贝尔", "[na2]")
				if strings.HasPrefix(row[2], "\\n") {
					matchText = "\\n " + matchText
				}
				if strings.HasSuffix(row[2], "{w}{nw}") {
					matchText += "{w}{nw}"
				}
			}
			// matchText改造
			err = stream.SetRow(cell, []interface{}{
				excelize.Cell{StyleID: styleID, Value: row[2]},
				excelize.Cell{StyleID: styleID, Value: simText},
				excelize.Cell{StyleID: styleID, Value: matchText},
				excelize.Cell{StyleID: styleID, Formula: formula},
			})
			//simText = strings.TrimLeft(simText, `\n `)
			//log.Println(simText)
		}
		_ = stream.Flush()
	}
	e.Output.DeleteSheet("Sheet1")
	if err := e.Output.SaveAs("hssh新旧文本差异.xlsx"); err != nil {
		fmt.Println(err)
	}
	log.Println(fmt.Sprintf("total:%d, unmatch:%d, tmp:%d", totalCount, unmatchedCount, tmpCount))
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

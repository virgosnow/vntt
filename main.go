package main

import (
	"awesomeProject/excel"
	"log"
)

func main() {
	e := excel.NewExcel()
	log.Println(e.Main())
}

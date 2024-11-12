package main

import "toJson/utils"

func main() {
	var fileName string = "quotes.xlsx"

	if err := utils.ReadQuotesFromExcel(fileName); err != nil {
		panic(err)
	}
}
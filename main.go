package main

import "toJson/utils"

func main() {
	var fileName string = "quotes.xlsx"

	// reads quotes from excel and converts in to json format
	if err := utils.ReadQuotesFromExcel(fileName); err != nil {
		panic(err)
	}
}
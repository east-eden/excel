v ?= latest

.PHONY: gen
gen:
	./bin/code_generator_mac -relocatePath="/excel" -readExcelPath="./" -exportPath="../server/excel/auto"
	cp *.xlsx ../server/config/excel/
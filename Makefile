v ?= latest

.PHONY: gen
gen:
	./bin/code_generator -relocatePath="" -readExcelPath="./" -exportPath="../server/excel/auto"
	cp *.xlsx ../server/config/excel/

.PHONY: gen_mac
gen_mac:
	./bin/code_generator_mac -relocatePath="/excel" -readExcelPath="./" -exportPath="../server/excel/auto"
	cp *.xlsx ../server/config/excel/
v ?= latest

.PHONY: gen
gen:
	./bin/code_generator -relocatePath="" -readExcelPath="./global/" -exportPath="../server/excel/auto"
	dotnet ./bin/excel.dll ./ ../client/Scripts
	cp ./global/*.xlsx ../server/config/excel/

.PHONY: gen_mac
gen_mac:
	./bin/code_generator_mac -relocatePath="/excel" -readExcelPath="./global/" -exportPath="../server/excel/auto"
	dotnet ./bin/excel.dll ./ ../client/Scripts
	cp ./global/*.xlsx ../server/config/excel/
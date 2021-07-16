v ?= latest

.PHONY: gen_server
gen_server:
	rm -f ../server/excel/auto/*_entry.go || true
	./bin/code_generator -relocatePath="" -readExcelPath="./global/" -exportGoPath="../server/excel/auto" -exportCsvPath="../server/config/csv/"
	gofmt -w ../server/excel/auto/*.go
	goimports -w ../server/excel/auto/*.go

.PHONY: gen_client
gen_client:
	dotnet ./bin/excel.dll ./ ../ee_client/scripts

.PHONY: gen_mac
gen_mac:
	rm -f ../server/excel/auto/*_entry.go || true
	rm -f ../server/config/csv/*.csv || true
	rm -r ../server_bin/config/csv/*.csv || true
	./bin/code_generator_mac -relocatePath="/excel" -readExcelPath="./global/" -exportGoPath="../server/excel/auto" -exportCsvPath="../server/config/csv/"
	# dotnet ./bin/excel.dll ./ ../ee_client/scripts
	cp ../server/config/csv/*.csv ../server_bin/config/csv/

	gofmt -w ../server/excel/auto/*.go
	goimports -w ../server/excel/auto/*.go
v ?= latest

.PHONY: gen
gen:
	./bin/code_generator
	cp ../excel/*.xlsx ../server/config/excel/

.PHONY: gen_mac
gen_mac:
	./bin/code_generator_mac
	cp ../excel/*.xlsx ../server/config/excel/
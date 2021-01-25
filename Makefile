v ?= latest

.PHONY: gen
gen:
	./bin/code_generator

.PHONY: gen_mac
gen_mac:
	./bin/code_generator_mac
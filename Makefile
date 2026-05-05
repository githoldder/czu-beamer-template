.PHONY: all build clean pptx watch

MAIN := beamersapienza

all: build

build:
	latexmk -xelatex -interaction=nonstopmode -synctex=1 $(MAIN).tex

clean:
	latexmk -c $(MAIN).tex
	rm -f *.xdv *.nav *.snm *.vrb

pptx: build
	python3 make_pptx.py

watch:
	latexmk -pvc -xelatex -interaction=nonstopmode $(MAIN).tex

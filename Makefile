all:
	python3 main.py
convert:
	abiword --to=pdf news.docx
view:
	libreoffice news.docx
clean:
	rm -rf *.docx *.pdf

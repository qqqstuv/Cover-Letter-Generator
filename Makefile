all:
	python example.py
convert:
	abiword --to=pdf news.docx

view:
	libreoffice news.docx

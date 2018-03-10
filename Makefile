all:
	python main.py
convert:
	abiword --to=pdf news.docx
view:
	libreoffice news.docx

#!/usr/bin/python
import imgkit
import pdfkit
import os
import sys
from pwn import log

opt = {
	'quiet': '',
	'enable-local-file-access': ''
}
pdfopt = dict(opt)
pdfopt.update({
	'page-size': 'A4',
    'margin-top': '2cm',
    'margin-bottom': '2cm',
    'margin-left': '3cm',
    'margin-right': '1cm',
    'encoding': 'UTF-8'
})

def clean():
	if os.path.exists('img'):
		for f in os.listdir('img'):
			os.remove(os.path.join('img', f))

def gen_image(path):
	files = os.listdir(path)
	css = [f for f in files if f.endswith('.css')]
	fp = os.path.join(path, 'index.html')
	try:
		imgkit.from_file(fp, os.path.join('img', f'{path.split("/")[-1]}.jpg'), options=opt, css=css)
	except Exception as e:
		log.warn(f'exception in rendering {path.split("/")[-1]} html')

def read_tasks(path):
	task = []
	with open(path) as f:
		for part in f.read().split('<h3>')[1:]:
			task.append([a.strip() for a in part.split('</h3>')])
	return task

text_to_html = {
	'&': '&amp;',
	'<': '&lt;',
	'>': '&gt;',
	' ': '&nbsp;',
	'\t': '&emsp;',
	'\n': '<br>'
}
def to_html(text):
	newtext = text
	for i in text_to_html:
		newtext = newtext.replace(i, text_to_html[i])
	return newtext

def build_all(path, overwrite=False):
	if not os.path.exists('img'):
		os.makedirs('img')
	codes = []
	images = [f for f in os.listdir('img') if f.endswith('.jpg')]
	for (dirpath, dirnames, filenames) in os.walk(path):
		dirnames.sort(key=lambda x: int(x.split('task')[1]))
		for d in dirnames:
			dp = os.path.join(dirpath, d)
			lp = log.progress(f'building {d}')
			if overwrite or f'{d}.jpg' not in images:
				gen_image(dp)
			code = []
			for f in os.listdir(dp):
				if f.split('.')[-1] in ['html', 'js', 'css']:
					with open(os.path.join(dirpath, d, f)) as file:
						code.append((f, to_html(file.read())))
			codes.append(code)
			lp.success('done.')
		break
	return codes

from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH

document = Document()
style = document.styles['Normal']
font = style.font
font.name = 'Times New Roman'
font.size = Pt(12)

def add_toc(document):
	# https://stackoverflow.com/questions/18595864/python-create-a-table-of-contents-with-python-docx-lxml
	ptitle = document.add_paragraph()
	ptitle.add_run('Оглавление').bold = True
	ptitle.align = WD_ALIGN_PARAGRAPH.CENTER

	paragraph = document.add_paragraph()
	#ptitile = paragraph.add_run('Оглавление')
	#ptitile.align = WD_ALIGN_PARAGRAPH.CENTER
	run = paragraph.add_run()
	fldChar = OxmlElement('w:fldChar')  # creates a new element
	fldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element
	instrText = OxmlElement('w:instrText')
	instrText.set(qn('xml:space'), 'preserve')  # sets attribute on element
	instrText.text = 'TOC \\o "1-3" \\h \\z \\u'   # change 1-3 depending on heading levels you need

	fldChar2 = OxmlElement('w:fldChar')
	fldChar2.set(qn('w:fldCharType'), 'separate')
	fldChar3 = OxmlElement('w:t')
	fldChar3.text = "Right-click to update field."
	fldChar2.append(fldChar3)

	fldChar4 = OxmlElement('w:fldChar')
	fldChar4.set(qn('w:fldCharType'), 'end')

	r_element = run._r
	r_element.append(fldChar)
	r_element.append(instrText)
	r_element.append(fldChar2)
	r_element.append(fldChar4)

def build_report(path):
	add_toc(document)
	#Giving headings that need to be included in Table of contents

	document.add_heading("Network Connectivity")
	document.add_heading("Weather Stations")

	document.save('demo.docx')
	"""codes = build_all(path, len(sys.argv) > 1 and sys.argv[1] == 'rebuild')
	with open('template.html') as f:
		doc = f.read()
	html = ''
	images = os.listdir('img')
	tasks = read_tasks(os.path.join(path, 'task.html'))
	log.info(f'images amount = {len(images)}')
	log.info(f'tasks amount = {len(tasks)}')
	lp = log.progress('generatig document')
	for i, task in enumerate(tasks):
		html += f'<h3 class="new-page">{task[0]}</h3>\n<p><b>Задание:</b><br>\n{task[1]}<br>\n'
		for ci, c in enumerate(codes[i]):
			part = f''
			html += f'<p>Листинг {i+1}.{ci+1} – {c[0]} <br><p><a>{c[1]}</a></p></p>'
		image = next((a for a in images if str(i+1) in a), None)
		if image:
			html += f'<p><img align="center" src=\"{os.path.join("img", image)}\"><br>\n'
			html += f'Рисунок {i+1} - Результат выполнения кода на странице</p>'
		else:
			log.warn(f'image not found for task {str(i+1)}')
		html += '</p>\n'
	lp.success('done!')
	lp = log.progress('converting to pdf')
	doc = doc.replace('#tasks', html)
	with open('out.html', 'w') as f:
		f.write(doc)
	pdfkit.from_file('out.html', 'out.pdf', options=pdfopt, toc={})
	lp.success('done!')"""


if len(sys.argv) > 1 and sys.argv[1] == 'clean':
	clean()
else:
	build_report('./mirea-frontend/p1')
	
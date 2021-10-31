import os
import re
# import xml.etree.ElementTree as ET
from lxml import etree as ET
from docx import Document

tei_ns = '{http://www.tei-c.org/ns/1.0}'
xml_ns = '{http://www.w3.org/XML/1998/namespace}'

main_dir = re.sub(r"\\", "/", os.getcwd())
cx_fname = 'Rom14_April-30-21.xml'

with open(cx_fname, 'r', encoding='utf-8') as file:
    tree = file.read()
tree = re.sub("xml:id", "verse", tree)
tree = re.sub("<TEI xmlns=\"http://www.tei-c.org/ns/1.0\">", "<TEI>", tree)
tree = re.sub("<?xml version='1.0' encoding='UTF-8'?>", "", tree)
with open('temp.xml', 'w', encoding='utf-8') as file:
    file.write(tree)

parser = ET.XMLParser(remove_blank_text=True)
tree = ET.parse('temp.xml', parser)
root = tree.getroot()

document = Document("template.docx")


# document.add_heading('Critical Apparatus\n', 0)

ab_elements = root.findall("ab")
for ab in ab_elements:
    apps = ab.findall('app')
    verse = ab.get("verse")
    verse = verse.replace('-APP', '')

    full_verse = verse.replace('Rom', 'Romans ')
    full_verse = full_verse.replace('.', ':')

    reference = document.add_paragraph(full_verse)
    reference.style = document.styles['reference']

    ref = re.sub(r"\.", " ", verse)
    ref = ref.split()
    chp = ref[0]
    vrs = ref[1]

    # Get RP verse for display and also the ECM style index nums
    
    # it is insane that I did this!
    # with open(rp_fname, 'r', encoding='utf-8') as file:
    #     basetext = file.readlines()

    basetext = 'This is filler text. I wrote this along time ago and did not have a good way of getting the basetext'
    basetext = basetext.split()

    for line in basetext:
        if line.startswith(vrs):
            line = re.sub(vrs, "", line)
            basetext = line.strip().split()

    index = []
    count = 2

    for i in range(len(basetext)):
        count_str = str(count)
        index.append(count_str)
        count += 2

    verse_length = len(basetext)

    cell = 0
    if verse_length <= 15:
        table = document.add_table(rows=0, cols=verse_length)
        row_cells = table.add_row().cells
        for x, y in zip(index, basetext):
            row_cells[cell].text = f"{x}\n{y}"
            row_cells[cell].style = document.styles['table cell']
            cell += 1

    elif verse_length <= 30:
        index_a = index[:15]
        basetext_a = basetext[:15]

        table = document.add_table(rows=0, cols=15)
        row_cells = table.add_row().cells

        for x, y in zip(index[:15], basetext[:15]):
            row_cells[cell].text = f"{y}\n{x}"
            row_cells[cell].paragraphs[0].style = document.styles['table cell']
            cell += 1
        cell = 0
        row_cells = table.add_row().cells
        for x, y in zip(index[15:], basetext[15:]):
            row_cells[cell].text = f"{y}\n{x}"
            row_cells[cell].paragraphs[0].style = document.styles['table cell']
            cell += 1

    else:
        table = document.add_table(rows=0, cols=15)
        row_cells = table.add_row().cells

        for x, y in zip(index[:15], basetext[:15]):
            row_cells[cell].text = f"{y}\n{x}"
            row_cells[cell].paragraphs[0].style = document.styles['table cell']
            cell += 1

        cell = 0
        row_cells = table.add_row().cells
        for x, y in zip(index[15:30], basetext[15:30]):
            row_cells[cell].text = f"{y}\n{x}"
            row_cells[cell].paragraphs[0].style = document.styles['table cell']
            cell += 1
        
        cell = 0
        row_cells = table.add_row().cells
        for x, y in zip(index[30:], basetext[30:]):
            row_cells[cell].text = f"{y}\n{x}"
            row_cells[cell].paragraphs[0].style = document.styles['table cell']
            cell += 1

    for app in apps:
        try:
            app_from = app.get('from')
            app_to = app.get('to')
            if app_from == app_to:
                index = app_from
            else:
                index = f'{app_from}–{app_to}'
            p = document.add_paragraph(index)
            p.style = document.styles['index']
            rdgs = app.findall('rdg')
            for rdg in rdgs:
                if rdg.text:
                    greek_text = rdg.text
                    p = document.add_paragraph()
                    p.style = document.styles['reading']
                    p.add_run(f"{rdg.get('n')}: ").italic = True
                    p.add_run(greek_text).bold = True
                    p.add_run(f" // {rdg.get('wit')}")
                else:
                    greek_text = rdg.get('type')
                    p = document.add_paragraph(f"Reading {rdg.get('n')}: ")
                    p.style = document.styles['reading']
                    p.add_run(greek_text).bold = True
                    p.add_run(f" // {rdg.get('wit')}")
        except:
            pass

document.save("apparatus.docx")

print('Done')
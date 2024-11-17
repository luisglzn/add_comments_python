from docx import Document
from docx.oxml.shared import qn
from docx.oxml import OxmlElement
from datetime import datetime
import random
import os
import numpy as np
import new
import xml.etree.ElementTree as ET
import uuid
import zipfile
import io
from lxml import etree
import re

def split_xml_by_elements(xml):
    root = etree.fromstring(xml)
    paragraphs = root.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    return [etree.tostring(p, encoding='unicode') for p in paragraphs]

def replace_tags(paragraphs, tag):
    txt_with_tags = []
    tags_list = []
    for p in paragraphs:
        parts = re.split(f'(<{tag}.*?>{tag}|</{tag}>)', p)
        txt_with_tags.append(''.join(parts[1::2]))
        tags_list.append(parts[::2])
    return txt_with_tags, tags_list

def build_txt(paragraphs):
    return [re.sub('<.*?>', '', p) for p in paragraphs]

def add_comments_to_docx(doc, xml, paragraphs, txt_with_tags, tags_list, txt, phrase, author):
    comment_count = 0
    modified_paragraphs = []
    phrase_regex = re.compile(re.escape(phrase), re.IGNORECASE)
    
    for i, paragraph in enumerate(paragraphs):
        paragraph_text = txt[i]
        matches = list(phrase_regex.finditer(paragraph_text))
        
        if matches:
            modified_paragraph, comments = add_comments_to_paragraph(
                paragraph, matches, phrase, author, comment_count
            )
            modified_paragraphs.append(modified_paragraph)
            for comment in comments:
                add_comment_to_document(doc, comment)
            comment_count += len(matches)
        else:
            modified_paragraphs.append(paragraph)
    
    new_xml = rebuild_xml(xml, modified_paragraphs)
    return new_xml

def add_comments_to_paragraph(paragraph, matches, phrase, author, comment_count):
    root = etree.fromstring(paragraph)
    comments = []
    
    for match in matches:
        comment_id = str(random.randint(0, 9999))
        comment_text = f"Comentario sobre: {phrase}"
        
        comment_range_start = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeStart')
        comment_range_start.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', comment_id)
        
        comment_range_end = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeEnd')
        comment_range_end.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', comment_id)
        
        comment_reference = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentReference')
        comment_reference.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', comment_id)
        
        insert_comment_elements(root, match.start(), match.end(), comment_range_start, comment_range_end, comment_reference)
        
        comment = create_comment(comment_id, author, comment_text)
        comments.append(comment)
        
        comment_count += 1
    
    return etree.tostring(root, encoding='unicode'), comments

def insert_comment_elements(root, start, end, comment_range_start, comment_range_end, comment_reference):
    w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    current_pos = 0
    
    for run in root.findall(f'.//{w_ns}r'):
        for text in run.findall(f'.//{w_ns}t'):
            text_length = len(text.text)
            if current_pos <= start < current_pos + text_length:
                split_pos = start - current_pos
                if split_pos > 0:
                    new_run = etree.Element(f'{w_ns}r')
                    new_text = etree.SubElement(new_run, f'{w_ns}t')
                    new_text.text = text.text[split_pos:]
                    text.text = text.text[:split_pos]
                    run.addnext(new_run)
                    run = new_run
                run.addprevious(comment_range_start)
                break
            current_pos += text_length
    
    current_pos = 0
    for run in root.findall(f'.//{w_ns}r'):
        for text in run.findall(f'.//{w_ns}t'):
            text_length = len(text.text)
            if current_pos <= end <= current_pos + text_length:
                split_pos = end - current_pos
                if split_pos < text_length:
                    new_run = etree.Element(f'{w_ns}r')
                    new_text = etree.SubElement(new_run, f'{w_ns}t')
                    new_text.text = text.text[split_pos:]
                    text.text = text.text[:split_pos]
                    run.addnext(new_run)
                run.addnext(comment_reference)
                run.addnext(comment_range_end)
                return
            current_pos += text_length

def create_comment(comment_id, author, comment_text):
    comment = OxmlElement('w:comment')
    comment.set(qn('w:id'), comment_id)
    comment.set(qn('w:author'), author)
    comment.set(qn('w:date'), datetime.now().isoformat())
    comment.set(qn('w:initials'), ''.join([name[0].upper() for name in author.split() if name]))
    
    p = OxmlElement('w:p')
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = comment_text
    r.append(t)
    p.append(r)
    comment.append(p)
    
    return comment

def add_comment_to_document(doc, comment):
    comments_part = doc.part.comments_part
    if comments_part is None:
        comments_part = doc.part.add_comments_part()
    comments_part._element.append(comment)

def rebuild_xml(original_xml, modified_paragraphs):
    root = etree.fromstring(original_xml)
    body = root.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body')
    
    if body is None:
        raise ValueError("No se pudo encontrar el elemento body en el XML")
    
    for p in body.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
        body.remove(p)
    
    for paragraph in modified_paragraphs:
        try:
            p_element = etree.fromstring(paragraph)
            body.append(p_element)
        except TypeError:
            body.append(paragraph)
        except etree.XMLSyntaxError as e:
            print(f"Error al parsear el párrafo: {e}")
            print(f"Párrafo problemático: {paragraph[:100]}...")
    
    return etree.tostring(root, encoding='unicode', pretty_print=True)

def main(doc_path, phrase, author):
    doc = Document(doc_path)
    xml = doc.element.xml
    paragraphs = split_xml_by_elements(xml)
    txt_with_tags, tags_list = replace_tags(paragraphs, '<#>')
    txt = build_txt(paragraphs)
    
    new_xml = add_comments_to_docx(doc, xml, paragraphs, txt_with_tags, tags_list, txt, phrase, author)
    
    new_root = etree.fromstring(new_xml)
    doc._element.clear()
    for child in new_root:
        doc._element.append(child)
    
    output_path = doc_path.replace('.docx', '_comentado.docx')
    doc.save(output_path)
    print(f"Documento con comentarios guardado como: {output_path}")

if __name__ == "__main__":
    doc_path = r"C:\Users\luisg\OneDrive\Escritorio\Trabajo\Add comments\add_comments_python\Script_TestFile - copia.docx"
    phrase = "división física"
    author = "Luis"
    main(doc_path, phrase, author)
import xml.etree.ElementTree as ET
import json
import zipfile
import os
import shutil
import uuid
from datetime import datetime

def add_comments_to_docx(input_docx, comments_json, output_docx):
    temp_dir = "temp_docx"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.mkdir(temp_dir)

    with zipfile.ZipFile(input_docx, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    with open(comments_json, 'r') as f:
        comments_data = json.load(f)

    document_tree = ET.parse(os.path.join(temp_dir, 'word', 'document.xml'))
    document_root = document_tree.getroot()

    comments_tree, comments_root = create_or_update_xml(temp_dir, 'word/comments.xml', 'comments')
    comments_extended_tree, comments_extended_root = create_or_update_xml(temp_dir, 'word/commentsExtended.xml', 'commentsExtended')
    people_tree, people_root = create_or_update_xml(temp_dir, 'word/people.xml', 'people')

    namespaces = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    }

    for ns, uri in namespaces.items():
        ET.register_namespace(ns, uri)

    comment_id = 0
    for comment_data in comments_data:
        quote = comment_data['quote']
        comment_text = comment_data['comment']
        author = comment_data['author']

        for paragraph in document_root.findall('.//w:p', namespaces):
            text_elements = paragraph.findall('.//w:t', namespaces)
            paragraph_text = ''.join(elem.text for elem in text_elements if elem.text)
            
            if quote in paragraph_text:
                comment_id += 1
                add_comment_to_paragraph(paragraph, comment_id, quote, comment_text, author, namespaces)
                add_comment_to_comments_xml(comments_root, comment_id, comment_text, author, namespaces)
                add_comment_to_comments_extended(comments_extended_root, comment_id, namespaces)
                add_person_to_people_xml(people_root, author, namespaces)
                break

    document_tree.write(os.path.join(temp_dir, 'word', 'document.xml'), encoding='UTF-8', xml_declaration=True)
    comments_tree.write(os.path.join(temp_dir, 'word', 'comments.xml'), encoding='UTF-8', xml_declaration=True)
    comments_extended_tree.write(os.path.join(temp_dir, 'word', 'commentsExtended.xml'), encoding='UTF-8', xml_declaration=True)
    people_tree.write(os.path.join(temp_dir, 'word', 'people.xml'), encoding='UTF-8', xml_declaration=True)

    update_content_types(temp_dir)
    update_document_rels(temp_dir)
    update_settings(temp_dir)

    shutil.make_archive(output_docx, 'zip', temp_dir)
    os.rename(output_docx + '.zip', output_docx)

    shutil.rmtree(temp_dir)

def create_or_update_xml(temp_dir, file_path, root_element):
    full_path = os.path.join(temp_dir, file_path)
    if os.path.exists(full_path):
        tree = ET.parse(full_path)
        root = tree.getroot()
    else:
        root = ET.Element(f'{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{root_element}')
        tree = ET.ElementTree(root)
    return tree, root

def add_comment_to_paragraph(paragraph, comment_id, quote, comment_text, author, namespaces):
    w_ns = namespaces['w']
    
    # Encontrar la posición del texto citado
    start_index = 0
    for run in paragraph.findall(f'.//{{{w_ns}}}r'):
        text_elem = run.find(f'{{{w_ns}}}t')
        if text_elem is not None and text_elem.text:
            if quote in text_elem.text:
                break
            start_index += len(text_elem.text)

    # Insertar commentRangeStart
    comment_range_start = ET.Element(f'{{{w_ns}}}commentRangeStart')
    comment_range_start.set(f'{{{w_ns}}}id', str(comment_id))
    paragraph.insert(start_index, comment_range_start)

    # Insertar commentRangeEnd
    comment_range_end = ET.Element(f'{{{w_ns}}}commentRangeEnd')
    comment_range_end.set(f'{{{w_ns}}}id', str(comment_id))
    paragraph.insert(start_index + 2, comment_range_end)

    # Insertar commentReference
    comment_reference = ET.Element(f'{{{w_ns}}}r')
    comment_reference_elem = ET.SubElement(comment_reference, f'{{{w_ns}}}commentReference')
    comment_reference_elem.set(f'{{{w_ns}}}id', str(comment_id))
    paragraph.insert(start_index + 3, comment_reference)

def add_comment_to_comments_xml(comments_root, comment_id, comment_text, author, namespaces):
    w_ns = namespaces['w']
    comment = ET.SubElement(comments_root, f'{{{w_ns}}}comment')
    comment.set(f'{{{w_ns}}}id', str(comment_id))
    comment.set(f'{{{w_ns}}}author', author)
    comment.set(f'{{{w_ns}}}date', datetime.now().isoformat())

    para = ET.SubElement(comment, f'{{{w_ns}}}p')
    run = ET.SubElement(para, f'{{{w_ns}}}r')
    text = ET.SubElement(run, f'{{{w_ns}}}t')
    text.text = comment_text

def add_comment_to_comments_extended(comments_extended_root, comment_id, namespaces):
    w15_ns = namespaces['w15']
    comment_ex = ET.SubElement(comments_extended_root, f'{{{w15_ns}}}commentEx')
    comment_ex.set(f'{{{w15_ns}}}paraId', str(uuid.uuid4()))
    comment_ex.set(f'{{{w15_ns}}}paraIdParent', "00000000")
    comment_ex.set(f'{{{w15_ns}}}done', "0")

def add_person_to_people_xml(people_root, author, namespaces):
    w_ns = namespaces['w']
    person = ET.SubElement(people_root, f'{{{w_ns}}}person')
    person_author = ET.SubElement(person, f'{{{w_ns}}}author')
    person_author.text = author

def update_content_types(temp_dir):
    content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
    tree = ET.parse(content_types_path)
    root = tree.getroot()

    needed_types = [
        ('comments.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'),
        ('commentsExtended.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml'),
        ('people.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml')
    ]

    for file_name, content_type in needed_types:
        if not root.find(f".//*[@PartName='/word/{file_name}']"):
            override = ET.SubElement(root, 'Override')
            override.set('PartName', f'/word/{file_name}')
            override.set('ContentType', content_type)

    tree.write(content_types_path, encoding='UTF-8', xml_declaration=True)

def update_document_rels(temp_dir):
    rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
    tree = ET.parse(rels_path)
    root = tree.getroot()

    needed_rels = [
        ('comments.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'),
        ('commentsExtended.xml', 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended'),
        ('people.xml', 'http://schemas.microsoft.com/office/2011/relationships/people')
    ]

    for target, type_value in needed_rels:
        if not root.find(f".//*[@Target='{target}']"):
            relationship = ET.SubElement(root, 'Relationship')
            relationship.set('Type', type_value)
            relationship.set('Target', target)
            relationship.set('Id', f'rId{len(root) + 1}')

    tree.write(rels_path, encoding='UTF-8', xml_declaration=True)

def update_settings(temp_dir):
    settings_path = os.path.join(temp_dir, 'word', 'settings.xml')
    if os.path.exists(settings_path):
        tree = ET.parse(settings_path)
        root = tree.getroot()
        
        w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        if not root.find(f'.//{w_ns}commentsEx'):
            comments_ex = ET.SubElement(root, f'{w_ns}commentsEx')
        
        tree.write(settings_path, encoding='UTF-8', xml_declaration=True)
# Uso de la función
input_docx = 'EP3567950-B1__seprotec_es.docx'
comments_json = 'errors.json'
output_docx = 'EP3567950-B1__seprotec_es_suggestions.docx'

add_comments_to_docx(input_docx, comments_json, output_docx)

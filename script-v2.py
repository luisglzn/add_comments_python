import xml.etree.ElementTree as ET
import zipfile
import json
import uuid
import os
import shutil
from datetime import datetime


def save_xml_structure_to_file(root, file_path):
    """Función para guardar la estructura del XML en un archivo."""
    with open(file_path, 'w', encoding='utf-8') as file:
        def write_structure(node, level=0):
            indent = "  " * level
            file.write(f"{indent}<{node.tag}>\n")
            for child in node:
                write_structure(child, level + 1)
            file.write(f"{indent}</{node.tag}>\n")
        write_structure(root)

def add_comments_to_docx(docx_path, json_path, output_path):
    # Crear una copia temporal del archivo DOCX
    temp_dir = "temp_docx"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.mkdir(temp_dir)
    
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    
    # Cargar los comentarios del archivo JSON
    with open(json_path, 'r') as json_file:
        comments_data = json.load(json_file)
    
    # Modificar document.xml
    document_tree = ET.parse(os.path.join(temp_dir, 'word', 'document.xml'))
    document_root = document_tree.getroot()

    # Imprimir la estructura inicial de document.xml
    print("Estructura inicial de document.xml:")
    save_xml_structure_to_file(document_root, "estructura_inicial.xml")
    
    # Modificar o crear comments.xml
    comments_path = os.path.join(temp_dir, 'word', 'comments.xml')
    if not os.path.exists(comments_path):
        comments_root = ET.Element('w:comments', {
            'xmlns:w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            'xmlns:w14': "http://schemas.microsoft.com/office/word/2010/wordml",
            'xmlns:w15': "http://schemas.microsoft.com/office/word/2012/wordml",
            'xmlns:w16se': "http://schemas.microsoft.com/office/word/2015/wordml/symex",
            'xmlns:wp14': "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
        })
    else:
        comments_tree = ET.parse(comments_path)
        comments_root = comments_tree.getroot()
    
    # Modificar o crear commentsExtended.xml
    comments_extended_path = os.path.join(temp_dir, 'word', 'commentsExtended.xml')
    if not os.path.exists(comments_extended_path):
        comments_extended_root = ET.Element('w15:commentsEx', {'xmlns:w15': "http://schemas.microsoft.com/office/word/2012/wordml"})
    else:
        comments_extended_tree = ET.parse(comments_extended_path)
        comments_extended_root = comments_extended_tree.getroot()
    
    # Modificar o crear commentsIds.xml
    comments_ids_path = os.path.join(temp_dir, 'word', 'commentsIds.xml')
    if not os.path.exists(comments_ids_path):
        comments_ids_root = ET.Element('w16cid:commentsIds', {'xmlns:w16cid': "http://schemas.microsoft.com/office/word/2016/wordml/cid"})
    else:
        comments_ids_tree = ET.parse(comments_ids_path)
        comments_ids_root = comments_ids_tree.getroot()
    
    # Modificar o crear people.xml
    people_path = os.path.join(temp_dir, 'word', 'people.xml')
    if not os.path.exists(people_path):
        people_root = ET.Element('w:people', {'xmlns:w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
    else:
        people_tree = ET.parse(people_path)
        people_root = people_tree.getroot()
    
    def find_text_in_document(root, text):
        for elem in root.iter():
            if elem.tag.endswith('}t') and elem.text and text.lower() in elem.text.lower():
                return elem
        return None
    
    # Añadir comentarios
    for comment_data in comments_data:
        quote = comment_data['quote']
        comment_text = comment_data['comment']
        author = comment_data['author']
        
        # Generar IDs únicos
        comment_id = str(uuid.uuid4())
        person_id = str(uuid.uuid4())
        
        # Encontrar el texto en el documento
        text_elem = find_text_in_document(document_root, quote)
        if text_elem is None:
            print(f"No se encontró el texto: {quote}")
            continue
        
        # Añadir marcadores de comentario en document.xml
        parent_map = {c: p for p in document_root.iter() for c in p}
        parent = parent_map.get(text_elem)
        if parent is None:
            print(f"No se pudo encontrar el padre del elemento de texto para: {quote}")
            continue
        
        index = list(parent).index(text_elem)
        
        comment_range_start = ET.Element('w:commentRangeStart', {'w:id': comment_id})
        comment_range_end = ET.Element('w:commentRangeEnd', {'w:id': comment_id})
        comment_reference = ET.Element('w:commentReference', {'w:id': comment_id})
        
        parent.insert(index, comment_range_start)
        parent.insert(index + 2, comment_range_end)
        parent.insert(index + 3, comment_reference)
        
        # Añadir comentario en comments.xml
        comment = ET.SubElement(comments_root, 'w:comment', {
            'w:id': comment_id,
            'w:author': author,
            'w:date': datetime.now().isoformat(),
            'w:initials': author[:2].upper()
        })
        para = ET.SubElement(comment, 'w:p')
        pPr = ET.SubElement(para, 'w:pPr')
        ET.SubElement(pPr, 'w:pStyle', {'w:val': 'CommentText'})
        run = ET.SubElement(para, 'w:r')
        rPr = ET.SubElement(run, 'w:rPr')
        ET.SubElement(rPr, 'w:rStyle', {'w:val': 'CommentReference'})
        ET.SubElement(run, 'w:annotationRef')
        run = ET.SubElement(para, 'w:r')
        text = ET.SubElement(run, 'w:t')
        text.text = comment_text
        
        # Añadir entrada en commentsExtended.xml
        comment_extended = ET.SubElement(comments_extended_root, 'w15:commentEx', {'w15:paraId': str(uuid.uuid4()), 'w15:done': "0"})
        
        # Añadir entrada en commentsIds.xml
        comment_id_entry = ET.SubElement(comments_ids_root, 'w16cid:commentId', {'w16cid:id': comment_id})
        
        # Añadir autor en people.xml
        person = ET.SubElement(people_root, 'w:person', {'w:author': author})
        person_data = ET.SubElement(person, 'w:presenceInfo', {'w:providerId': "None", 'w:userId': person_id})
    
    # Imprimir la estructura final de document.xml
    print("Estructura final de document.xml:")
    save_xml_structure_to_file(document_root, "estructura_final.xml")
    

    # Guardar los archivos modificados
    document_tree.write(os.path.join(temp_dir, 'word', 'document.xml'), encoding='UTF-8', xml_declaration=True)
    ET.ElementTree(comments_root).write(os.path.join(temp_dir, 'word', 'comments.xml'), encoding='UTF-8', xml_declaration=True)
    ET.ElementTree(comments_extended_root).write(os.path.join(temp_dir, 'word', 'commentsExtended.xml'), encoding='UTF-8', xml_declaration=True)
    ET.ElementTree(comments_ids_root).write(os.path.join(temp_dir, 'word', 'commentsIds.xml'), encoding='UTF-8', xml_declaration=True)
    ET.ElementTree(people_root).write(os.path.join(temp_dir, 'word', 'people.xml'), encoding='UTF-8', xml_declaration=True)
    
    # Actualizar [Content_Types].xml
    content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
    content_types_tree = ET.parse(content_types_path)
    content_types_root = content_types_tree.getroot()
    
    # Añadir tipos de contenido si no existen
    types_to_add = [
        ('comments.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'),
        ('commentsExtended.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml'),
        ('commentsIds.xml', 'application/vnd.microsoft.word.comments.ids+xml'),
        ('people.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml')
    ]
    
    for file_name, content_type in types_to_add:
        if not any(override.get('PartName') == f'/word/{file_name}' for override in content_types_root.findall('{*}Override')):
            ET.SubElement(content_types_root, 'Override', {'PartName': f'/word/{file_name}', 'ContentType': content_type})
    
    content_types_tree.write(content_types_path, encoding='UTF-8', xml_declaration=True)
    
    # Actualizar word/_rels/document.xml.rels
    rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
    rels_tree = ET.parse(rels_path)
    rels_root = rels_tree.getroot()
    
    # Añadir relaciones si no existen
    rels_to_add = [
        ('comments.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'),
        ('commentsExtended.xml', 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended'),
        ('commentsIds.xml', 'http://schemas.microsoft.com/office/2016/relationships/commentsIds'),
        ('people.xml', 'http://schemas.microsoft.com/office/2011/relationships/people')
    ]
    
    for file_name, rel_type in rels_to_add:
        if not any(rel.get('Target') == file_name for rel in rels_root.findall('{*}Relationship')):
            ET.SubElement(rels_root, 'Relationship', {
                'Id': f'rId{len(rels_root) + 1}',
                'Type': rel_type,
                'Target': file_name
            })
    
    rels_tree.write(rels_path, encoding='UTF-8', xml_declaration=True)
    
    # Crear el nuevo archivo DOCX
    with zipfile.ZipFile(output_path, 'w') as zipf:
        for root, _, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, temp_dir)
                zipf.write(file_path, arcname)
    
    # Limpiar archivos temporales
    shutil.rmtree(temp_dir)

# Uso de la función
add_comments_to_docx('EP3567950-B1__seprotec_es.docx', 'errors.json', 'EP3567950-B1__seprotec_es_suggestions.docx')

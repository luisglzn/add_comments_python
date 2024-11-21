import xml.etree.ElementTree as ET
import zipfile
import json
import uuid
import os
import shutil
from datetime import datetime

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
    
    # Modificar o crear comments.xml
    comments_path = os.path.join(temp_dir, 'word', 'comments.xml')
    if not os.path.exists(comments_path):
        comments_root = ET.Element('w:comments', {'xmlns:w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
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
    
    # Modificar o crear commentsExtensible.xml
    comments_extensible_path = os.path.join(temp_dir, 'word', 'commentsExtensible.xml')
    if not os.path.exists(comments_extensible_path):
        comments_extensible_root = ET.Element('w16cex:commentsExtensible', {'xmlns:w16cex': "http://schemas.microsoft.com/office/word/2018/wordml/cex"})
    else:
        comments_extensible_tree = ET.parse(comments_extensible_path)
        comments_extensible_root = comments_extensible_tree.getroot()
    
    # Modificar o crear people.xml
    people_path = os.path.join(temp_dir, 'word', 'people.xml')
    if not os.path.exists(people_path):
        people_root = ET.Element('w:people', {'xmlns:w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
    else:
        people_tree = ET.parse(people_path)
        people_root = people_tree.getroot()
    
    # Función para encontrar el texto en el documento
    def find_text_in_document(root, text):
        for elem in root.iter():
            if elem.tag.endswith('}t') and elem.text and text in elem.text:
                return elem
        return None
    
    def find_parent(root, element):
        for parent in root.iter():
            for child in parent:
                if child == element:
                    return parent
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
        parent = find_parent(document_root, text_elem)
        index = list(parent).index(text_elem)
        
        comment_range_start = ET.Element('w:commentRangeStart', {'w:id': comment_id})
        comment_range_end = ET.Element('w:commentRangeEnd', {'w:id': comment_id})
        comment_reference = ET.Element('w:commentReference', {'w:id': comment_id})
        
        parent.insert(index, comment_range_start)
        parent.insert(index + 2, comment_range_end)
        parent.insert(index + 3, comment_reference)
        
        # Añadir comentario en comments.xml
        comment = ET.SubElement(comments_root, 'w:comment', {'w:id': comment_id, 'w:author': author, 'w:date': datetime.now().isoformat()})
        para = ET.SubElement(comment, 'w:p')
        run = ET.SubElement(para, 'w:r')
        text = ET.SubElement(run, 'w:t')
        text.text = comment_text
        
        # Añadir entrada en commentsExtended.xml
        comment_extended = ET.SubElement(comments_extended_root, 'w15:commentEx', {'w15:paraId': str(uuid.uuid4()), 'w15:done': "0"})
        
        # Añadir entrada en commentsExtensible.xml
        comment_extensible = ET.SubElement(comments_extensible_root, 'w16cex:commentExtensible', {'w16cex:paraId': str(uuid.uuid4())})
        
        # Añadir autor en people.xml
        person = ET.SubElement(people_root, 'w:person', {'w:author': author})
        person_data = ET.SubElement(person, 'w:presenceInfo', {'w:providerId': "None", 'w:userId': person_id})
    
    # Guardar los archivos modificados
    document_tree.write(os.path.join(temp_dir, 'word', 'document.xml'), encoding='UTF-8', xml_declaration=True)
    ET.ElementTree(comments_root).write(os.path.join(temp_dir, 'word', 'comments.xml'), encoding='UTF-8', xml_declaration=True)
    ET.ElementTree(comments_extended_root).write(os.path.join(temp_dir, 'word', 'commentsExtended.xml'), encoding='UTF-8', xml_declaration=True)
    ET.ElementTree(comments_extensible_root).write(os.path.join(temp_dir, 'word', 'commentsExtensible.xml'), encoding='UTF-8', xml_declaration=True)
    ET.ElementTree(people_root).write(os.path.join(temp_dir, 'word', 'people.xml'), encoding='UTF-8', xml_declaration=True)
    
    # Actualizar [Content_Types].xml
    content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
    content_types_tree = ET.parse(content_types_path)
    content_types_root = content_types_tree.getroot()
    
    # Añadir tipos de contenido si no existen
    types_to_add = [
        ('comments.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'),
        ('commentsExtended.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml'),
        ('commentsExtensible.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml'),
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
        ('commentsExtensible.xml', 'http://schemas.microsoft.com/office/2018/relationships/commentsExtensible'),
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

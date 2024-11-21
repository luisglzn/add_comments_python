import xml.etree.ElementTree as ET
import json
import zipfile
import os
import shutil
import uuid

def add_comments_to_docx(input_docx, comments_json, output_docx):
    # Crear un directorio temporal para trabajar con los archivos
    temp_dir = "temp_docx"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.mkdir(temp_dir)

    # Extraer el contenido del archivo .docx
    with zipfile.ZipFile(input_docx, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # Cargar los comentarios desde el archivo JSON
    with open(comments_json, 'r') as f:
        comments_data = json.load(f)

    # Procesar el archivo document.xml
    document_tree = ET.parse(os.path.join(temp_dir, 'word', 'document.xml'))
    document_root = document_tree.getroot()

    # Crear o actualizar los archivos relacionados con los comentarios
    comments_tree, comments_root = create_or_update_xml(temp_dir, 'word/comments.xml')
    comments_extended_tree, comments_extended_root = create_or_update_xml(temp_dir, 'word/commentsExtended.xml')
    people_tree, people_root = create_or_update_xml(temp_dir, 'word/people.xml')

    # Namespaces necesarios
    namespaces = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    }

    for ns, uri in namespaces.items():
        ET.register_namespace(ns, uri)

    # Procesar cada comentario
    for comment_data in comments_data:
        quote = comment_data['quote']
        comment_text = comment_data['comment']
        author = comment_data['author']

        # Buscar el texto en el documento
        for paragraph in document_root.findall('.//w:p', namespaces):
            text_elements = paragraph.findall('.//w:t', namespaces)
            paragraph_text = ''.join(elem.text for elem in text_elements if elem.text)
            
            if quote in paragraph_text:
                # Generar un ID único para el comentario
                comment_id = str(uuid.uuid4())[:8]

                # Añadir el marcador de comentario en el documento
                start_range = ET.SubElement(paragraph, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeStart')
                start_range.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', comment_id)

                end_range = ET.SubElement(paragraph, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeEnd')
                end_range.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', comment_id)

                reference = ET.SubElement(paragraph, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                comment_reference = ET.SubElement(reference, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentReference')
                comment_reference.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', comment_id)

                # Añadir el comentario al archivo comments.xml
                comment = ET.SubElement(comments_root, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comment')
                comment.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', comment_id)
                comment.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', author)
                comment.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '2024-11-21T10:00:00Z')

                comment_para = ET.SubElement(comment, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
                comment_run = ET.SubElement(comment_para, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                comment_text_elem = ET.SubElement(comment_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                comment_text_elem.text = comment_text

                # Añadir entrada en commentsExtended.xml
                comment_extended = ET.SubElement(comments_extended_root, '{http://schemas.microsoft.com/office/word/2012/wordml}commentEx')
                comment_extended.set('{http://schemas.microsoft.com/office/word/2012/wordml}paraId', str(uuid.uuid4()))
                comment_extended.set('{http://schemas.microsoft.com/office/word/2012/wordml}done', "0")

                # Añadir autor a people.xml si no existe
                person_exists = people_root.find(f".//w:person[w:author='{author}']", namespaces)
                if person_exists is None:
                    person = ET.SubElement(people_root, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}person')
                    person_author = ET.SubElement(person, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author')
                    person_author.text = author

                break  # Salir después de encontrar y procesar la primera ocurrencia

    # Guardar los cambios en los archivos XML
    document_tree.write(os.path.join(temp_dir, 'word', 'document.xml'), encoding='UTF-8', xml_declaration=True)
    comments_tree.write(os.path.join(temp_dir, 'word', 'comments.xml'), encoding='UTF-8', xml_declaration=True)
    comments_extended_tree.write(os.path.join(temp_dir, 'word', 'commentsExtended.xml'), encoding='UTF-8', xml_declaration=True)
    people_tree.write(os.path.join(temp_dir, 'word', 'people.xml'), encoding='UTF-8', xml_declaration=True)

    # Actualizar [Content_Types].xml para incluir los nuevos archivos
    update_content_types(temp_dir)

    # Actualizar word/_rels/document.xml.rels
    update_document_rels(temp_dir)

    # Crear el nuevo archivo .docx
    shutil.make_archive(output_docx, 'zip', temp_dir)
    os.rename(output_docx + '.zip', output_docx)

    # Limpiar el directorio temporal
    shutil.rmtree(temp_dir)

def create_or_update_xml(temp_dir, file_path):
    full_path = os.path.join(temp_dir, file_path)
    if os.path.exists(full_path):
        tree = ET.parse(full_path)
        root = tree.getroot()
    else:
        root = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comments')
        tree = ET.ElementTree(root)
    return tree, root

def update_content_types(temp_dir):
    content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
    tree = ET.parse(content_types_path)
    root = tree.getroot()

    # Asegurarse de que los tipos de contenido necesarios estén presentes
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

    # Asegurarse de que las relaciones necesarias estén presentes
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

# Uso de la función
input_docx = 'EP3567950-B1__seprotec_es.docx'
comments_json = 'errors.json'
output_docx = 'EP3567950-B1__seprotec_es_suggestions.docx'

add_comments_to_docx(input_docx, comments_json, output_docx)

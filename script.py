import zipfile
import xml.etree.ElementTree as ET
import uuid
import tempfile
import uuid
import os
from datetime import datetime

def add_comment_to_docx(docx_path, comment_text, author, paragraph_index):
    # Generar un ID único para el comentario
    comment_id = str(uuid.uuid4())
    # Crear un directorio temporal único
    with tempfile.TemporaryDirectory() as temp_dir:
        # Paso 1: Descomprimir el archivo DOCX
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Paso 2: Crear los archivos de comentarios
        create_comments_xml(temp_dir, comment_text, author, comment_id)
        create_comments_extended_xml(temp_dir, comment_id)
        create_comments_extensible_xml(temp_dir, comment_id)
        create_comments_ids_xml(temp_dir, comment_id)

        # Paso 3: Modificar los archivos existentes
        update_content_types_xml(temp_dir)
        update_document_xml_rels(temp_dir)
        update_document_xml(temp_dir, paragraph_index, comment_id)

        # Paso 4: Volver a comprimir el archivo DOCX
        new_docx_path = docx_path.replace('.docx', '_with_comments.docx')
        with zipfile.ZipFile(new_docx_path, 'w') as zipf:
            for foldername, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, arcname)

    print(f"Documento con comentarios creado: {new_docx_path}")

def create_comments_xml(temp_dir, comment_text, author, comment_id):
    now = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    
    root = ET.Element("w:comments")
    root.set("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    
    comment = ET.SubElement(root, "w:comment")
    comment.set("w:id", comment_id)
    comment.set("w:author", author)
    comment.set("w:date", now)
    
    p = ET.SubElement(comment, "w:p")
    r = ET.SubElement(p, "w:r")
    t = ET.SubElement(r, "w:t")
    t.text = comment_text
    
    tree = ET.ElementTree(root)
    tree.write(os.path.join(temp_dir, "word/comments.xml"), encoding="UTF-8", xml_declaration=True)


def create_comments_extended_xml(temp_dir, comment_id):
    root = ET.Element("w15:commentsEx")
    root.set("xmlns:w15", "http://schemas.microsoft.com/office/word/2012/wordml")
    
    comment_ex = ET.SubElement(root, "w15:commentEx")
    comment_ex.set("w15:paraId", str(uuid.uuid4()))  # Genera un paraId único
    comment_ex.set("w15:done", "0")
    comment_ex.set("w:id", comment_id)
    
    tree = ET.ElementTree(root)
    tree.write(os.path.join(temp_dir, "word/commentsExtended.xml"), encoding="UTF-8", xml_declaration=True)

def create_comments_extensible_xml(temp_dir, comment_id):
    root = ET.Element("w16cex:commentsExtensible")
    root.set("xmlns:w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex")
    
    comment_extensible = ET.SubElement(root, "w16cex:commentExtensible")
    comment_extensible.set("w16cex:paraId", str(uuid.uuid4()))  # Genera un paraId único
    comment_extensible.set("w:id", comment_id)
    
    tree = ET.ElementTree(root)
    tree.write(os.path.join(temp_dir, "word/commentsExtensible.xml"), encoding="UTF-8", xml_declaration=True)

def create_comments_ids_xml(temp_dir, comment_id):
    root = ET.Element("w16cid:commentsIds")
    root.set("xmlns:w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid")
    
    comment_id_element = ET.SubElement(root, "w16cid:commentId")
    comment_id_element.set("w16cid:id", comment_id)
    
    tree = ET.ElementTree(root)
    tree.write(os.path.join(temp_dir, "word/commentsIds.xml"), encoding="UTF-8", xml_declaration=True)

def update_content_types_xml(temp_dir):
    tree = ET.parse(f"{temp_dir}/[Content_Types].xml")
    root = tree.getroot()
    
    new_overrides = [
        ("application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml", "/word/comments.xml"),
        ("application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml", "/word/commentsExtended.xml"),
        ("application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml", "/word/commentsExtensible.xml"),
        ("application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml", "/word/commentsIds.xml")
    ]
    
    for content_type, part_name in new_overrides:
        override = ET.Element("Override")
        override.set("ContentType", content_type)
        override.set("PartName", part_name)
        root.append(override)
    
    tree.write(f"{temp_dir}/[Content_Types].xml", encoding="UTF-8", xml_declaration=True)

def update_document_xml_rels(temp_dir):
    tree = ET.parse(os.path.join(temp_dir, "word/_rels/document.xml.rels"))
    root = tree.getroot()
    
    new_relationships = [
        ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", "comments.xml"),
        ("http://schemas.microsoft.com/office/2011/relationships/commentsExtended", "commentsExtended.xml"),
        ("http://schemas.microsoft.com/office/2018/wordml/cex", "commentsExtensible.xml"),
        ("http://schemas.microsoft.com/office/2016/wordml/cid", "commentsIds.xml")
    ]
    
    existing_ids = set(rel.get("Id") for rel in root)
    next_id = max(int(id[3:]) for id in existing_ids if id.startswith("rId")) + 1 if existing_ids else 1
    
    for rel_type, target in new_relationships:
        if not any(rel.get("Type") == rel_type for rel in root):
            relationship = ET.Element("Relationship")
            relationship.set("Type", rel_type)
            relationship.set("Target", target)
            relationship.set("Id", f"rId{next_id}")
            root.append(relationship)
            next_id += 1
    
    tree.write(os.path.join(temp_dir, "word/_rels/document.xml.rels"), encoding="UTF-8", xml_declaration=True)

def update_document_xml(temp_dir, paragraph_index, comment_id):
    tree = ET.parse(os.path.join(temp_dir, "word/document.xml"))
    root = tree.getroot()
    
    body = root.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body")
    paragraphs = body.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")
    
    if paragraph_index < len(paragraphs):
        paragraph = paragraphs[paragraph_index]
        
        comment_range_start = ET.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeStart")
        comment_range_start.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id", comment_id)
        
        comment_range_end = ET.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeEnd")
        comment_range_end.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id", comment_id)
        
        comment_reference = ET.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r")
        comment_reference_element = ET.SubElement(comment_reference, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentReference")
        comment_reference_element.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id", comment_id)
        
        # Insertar los elementos de comentario alrededor del contenido existente
        paragraph.insert(0, comment_range_start)
        paragraph.append(comment_range_end)
        paragraph.append(comment_reference)
    
    tree.write(os.path.join(temp_dir, "word/document.xml"), encoding="UTF-8", xml_declaration=True)

# Uso de la función
add_comment_to_docx(r"C:\Users\luisg\OneDrive\Escritorio\Trabajo\Add comments\add_comments_python\Script_TestFile.docx", "Este es un comentario de prueba", "Seprotec.AI", 0)

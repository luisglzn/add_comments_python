from docx import Document
from docx.oxml.shared import qn
from docx.oxml import OxmlElement
from datetime import datetime
import re
from lxml import etree

def split_xml_by_elements(xml):
    root = etree.fromstring(xml)
    paragraphs = root.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    return [etree.tostring(p, encoding='unicode') for p in paragraphs]

def build_txt(paragraphs):
    return [re.sub('<.*?>', '', p) for p in paragraphs]

def add_comment_to_phrase(doc_path, phrase, comment_text, author="Anonymous"):
    doc = Document(doc_path)
    xml = doc.element.xml
    paragraphs = split_xml_by_elements(xml)
    txt = build_txt(paragraphs)
    
    comment_count = 0
    phrase_regex = re.compile(re.escape(phrase), re.IGNORECASE)
    
    # Obtener el último ID de comentario utilizado
    last_comment_id = get_last_comment_id(doc)
    
    for i, paragraph in enumerate(paragraphs):
        paragraph_text = txt[i]
        if phrase_regex.search(paragraph_text):
            modified_paragraph, comment = add_comment_to_paragraph_end(
                paragraph, comment_text, author, last_comment_id + comment_count + 1
            )
            # Modificar directamente el XML del párrafo
            update_paragraph_xml(doc, i, modified_paragraph)
            add_comment_to_document(doc, comment)
            comment_count += 1
    
    # Guardar el documento en el archivo original
    doc.save(doc_path)
    print(f"Documento modificado con {comment_count} comentarios")

def get_last_comment_id(doc):
    comments_part = doc.part.comments_part
    if comments_part is None:
        return 0
    comments = comments_part._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comment')
    if not comments:
        return 0
    last_comment = comments[-1]
    return int(last_comment.get(qn('w:id')))

def add_comment_to_paragraph_end(paragraph, comment_text, author, comment_id):
    root = etree.fromstring(paragraph)
    
    comment_range_start = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeStart')
    comment_range_start.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(comment_id))
    
    comment_range_end = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeEnd')
    comment_range_end.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(comment_id))
    
    comment_reference = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentReference')
    comment_reference.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(comment_id))
    
    # Insertar los elementos de comentario al final del párrafo
    root.append(comment_range_start)
    root.append(comment_range_end)
    root.append(comment_reference)
    
    comment = create_comment(str(comment_id), author, comment_text)
    
    return etree.tostring(root, encoding='unicode'), comment

def update_paragraph_xml(doc, index, modified_paragraph):
    # Actualizar el XML del párrafo en el documento
    body = doc.element.body
    paragraph_elements = body.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    if index < len(paragraph_elements):
        paragraph_elements[index].clear()
        new_paragraph_element = etree.fromstring(modified_paragraph)
        for child in new_paragraph_element:
            paragraph_elements[index].append(child)

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

# Ejemplo de uso
if __name__ == "__main__":
    doc_path = r"C:\Users\luisg\OneDrive\Escritorio\Trabajo\Add comments\add_comments_python\translated - copia.docx"
    phrase = "CLDN18"
    comment_text = "Prueba de comentario, se ha probado CLDN18"
    author = "Luis"
    add_comment_to_phrase(doc_path, phrase, comment_text, author)
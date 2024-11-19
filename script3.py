from docx import Document
from docx.oxml.shared import qn
from docx.oxml import OxmlElement
from datetime import datetime
from docx.enum.text import WD_COLOR_INDEX
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
    phrase_regex = re.compile(r'(^|\s)' + re.escape(phrase) + r'(?=\s|$|\.|,)', re.IGNORECASE)
    
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
    first_occurrence = True
    
    for elem in root.iter():
        if elem.tag.endswith('t') and elem.text and phrase in elem.text:
            # Usar una expresión regular para encontrar las posiciones de la frase
            pattern = re.escape(phrase)
            matches = list(re.finditer(pattern, elem.text))
            
            # Si no hay coincidencias, continuar
            if not matches:
                continue
            
            # Crear una lista para almacenar los nuevos elementos
            new_elements = []
            last_end = 0
            
            for match in matches:
                start, end = match.span()
                
                # Añadir texto antes de la coincidencia
                if last_end < start:
                    before_text = elem.text[last_end:start]
                    if before_text:
                        before_run = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                        before_text_elem = etree.SubElement(before_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                        before_text_elem.text = before_text
                        new_elements.append(before_run)
                
                # Añadir la frase exacta
                phrase_run = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                phrase_text = etree.SubElement(phrase_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                phrase_text.text = elem.text[start:end]
                
                # Añadir comentario solo en la primera ocurrencia
                if first_occurrence:
                    comment_range_start = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeStart')
                    comment_range_start.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(comment_id))
                    new_elements.append(comment_range_start)

                    new_elements.append(phrase_run)

                    comment_range_end = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeEnd')
                    comment_range_end.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(comment_id))
                    new_elements.append(comment_range_end)

                    comment_reference = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentReference')
                    comment_reference.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(comment_id))
                    new_elements.append(comment_reference)
                    
                    first_occurrence = False
                else:
                    new_elements.append(phrase_run)
                
                last_end = end
            
            # Añadir texto después de la última coincidencia
            if last_end < len(elem.text):
                after_text = elem.text[last_end:]
                if after_text:
                    after_run = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                    after_text_elem = etree.SubElement(after_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                    after_text_elem.text = after_text
                    new_elements.append(after_run)
            
            # Reemplazar el elemento original con los nuevos elementos
            elem.clear()
            for new_elem in new_elements:
                elem.addnext(new_elem)
                elem = new_elem

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

def highlight_repeated_phrases(doc_path, phrase):
    doc = Document(doc_path)
    phrase_regex = re.compile(rf'\b{re.escape(phrase)}\b', re.IGNORECASE)
    
    for paragraph in doc.paragraphs:
        if phrase_regex.search(paragraph.text):
            first_occurrence = True
            for run in paragraph.runs:
                if phrase_regex.search(run.text):
                    if not first_occurrence:
                        # Resaltar las ocurrencias adicionales
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    else:
                        first_occurrence = False

    # Guardar el documento en el archivo original
    doc.save(doc_path)
    print("Documento modificado con resaltado en las ocurrencias adicionales")

# Ejemplo de uso
if __name__ == "__main__":
    doc_path =r"C:\Users\luisg\OneDrive\Escritorio\Trabajo\Add comments\add_comments_python\translated - copia.docx"
    phrase = "CLDN18"
    #highlight_repeated_phrases(doc_path, phrase)
    comment_text = "Prueba de comentario, se ha probado CLDN18"
    author = "Luis"
    add_comment_to_phrase(doc_path, phrase, comment_text, author)
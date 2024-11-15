from docx import Document
from docx.oxml.shared import qn
from docx.oxml import OxmlElement
from datetime import datetime
import random
import re
import numpy as np
import new

def comment(doc_path):
    doc = Document(doc_path)
    comments_dict = {}
    comments_of_dict = {}
    author_dict = {}
    date_dict = {}

    # Obtener la parte de comentarios
    comments_part = doc.part.comments_part
    if comments_part is None:
        return []  # No hay comentarios en el documento

    # Iterar sobre los comentarios
    for comment in comments_part.element.findall(qn('w:comment')):
        comment_id = comment.get(qn('w:id'))
        # Obtener el texto del comentario
        comment_text = ''
        for paragraph in comment.findall(qn('w:p')):
            for run in paragraph.findall(qn('w:r')):
                for text_node in run.findall(qn('w:t')):
                    if text_node.text:
                        comment_text += text_node.text
        comments_dict[comment_id] = comment_text
        author_dict[comment_id] = comment.get(qn('w:author'))
        date_dict[comment_id] = comment.get(qn('w:date'))

    # Iterar sobre el contenido del documento para encontrar el texto comentado
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            comment_reference = run._element.find(qn('w:commentReference'))
            if comment_reference is not None:
                comment_id = comment_reference.get(qn('w:id'))
                comments_of_dict[comment_id] = paragraph.text

    data = []
    for k in set(comments_dict.keys()) & set(comments_of_dict.keys()) & set(author_dict.keys()) & set(date_dict.keys()):
        data.append({
            "comment": comments_dict[k],
            "text_selected": comments_of_dict[k],
            "commented_on": date_dict[k],
            "author": author_dict[k]
        })

    return data

def add_comment_to_document(doc_path, comment_text, author="Anonymous"):
    doc = Document(doc_path)
    paragraph = doc.paragraphs[0]
    run = paragraph.add_run()
    
    # Generar un ID único para el comentario
    comment_id = str(random.randint(0, 9999))
    
    # Crear el elemento de referencia del comentario
    comment_reference = OxmlElement('w:commentReference')
    comment_reference.set(qn('w:id'), comment_id)
    run._element.append(comment_reference)
    
    # Crear el comentario
    comment = OxmlElement('w:comment')
    comment.set(qn('w:id'), comment_id)
    comment.set(qn('w:author'), author)
    comment.set(qn('w:date'), datetime.now().isoformat())
    comment.set(qn('w:initials'), ''.join([name[0].upper() for name in author.split() if name]))
    
    # Añadir el texto del comentario
    p = OxmlElement('w:p')
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = comment_text
    r.append(t)
    p.append(r)
    comment.append(p)
    
    # Añadir el comentario al documento
    comments_part = doc.part.comments_part
    if comments_part is None:
        comments_part = doc.part.add_comments_part()
    comments_part._element.append(comment)
    
    doc.save(doc_path)

def add_comment_to_phrase(doc_path, phrase, comment_text, author="Anonymous"):
    doc = Document(doc_path)
    xml = doc.element.xml
    paragraphs = new.split_xml_by_elements(xml)
    txt_with_tags, tags_list = new.replace_tags(paragraphs, '<#>')
    txt = new.build_txt(paragraphs)
    print("Phrase: ", phrase)
    print("Txt with tags: ", txt_with_tags)
    comment_count = 0
    # Crear una lista de párrafos del documento
    doc_paragraphs = list(doc.paragraphs)
    for i, paragraph_text in enumerate(txt_with_tags):
        # Buscar todas las ocurrencias de la frase en el párrafo
        start = 0
        while True:
            loc = new.localize_substring_ignoring_separator(paragraph_text[start:], phrase)
            if not loc:
                break
            start_index, end_index = loc
            start_index += start  # Ajustar el índice de inicio para la posición real en el párrafo
            end_index += start
            print(f"Found phrase in paragraph {i} at positions: {start_index}, {end_index}")
            # Generate a unique ID for the comment
            comment_id = str(random.randint(0, 9999))
            # Create comment range start element
            comment_range_start = OxmlElement('w:commentRangeStart')
            comment_range_start.set(qn('w:id'), comment_id)
            # Create comment range end element
            comment_range_end = OxmlElement('w:commentRangeEnd')
            comment_range_end.set(qn('w:id'), comment_id)
            # Create comment reference element
            comment_reference = OxmlElement('w:commentReference')
            comment_reference.set(qn('w:id'), comment_id)
            # Obtener el párrafo correcto del documento
            p = doc_paragraphs[i]._p
            # Insert elements into the paragraph
            p.insert(start_index, comment_range_start)
            p.insert(end_index + 1, comment_range_end)
            # Insertar el comentario al final del párrafo actual
            last_run = p.xpath('.//w:r')
            if last_run:
                last_run[-1].addnext(comment_reference)
            else:
                # Si no hay runs, añadir el comentario directamente al párrafo
                p.append(comment_reference)
            # Create the comment
            comment = OxmlElement('w:comment')
            comment.set(qn('w:id'), comment_id)
            comment.set(qn('w:author'), author)
            comment.set(qn('w:date'), datetime.now().isoformat())
            comment.set(qn('w:initials'), ''.join([name[0].upper() for name in author.split() if name]))
            # Add comment text
            p = OxmlElement('w:p')
            r = OxmlElement('w:r')
            t = OxmlElement('w:t')
            t.text = comment_text
            r.append(t)
            p.append(r)
            comment.append(p)
            # Add comment to the document
            comments_part = doc.part.comments_part
            if comments_part is None:
                comments_part = doc.part.add_comments_part()
            comments_part._element.append(comment)
            comment_count += 1
            print(f"Comment added to phrase in paragraph {i}")
            # Mover el inicio para la próxima búsqueda
            start = end_index + 1
    doc.save(doc_path)
    print(f"Document saved with {comment_count} comments")

# Ejemplo de uso
doc_path = r"C:\Users\luisg\OneDrive\Escritorio\Trabajo\Add comments\add_comments_python\BQER-AAP-2024.docx"

# Añadir un comentario al documento
#add_comment_to_document(doc_path, "Este es un comentario general del documento", "Usuario1")

# Añadir un comentario a una frase específica
add_comment_to_phrase(doc_path, "préprofessionnelles", "Caso de error en la traducción","Luis")

data = comment(doc_path)

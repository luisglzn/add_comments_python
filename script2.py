import xml.etree.ElementTree as ET
import re

def add_comment(xml_file, search_text, comment_text):
    # Parsear el archivo XML
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Función para buscar y agregar comentarios recursivamente
    def search_and_add_comment(element):
        if element.text and search_text in element.text:
            # Crear el comentario
            comment = ET.Comment(comment_text)
            # Insertar el comentario antes del elemento actual
            parent = element.find("..")
            if parent is not None:
                index = list(parent).index(element)
                parent.insert(index, comment)
        
        # Buscar en los elementos hijos
        for child in element:
            search_and_add_comment(child)

    # Iniciar la búsqueda desde la raíz
    search_and_add_comment(root)

    # Guardar el archivo modificado
    tree.write(xml_file, encoding="utf-8", xml_declaration=True)

# Uso de la función
xml_file = r"C:\Users\luisg\OneDrive\Escritorio\Trabajo\Add comments\add_comments_python\EP3567950-B1__seprotec_es\word\document.xml"
search_text = "procesamiento del grupo de recursos"
comment_text = "Comentario a agregar"
add_comment(xml_file, search_text, comment_text)
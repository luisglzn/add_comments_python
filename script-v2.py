import uno
import json
import os
from com.sun.star.beans import PropertyValue
from com.sun.star.text.TextContentAnchorType import AS_CHARACTER
from com.sun.star.awt import Size


def create_property(name, value):
    property = PropertyValue()
    property.Name = name
    property.Value = value
    return property

def open_document(file_path):
    file_path = os.path.abspath(file_path)
    
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context)
    
    try:
        context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    except:
        os.system('soffice --headless --accept="socket,host=localhost,port=2002;urp;"')
        import time
        time.sleep(3)  # Espera 3 segundos para que LibreOffice se inicie
        context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")

    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    
    url = uno.systemPathToFileUrl(file_path)
    
    properties = (
        create_property("Hidden", True),
        create_property("ReadOnly", False),
        create_property("FilterName", "MS Word 2007 XML")
    )
    
    document = desktop.loadComponentFromURL(url, "_blank", 0, properties)
    return document

def add_comments(document, comments_data):
    print("Iniciando la adición de comentarios")
    text = document.getText()
    cursor = text.createTextCursor()
    
    search = document.createSearchDescriptor()
    
    for item in comments_data:
        quote = item['quote']
        comment = item['comment']
        author = item['author']
        
        print(f"Buscando: '{quote}'")

        search.setSearchString(quote)
        search.SearchCaseSensitive = False
        search.SearchWords = True

        found = document.findFirst(search)
        if found:
            print(f"Encontrado '{quote}'. Añadiendo comentario.")
        else:
            print(f"No se encontró '{quote}'")

        while found:
            # Crear un marco de texto para el comentario
            text_frame = document.createInstance("com.sun.star.text.TextFrame")
            
            # Establecer el tamaño del marco
            size = Size()
            size.Width = 3000
            size.Height = 1500
            text_frame.setSize(size)
            
            text_frame.AnchorType = AS_CHARACTER

            # Mover el cursor al final del texto encontrado
            cursor.gotoRange(found.getEnd(), False)
            
            # Insertar el marco de texto
            text.insertTextContent(cursor, text_frame, False)

            # Añadir el comentario al marco de texto
            frame_text = text_frame.getText()
            frame_cursor = frame_text.createTextCursor()
            frame_text.insertString(frame_cursor, f"{comment}\n\nAutor: {author}", 0)

            # Buscar la siguiente ocurrencia
            found = document.findNext(found.getEnd(), search)

    print("Finalizada la adición de comentarios")

def main():
    # Rutas de los archivos (ajusta estas rutas según tu configuración)
    document_path = r"C:\Users\luisg\OneDrive\Escritorio\Trabajo\Add comments\add_comments_python\EP3567950-B1__seprotec_es.docx"
    json_path = r"C:\Users\luisg\OneDrive\Escritorio\Trabajo\Add comments\add_comments_python\errors.json"

    # Abrir el documento
    document = open_document(document_path)

    # Cargar los datos del JSON
    with open(json_path, 'r', encoding='utf-8') as f:
        comments_data = json.load(f)

    # Añadir los comentarios
    add_comments(document, comments_data)

    # Guardar y cerrar el documento
    document.store()
    document.close(True)

    print("Comentarios añadidos con éxito")

if __name__ == "__main__":
    main()
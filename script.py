import uno
import json
import os
from com.sun.star.beans import PropertyValue
from com.sun.star.text.TextContentAnchorType import AS_CHARACTER
from com.sun.star.util import XReplaceable

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
        time.sleep(3)
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
    
    for item in comments_data:
        quote = item['quote']
        comment = item['comment']
        author = item['author']
        
        print(f"Buscando: '{quote}'")

        # Crear un descriptor de búsqueda
        search_descriptor = document.createSearchDescriptor()
        search_descriptor.setSearchString(quote)
        search_descriptor.SearchCaseSensitive = False
        search_descriptor.SearchWords = True

        # Realizar la búsqueda
        found = document.findFirst(search_descriptor)

        if found:
            print(f"Encontrado '{quote}'. Añadiendo comentario.")
            
            # Seleccionar el texto encontrado
            cursor.gotoRange(found.getStart(), False)
            cursor.gotoRange(found.getEnd(), True)
            
            # Crear y añadir el comentario
            annotation = document.createInstance("com.sun.star.text.textfield.Annotation")
            annotation.Author = author
            annotation.Content = comment
            
            # Insertar el comentario
            cursor.Text.insertTextContent(cursor, annotation, False)
        else:
            print(f"No se encontró '{quote}'")

    print("Finalizada la adición de comentarios")

def get_available_services(document):
    ctx = document.getComponentContext()
    sm = ctx.getServiceManager()
    services = sm.getAvailableServiceNames()
    return services

def main():
    
    # Rutas de los archivos (ajusta estas rutas según tu configuración)
    document_path = r"C:\Users\luisg\OneDrive\Escritorio\Trabajo\Add comments\add_comments_python\EP3567950-B1__seprotec_es (1).docx"
    json_path = r"C:\Users\luisg\OneDrive\Escritorio\Trabajo\Add comments\add_comments_python\errors.json"

    # Abrir el documento
    document = open_document(document_path)
    #services = get_available_services(document)
    #print("Servicios disponibles:", services)

    
    # Cargar los datos del JSON
    with open(json_path, 'r', encoding='utf-8') as f:
        comments_data = json.load(f)

    # Añadir los comentarios
    add_comments(document, comments_data)

    # Guardar y cerrar el documento
    file_name, file_extension = os.path.splitext(document_path)
    new_file_path = f"{file_name}_suggestions{file_extension}"
     # Guardar el documento con el nuevo nombre
    properties = (
        PropertyValue("Overwrite", 0, True, 0),
        PropertyValue("FilterName", 0, "MS Word 2007 XML", 0)
    )
    document.storeAsURL(uno.systemPathToFileUrl(new_file_path), properties)

    document.close(True)

    print(f"Comentarios añadidos con éxito. Documento guardado como: {new_file_path}")

if __name__ == "__main__":
    main()

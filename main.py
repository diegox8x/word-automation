import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate

def obtener_datos_usuario():
    nombre = "Jorge Garzon"
    telefono = "(123)4-56-61-77"
    correo = "jorge@gmail.com"
    fecha = datetime.today().strftime("%d/%m/%Y")
    return {'nombre': nombre, 'telefono': telefono, 'correo': correo, 'fecha': fecha}

def generar_documento(datos, plantilla, salida):
    try:
        doc = DocxTemplate(plantilla)
        doc.render(datos)
        doc.save(salida)
        print(f"Documento guardado como {salida}")
    except Exception as e:
        print(f"Error al generar el documento: {e}")

def procesar_datos_excel(ruta_excel, plantilla):
    df = pd.read_excel(ruta_excel)
    print(df)
    for indice, file in df.iterrows():
        datos_alumno = {
            'nombre_alumno':file["Nombre completo"], 
            'nota_mat':file["Matematica"],
            'nota_fis':file["Fisica"],
            'nota_qui':file["Quimica"]
        }
        datos_completos = datos_alumno.copy()
        datos_completos.update(obtener_datos_usuario())
        salida = f"Notas_{file['Nombre completo']}.docx"
        generar_documento(datos_completos, plantilla, salida)
        print(datos_completos)


def main():
    plantilla = "plantilla.docx"
    ruta_excel = "kan.xlsx"
    procesar_datos_excel(ruta_excel, plantilla)


if __name__ == "__main__":
    main()


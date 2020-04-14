#! python3

from mailmerge import MailMerge
from datetime import date
from docx2pdf import convert
import openpyxl
import os

'''
Genera todas las TDS entre las lineas especificadas en el archivo "base_datos_TDS.xslx"
Extrae todos os datos del archivo excel y rellena los distintos campos de la TDS
'''

# Crear carpeta donde guardaremos nuestras hojas
os.mkdir('M:\\I+D\\DOCUMENTACION TECNICA\\00-TDS BRENDLE\\Ingles\\TDS creadas el {:%Y%m%d}'.format(date.today()))

# abrir base de datos Excel "base_datos_TDS.xslx"

base_datos_excel = openpyxl.load_workbook('base_datos_TDS.xlsx')

# situarnos en la primera pagina del excel

sheet = base_datos_excel["Sheet1"]

# Hace un ciclo para cada una de las líneas del excel. Si la línea está vacia, no hace nada.
for rownum in range(2, len(sheet['A'])):
    # Si tiene todos los campos completos (no vacíos)
    if (
            (sheet[f'A{rownum}'].value != None) and
            (sheet[f'B{rownum}'].value != None) and
            (sheet[f'C{rownum}'].value != None) and
            (sheet[f'D{rownum}'].value != None) and
            (sheet[f'E{rownum}'].value != None) and
            (sheet[f'F{rownum}'].value != None) and
            (sheet[f'G{rownum}'].value != None) and
            (sheet[f'H{rownum}'].value != None) and
            (sheet[f'I{rownum}'].value != None) and
            (sheet[f'J{rownum}'].value != None) and
            (sheet[f'K{rownum}'].value != None) and
            (sheet[f'L{rownum}'].value != None) and
            (sheet[f'M{rownum}'].value != None) and
            (sheet[f'N{rownum}'].value != None) and
            (sheet[f'O{rownum}'].value != None)):

        # De los productos que están rellenos, diferenciar si es Líquido, Pasta o TH
        ## En funcion de lo que sea,se define la plantilla a tomar

        if "Thermoplastic" in sheet[f'C{rownum}'].value:
            template = "th_template_EN.docx"
        elif "Paste" in sheet[f'C{rownum}'].value:
            template = "paste_template_EN.docx"
        elif "Liquid" or "Banding" in sheet[f'C{rownum}'].value:
            template = "liquid_template_EN.docx"

        # Crea el documento provisional de salida
        TDS_salida = MailMerge(template)

        # Si quetemos ver los campos a rellenar que el programa encuentra
        # print("Fields included in {}: {}".format(template, TDS_salida.get_merge_fields()))

        # Para diferenciar valor a meter en el campo de resistencia a MW:
        # Si tiene frase de resistencia a MW
        if str(sheet[f'P{rownum}'].value) != 'None':
            mw = ". " + str(sheet[f'P{rownum}'].value)
        # Si no tiene resistencia a MW introduce "."
        else:
            mw = "."

        # Introduce los valores de la línea de Excel encontrados en nuestra plantilla.
        TDS_salida.merge(
            familia_producto=str(sheet[f'C{rownum}'].value.upper()),
            familia_producto_minusculas=str(sheet[f'C{rownum}'].value.lower()),
            nombre_producto=str(sheet[f'A{rownum}'].value),
            metodo_aplicacion=str(sheet[f'D{rownum}'].value),
            sustrato_aplicacion=str(sheet[f'E{rownum}'].value),
            bright_matt=str(sheet[f'F{rownum}'].value),
            tono_color_cocido=str(sheet[f'G{rownum}'].value),
            viscosidad=str(sheet[f'H{rownum}'].value),
            consistencia=str(sheet[f'I{rownum}'].value),
            shelf_life=str(sheet[f'J{rownum}'].value),
            color_crudo=str(sheet[f'K{rownum}'].value),
            frase_dilucion=str(sheet[f'L{rownum}'].value),
            disolvente_recomendado=str(sheet[f'M{rownum}'].value),
            dilucion_maxima=str(sheet[f'N{rownum}'].value),
            temperatura_coccion=str(sheet[f'O{rownum}'].value),
            frase_mw=mw)

        # Guarda el documento con el nombre del producto, numero de articulo y fecha del dia producido en la carpeta creada
        TDS_salida.write(
            'M:\\I+D\\DOCUMENTACION TECNICA\\00-TDS BRENDLE\\Ingles\\TDS creadas el {:%Y%m%d}\\{}_{}_TDS_EN_{:%d%m%Y}.docx'.format(
                date.today(), sheet[f'A{rownum}'].value, sheet[f'B{rownum}'].value, date.today()))

        # Convertir todos los docx a PDF tambien

        convert(
            'M:\\I+D\\DOCUMENTACION TECNICA\\00-TDS BRENDLE\\Ingles\\TDS creadas el {:%Y%m%d}\\{}_{}_TDS_EN_{:%d%m%Y}.docx'.format(
                date.today(), sheet[f'A{rownum}'].value, sheet[f'B{rownum}'].value, date.today()),
            'M:\\I+D\\DOCUMENTACION TECNICA\\00-TDS BRENDLE\\Ingles\\TDS creadas el {:%Y%m%d}\\{}_{}_TDS_EN_{:%d%m%Y}.pdf'.format(
                date.today(), sheet[f'A{rownum}'].value, sheet[f'B{rownum}'].value, date.today()))


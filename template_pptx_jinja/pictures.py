import hashlib
from PIL import Image  # Pillow: permite abrir imágenes y acceder a sus propiedades.
from pptx.util import Inches  # Utilidad de python-pptx para convertir pulgadas a unidades EMU.

# --------------------------------------
# Función que calcula el hash SHA-1 de una imagen.
# Esto se usa para identificar imágenes dentro de la presentación y saber cuál hay que reemplazar.
def get_hash(filename):
    with open(filename, "rb") as f:
        blob = f.read()  # Leemos el archivo binario
    sha1 = hashlib.sha1(blob)  # Calculamos su hash SHA-1
    return sha1.hexdigest()  # Devolvemos el hash como cadena hexadecimal
# 

# --------------------------------------
# Función que reemplaza una imagen existente en una diapositiva por otra nueva.
# Mantiene el tamaño aproximado y centra la imagen nueva en el mismo lugar.
def replace_img_slide(slide, img_shape, img_path):
    # Paso 1: Reemplazar el contenido de la imagen
    img_pic = img_shape._pic  # Accedemos al objeto <p:pic> XML de la imagen original
    img_rid = img_pic.xpath('./p:blipFill/a:blip/@r:embed')[0]  # Obtenemos el ID de recurso de imagen embebida
    img_part = slide.part.related_part(img_rid)  # Localizamos la parte correspondiente en la presentación

    with open(img_path, 'rb') as f:
        new_img_blob = f.read()  # Leemos la imagen nueva
    img_part._blob = new_img_blob  # Reemplazamos la imagen binaria en la presentación

    # Paso 2: Obtener dimensiones de la nueva imagen en píxeles y su DPI
    with Image.open(img_path) as im:
        px_w, px_h = im.size  # Anchura y altura en píxeles
        dpi = im.info.get("dpi", (96, 96))  # DPI (puntos por pulgada), por defecto 96 si no está definido
        inch_w = px_w / dpi[0]  # Convertimos ancho a pulgadas
        inch_h = px_h / dpi[1]  # Convertimos alto a pulgadas

    # Paso 3: Obtener dimensiones originales de la forma y su posición centrada
    orig_w = img_shape.width  # Ancho original (en EMU)
    orig_h = img_shape.height  # Alto original (en EMU)
    center_x = img_shape.left + orig_w // 2  # Centro X original
    center_y = img_shape.top + orig_h // 2  # Centro Y original

    # Paso 4: Escalado proporcional para que la imagen nueva quepa en el espacio de la original
    shape_w_inch = orig_w / 914400  # Convertimos EMU a pulgadas
    shape_h_inch = orig_h / 914400
    scale_w = shape_w_inch / inch_w
    scale_h = shape_h_inch / inch_h
    scale = min(scale_w, scale_h, 1.0)  # Escalado proporcional máximo (sin agrandar, solo reducir si hace falta)

    final_w = Inches(inch_w * scale)  # Dimensiones finales a aplicar
    final_h = Inches(inch_h * scale)

    # Paso 5: Aplicar las dimensiones a la nueva imagen
    img_shape.width = final_w
    img_shape.height = final_h

    # Paso 6: Reposicionar para mantener centrado
    img_shape.left = int(center_x - final_w / 2)
    img_shape.top = int(center_y - final_h / 2)
# 

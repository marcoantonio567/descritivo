from docx import Document
from docx.shared import Inches
from PIL import Image, ImageOps
import os

def colar_maps_e_croquii(doc_path, images_with_placeholders, width=Inches(5.91), height=Inches(4.18)):
  
    # Carrega o documento Word
    doc = Document(doc_path)

    for image_path, placeholder in images_with_placeholders:
        # Adiciona uma borda preta à imagem
        image_with_border_path = os.path.join(os.path.dirname(image_path), "bordered_" + os.path.basename(image_path))
        with Image.open(image_path) as img:
            bordered_image = ImageOps.expand(img, border=2, fill="black")
            bordered_image.save(image_with_border_path)

        for paragraph in doc.paragraphs:
            if placeholder in paragraph.text:
                # Remove o marcador de posição
                paragraph.text = paragraph.text.replace(placeholder, "")

                # Adiciona a imagem abaixo do parágrafo
                run = paragraph.add_run()
                run.add_picture(image_with_border_path, width=width, height=height)

        # Remove a imagem temporária
        if os.path.exists(image_with_border_path):
            os.remove(image_with_border_path)

    # Salva o documento atualizado
    doc.save(doc_path)



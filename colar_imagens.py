from docx import Document
from docx.shared import Inches
from PIL import Image, ImageOps
import tempfile
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



def colar_imagens_documentos(doc_path, images_by_code, width=Inches(6), height=Inches(8.32)):
    """
    Insere várias imagens em diferentes locais no documento Word conforme um dicionário de códigos.

    :param doc_path: Caminho para o arquivo Word.
    :param images_by_code: Dicionário onde as chaves são códigos únicos e os valores são listas de caminhos de imagens.
    :param width: Largura das imagens (padrão em polegadas).
    :param height: Altura das imagens (padrão em polegadas).
    """
    # Carrega o documento Word
    doc = Document(doc_path)

    for code, image_paths in images_by_code.items():
        for paragraph in doc.paragraphs:
            if code in paragraph.text:
                # Remove o marcador de posição
                paragraph.text = paragraph.text.replace(code, "")

                for image_path in image_paths:
                    # Cria um arquivo temporário para a imagem com borda
                    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
                        temp_image_path = temp_file.name

                    with Image.open(image_path) as img:
                        # Converte para RGB se necessário
                        if img.mode != "RGB":
                            img = img.convert("RGB")
                        bordered_image = ImageOps.expand(img, border=2, fill="black")
                        bordered_image.save(temp_image_path)

                    # Adiciona a imagem abaixo do parágrafo
                    run = paragraph.add_run()
                    run.add_picture(temp_image_path, width=width, height=height)

                    # Remove a imagem temporária
                    if os.path.exists(temp_image_path):
                        os.remove(temp_image_path)

    # Salva o documento atualizado
    doc.save(doc_path)


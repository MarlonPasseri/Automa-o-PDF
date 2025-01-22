import os
import requests
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from PIL import Image
import pandas as pd

# Função para obter o link direto de download do Google Drive
def get_google_drive_download_url(file_id):
    return f"https://drive.google.com/uc?export=download&id={file_id}"

# Função para baixar a imagem
def download_image(url):
    try:
        response = requests.get(url)
        response.raise_for_status()  # Verifica se a requisição foi bem-sucedida
        img = Image.open(BytesIO(response.content))  # Abre a imagem
        img.verify()  # Verifica se a imagem é válida
        return img.convert('RGB')
    except Exception as e:
        print(f"Erro ao baixar a imagem: {str(e)}")
        return None

def criar_pdf(dados, nome_pdf, logotipo=None):
    c = canvas.Canvas(nome_pdf, pagesize=A4)
    largura, altura = A4

    # Cabeçalho com logotipo
    if logotipo:
        c.drawImage(logotipo, 50, altura - 150, width=120, height=120, preserveAspectRatio=True, mask='auto')
    
    # Título centralizado
    c.setFont("Helvetica-Bold", 18)
    c.drawString(200, altura - 180, "   Ficha de Campo Geoprojetos")
    c.line(50, altura - 190, largura - 50, altura - 190)

    # Espaço para as informações
    y = altura - 220  # Ajustado para dar mais espaço ao logo e ao título

    # Estilo de texto para o conteúdo
    styles = getSampleStyleSheet()
    style_normal = styles['Normal']
    style_normal.fontName = 'Helvetica'
    style_normal.fontSize = 10
    style_normal.leading = 14

    # Adicionar cada campo de dados com um título e valor em formato de parágrafo
    for campo, valor in dados.items():
        texto = f"<b>{campo}:</b> {valor}"
        p = Paragraph(texto, style_normal)
        p_width, p_height = p.wrap(largura - 100, altura)
        
        # Criar uma caixa para o campo de dados
        c.setStrokeColor(colors.black)
        c.setFillColor(colors.whitesmoke)
        c.rect(50, y - p_height - 5, largura - 100, p_height + 10, fill=1)

        # Desenhar o parágrafo dentro da caixa
        p.drawOn(c, 50, y - p_height)
        y -= p_height + 25  # Espaço entre os campos

        if y < 50:  # Adiciona nova página se necessário
            c.showPage()
            c.setFont("Helvetica-Bold", 18)
            c.drawString(200, altura - 180, "   Ficha de Campo Geoprojetos")
            c.line(50, altura - 190, largura - 50, altura - 190)
            y = altura - 220

    # Se houver link de imagem, tentar baixar e inserir a imagem no PDF
    if "Foto da Ocorrência/ Foto do Talude" in dados and dados["Foto da Ocorrência/ Foto do Talude"]:
        link_imagem = dados["Foto da Ocorrência/ Foto do Talude"]
        if link_imagem.startswith("https://drive.google.com/open?id="):
            try:
                # Extrair o ID do arquivo do link
                file_id = link_imagem.split('id=')[-1]
                download_url = get_google_drive_download_url(file_id)
                img = download_image(download_url)
                
                if img:
                    # Salvar a imagem temporariamente para inserir no PDF
                    img_path = "temp_image.jpg"
                    img.save(img_path)
                    c.drawImage(img_path, 50, y - 50, width=100, height=100)  # Ajuste conforme necessário
                    y -= 120  # Espaço após a imagem
                else:
                    c.setFont("Helvetica", 8)
                    c.drawString(50, y, "Erro ao carregar a imagem.")
                    y -= 15
            except Exception as e:
                c.setFont("Helvetica", 8)
                c.drawString(50, y, f"Erro ao carregar a imagem: {str(e)}")
                y -= 15

    # Rodapé com detalhes
    def rodape(canvas_obj):
        canvas_obj.setFont("Helvetica", 8)
        canvas_obj.setFillColor(colors.grey)
        canvas_obj.drawString(50, 30, "Geoprojetos Engenharia Civil | www.geoprojetos.com.br | contato@geoprojetos.com.br")
        canvas_obj.drawRightString(largura - 50, 30, f"Página {canvas_obj.getPageNumber()}")
    rodape(c)

    # Salvar PDF
    c.save()

# Carregar os dados do Excel
file_path = r'C:\Users\marlon.soares\Desktop\Ficha Teste.xlsx'
df = pd.read_excel(file_path, sheet_name='Respostas do Formulário 1')

# Caminho para o logotipo da empresa (opcional)
logotipo = r'C:\Users\marlon.soares\Desktop\logo.jpg'  # Substitua pelo caminho do logotipo

# Gerar um PDF para cada linha
output_dir =r'C:\Users\marlon.soares\Desktop\Fichas Obra Arthur'
os.makedirs(output_dir, exist_ok=True)

for index, row in df.iterrows():
    dados = row.to_dict()
    nome_pdf = os.path.join(output_dir, f"Ficha De Campo-{index + 1}.pdf")
    criar_pdf(dados, nome_pdf, logotipo)

print("PDFs empresariais criados com sucesso!")

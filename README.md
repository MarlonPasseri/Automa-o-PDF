# Documentação do Código: Geração de PDFs a partir de uma Planilha

Este código tem como objetivo gerar arquivos PDF com base nos dados de uma planilha do Excel. Cada linha da planilha gera um PDF individual contendo as informações organizadas e estilizadas.

## Dependências Necessárias
Para executar o código, é necessário instalar as seguintes bibliotecas:

- `pandas`: Para leitura de arquivos Excel.
- `requests`: Para realizar downloads de imagens da internet.
- `reportlab`: Para criação de PDFs.
- `Pillow`: Para manipulação de imagens.

Instale as dependências com o comando:
```bash
pip install pandas requests reportlab pillow
```

---

## Estrutura do Código

### 1. **Função `get_google_drive_download_url`**
Converte um link do Google Drive no formato padrão para um link direto de download.

#### Parâmetros:
- `file_id` (str): ID do arquivo no Google Drive.

#### Retorno:
- `str`: URL direta para download do arquivo.

### 2. **Função `download_image`**
Baixa uma imagem de uma URL e retorna o objeto da imagem em formato RGB.

#### Parâmetros:
- `url` (str): URL da imagem.

#### Retorno:
- `Image` (Pillow): Objeto da imagem convertida para RGB.
- `None`: Em caso de erro no download.

#### Exceções Tratadas:
- Erros de conexão ou URL inválida.

### 3. **Função `criar_pdf`**
Gera um PDF a partir de um dicionário de dados.

#### Parâmetros:
- `dados` (dict): Dados para preenchimento do PDF.
- `nome_pdf` (str): Caminho onde o PDF será salvo.
- `logotipo` (str, opcional): Caminho para o arquivo de logotipo a ser incluído no cabeçalho.

#### Funcionalidades:
1. Insere um cabeçalho com logotipo (se fornecido).
2. Adiciona título e linhas divisórias.
3. Organiza os dados em parágrafos com estilo.
4. Adiciona um rodapé com informações da empresa.

#### Rodapé:
- Inclui nome da empresa, site e contato.
- Exibe o número da página.

### 4. **Leitura da Planilha**
O arquivo Excel é carregado usando o `pandas`. Cada linha da planilha é convertida em um dicionário para gerar PDFs individuais.

#### Configurações:
- `file_path`: Caminho do arquivo Excel.
- `sheet_name`: Nome da aba a ser lida (neste caso, “Respostas do Formulário 1”).

### 5. **Geração de PDFs**
Para cada linha da planilha:
1. Cria um PDF com as informações da linha.
2. Salva o PDF na pasta de saída.

#### Configurações:
- `output_dir`: Diretório onde os PDFs serão salvos.
- `logotipo`: Caminho para o arquivo de logotipo.

---

## Exemplo de Uso
1. Configure os caminhos para:
   - Planilha Excel.
   - Logotipo da empresa.
   - Pasta de saída dos PDFs.

2. Execute o script:
```bash
python gerar_pdfs.py
```

3. Os PDFs serão gerados no diretório especificado.

---

## Estrutura de Saída
Os PDFs gerados seguem o formato:
- Nome: `Ficha_{número_da_linha}.pdf`.
- Exemplo: `Ficha_1.pdf`, `Ficha_2.pdf`.

Cada PDF inclui:
1. Logotipo no cabeçalho (se fornecido).
2. Título “Ficha de Campo Geoprojetos”.
3. Dados formatados em parágrafos com título e valor.
4. Rodapé com informações da empresa e número da página.

---

## Possíveis Erros e Soluções
1. **Erro ao carregar logotipo:**
   - Certifique-se de que o caminho do logotipo está correto e que o arquivo existe.

2. **Erro ao carregar a planilha:**
   - Verifique se o arquivo Excel está no caminho correto e se o nome da aba está correto.

3. **Erro ao baixar imagem:**
   - Confirme se a URL é válida e se a imagem é acessível.

---

## Observações Finais
- Personalize os textos, estilos e espaçamentos no PDF conforme necessário.
- Adapte o caminho do arquivo Excel e do logotipo para seu ambiente local.

Este código é altamente flexível e pode ser ajustado para diferentes tipos de relatórios ou formulários.


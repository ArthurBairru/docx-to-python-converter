<h1 align="center">Docx to Python Converter - Uma ferramenta para recriar documentos Word usando Python</h1> <p align="center"><br><br></b> <img src="https://upload.wikimedia.org/wikipedia/commons/c/c3/Python-logo-notext.svg" alt="Word Document Recreator Logo" width="120px" height="120px"/> <br><br> <i>Este script Python é uma ferramenta para analisar documentos Word existentes e gerar código Python <br>que pode recriar o documento de forma programática, preservando formatação, tabelas e estilos.</i> </p><p align="center"> </p>

<br>

## Descrição

Ferramenta desenvolvida em Python para automatizar a análise e reconstrução de documentos Word (.docx). O script examina minuciosamente a estrutura do documento original - incluindo parágrafos, tabelas, formatação, bordas, espaçamento e estilos - e gera um script Python completo que pode reproduzir fielmente o documento.

Desenvolvido com foco em precisão e automação, utilizando a biblioteca python-docx para manipulação avançada de documentos Word e processamento XML para capturar detalhes complexos de formatação.

![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)
![Office](https://img.shields.io/badge/Microsoft_Word-2B579A?style=for-the-badge&logo=microsoft-word&logoColor=white)

<br>

## Funcionalidades

- **Análise Completa de Documentos**: Extrai parágrafos, tabelas, estilos e formatação
- **Preservação de Formatação**: Mantém alinhamento, espaçamento, fontes e cores
- **Detecção de Bordas**: Identifica e replica bordas de tabelas e parágrafos
- **Processamento de Listas**: Detecta automaticamente listas com bullet points
- **Controle de Espaçamento**: Preserva espaçamento entre linhas e parágrafos
- **Geração de Código**: Produz script Python executável para recriação do documento
- **Suporte a Tabelas Complexas**: Replica estrutura e formatação de tabelas

<br>

## Instalação e Utilização

#### Passo 1: Clonar o Projeto

```bash
git clone https://github.com/ArthurBairru/docx-to-python-converter.git
```

#### Passo 2: Configurar Ambiente Virtual

```python
# Criação Ambiente Virtual
python -m venv .venv

# Ativar (Windows)
.venv\Scripts\activate

# Ativar (Linux/Mac)
source .venv/bin/activate
```

#### Passo 3: Instalar Dependências

```
pip install -r requirements.txt
```

#### Passo 4: Preparar Arquivos Necessários

- Colocar arquivo .docx a ser convertido no root do projeto, com o nome input.docx


#### Passo 5: Executar o Script

```
python main.py
```

#### Passo 6: Verificar Resultados

- O script que recria/recriará seu arquivo word será outputtado no arquivo "recreate_docx.py". O código de recriação do arquivo word pode ser usado para diversas finalidades, como implementação de contratos dinâmicos e etc;
- Para testar a recriação com o script recém gerado, basta executá-lo com `python recreate_docx.py`;
- O arquivo criado pelo script recreate_docx.py aparecerá no root do projeto, sob o nome "output.docx";
- O Relatório detalhado será produzido na raiz do projeto, sob o no

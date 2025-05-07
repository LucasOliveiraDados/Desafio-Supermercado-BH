import pandas as pd
from datetime import datetime
from calendar import monthrange
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from num2words import num2words
import os
from reportlab.lib import colors

planilha = r"C:\Users\lucas\OneDrive\Área de Trabalho\python\Desafio BH\TESTE_ VAGA.xlsx"
pdfs_colaboradores = r"C:\Users\lucas\OneDrive\Área de Trabalho\python\Desafio BH\pdfs_colaboradores"

# Função para ler o excel
def processar_planilha(caminho_excel):
    df = pd.read_excel(caminho_excel, dtype=str)
    df = df[df['STATUS'].str.strip().str.upper() != "NÃO RECEBE"]

    for index, linha in df.iterrows():
        gerar_pdf(linha)

# Função para converter valores para extenso e em português
def valor_por_extenso(valor):
    return num2words(round(valor, 2), lang='pt_BR').replace(" e zero centavos", "")

# Função para extrair o a data com o ultimo dia do mês
def ultima_data_mes(data):
    if isinstance(data, str):
        data = pd.to_datetime(data)
    ultimo_dia = monthrange(data.year, data.month)[1]
    return datetime(data.year, data.month, ultimo_dia).strftime("%d/%m/%Y")

# Função para gerar PDF para cada um colaborador
def gerar_pdf(linha):
    nome = linha['NOME'].strip()
    matricula = str(linha['MATRÍCULA']).strip()
    data_admissao = pd.to_datetime(linha['DATA ADMISSÃO'])
    data_admissao_str = data_admissao.strftime("%d/%m/%Y")
    filial = linha['FILIAL'].strip()

    valor_original = float(str(linha['VALOR (R$)']).replace("R$", "").replace(",", ".").strip())
    desconto = float(str(linha['DESCONTO (R$)']).replace("R$", "").replace(",", ".").strip())
    valor_receber = float(str(linha['VALOR A RECEBER']).replace("R$", "").replace(",", ".").strip())

    fim_mes = ultima_data_mes(data_admissao)

    nome_arquivo = f"{filial}_{nome}.pdf"
    caminho_arquivo = os.path.join(pdfs_colaboradores, nome_arquivo)

    c = canvas.Canvas(caminho_arquivo, pagesize=A4)
    width, height = A4

    margem = 20  
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)
    c.rect(margem, margem, width - 2 * margem, height - 2 * margem)
    c.setFont("Helvetica", 10)
    y = height - 40 

    texto = f"""
Sr(a) Gerente, está autorizado ao(à) colaborador(a) {nome} matrícula: {matricula},
admitido(a) em {data_admissao_str}, receber a quantia de R$ {valor_receber:.2f} ({valor_por_extenso(valor_receber).capitalize()} Reais)
referente ao Vale Alimentação do período de {data_admissao_str} a {fim_mes}, já aplicado o
desconto de 20% (vinte por cento) permitido por lei.

A seguir, descritivo de valores:

Valor Original = R$ {valor_original:.2f} ({valor_por_extenso(valor_original).capitalize()} Reais)
Valor do Desconto de 20% = R$ {desconto:.2f} ({valor_por_extenso(desconto).capitalize()} Reais)
Valor Líquido = R$ {valor_receber:.2f} ({valor_por_extenso(valor_receber).capitalize()} Reais)

Este documento deverá ser datado e assinado pelo colaborador.
Deverá ser devolvido ao Setor de Benefícios, no seguinte e-mail:
fdsdfsdf@supermercadosbh.com.br

Assinatura Colaborador: _________________________________________
Data: ______/____________/__________
"""
    for linha_texto in texto.strip().split("\n"):
        c.drawString(80, y, linha_texto.strip())
        y -= 20

    c.save()

# execução do código
if __name__ == "__main__":
    processar_planilha(planilha)
    print("Arquivos gerados com sucesso.")
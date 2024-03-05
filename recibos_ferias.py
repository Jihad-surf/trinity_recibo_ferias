import os
import PyPDF2
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import messagebox

mapa_numeros = {
    "zero": 0, "um": 1, "dois": 2, "três": 3, "tres":3,  "quatro": 4,
    "cinco": 5, "seis": 6, "sete": 7, "oito": 8, "nove": 9,
    "dez": 10, "onze": 11, "doze": 12, "treze": 13, "quatorze": 14,
    "quinze": 15, "dezesseis": 16, "dezessete": 17, "dezoito": 18, "dezenove": 19,
    "vinte": 20, "trinta": 30, "quarenta": 40, "cinquenta": 50,
    "sessenta": 60, "setenta": 70, "oitenta": 80, "noventa": 90,
    "cem": 100, "cento": 100, "duzentos": 200, "trezentos": 300,
    "quatrocentos": 400, "quinhentos": 500, "seiscentos": 600,
    "setecentos": 700, "oitocentos": 800, "novecentos": 900,
    "mil": 1000, "milhão": 1000000, "milhões": 1000000,
    "bilhão": 1000000000, "bilhões": 1000000000, '':0
}

meses = {
    "Janeiro": "January",
    "Fevereiro": "February",
    "Março": "March",
    "Abril": "April",
    "Maio": "May",
    "Junho": "June",
    "Julho": "July",
    "Agosto": "August",
    "Setembro": "September",
    "Outubro": "October",
    "Novembro": "November",
    "Dezembro": "December"
}

def main():
    root = tk.Tk()
    root.withdraw()
    try:
        files = get_files()
        texto = get_text_on_pdf(files)
        nomes = get_name(texto)
        datas = get_data(texto)
        valores = get_valor(texto)
        salvar_arquivo(nomes, datas, valores)
        messagebox.showinfo("Processo Finalizado", "O processo foi concluído com sucesso!")

    except Exception as e:
        print('Erro no processo {e}')
        messagebox.showerror("Erro", e)
    

def get_files():
    """Retorna uma lista com os nomes dos arquivos PDF no diretório atual."""
    pdf_files = []
    for file in os.listdir():
        if file.endswith('.pdf'):
            pdf_files.append(file)

    if not pdf_files:
        raise ValueError("Nenhum arquivo PDF encontrado no diretório atual.")   
    print('Arquivos encontrados: ', pdf_files)
    return pdf_files


def get_text_on_pdf(files):
    """Retorna uma lista com o texto extraído de cada página de cada arquivo PDF."""
    texto = []
    for file in files:
        with open(file, 'rb') as arquivo_pdf:
            leitor_pdf = PyPDF2.PdfReader(arquivo_pdf)
            num_paginas = len(leitor_pdf.pages)
            for pagina_numero in range(num_paginas):
                pagina = leitor_pdf.pages[pagina_numero]
                texto.append(pagina.extract_text().split('\n'))
    return texto


def get_name(texto):
    """Retorna uma lista com os nomes extraídos do texto."""
    nomes = []
    for page in texto:
        nomes.append(page[-2].strip())
    return nomes


def get_data(texto):
    """Retorna uma lista com as datas extraídas do texto."""
    datas = []
    for page in texto:
        data = page[-3].split(',')[-1].strip()
        for pt, en in meses.items():
            data = data.replace(pt, en)
        data_objeto = datetime.strptime(data, "%d de %B de %Y")

        data_formatada = data_objeto.strftime("%d/%m/%Y")
        datas.append(data_formatada)
    return datas


def get_valor(texto):
    """Retorna uma lista com os valores extraídos do texto."""
    valores = []
    for page in texto:
        valor = page[6].split('Valor')[0]
        valor = extenso_para_numero(valor.strip())
        valores.append("{:.2f}".format(valor))
    return valores


def salvar_arquivo(nomes, datas, valores):
    """Salva um arquivo Excel com os dados extraídos."""
    df = pd.DataFrame({
            'datas': datas,
            'debito': 553,
            'credito': 5,
            'complemento': nomes,
            'historico': 323,
            'valores': valores
            })
    
    df.drop_duplicates(inplace=True)

    df.columns = ['Data', 'Complemento', 'Valor', 'Debito', 'Credito', 'Historico']
    data_horario = datetime.now().strftime('%d%m_%H%M')
    nome_arquivo = f"LANÇAMENTO PAGAMENTO FÉRIAS{data_horario}.xlsx"

    df.to_excel(nome_arquivo, index=False, header=False, engine='openpyxl')

    print(f"Arquivo '{nome_arquivo}' criado com sucesso!")


#funcao auxiliar
def extenso_para_numero(texto):
    """Converte um número por extenso para um valor numerico."""
    texto = texto.lower()
    texto = texto.replace(' e ', 'E').replace(' ', '')

    # acha quantos mil tem
    valor_mil = 0
    milhares = texto.split('mil')[0].split('E')
    for numero in milhares:
        valor = mapa_numeros.get(numero, None)
        valor_mil += valor

    valor_mil = valor_mil * 1000

    # acha entre 0 e 999
    valor_centenas = 0
    centenas = texto.split('mil')[1].split('reais')[0].split('E')
    for numero in centenas:
        valor = mapa_numeros.get(numero, None)
        valor_centenas += valor

    valor_centenas = valor_centenas

    # acha os centavos
    valor_centavos = 0
    centavos = texto.split('reais')[1].split('centavos')[0].split('E')
    if centavos[0] == '':
        centavos = centavos[1:]
    for numero in centavos:
        valor = mapa_numeros.get(numero, None)
        valor_centavos += valor

    valor_centavos = valor_centavos/100
    
    total = valor_mil + valor_centenas + valor_centavos
    return total


if __name__ == '__main__':
    main()
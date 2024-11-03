import pdfplumber
import pandas as pd
import os
import re

def main():
    pdf_files = os.listdir("../resources/pdfs")
    for file in pdf_files:
        pdf_path = f"../resources/pdfs/{file}"
        xlsx_path = pdf_path.replace('.pdf', '-Dani.xlsx')

        data = []

        # Abrir o arquivo PDF
        with pdfplumber.open(pdf_path) as pdf:
            # Loop através das páginas e extrair texto
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    # Dividir o texto em linhas
                    lines = text.split('\n')
                    for line in lines:
                        # Verificar se a linha contém a data e a descrição
                        match = re.match(r'(\d{2} \w{3}) (.+) R\$ ([\d,.]+)', line)
                        if match:
                            date, description, amount = match.groups()
                            # Adicionar à lista de dados
                            data.append([date, '', '', '', description.strip(), '', '', '', float(amount.replace('.', '').replace(',', '.'))])

        # Criar um DataFrame do pandas
        df = pd.DataFrame(data, columns=['Data de compra','','','','Descrição','','','','Valor'])

        # Salvar o DataFrame como um arquivo Excel
        df.to_excel(xlsx_path.replace("pdfs", "faturas"), index=False)

if __name__ == "__main__":
    main()
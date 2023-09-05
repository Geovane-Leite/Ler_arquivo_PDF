# Entrar e editar o arquivo com a informação que tem dentro e  Atualizar para nome Padrão
from datetime import datetime
import os
import re
import shutil
from tkinter import Tk
from tkinter.filedialog import askdirectory, askopenfilename
import PyPDF2
import pandas as pd
from tqdm import tqdm



diretorio_inicial = r'C:\Users'              # diretorio do arquivo .XLSX
diretorio_inicial2 = r'C:\Users\Publics'     # diretorio do arquivos .PDF
# Carregar o arquivo Excel

janela = Tk()
janela.withdraw()
janela.attributes("-topmost", True) #manter no topo

# Exiba a caixa de diálogo para selecionar o arquivo
Arquivo_Excel = askopenfilename(title="Selecionar Arquivo com Lote de Mensalidade", initialdir=diretorio_inicial, filetypes=[('Arquivo Excel', '*.xlsx')])

janela.destroy()

# Abrir a janela de seleção de pasta
janela = Tk()
janela.withdraw()
janela.attributes("-topmost", True) # Manter no topo

# Exibir a caixa de diálogo para selecionar a pasta
#pasta = askdirectory(title="Selecionar Pasta para Anexo")
pasta_raiz = askdirectory(title="Selecionar Pasta dos Arquivos PDF", initialdir=diretorio_inicial2)

df = pd.read_excel(Arquivo_Excel)
quantidade_de_linhas = df.shape[0]
print(quantidade_de_linhas)

initial_count = 0


# Percorrer todas as pastas e subpastas9985623.pdf

for diretorio, subpastas, arquivos in os.walk(pasta_raiz):
    for arquivo in arquivos:
        if arquivo.endswith('.pdf'):
            initial_count += 1
        #nome_da_pasta = os.path.basename(os.path.normpath(foldername))
        #if nome_da_pasta.startswith('FATURA'):
        

with tqdm(total=initial_count) as pbar:
    # completa a primeira tarefa
    # atualiza a barra de progresso
    
    for diretorio, subpastas, arquivos in os.walk(pasta_raiz):

        for arquivo in arquivos:
            # Verificar se o arquivo é um arquivo de texto
            nome_da_pasta = os.path.basename(os.path.normpath(diretorio))
            
            #if nome_da_pasta.startswith('*'):
            if arquivo.endswith('.pdf'):
                arquivo_pdf = os.path.join(diretorio,arquivo)

                pbar.update(1)
                
                

                #print(filename)
                # Abrir o arquivo PDF em modo de leitura binário
                todas_as_paginas = ''
                with open(arquivo_pdf, 'rb') as pdf_file:
                    # Criar um objeto do tipo PdfFileReader
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    
                    # Obter o número de páginas do arquivo PDF
                    num_pages = len(pdf_reader.pages) #reader.getPage(pageNumber) is deprecated and was removed in PyPDF2 3.0.0. Use reader.pages[page_number] instead.

                    # Ler o conteúdo de cada página do arquivo PDF
                    ## for page_num in range(num_pages):
                    paginas = 3
                    if num_pages < paginas:
                        paginas = num_pages

                    for page_num in range(paginas):
                        # Obter o objeto da página atual
                        page_obj = pdf_reader.pages[page_num]
                        
                        # Extrair o texto da página
                        page_text = page_obj.extract_text()
                        todas_as_paginas += page_text
                    
                        
                
                # 4 Tipos de Arquivos serão lidos e alterados

                if 'OPS - Coparticipação por mensalidade' in todas_as_paginas:
                    Pagador = todas_as_paginas.split("Pagador: ")[1].split("\n")[0]
                    Pagador = Pagador.replace('/','-')
                    new_arquivo_pdf = os.path.join(diretorio,Pagador+' - EXT.pdf')

                elif 'Local de Pagamento' in todas_as_paginas: #BOLETO
                    Plano = todas_as_paginas.split("\nPlano: ")[1].split("\n")[0] if "\nPlano: " in todas_as_paginas else 'Plano Coparticipativo'
                    Documento = todas_as_paginas.split("Número do Documento\n")[1].split("Espécie")[0]
                    Sacado = todas_as_paginas.split("Sacado\n")[1].split("\n")[0]
                    Sacado_nome = Sacado.split("........ ")[1].split(".")[0]
                    Sacado_nome = Sacado_nome.replace('/','-')
                    #print(Sacado_nome)
                    new_arquivo_pdf = os.path.join(diretorio,Sacado_nome+' - '+Plano+' - '+Documento+' - BOL.pdf')

                elif 'NOTA FISCAL DE SERVIÇOS ELETRÔNICA - NFS-e\n' in todas_as_paginas:
                    RPS = todas_as_paginas.split("RPS número ")[1].split(" Série")[0]
                    Pagador_NF = todas_as_paginas.split("Nome/R azão Social\n")[1].split("\nCPF/CNPJ")[0]
                    Pagador_NF = Pagador_NF.replace('/','-')
                    new_arquivo_pdf = os.path.join(diretorio,Pagador_NF+' - RPS '+RPS+' - NF.pdf')

                elif 'NOTA FISCAL DE SERVIÇOS ELETRÔNICA - NFS-eNúmero RPS\n' in todas_as_paginas:
                    RPS2 = todas_as_paginas.split("Número RPS\n")[1].split("Número da Nota")[0]
                    Nota = todas_as_paginas.split("Número da Nota\n")[1].split("\n")[0]
                    Pagador_NF2 = todas_as_paginas.split("\nRazão Social:\n")[1].split("\n")[0]
                    Pagador_NF2 = Pagador_NF2.replace('/','-')

                    new_arquivo_pdf = os.path.join(diretorio,Pagador_NF2+' - RPS '+RPS2+' NF '+Nota+' - NF.pdf')
                
               
                if os.path.exists(arquivo_pdf):
                    shutil.move(arquivo_pdf, new_arquivo_pdf)

             

# Apos a alteração de dentro do arquivo consegui obter as iformações
# necessarias para uma nova renomeação de acordo com a Planlha

for index, row in df.iterrows():
    rps = str(row['RPS'])  # Valor do RPS na coluna 'RPS'
    pagador = str(row['PAGADOR']).replace(' - C','')  # Valor do Pagador na coluna 'PAGADOR'
    contrato = str(row['NR_SEQ_CONTRATO'])
    competencia = str(row['DT_REFERENCIA'])
    titulo = str(row['TITULO'])
    Contrato_Relatorio = str(row['Contrato_Relatorio'])
    # contrato = 'Contrato: '+contrato
    # Contrato_Relatorio = 'Pagador: '+Contrato_Relatorio

    #formatação:
    
    pagador = pagador.replace(' - C','').replace('/','-')
    competencia = datetime.strptime(competencia, '%Y-%m-%d %H:%M:%S').strftime('%m_%Y')

    
    # Iterar pelos arquivos no diretório
    
    for filename in os.listdir(pasta_raiz):
        extensao = os.path.splitext(filename)[1]
        diretorio_completo = os.path.join(pasta_raiz,filename)
        
        if 'RPS '+rps in filename:  # Verificar se o RPS está na descrição do arquivo
            # Construir novo nome de arquivo com base no valor do Pagador
            novo_nome = os.path.join(pasta_raiz,'Contrato - '+contrato + ' - ' +'Pagador - '+Contrato_Relatorio+' - '+ pagador + ' - ' + competencia+' - NF'+extensao)
     
            if not os.path.exists(novo_nome):
                if os.path.exists(diretorio_completo):
                    shutil.move(diretorio_completo, novo_nome)
            else:
                novo_nome = os.path.join(pasta_raiz,'Contrato - '+contrato + ' - ' +'Pagador - '+Contrato_Relatorio+' - '+ pagador + ' - ' + competencia+' - NF '+f'{index} DUPLICIDADE'+extensao)
    
                if not os.path.exists(novo_nome):
                    if os.path.exists(diretorio_completo):
                        shutil.move(diretorio_completo, novo_nome)
            
        elif titulo+' - BOL' in filename:
            novo_nome = os.path.join(pasta_raiz,'Contrato - '+contrato + ' - ' +'Pagador - '+Contrato_Relatorio+' - '+ pagador + ' - ' + competencia+' - BOL'+extensao)
        
        
            if not os.path.exists(novo_nome):
                if os.path.exists(diretorio_completo):
                    shutil.move(diretorio_completo, novo_nome)
            else:
                novo_nome2 = os.path.join(pasta_raiz,'Contrato - '+contrato + ' - ' +'Pagador - '+Contrato_Relatorio+' - '+ pagador + ' - ' + competencia+' - BOL '+f'{index} DUPLICIDADE'+extensao)
        
                if not os.path.exists(novo_nome2):
                    if os.path.exists(diretorio_completo):
                        shutil.move(diretorio_completo, novo_nome2)


        elif filename.startswith(Contrato_Relatorio):
            novo_nome = os.path.join(pasta_raiz,'Contrato - '+contrato + ' - ' +'Pagador - '+Contrato_Relatorio+' - '+ pagador + ' - ' + competencia+' - EXT'+extensao)
            
        
            if not os.path.exists(novo_nome):
                if os.path.exists(diretorio_completo):
                    shutil.move(diretorio_completo, novo_nome)
            else:
                novo_nome2 = os.path.join(pasta_raiz,'Contrato - '+contrato + ' - ' +'Pagador - '+Contrato_Relatorio+' - '+ pagador + ' - ' + competencia+' - EXT '+f'{index} DUPLICIDADE'+extensao)
            
                if not os.path.exists(novo_nome2):
                    if os.path.exists(diretorio_completo):
                        shutil.move(diretorio_completo, novo_nome2)
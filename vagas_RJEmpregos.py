import requests, pandas as pd, openpyxl
from bs4 import BeautifulSoup
from datetime import datetime, date, timedelta
import locale, pyttsx3
from tkinter import *
from tkinter import ttk
#################  DEFINE CONFIGURAÇÃO LOCAL  ################################
locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')

#################  VARIÁVEIS INICIAIS  ################################
status='atual'
pagina=1
listaVagas = []

#################  TRATA DATA ATUAL  ################################
data_atual = datetime.now().strftime('%d %B, %Y')
hora_atual = datetime.now().strftime('%H:%M')

data_tempo = datetime.now().date()
hora_tempo = datetime.now().time()

data_hora_atual = datetime.now().strftime('%d %B, %Y %H:%M')
data_hora_atual_tempo = datetime.now().strptime(data_hora_atual,'%d %B, %Y %H:%M')

#################  TRATA DADOS DA PLANILHA  ################################
dados = pd.read_excel('vagas.xlsx')
# print('################# DADOS DA PLANILHA ########################################')
# listaVagas.append([hora_atual, pagina, data_anuncio, hora_anuncio, tituloAnuncio.text, texto, link, texto_leiaMais])
listaVagas = []
# print('################# DADOS NA LISTA ########################################')
try:
    hora_mais_recente = max(dados['Hora_Anúncio'])
    data_mais_recente = max(dados['Data'])
except:
    # print('Falha na data da planilha!')
    hora_mais_recente = hora_tempo
    data_mais_recente = data_atual
finally:
    data_hora_recente = f"{data_mais_recente} {hora_mais_recente}"
    data_hora_recente_tempo = data_hora_atual_tempo

# #################  CONFIGURAR JANELA  ################################
janela = Tk()
exibir = ttk.Treeview(janela, selectmode='browse', column=('a','b','c','d'), show='headings')
exibir.column('a', width=100)
exibir.heading('#1',text='Hora')
exibir.column('b', width=500)
exibir.heading('#1',text='Título')
exibir.column('c', width=600)
exibir.heading('#2',text='Descrição')
exibir.grid(row=0,column=0)
exibir.column('d', width=100)
exibir.heading('#3',text='Sit')

################  BUSCA ANÚNCIOS DO DIA EM RJ EMPREGOS  #################
# percorre os dados enquanto a publicação for da data atual
# while pagina < 50:
while status == 'atual':
    #################  BUSCA DADOS ################################
    resposta = requests.get(f'https://rjempregos.net/page/{pagina}/')
    conteudo = resposta.content
    conteudo = BeautifulSoup(conteudo, 'html.parser')
    anuncios = conteudo.findAll('article')
    print(pagina) # Apenas para acompanhar o andamento, imprimindo o número da página atual
#################  PERCORRE CADA BLOCO DE DADOS BUSCANDO AS INFORMAÇÕES  ################################
    for anuncio in anuncios:
        data_anuncio = anuncio.find('time', attrs={"class": "entry-date"}).text.split(', - ')[0] # TEXTO
        data_anuncio_tempo = datetime.strptime(data_anuncio, '%d %B, %Y').date()
        hora_anuncio = anuncio.find('time', attrs={"class": "entry-date"}).text.split(', - ')[1] # TEXTO
        hora_anuncio_tempo = datetime.strptime(hora_anuncio, '%H:%M').time()
        data_hora_anuncio = f"{data_anuncio} {hora_anuncio}" # TEXTO
        data_hora_anuncio_tempo = datetime.strptime(data_hora_anuncio, '%d %B, %Y %H:%M')
        # Se a data do anúncio é diferente a data atual
        if (data_tempo != data_anuncio_tempo):
            status='Antiga'
            # print("Vaga de ontem")
            break
        else:
            # Compara a data/hora do anúncio com o registro mais recente na planilha para identificar o que já havia sido pego anteriormente
            if (data_hora_anuncio_tempo < data_hora_recente_tempo):
                sit = "Anterior"
            else:
                sit = "Nova"
            tituloAnuncio = anuncio.find('h2', attrs={"class":"entry-title"})
            textoBruto = anuncio.find('div', attrs={"class":"entry-content"})
            texto = textoBruto.text
            texto = texto.split('\n')[1]
            link = anuncio.find('a', attrs={"class":"read-more"})
            if link:
                leiaMais = link['href']
                link = link['href']
                pagina_leiaMais = requests.get(leiaMais).content
                pagina_leiaMais = BeautifulSoup(pagina_leiaMais, 'html.parser')
                texto_leiaMais = pagina_leiaMais.find('div', attrs={"class": "job_description"})
                texto_leiaMais = texto_leiaMais.prettify()
            else:
                texto_leiaMais= '-'
                link = '-'
#################  SALVA NA LISTA ################################
            # print(sit, hora_atual, pagina, data_anuncio, hora_anuncio, tituloAnuncio.text, texto, link, texto_leiaMais)
            listaVagas.append([sit, hora_atual, pagina, data_anuncio, hora_anuncio, tituloAnuncio.text, texto, link, texto_leiaMais])
            # print(listaVagas)
#################  PREPARA PARA EXIBIR O RESULTADO NA JANELA ABERTA ################################
            exibir.insert('', END, values=(hora_anuncio, tituloAnuncio.text, texto), tag='0')
    pagina = pagina + 1

# ABRE E MANTEM A JANELA ABERTA
# janela.mainloop()
coletaVagas = pd.DataFrame(listaVagas, columns=['Status','Hora_Busca','Página','Data','Hora_Anúncio','Titulo','Descricao', 'Link', 'Leia Mais'])
pd.set_option('display.max_columns', None)

# SALVAR EM EXCEL
coletaVagas.to_excel('vagas.xlsx', index=False)

# pd.set_option('display.max_columns', None)
# Exibe a tabela na tela de comando
# print(coletaVagas[['Status','Hora_Busca','Página','Data','Hora_Anúncio','Titulo','Descricao', 'Link', 'Leia Mais']])

# Faz leitura do texto
# print(texto_leitura)
# leitor = pyttsx3.init()
# leitor.save_to_file(texto_leitura, 'lido.mp3')
# leitor.runAndWait()

# Cria executável do código, usado no terminal, o '-w' é quando tem uma tela para ser aberta, '--noconsole' não exibe nada
# --onefile cria um único arquivo
# pyinstaller --onefile -w vagas_RJEmpregos.py


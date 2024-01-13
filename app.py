#Feito para cadastrar produtos ficticios no site https://cadastro-produtos-devaprender.netlify.app/

import openpyxl
import pyperclip as pc
import pyautogui as py
import time
# # Entrar na planilha
workbook = openpyxl.load_workbook("produtos_ficticios.xlsx")
sheet_produtos = workbook["Produtos"]
# Copiar informação de um capo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
  #Etapa 1
  nome_produto = linha[0].value
  pc.copy(nome_produto)
  py.click(1124,342,duration=0.3)
  py.hotkey('ctrl','v')

  descricao = linha[1].value
  pc.copy(descricao)
  py.click(1115,434, duration=0.3)
  py.hotkey('ctrl','v')

  categoria = linha[2].value
  pc.copy(categoria)
  py.click(1138,561, duration=0.3)
  py.hotkey('ctrl','v')

  codigo_produto = linha[3].value
  pc.copy(codigo_produto)
  py.click(1136,644, duration=0.3)
  py.hotkey('ctrl','v')
  
  peso = linha[4].value
  pc.copy(peso)
  py.click(1131,729, duration=0.3)
  py.hotkey('ctrl','v')

  dimensoes = linha[5].value
  pc.copy(dimensoes)
  py.click(1151,824, duration=0.3)
  py.hotkey('ctrl','v')

  #Clicar em proxima
  py.click(1119,875, duration=0.3)
  time.sleep(2)
  #Etapa 2
  preco = linha[6].value
  pc.copy(preco)
  py.click(1135,364, duration=0.3)
  py.hotkey('ctrl','v')

  quantidade_em_estoque = linha[7].value
  pc.copy(quantidade_em_estoque)
  py.click(1155,456,duration=0.3)
  py.hotkey('ctrl','v')

  data_de_validade = linha[8].value
  pc.copy(data_de_validade)
  py.click(1122,537,duration=0.3)
  py.hotkey('ctrl','v')

  cor = linha[9].value
  pc.copy(cor)
  py.click(1166,632,duration=0.3)
  py.hotkey('ctrl','v')

  py.click(1157,709,duration=0.5)
  tamanho = linha[10].value
  if tamanho == 'Pequeno':
      py.click(1128,747, duration=0.3)
  elif tamanho == 'Medio':
      py.click(1121,777)
  else:
     py.click(1137,811)

  material = linha[11].value
  pc.copy(material)
  py.click(1115,799, duration=0.3)
  py.hotkey('ctrl','v')

  #Clicar em proxima
  py.click(1132,859)

  fabricante = linha[12].value
  pc.copy(fabricante)
  py.click(1153,404, duration=0.3)
  py.hotkey('ctrl','v')

  pais_origem = linha[13].value
  pc.copy(pais_origem)
  py.click(1154,493,duration=0.3)
  py.hotkey('ctrl','v')

  observacoes = linha[14].value
  pc.copy(observacoes)
  py.click(1154,580,duration=0.3)
  py.hotkey('ctrl','v')

  codigo_de_barras = linha[15].value
  pc.copy(codigo_de_barras)
  py.click(1159,701,duration=0.3)
  py.hotkey('ctrl','v')

  localizacao_armazem = linha[16].value
  pc.copy(localizacao_armazem)
  py.click(1165,790,duration=0.3)
  py.hotkey('ctrl','v')

 #Clicar em Concluir
  py.click(1135,861,duration=0.3)
  time.sleep(2)
  #Clicar em POPUP
  py.click(1605,608,duration=0.3)
  #Clicar para Adicionar mais um
  py.click(1435,616,duration=0.3)

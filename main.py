# Importando as Bibliotecas

from time import sleep
import pyautogui as pg
from datetime import date
import xlwings as xw

# Variáveis
usuario = 'usuario' # informar antes de rodar
senha = 'senha' # informar antes de rodar
venda_dia_anterior = (date.today().day) - 1
app = xw.apps.active 

# Abrindo o Navegador
pg.press('win')
pg.write('Microsoft Edge')
sleep(1)
pg.press('enter')
sleep(1.5)

# Entrando no B.I
pg.click(x=1465, y=84)
sleep(1.5)
pg.write(usuario)
sleep(1)
pg.press('tab')
sleep(1)
pg.write(senha)
sleep(1)
pg.press('tab')
sleep(1)
pg.press('enter')
sleep(5)

# Entrando no Aplicativo e Dashboard
pg.click(x=1686, y=257) # Clica no Aplicativo
sleep(16)
pg.click(x=1593, y=221) # Clina no filtro de dia
sleep(2)
pg.write(str(venda_dia_anterior)) # Escreve o dia
sleep(2)
pg.press('enter')
sleep(2)
pg.click(x=1786, y=260)
sleep(1)
pg.click(x=1475, y=140) # Clina na Pastas
sleep(5)
pg.click(x=1549, y=646) # Seleciona a 3º Pasta (Dashboard de Visão Faturada)
sleep(3)
pg.click(x=2719, y=617,clicks=3) # Clica 3x na barra de rolagem para ver a Tabela
sleep(1.5)
pg.click(x=2013, y=418,button='right') # Clica na tabela para extrair os dados
sleep(1)
pg.click(x=2100, y=438) # 1º Etapa para extrair
sleep(1)
pg.click(x=2123, y=619) # 2º Etapa para extrair
sleep(1)
pg.click(x=2052, y=513) # 3º Etapa para extrair
sleep(1)
pg.click(x=1904, y=408) # 4º Etapa para extrair
sleep(1.5)
pg.click(x=2320, y=148) # Abri a Planilha de Venda extraida
sleep(7)
pg.click(x=2528, y=91)
sleep(2)

# Abrindo a Planilha Relatório FINAL
coord_pasta_vendas = 1462, 562
coord_pasta_vendas_mes = 1797, 171 # -> Coordenada do mês de Janeiro - 2026
coord_planilha_vendas_mes = 1845, 360
coord_opcao_atualizar_planilha = 1944, 429

pg.click(x=2071, y=740) # Abrindo o explorador de arquivos
sleep(2.5)
pg.click(coord_pasta_vendas)
sleep(2)
pg.click(coord_pasta_vendas_mes,clicks=2)
sleep(2)
pg.click(coord_planilha_vendas_mes,clicks=2)
sleep(15)
pg.click(coord_opcao_atualizar_planilha)
sleep(8)

# Voltando para a Planilha de Venda Extraida para Tratar
pg.click(x=2247, y=744)
sleep(2)
pg.click(x=2108, y=627)
sleep(2)

# Tratando a Planilha de Venda Extraida

planilha_aberta = xw.books.active
aba_atual = planilha_aberta.sheets.active
app = xw.apps.active
campos_outros = [28.0,21.0,14.0,7.0,5.0]
ultima_linha = aba_atual.range('A' + str(aba_atual.cells.last_cell.row)).end('up').row
valores_a = aba_atual.range(f'A1:A{ultima_linha}').value

# Escrevendo TD 
aba_atual.range('C2').value = 'TD' 
# Copiando a Coluna C & Colando na Coluna A
aba_atual.range('C:C').copy(destination=aba_atual.range('A:A'))
# Quebrando o Texto da Coluna A
app.display_alerts = False
intervalo = aba_atual.range('A:A').api
intervalo.TextToColumns(
    Destination=aba_atual.range('A1').api,
    DataType=1, 
    Other=True, 
    OtherChar='-'
    )
app.display_alerts = True
sleep(1)
aba_atual.range('B:B').delete()
sleep(3)
valores_a = aba_atual.range(f'A1:A{ultima_linha}').value
for i, valor in enumerate(valores_a,start=1):
    if valor in campos_outros:
        aba_atual.range(f'A{i}').value = 'Outros'
sleep(1)
aba_atual.used_range.copy()
sleep(1.5)

# Colando informação na Planilha Relatório FINAL
pg.click(x=2251, y=744)
sleep(1)
pg.click(x=2332, y=629)
sleep(1)
planilha_final_vendas = xw.books['Vendas Jan_2026 - BI']
for aba in planilha_final_vendas.sheets:
    if aba.name == str(venda_dia_anterior):
        aba.range('AI48').paste()
        

        

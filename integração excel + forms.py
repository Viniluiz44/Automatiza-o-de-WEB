import openpyxl
import pyperclip
import pyautogui

planilha = openpyxl.load_workbook("produtos_ficticios.xlsx")
aba_produtos = planilha['Produtos']

for linha in aba_produtos.iter_rows(min_row=2):
   nome_produto = (linha[0].value)
   pyperclip.copy(nome_produto)
   pyautogui.click(1321,326,duration=1)
   pyautogui.hotkey('ctrl','v')

   descricao = (linha[1].value)
   pyperclip.copy(descricao)
   pyautogui.click(1305,448,duration=1)
   pyautogui.hotkey('ctrl','v')

   categoria = (linha[2].value)
   pyperclip.copy(categoria)
   pyautogui.click(1282,564,duration=1)
   pyautogui.hotkey('ctrl','v')

   codigo_produto = (linha[3].value)
   pyperclip.copy(codigo_produto)
   pyautogui.click(1319,681,duration=1)
   pyautogui.hotkey('ctrl','v')

   peso = (linha[4].value)
   pyperclip.copy(peso)
   pyautogui.click(1293,801,duration=1)
   pyautogui.hotkey('ctrl','v')

   dimensoes = (linha[5].value)
   pyperclip.copy(dimensoes)
   pyautogui.click(1294,918,duration=1)
   pyautogui.hotkey('ctrl','v')

   pyautogui.click(1239,989,duration=1)

   pyautogui.click(1270,245,duration=3)
 
# caso necessite de mais informações na tabela/planilha
   # apenas teste

#    Preço = (linha[6].value)
#    Quantidade_em_estoque = (linha[7].value)
#    Data_de_validade = (linha[8].value)
#    Cor = (linha[9].value)
#    Tamanho = (linha[10].value)
#    Material = (linha[11].value)
#    Fabricante = (linha[12].value)
#    País_de_origem = (linha[13].value)
#    Observações = (linha[14].value)
#    Código_de_barras = (linha[15].value)
#    Localização_no_armazém = (linha[16].value)

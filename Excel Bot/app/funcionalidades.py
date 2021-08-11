import xlsxwriter
import os
import sys
import pyautogui
import time

def PaginaPrincipal():
    print("="*35)
    print("Sejam bem-vindos(as) ao Excel Bot.")
    print("A forma mais fácil de criar suas planilhas automatizadas")
    print("="*35)
    pass

def ExcelBotOpcoes():
    print("[1] - Criar Planilhas")
    print("[2] - Fechar Excel Bot")
    print("="*20)
    pass

def OpcoesColunasLinhas():
    os.system('cls')
    print("="*45)
    print("OPÇÕES DISPONIVEIS - TAMANHO DE PLANILHAS")
    print("="*45)
    print("[1] - Colunas [3] & Linhas [3]")
    print("[2] - Colunas [5] & Linhas [5]")
    print("="*45)
    pass

def PlanilhaPrimeiraOpcao():
    os.system('cls')
    pyautogui.alert("Após criar a planilha, não mexa no teclado ou mouse, iremos abrir o projeto após a conclusão.")
    print("="*30)
    print("CRIANDO PLANILHA...")
    print("="*30)
    nomeArquivo = input("Nome do Arquivo: ")
    file = xlsxwriter.Workbook(f'{nomeArquivo}.xlsx')
    table = file.add_worksheet()
    negrito = file.add_format({"bold": True})
    
    #SESSÃO A1
    tituloA1 = input("Titulo de A1: ")
    conteudoA2 = input("Conteudo A2: ")
    conteudoA3 = input("Conteudo A3: ")
    conteudoA4 = input("Conteudo A4: ")
    print("\n")
    #SESSÃO A1

    #SESSÃO B1
    tituloB1 = input("Titulo de B1: ")
    conteudoB2 = input("Conteudo B2: ")
    conteudoB3 = input("Conteudo B3: ")
    conteudoB4 = input("Conteudo B4: ")
    print("\n")
    #SESSÃO B1

    #SESSÃO C1
    tituloC1 = input("Titulo de C1: ")
    conteudoC2 = input("Conteudo C2: ")
    conteudoC3 = input("Conteudo C3: ")
    conteudoC4 = input("Conteudo C4: ")
    #SESSÃO C1

    table.write("A1",f'{tituloA1}', negrito)
    table.write("A2",f'{conteudoA2}')
    table.write("A3",f'{conteudoA3}')
    table.write("A4",f'{conteudoA4}')

    table.write("B1",f'{tituloB1}', negrito)
    table.write("B2",f'{conteudoB2}')
    table.write("B3",f'{conteudoB3}')
    table.write("B4",f'{conteudoB4}')

    table.write("C1",f'{tituloC1}', negrito)
    table.write("C2",f'{conteudoC2}')
    table.write("C3",f'{conteudoC3}')
    table.write("C4",f'{conteudoC4}')

    file.close()
    pyautogui.press('win')
    time.sleep(1)
    pyautogui.write(f'{nomeArquivo}')
    pass

def PlanilhaSegundaOpcao():
    os.system('cls')
    pyautogui.alert("Após criar a planilha, não mexa no teclado ou mouse, iremos abrir o projeto após a conclusão.")
    print("="*30)
    print("CRIANDO PLANILHA...")
    print("="*30)
    nomeArquivo = input("Nome do Arquivo: ")
    file = xlsxwriter.Workbook(f'{nomeArquivo}.xlsx')
    table = file.add_worksheet()
    negrito = file.add_format({"bold": True})
    
    #SESSÃO A1
    tituloA1 = input("Titulo de A1: ")
    conteudoA2 = input("Conteudo A2: ")
    conteudoA3 = input("Conteudo A3: ")
    conteudoA4 = input("Conteudo A4: ")
    conteudoA5 = input("Conteudo A5: ")
    conteudoA6 = input("Conteudo A6: ")
    print("\n")
    #SESSÃO A1

    #SESSÃO B1
    tituloB1 = input("Titulo de B1: ")
    conteudoB2 = input("Conteudo B2: ")
    conteudoB3 = input("Conteudo B3: ")
    conteudoB4 = input("Conteudo B4: ")
    conteudoB5 = input("Conteudo B5: ")
    conteudoB6 = input("Conteudo B6: ")
    print("\n")
    #SESSÃO B1

    #SESSÃO C1
    tituloC1 = input("Titulo de C1: ")
    conteudoC2 = input("Conteudo C2: ")
    conteudoC3 = input("Conteudo C3: ")
    conteudoC4 = input("Conteudo C4: ")
    conteudoC5 = input("Conteudo C5: ")
    conteudoC6 = input("Conteudo C6: ")
    print("\n")
    #SESSÃO C1

    #SESSÃO D1
    tituloD1 = input("Titulo de D1: ")
    conteudoD2 = input("Conteudo D2: ")
    conteudoD3 = input("Conteudo D3: ")
    conteudoD4 = input("Conteudo D4: ")
    conteudoD5 = input("Conteudo D5: ")
    conteudoD6 = input("Conteudo D6: ")
    print("\n")
    #SESSÃO D1

    #SESSÃO E1
    tituloE1 = input("Titulo de E1: ")
    conteudoE2 = input("Conteudo E2: ")
    conteudoE3 = input("Conteudo E3: ")
    conteudoE4 = input("Conteudo E4: ")
    conteudoE5 = input("Conteudo E5: ")
    conteudoE6 = input("Conteudo E6: ")
    print("\n")
    #SESSÃO E1
    
    #SESSÃO F1
    tituloF1 = input("Titulo de F1: ")
    conteudoF2 = input("Conteudo F2: ")
    conteudoF3 = input("Conteudo F3: ")
    conteudoF4 = input("Conteudo F4: ")
    conteudoF5 = input("Conteudo F5: ")
    conteudoF6 = input("Conteudo F6: ")
    print("\n")
    #SESSÃO F1


    table.write("A1",f'{tituloA1}', negrito)
    table.write("A2",f'{conteudoA2}')
    table.write("A3",f'{conteudoA3}')
    table.write("A4",f'{conteudoA4}')
    table.write("A5",f'{conteudoA5}')
    table.write("A6",f'{conteudoA6}')

    table.write("B1",f'{tituloB1}', negrito)
    table.write("B2",f'{conteudoB2}')
    table.write("B3",f'{conteudoB3}')
    table.write("B4",f'{conteudoB4}')
    table.write("B5",f'{conteudoB5}')
    table.write("B6",f'{conteudoB6}')

    table.write("C1",f'{tituloC1}', negrito)
    table.write("C2",f'{conteudoC2}')
    table.write("C3",f'{conteudoC3}')
    table.write("C4",f'{conteudoC4}')
    table.write("C5",f'{conteudoC5}')
    table.write("C6",f'{conteudoC6}')

    table.write("D1",f'{tituloD1}', negrito)
    table.write("D2",f'{conteudoD2}')
    table.write("D3",f'{conteudoD3}')
    table.write("D4",f'{conteudoD4}')
    table.write("D5",f'{conteudoD5}')
    table.write("D6",f'{conteudoD6}')

    table.write("E1",f'{tituloE1}', negrito)
    table.write("E2",f'{conteudoE2}')
    table.write("E3",f'{conteudoE3}')
    table.write("E4",f'{conteudoE4}')
    table.write("E5",f'{conteudoE5}')
    table.write("E6",f'{conteudoE6}')

    table.write("F1",f'{tituloF1}', negrito)
    table.write("F2",f'{conteudoF2}')
    table.write("F3",f'{conteudoF3}')
    table.write("F4",f'{conteudoF4}')
    table.write("F5",f'{conteudoF5}')
    table.write("F6",f'{conteudoF6}')

    file.close()
    pyautogui.press('win')
    time.sleep(1)
    pyautogui.write(f'{nomeArquivo}')
    pass
    
import funcionalidades
import sys

funcionalidades.PaginaPrincipal()
funcionalidades.ExcelBotOpcoes()
opcaoSelecionada = int(input("Opção selecionada: "))
if opcaoSelecionada == 1:
    funcionalidades.OpcoesColunasLinhas()
    qualOpcaoSelecionada = int(input("Opção selecionada: "))
    
    if qualOpcaoSelecionada == 1:
        funcionalidades.PlanilhaPrimeiraOpcao()
    elif qualOpcaoSelecionada == 2:
        funcionalidades.PlanilhaSegundaOpcao()
elif opcaoSelecionada == 2:
    sys.exit()
import win32com.client as win
import pandas as pd
import graphics as graph
from funcoes import func
from variables import var
##################################################
##    ESTA AUTOMAÇÃO ESTÁ EM DESENVOLVIMENTO    ## 
##################################################
# tipos de email criados até o momento:
#  - validacao
#  - cancelada
#  - encerrada
#  - inssucesso 
#
# Projeto para integração com planilha para organização de dados automatica, sendo necessário apenas passar o path da planilha manualmente!
# 
# Pendencias:
#  - interface para melhor visualização dos dados passados
#  - integração com excel via pandas



tipo = input('estado de encerramento:\n Opções: validacao, cancelada, inssucesso e encerrada \n')
para = input('email de validação:\n EX: *****@email.com; *****@outroemail.com.br\n')
gmud = input('Numero da GMUD\n EX: CHG**** \n')
host_application = input('Nome do host e a aplicação rodando:\n EX: GNCANNL5605 - linux/database\n')
desc = input('Escreva a descrição do encerramento/motivos:\n')


#-tipo de email-|-para quem sera enviado-|-GMUD-|-Nome da maquina e aplicação-|-Descrição de ações realizadas e motivos-                
func.cria_email(tipo,para,gmud,host_application,desc)

print('tarefa comcluida!')

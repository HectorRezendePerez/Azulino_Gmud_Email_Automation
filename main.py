import win32com.client as win
from funcoes import func
##################################################
##    ESTA AUTOMAÇÃO ESTÁ EM DESENVOLVIMENTO    ## 
##################################################
# tipos de email criados até o momento:
#  - validacao
#  - cancelada
#  - encerrada
#  - inssucesso 
#
#             -tipo de email-    -para quem sera enviado-       -GMUD-      -Nome da maquina e aplicação-              -Descrição de ações realizadas e motivos-                
func.cria_email('inssucesso' ,'hectordavidrezende@gmail.com', 'CHGTANANAN', 'GNCANNL5605 - linux/database', ' A aplicação de patches de S.O. foi realizada com sucesso. Reboot realizado com sucesso e nenhum patch pendente.')
print('tarefa comcluida!')

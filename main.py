import win32com.client as win
from funcoes import func
#############################################
##    ESTA AUTOMAÇÃO EM DESENVOLVIMENTO    ## 
#############################################
# tipos de email criados até o momento:
#  - validacao
#  - cancelada
#  - encerrada
#
#             -tipo de email-    -para quem sera enviado-       -Copia de email-          -GMUD-      -Nome da maquina e aplicação-              -Descrição de ações realizadas e motivos-                
func.cria_email('validacao' ,'hectordavidrezende@gmail.com', 'hector.perez@oraex.com', 'CHGTANANAN', 'GNCANNL5605 - linux/database', ' A aplicação de patches de S.O. foi realizada com sucesso. Reboot realizado com sucesso e nenhum patch pendente.')
print('tarefa comcluida!')
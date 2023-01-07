import win32com.client as win
import pandas as pd
from variables import Var
df = pd.read_excel("c:/email_gmud_win32/git_repo/automation_email_GMUD/Pasta1.xlsx")
print(df)
for ind, host_value in enumerate(df['MÁQUINAS']):
    #DADOS DE ENCAMINHAMENTO
    emailTo = df.loc[ind, 'DESTINATÁRIO']
    #DADOS DO CORPO
    closure = df.loc[ind,'ENCERRAMENTO']
    gmud = df.loc[ind,'GMUD']
    adequacy = df.loc[ind,'NOME']
    host = host_value
    primaryContactEmail = df.loc[ind,'RESPONSÁVEL']
    description = df.loc[ind, 'DESCRIÇÃO']
    preCheck = df.loc[ind, 'PRÉ ALERTA']
    afterCheck = df.loc[ind, 'PÓS ALERTA']

    index_closure = closure.find('-')
    if index_closure <= 9:
        index_closure *= -20
    else:
        pass
    treated_closure_str = closure[:index_closure]

    #print(treated_closure_str, ' - ', closure)
    #print(index_closure)
    outlook = win.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    email.to = emailTo
    email.Subject = f'[ {treated_closure_str} ] - {gmud} - {adequacy}'
    email.HTMLBody = f'''
    <p>Prezados,&nbsp;</p>
    
    <p>&nbsp;</p>
    <p>Segue o status da mudan&ccedil;a:&nbsp;<strong>[ {treated_closure_str} ] - {gmud} - {adequacy}</strong></p>
    <p>&nbsp;</p>
    <p>&Aacute;rea respons&aacute;vel para valida&ccedil;&atilde;o:&nbsp;<strong>{primaryContactEmail}</strong></p>
    <p>Servidor:&nbsp;<strong>{host}</strong></p>
    <p>Contato plantonista:&nbsp;<strong>{emailTo}</strong></p>
    <p>Alertas Zabbix antes execu&ccedil;&atilde;o:&nbsp;<strong>{preCheck}</strong>&nbsp;&nbsp;(Anexar o print com a evid&ecirc;ncia)</p>
    <p>Alertas Zabbix p&oacute;s execu&ccedil;&atilde;o:&nbsp;<strong>{afterCheck}</strong>&nbsp;&nbsp;(Anexar o print com a evid&ecirc;ncia)</p>
    <p>Observa&ccedil;&otilde;es:&nbsp;<strong>{description}</strong><br /p>
    <br />
    <p><strong>Por gentileza, validar o funcionamento do ambiente.</strong></p>
    <p>{str(Var.rodape)}
    <img alt="" src="https://i.ibb.co/s1vcggZ/thumbnail-Outlook-ir4tv3sj.png">
    </p>   
    '''
    email.Send()
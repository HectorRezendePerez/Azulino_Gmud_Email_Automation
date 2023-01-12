import win32com.client as win
import pandas as pd
import os, fnmatch
from variables import Var
import time
print('_AZULINO_GMUD_EMAIL_AUTOMATION_')
print('!!! WELCOME TO AZULINO CLI !!!')
while True:
    option = input('azulino ~> ')
    df = pd.read_excel(Var.configYaml['config']['paths']['path_to_excel'])
    if option == 'check excel':
        print(df)
        time.sleep(2)
    elif option == 'azulino --help':
        print('send        -> envia os emails \ncheck excel -> verifica o teu excel (printa na tela o dataframe) \nexit        -> encerra o processo')
    elif option == 'send':        
        for ind, host_value in enumerate(df[str(Var.configYaml['config']['excel_columns_config']['host'])]):
            #DADOS DE ENCAMINHAMENTO
            emailTo = df.loc[ind, str(Var.configYaml['config']['excel_columns_config']['emailTo'])]
            #DADOS DO CORPO
            closure = df.loc[ind,str(Var.configYaml['config']['excel_columns_config']['closure'])]
            gmud = df.loc[ind,str(Var.configYaml['config']['excel_columns_config']['gmud'])]
            adequacy = df.loc[ind,str(Var.configYaml['config']['excel_columns_config']['adequacy'])]
            host = host_value
            primaryContactEmail = df.loc[ind,str(Var.configYaml['config']['excel_columns_config']['primaryContactEmail'])]
            description = df.loc[ind, str(Var.configYaml['config']['excel_columns_config']['description'])]
            preCheck = df.loc[ind, str(Var.configYaml['config']['excel_columns_config']['preCheck'])]
            afterCheck = df.loc[ind, str(Var.configYaml['config']['excel_columns_config']['afterCheck'])]

            # index_closure = closure.find('-')
            # if index_closure <= 9:
            #     index_closure *= -20
            # else:
            #     pass
            # treated_closure_str = closure[:index_closure]

            #print(treated_closure_str, ' - ', closure)
            #print(index_closure)
            outlook = win.Dispatch('outlook.application')
            email = outlook.CreateItem(0)

            email.to = emailTo
            if closure == 'sucesso':
                email.Cc = Var.configYaml['config']['email_configure']['success_copy']
            elif closure == 'insucesso' or closure == 'cancelada':
                email.Cc = Var.configYaml['config']['email_configure']['canceled_insuccesss_copy']
            else:
                pass
            email.Subject = f'[{closure}] - {gmud} - {adequacy}'
            email.HTMLBody = f'''
            <p>Prezados,&nbsp;</p>
            <p>Segue o status da mudan&ccedil;a:&nbsp;<strong>[{closure}] - {gmud} - {adequacy}</strong></p>
            <p>&nbsp;</p>
            <p>Respons&aacute;vel para valida&ccedil;&atilde;o:&nbsp;<strong>{primaryContactEmail}</strong></p>
            <p>Servidor:&nbsp;<strong>{host}</strong></p>
            <p>Contato plantonista:&nbsp;<strong>{emailTo}</strong></p>
            <p>Alertas Zabbix antes execu&ccedil;&atilde;o:&nbsp;<strong>{preCheck}</strong></p>
            <p>Alertas Zabbix p&oacute;s execu&ccedil;&atilde;o:&nbsp;<strong>{afterCheck}</strong></p>
            <p>Observa&ccedil;&otilde;es:&nbsp;<strong>{description}</strong><br /p>
            <br />
            <p><strong>Por gentileza, validar o funcionamento do ambiente.</strong></p>
            <p>{str(Var.rodape)}
            <img alt="" src="https://i.ibb.co/s1vcggZ/thumbnail-Outlook-ir4tv3sj.png">
            </p>   
            '''
            path = str(Var.configYaml['config']['paths']['path_to_prints'])
            pattern = f'*{gmud}*'
            filesListArray = []

            for dName, sdName, fList in os.walk(path):
                for fileName in fList:
                    if fnmatch.fnmatch(fileName, pattern):
                        filesListArray.append(os.path.join(dName, fileName))
            filesListArray = sorted(filesListArray)
            print(filesListArray)
            for i in filesListArray:
                email.Attachments.Add(i)
            email.Send()
    elif option == 'exit':
        break
    elif option == '':
        pass
    else:
        print("[ERROR] - Invalid command, type 'azulino --help' to see more options")
print('hasta luego!')
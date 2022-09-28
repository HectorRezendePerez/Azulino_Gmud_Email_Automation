import win32com.client as win
from variables import var
class func:
    '''
    def To_(self, qntd_email):
        To_emails = []
        for ind in (0, qntd_email):
            email = input('para quem é o email?')
            To_emails.append(email)
        return To_emails
    '''
    def cria_email(tipo ,para ,copia , GMUD ,HOST_service ,descricao ):
        outlook = win.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        if tipo == 'validacao':
            email.To = para
            email.CC = copia
            email.Subject = f'[VALIDAÇÃO] - GMUD: {GMUD} - Adequações de Segurança - {HOST_service}'
            email.HTMLBody = f"""
            <p>Prezados,</p>
            <p class="elementToProof"><br></p>
            <p>Implantada GMUD:</p>
            <p class="elementToProof"><br></p>
            <p>{GMUD} - Adequações de Segurança - {HOST_service} </p>
            <p class="elementToProof"><br></p>
            <p>{descricao}</p>
            <p class="elementToProof"><br></p>
            <p>Favor validar o funcionamento do ambiente.</p>
            <p class="elementToProof"><br></p>
            {var.rodape}
            """
        elif tipo == 'encerrada':
            email.To = para
            email.CC = copia
            email.Subject = f'[ENCERRADA] - GMUD: {GMUD} - Adequações de Segurança - {HOST_service}'
            email.HTMLBody = f"""
            <p>Prezados,</p>
            <p class="elementToProof"><br></p>
            <p>Implantada GMUD:</p>
            <p class="elementToProof"><br></p>
            <p>{GMUD} - Adequações de Segurança - {HOST_service} </p>
            <p class="elementToProof"><br></p>
            <p>{descricao}</p>
            <p class="elementToProof"><br></p>
            {var.rodape}
            """
        elif tipo == 'cancelada':
            email.To = para
            email.CC = copia
            email.Subject = f'[CANCELAMENTO] - GMUD: {GMUD} - Adequações de Segurança - {HOST_service}'
            email.HTMLBody = f"""
            <p>Prezados,</p>
            <p class="elementToProof"><br></p>
            <p>GMUD:</p>
            <p class="elementToProof"><br></p>
            <p>{GMUD} - Adequações de Segurança - {HOST_service} </p>
            <p class="elementToProof"><br></p>
            <p>CANCELADA,</p>
            <p class="elementToProof"><br></p>
            <p>{descricao}</p>
            <p class="elementToProof"><br></p>
            {var.rodape}
            """
        email.Save()
        return print('email criado  e salvo para anexar evidendias')
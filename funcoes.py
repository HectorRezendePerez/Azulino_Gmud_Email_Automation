import win32com.client as win
from variables import var
class func:
    def cria_email(tipo ,para , GMUD ,HOST_service ,descricao ):
        outlook = win.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        if tipo == 'validacao':
            email.To = para
            email.CC = var.cc_teste
            email.Subject = f'[VALIDAÇÃO] - GMUD: {GMUD} - Adequações de Segurança - {HOST_service}'
            email.HTMLBody = f"""
            <p>Prezados,</p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            <p>Implantada GMUD:</p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            <p>
            <b>{GMUD}</b>
             - Adequações de Segurança - 
            <b>{HOST_service}</b> 
            </p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            <p>{descricao}</p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            <p>Favor validar o funcionamento do ambiente.</p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            {var.rodape}
            """
        elif tipo == 'encerrada':
            email.To = para
            email.CC = var.cc_teste
            email.Subject = f'[ENCERRADA] - GMUD: {GMUD} - Adequações de Segurança - {HOST_service}'
            email.HTMLBody = f"""
            <p>Prezados,</p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            <p>Implantada GMUD:</p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            <p>
            <b>{GMUD}</b>
             - Adequações de Segurança - 
            <b>{HOST_service}</b> 
            </p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            <p>{descricao}</p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            {var.rodape}
            """
        elif tipo == 'cancelada':
            email.To = para
            email.CC = var.cc_teste
            email.Subject = f'[CANCELAMENTO] - GMUD: {GMUD} - Adequações de Segurança - {HOST_service}'
            email.HTMLBody = f"""
            <p>Prezados,</p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            <p>GMUD:</p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            <p>
            <b>{GMUD}</b>
             - Adequações de Segurança - 
            <b>{HOST_service}</b> 
            </p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            <p>CANCELADA,</p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            <p>{descricao}</p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            {var.rodape}
            """
        elif tipo == 'inssucesso':
            email.To = para
            email.CC = var.cc_teste
            email.Subject = f'[INSSUCESSO] - GMUD: {GMUD} - Adequações de Segurança - {HOST_service}'
            email.HTMLBody = f"""
            <p>Prezados,</p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            <p>INSUCESSO da GMUD:</p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            <p>
            <b>{GMUD}</b>
             - Adequações de Segurança - 
            <b>{HOST_service}</b> 
            </p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            <p>{descricao}</p>
            <p class="elementToProof"><br></p>
            <p class="elementToProof"><br></p>
            {var.rodape}
            """
        else:
            print('tipo invalido, favor ajustar')
        email.Save()
        return print('email criado  e salvo para anexar evidendias')
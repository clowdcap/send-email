import win32com.client as win32
# para baixar a biblioteca -> pip install pywin32

# criar integração com o outlook
outlook = win32.Dispatch('outlook.application')

# criar um e-mail
email = outlook.CreateItem(0)

# configurar as informações
email.To = 'email@email.com' # <- e-mail diretório
email.Subject = 'Campos variáveis' # <- assunto do e-mail

# variváveis
nome = 'Nome' # <- nome da pessoa que vc vai mandar o e-mail
sobrenome = 'Sobrenome' # <- sobrenome da pessoa que vc vai mandar o e-mail

# adicionando anexo
nome_arquivo = 'db.xlsx' # <- nome do arquivo para o anexo
anexo = fr'E:\ ... \enviar-email\anexo\{nome_arquivo}' # <- copiar do diretório completo
# deixar o f para aceitar variável e r para que escreva o diretorio sem interferencias de manipulação de str.
if anexo == '':
    print('Sem Anexo...') # <- caso você não queria anexar arquivo, na variável anexo deixar em branco = ''
else:
    email.Attachments.Add(anexo) # <- Anexando o documento
    print('Anexado documento...')

# HTMLBody
# css
css = '''
<style>     
            /* Reset */
            * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }

            body {
                width: 100%;
                overflow-x: hidden;
            }

            .email {
                margin: 2%;
            }

            .topo {
                padding: 0.3rem 0; 
                background-color: brown;
                width: 100%;
                text-align: center !important;
                margin-bottom: 35px;
                font-family: Arial, Helvetica, sans-serif;
            }

            .topo h2 {
                font-size: 28px;
                color: white;
                padding-top: 25px;
            }

            img {
                width: 80px;
                height: 80px;
            }

            .capa {
                text-align: center;
                width: 100%;
                padding: 0.5rem 0;
                background-color: cadetblue;
            }
            .capa p {
                font-size: 16px;
                font-family: Arial, Helvetica, sans-serif;
                color: white;
                text-align: left !important;
                padding: 0 30px;
            }

            .conteudo {
                text-align: center;
                width: 100%;
                padding: 2rem;
                background-color: gainsboro;
            }

            .conteudo p{
                font-size: 20px;
                font-family: Arial, Helvetica, sans-serif;
            }

            .conteudo p a {
                text-decoration: none;
                color: red;
            }

            .assinatura {
                background-color: royalblue;
                color: white;
                font-family: Arial, Helvetica, sans-serif;
                margin-top: 20px;
            }
        </style>
'''

# HTML body <- você deve alterar os dados dentro das suas variáveis do body, proceso é simples.
# Você cria uma variável que armazene o conteudo HTML, podendo estilizar na variavel CSS
# Assim ter uma dinamica melhor na hora de montar o corpo do e-mail.

conteudo = f'''
<div class="conteudo">
                <p>Bom dia {nome} {sobrenome},</p>
                <br>
                <p>Esse é um email simples que pode ser aplicado por qualquer pessoa</p> 
                <br>
                <p>Crie um ambiente python, no caso desse código, eu realizei no PyCharm Community com Python
                3.9 instalado na máquina.</p> 
                <br>
                <p>Caso você deja mandar pelo email, um link, também da certo:</p>
                <p>Link: <a href="https://www.google.com/">
https://www.google.com/</a></p>
                <br>
                <p>Caso tenha alguma dúvida, estou à disposição</p>

                
</div> <!--conteudo-->
'''

assinatura = '''
 <div class="assinatura">
    <p>Atenciosamente,</p>
        <br>
        <br>
    <h3>Jose Marinho</h3>
        <br>
    <p>Whatsapp: (41) X XXXX-XXXX</p>
    <p>Telefone: (41) XXXX-XXXX - </p>
    <p>E-mail: clowdcap@hotmail.com</p>
    <p>Nome Completo da Sua Empresa LTDA.</p>
</div> <!--assinatura-->
'''

topo = '''
<div class="topo">
    <h2>Nome Completo da Sua Empresa LTDA.</h2>
</div> <!--topo-->
'''

capa = '''
<div class="capa">
    <p>Atendimento via E-mail - A/C: <b>José Marinho - Developer</b></p>
</div> <!--capa-->
'''

# Montagem do corpo do HTML, criado bloco com a semântica do HTML e integração das variáveis (por isso o f inicial)
email.HTMLBody = f''' 
<!DOCTYPE html>
<html>
    <head>
        <meta charset="utf-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        {css}
    </head>
    <body>
        <section class="email">
            
            {topo}

            {capa}
            
            {conteudo}
                
            {assinatura}
           
        </section> <!--email-->
    </body>
</html>

'''

# finalizando email
email.Send()
print('Email enviado !')

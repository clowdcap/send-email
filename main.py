import win32com.client as win32

# criar integração com o outlook
outlook = win32.Dispatch('outlook.application')

# criar um email
email = outlook.CreateItem(0)

# configurar as informações
email.To = 'email@email.com' # <- email diretório
email.Subject = 'Campos variáveis' # <- assunto do e-mail

# variváveis
nome = 'Nome' # <- nome da pessoa que vc vai mandar o email
sobrenome = 'Sobrenome' # <- sobrenome da pessoa que vc vai mandar o email

# adicionando anexo
nome_arquivo = 'db.xlsx' # <- nome do arquivo para o anexo
anexo = fr'E:\ ... \enviar-email\anexo\{nome_arquivo}' # <- copiar do diretório completo
# deixar o f para aceitar variável e r para que escreva o diretorio sem interferencias de manipulação de str.
if anexo == '':
    print('Sem Anexo...') # <- caso você não queria anexar arquivo, na variável anexo deixar em branco = ''
else:
    email.Attachments.Add(anexo) # <-
    print('Anexado documento...')

# HTMLBody
# css
css = '''
<style>
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

# HTML body
conteudo = f'''
<div class="conteudo">
                <p>Bom dia {nome} {sobrenome},</p>
                <br>
                <p>Como conversados na sua visita ao Setor de Urbanismo da prefeitura de Campo Magro, fiquei de 
responder algumas dúvidas suas e lhe retorno com essa mensagem para ajudá-lo.</p> 
                <br>
                <p>Sobre o loteamento de interesse social, é importante lembrar que deve ter um interesse público e 
privado, logo, pedimos que, após a idealização de locais para ter esse loteamento, deve ser aberto uma guia amarela, 
solicitando um "", para que possa ter uma análise. O critério é que a área estimada, deve ter acesso a coleta de 
esgoto e se caso não tiver, propor alguma solução de estratégia para levar esse serviço para área.</p> 
                <br>
                <p>E o mapa de zoneamento, pode ser pego no site da COMEC:</p>
                <p>Link: <a href="http://www.comec.pr.gov.br/Pagina/UTP-Campo-Magro">
http://www.comec.pr.gov.br/Pagina/UTP-Campo-Magro</a></p>
                <br>
                <p>Caso eu não tenha esclarecido totalmente a sua dúvida, estou à disposição</p>

                
</div>
'''

assinatura = '''
 <div class="assinatura">
    <p>Atenciosamente,</p>
        <br>
        <br>
    <h3>Jose Marinho</h3>
        <br>
    <p>Whatsapp: (41) 9 9272-5388</p>
    <p>Telefone: (41) 3677-4050 - Setor Urbanismo</p>
    <p>jm.arquiteturacwb@gmail.com</p>
    <p>Prefeitura Municipal de Campo Magro / PR</p>
</div>
'''

topo = '''
<div class="topo">
    <img src="https://leismunicipais.com.br/img/cidades/pr/campo-magro.png" alt="campo-magro">
    <h2>Prefeitura Municipal de Campo Magro</h2>
</div> <!--topo-->
'''

capa = '''
<div class="capa">
    <p>Atendimento via E-mail - A/C: <b>José Marinho - Estagiário</b></p>
</div> <!--capa-->
'''

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
           
        </section>
    </body>
</html>

'''

# finalizando email
email.Send()
print('Email enviado !')

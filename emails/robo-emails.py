import pandas as pd
import time
import smtplib
import email.message


try:
    file_path = r'C:\\Users\\Igor\\Desktop\\trabalho\\emails.xlsx' 
    df = pd.read_excel(file_path, engine='openpyxl')


    if 'email' not in df.columns:
        print("Erro: Coluna 'email' não encontrada na planilha.")
    else:

        for index, row in df.iterrows():
            email1 = row['email']
            def enviar_email():  
                corpo_email = """
                    <!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Email Inovador</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            background: linear-gradient(135deg, #007BFF, #00BFFF);
        }
        .container {
            max-width: 600px;
            margin: 20px auto;
            background-color: #ffffff;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
        }
        .header {
            background: linear-gradient(135deg, #911BFF, #00BFFF);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 {
            margin: 0;
            font-size: 28px;
        }
        .content {
            padding: 20px;
            text-align: center;
            line-height: 1.6;
        }
        .content h2 {
            color: #007BFF;
            margin-top: 0;
        }
        .cta {
            background-color: #100;
            color: white;
            border-radius: 5px;
            padding: 15px 30px;
            margin: 20px 0;
            display: inline-block;
            text-decoration: none;
            font-weight: bold;
            transition: background-color 0.3s;
        }
        .cta:hover {
            background-color: #004494;
        }
        .footer {
            background-color: #f9f9f9;
            text-align: center;
            padding: 15px;
            font-size: 12px;
            color: #555;
        }
        .social-icons img {
            width: 30px;
            margin: 0 5px;
        }
        .highlight {
            color: #00AFFF;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Inovação em Automação</h1>
        </div>
        <div class="content">
            <h2>Transforme sua Empresa!</h2>
            
            <p>Olá,</p>
            <p>Desde 2022, nossa missão é <span class="highlight">automatizar e dinamizar tarefas</span> para que empresas não sofram mais com desperdicio de tempo. Com o nosso sistema gerenciador de automações RPA. Acreditamos que a automação é o futuro!</p>
            <a href="#"><img src="https://1drv.ms/i/c/98feb852d4b09705/IQTkyw0XO1phTr2r9ERB8bT0AVL0Mof0CUDJTNIIXloGs8E?width=1024" alt="imagem exemplo"></a>
            <p>Junte-se a nós e descubra como podemos otimizar seus processos.</p>
            <a href="https://wa.me/message/FGIN3CMRMB75F1" class="cta">Saiba Mais</a>
        </div>
        <div class="footer">
            <p>&copy; 2024 Grupo Satelite. Todos os direitos reservados.</p>
            <div class="social-icons">
                <a href="#"><img src="https://wa.me/message/FGIN3CMRMB75F1" alt="Whatsapp"></a>
                <a href="#"><img src="https://via.placeholder.com/30x30?text=TW" alt="Instagram"></a>

            </div>
            <p><a href="#" style="color: #007BFF;">Cancelar inscrição</a></p>
        </div>
    </div>
</body>
</html>


                """

                msg = email.message.Message()
                msg['Subject'] = "Elimine as tarefas repetitivas AGORA!!"
                msg['From'] = '#coloque aqui o seu email'
                msg['To'] = email1
                password = '#aqui você precisa colocar a sua senha de aplicativo disponibilizada nas configurações do google' 
                msg.add_header('Content-Type', 'text/html')
                msg.set_payload(corpo_email)

                s = smtplib.SMTP('smtp.gmail.com: 587')
                s.starttls()
                
                s.login(msg['From'], password)
                s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
                print('Email enviado')


            # In[ ]:


            enviar_email()
        time.sleep(5)
        print(f"Enviando e-mail para: {email}")

except FileNotFoundError:
    print("Erro: Arquivo 'emails.xlsx' não encontrado. Verifique o caminho.")
except Exception as e:
        print(f"Ocorreu um erro: {e}")

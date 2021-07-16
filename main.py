# https://wordtohtml.net/
# https://www.python.org/downloads/
# pip install PySimpleGUI
# pip install python-dotenv


import PySimpleGUI as sg  # graphic
import time  # timer
import csv  # read and write csv
import smtplib  # email
from email.mime.multipart import MIMEMultipart  # email
from email.mime.text import MIMEText  # email
import os
from dotenv import load_dotenv


def app_interface():
    load_dotenv()
    MY_ENV_EMAIL = os.getenv('USER_EMAIL')
    MY_ENV_PASS = os.getenv('APP_PASSWORD')
    MY_ENV_HOST = os.getenv('SERVER_HOST')
    MY_ENV_PORT = os.getenv('SERVER_PORT')
    MY_ENV_FORM_EMAIL = os.getenv('FORM_EMAIL')
    MY_ENV_FORM_NAME = os.getenv('FORM_NAME')
    MY_ENV_FORM_CPF = os.getenv('FORM_CPF')

    sg.theme('Reddit')  # page themes
    # component distribution
    layout = [
        [sg.Text('E-mail', size=(10, 1), justification='left'), sg.Input(MY_ENV_EMAIL, key='email'),
         sg.Button('Testar Conexão', size=(18, 1))],
        [sg.Text('Senha', size=(10, 1), justification='left'), sg.Input(MY_ENV_PASS, key='pass')],  # ,password_char='*'
        [sg.Text('Tabela', size=(10, 1), auto_size_text=False, justification='left'),
         sg.InputText('Arquivo ".csv"', key='path'), sg.FileBrowse("Selecionar Arquivo", size=(18, 1))],
        [sg.Text('Assunto', size=(10, 1), justification='left'), sg.Input(key='subject')],
        [sg.Text('Conteúdo', size=(10, 1), justification='left'), sg.Input(key='content')],
        [sg.Button('Transmitir e-mails', size=(18, 1))],
        [sg.Output(key='text_area', size=(None, 20))],
        [sg.Text('Coluna Nome', size=(10, 1), justification='left'),
         sg.Input(MY_ENV_FORM_NAME, key='form_name', size=(4, 1), justification='center', disabled=True)],
        [sg.Text('Coluna E-mail', size=(10, 1), justification='left'),
         sg.Input(MY_ENV_FORM_EMAIL, key='form_email', size=(4, 1), justification='center', disabled=True)],
        [sg.Text('Coluna CPF', size=(10, 1), justification='left'),
         sg.Input(MY_ENV_FORM_CPF, key='form_cpf', size=(4, 1), justification='center', disabled=True)]
    ]
    my_window = sg.Window('Emissor de E-mails', layout)  # janela do windows
    status_system = True  # key system

    def select_host(email, my_window):
        host_email = ''
        if '@gmail.' in email:
            host_email = 'smtp.gmail.com'  # declaração do host (provedor) de email
        elif '@outlook.' in email or '@hotmail.' in email:
            host_email = 'smtp-mail.outlook.com'  # declaração do host (provedor) de email
        else:
            sg.popup_ok(
                "Verifique se o email foi preenchido corretamente\nCaso o erro continue, contate o Administrador")
            print('Provedor Host não definido, favor inserir manualmente!')
            my_window.Refresh()

        if MY_ENV_HOST != "":
            host_email = MY_ENV_HOST

        return host_email

    controller_print = True

    def introduction():
        print('Para personalizar seu texto, acesse o link: https://wordtohtml.net/')
        print('Para copiar corretamente o texto gerado utilize os comando:')
        print('Click com botão esquerdo na área  >  Ctrl+A  >  Ctrl+C')
        print('')
        print(
            'Legenda de dados:\n-> name (nome do convidado)\n-> email (email do convidado)\n-> cpf (cpf do convidado)')
        print('As legendas devem ser colocadas entre colchetes "{" "}", exemplo: {name}')
        print('')

    while status_system:  # enquanto não clicar no X
        event, values = my_window.read()

        if controller_print:
            introduction()
            my_window.Refresh()
            controller_print = False

            # encerra aplicação
        if event == sg.WINDOW_CLOSED:
            status_system = False
            my_window.close()
            quit()

        if event == 'Testar Conexão':
            try:
                try:
                    host_email = select_host(values['email'], my_window)
                    if host_email == '':
                        print('Erro ao definir provedor host, altere-o na pasta env')
                except:
                    print('Erro! Verifique se o email faz parte do dominio Gmail ou Outlook.')
                connection = smtplib.SMTP(host=host_email, port=MY_ENV_PORT)
                connection.starttls()
                connection.login(values['email'], values['pass'])
                connection.quit()
                print('Usuário autentincado! Prossiga')
                print('')
            except:
                print(
                    'Falha ao conectar com o servidor! Verifique os dados de usuário e as configurações do provedor de email.')
                print('')

        if event == 'Transmitir e-mails':
            # se campo email não for vazio
            if values['email'] != '':
                # se campo senha não for vazio
                if values['pass'] != '':
                    # se campo tabela for do tipo .cvs e não for vazio
                    if values['path'] != '' and values['path'][-4::] == '.csv':
                        # se campo de assunto ou conteudo tiverem dados
                        if values['subject'].replace(' ', '') != '' or values['content'].replace(' ', '') != '':
                            print('Iniciando teste de conexão com e-mail...')
                            my_window.Refresh()
                            connection_attempt_email = False

                            name_form = values['path']
                            name_list = []
                            email_list = []
                            cpf_list = []
                            # phone_list = []
                            # message_list = []

                            username = values['email']  # e-mail que será usado para encaminhar as mensagens
                            password = values['pass']  # senha de aplicativo que precisa ser gerada no provedor do email
                            mail_from = username

                            host_email = select_host(values['email'], my_window)

                            print('Configuração de porta de conexão foi um sucesso! Host: {}'.format(host_email))
                            my_window.Refresh()

                            # tentativa de coleta de dados da planila .csv
                            try:
                                print('Iniciando coleta dos dados da tabela .csv')
                                my_window.Refresh()
                                with open(name_form) as csv_file:
                                    csv_reader = csv.reader(csv_file, delimiter=',')
                                    csv_reader.__next__()
                                    for row in csv_reader:
                                        name_list.append(row[int(MY_ENV_FORM_NAME)].replace('Ã©', 'é'))
                                        email_list.append(row[int(MY_ENV_FORM_EMAIL)])
                                        cpf_list.append(row[int(MY_ENV_FORM_CPF)])
                                        # phone_list.append(row[4])
                                        # message_list.append(row[5])
                            except:
                                sg.popup_ok("Falha ao coletar dados da tabela .csv!")
                                print('Coleta de dados falhou!')
                                my_window.Refresh()
                                continue

                            # tentativa de encaminhamento de emails
                            try:
                                print('Iniciando envio dos e-mails. Total de {} emails a serem enviados'.format(
                                    len(email_list)))
                                print('')
                                my_window.Refresh()
                                for x in range(0, len(email_list), 1):
                                    try:
                                        mail_to = email_list[x]
                                        mail_subject = values['subject']
                                        # https://wordtohtml.net/
                                        message_html = '{content}'.format(content=values['content'])
                                        message_html = message_html.format(name=name_list[x], email=email_list[x],
                                                                           cpf=cpf_list[x])
                                        message_html.replace('Á', '&Aacute;').replace('É', '&Eacute;').replace('Í',
                                                                                                               '&Iacute;').replace(
                                            'Ó', '&Oacute;').replace('Ú', '&Uacute;').replace('á', '&aacute;').replace(
                                            'é', '&eacute;').replace('í', '&iacute;').replace('ó', '&oacute;').replace(
                                            'ú', '&uacute;').replace('Â', '&Acirc;').replace('Ê', '&Ecirc;').replace(
                                            'Ô', '&Ocirc;').replace('â', '&acirc;').replace('ê', '&ecirc;').replace('ô',
                                                                                                                    '&ocirc;').replace(
                                            'À', '&Agrave;').replace('à', '&agrave;').replace('Ü', '&Uuml;').replace(
                                            'ü', '&uuml;').replace('Ç', '&Ccedil;').replace('ç', '&ccedil;').replace(
                                            'Ã', '&Atilde;').replace('Õ', '&Otilde;').replace('ã', '&atilde;').replace(
                                            'õ', '&otilde;').replace('Ñ', '&Ntilde;').replace('ñ', '&ntilde;').replace(
                                            '&', '&amp;').replace('<', '&lt;').replace('>', '&gt')
                                        mimemsg = MIMEMultipart()
                                        mimemsg['From'] = mail_from
                                        mimemsg['To'] = mail_to
                                        mimemsg['Subject'] = mail_subject
                                        mimemsg.attach(MIMEText(message_html, 'html'))
                                        connection = smtplib.SMTP(host=host_email, port=MY_ENV_PORT)
                                        connection.starttls()
                                        connection.login(username, password)
                                        connection.send_message(mimemsg)
                                        connection.quit()
                                        time.sleep(1)
                                        print('{} / {} - Encaminhado email para {}'.format(len(email_list), x + 1,
                                                                                           email_list[x]))
                                        my_window.Refresh()
                                    except:
                                        print(
                                            'Atenção! {} / {} - Falha ao enviar email {}'.format(len(email_list), x + 1,
                                                                                                 email_list[x]))
                                        archive = open('log-error.txt', 'r')
                                        content = archive.readlines()
                                        content.append('{},'.format(email_list[x]))
                                        archive = open('log-error.txt', 'w')
                                        archive.writelines(content)
                                        archive.close()

                                print('')
                                print('Envio de emails finalizado!')
                                my_window.Refresh()

                            except:
                                sg.popup_ok(
                                    "Falha ao encaminhar e-mails, verifique se os dados\nde usuário estão corretos!")
                                print('Falha ao encaminhar e-mails, verifique os dados de usuário.')
                                my_window.Refresh()
                                continue
                        else:
                            sg.popup_ok("Preencha os campos de Assunto e Conteúdo!")
                            continue
                    else:
                        sg.popup_ok("Selecione o arquivo correto")
                        continue
                else:
                    sg.popup_ok("Preencha o campo senha")
                    continue
            else:
                sg.popup_ok("Preencha o campo e-mail")
                continue


if __name__ == '__main__':
    print('run code')
    try:
        app_interface()
    except:
        print('err')

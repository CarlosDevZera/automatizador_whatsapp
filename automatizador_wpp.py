import openpyxl
from urllib.parse import quote
import webbrowser
import pyautogui
from time import sleep
from datetime import datetime

workbook = openpyxl.load_workbook('D:/PyProjetos/pjt007/automatizador_whatsapp/clientes_exemplo.xlsx')
clientes = workbook['Folha1']

# Obter a data de hoje, sem a parte de horas/minutos/segundos
hoje = datetime.now().date()

for linha in clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento_excel = linha[2].value  # Data de vencimento lida do Excel

    data_vencimento = None
    # Converter a data de vencimento do Excel para um objeto date
    if isinstance(vencimento_excel, datetime):
        data_vencimento = vencimento_excel.date()
    else:
        print(
            f"AVISO: Data de vencimento para '{nome}' é inválida ou ausente: '{vencimento_excel}'. Pulando este contato.")
        # Adicionar ao log de erros, se desejar
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:  # Defina 'caminho_log_erros' antes
            arquivo.write(f'{nome},{telefone},Data de vencimento inválida\n')
        continue  # Pula para o próximo contato, crucial para evitar o erro

    # Calcular a diferença entre a data de vencimento e hoje
    diferenca_dias = (data_vencimento - hoje).days

    # Verificar se faltam exatamente 5 dias para o vencimento
    if diferenca_dias == 5:
        # Mensagem personalizada para o lembrete de 5 dias
        mensagem = f'Olá {nome}, tudo bem? Seu boleto conosco vence em 5 dias, no dia {data_vencimento.strftime('%d/%m/%Y')}.'

        try:
            link_mensagem = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
            webbrowser.open(link_mensagem)
            sleep(15)


            seta = None
            tentativas_localizacao = 0
            while seta is None and tentativas_localizacao < 5:  # Tenta por até 5 segundos
                seta = pyautogui.locateOnScreen('D:/PyProjetos/pjt007/automatizador_whatsapp/img/1.PNG', confidence=0.9)  # Adicione confidence
                sleep(1)
                tentativas_localizacao += 1

            if seta:
                pyautogui.click(seta[0], seta[1])
                sleep(5)
                pyautogui.hotkey('ctrl', 'w')
                print(f"Mensagem enviada para {nome} (vencimento em 5 dias).")

            else:
                print(f'Não foi possível localizar o botão de enviar para {nome}. Fechando a aba.')
                pyautogui.hotkey('ctrl', 'w')  # Fecha a aba mesmo se não encontrar o botão
                sleep(2)
                with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                    arquivo.write(f'{nome},{telefone},Botão de enviar não localizado\n')

        except Exception as e:
            print(f'Não foi possível enviar mensagem para {nome}. Erro: {e}')
            with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                arquivo.write(f'{nome},{telefone},{e}\n')


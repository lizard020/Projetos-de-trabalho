import win32com.client as win32
import pandas as pd

caminho = input("Insira o caminho da planilha: ")

# Lendo o arquivo Excel (ajuste o caminho do arquivo se necessário)
df = pd.read_excel(caminho)

# Inicializando o Outlook
outlook = win32.Dispatch('outlook.application')

print("Dados carregados.")

while True:
    msg = input("O que deseja fazer? (visualizar, enviar ou fechar) ")

    if msg == 'visualizar':
        print(df)

    elif msg == 'enviar':
        # Loop para cada linha da planilha
        for index, row in df.iterrows():
            # Criar o e-mail
            mail = outlook.CreateItem(0)
            
            # Definindo o destinatário
            mail.To = row['Email']
            
            # Definindo o assunto
            mail.Subject = f"Proposta de Acordo - {row['Processo']} - {row['Reclamante']}"
            
            # Definindo o corpo do e-mail
            mail.Body = (f"Sou o Fernando do time de acordos. Quero apresentar essa proposta:\n"
                        f"\n"
                        f"Nº do Processo: {row['Processo']}\n"
                        f"Reclamante: {row['Reclamante']}\n"
                        f"Valor do acordo: {row['Valor']}\n"
                        f"\n"
                        "Aguardo retorno.")
            
            # Exibir o e-mail (sem enviar)
            mail.Display()

    elif msg == 'fechar':
        break  # Encerra o loop

    else:
        print("Comando inválido. Insira: visualizar, enviar ou fechar.")

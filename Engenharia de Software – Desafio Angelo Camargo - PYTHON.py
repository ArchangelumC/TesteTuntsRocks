import gdown
import pandas as pd

# Specifying the Google Drive file URL
drive_url = "https://docs.google.com/spreadsheets/d/12ORB7DwpwZTpAs8aItQWBGIGYESGFCx1/edit?usp=sharing&ouid=101996123968694005289&rtpof=true&sd=true"

# Get the file ID from the Google Drive URL
file_id = drive_url.split('/')[-2]

# Specify the local file path where you want to save the downloaded Excel file
local_file_path = "Desafio.xlsx"

# Download the file using gdown
gdown.download(f"https://drive.google.com/uc?id={file_id}", local_file_path, quiet=False)

print(f"File downloaded successfully to {local_file_path}")

dados = pd.read_excel(local_file_path)

# Operation to verify the approval situation for each student
for indice, linha in dados.iterrows():
    faltas = linha['Faltas']
    Aluno = linha['Aluno']
    P1 = linha['P1']
    P2 = linha['P2']
    P3 = linha['P3']

    media = (P1 + P2 + P3) / 3
    #Operation to fill up the sheets with the results and situations
    if faltas > 15: #Coresponds at 25% of the Total number of Classes
        dados.at[indice, 'Situação'] = "Reprovado por Falta"
        dados.at[indice, 'Nota para Aprovação Final'] = 0
    elif media >= 70:
        dados.at[indice, 'Situação'] = "Aprovado"
        dados.at[indice, 'Nota para Aprovação Final'] = 0
    elif media >= 50 and media <70:
        dados.at[indice, 'Situação'] = "Exame Final"
        dados.at[indice, 'Nota para Aprovação Final'] = round(100 - media) #It cames from the equation 50 <= (Media + Nota para Aprovação Final)/2
    else:
        dados.at[indice, 'Situação'] = "Reprovado por Nota"
        dados.at[indice, 'Nota para Aprovação Final'] = 0

#Saves the DataFrame in the sheet
dados.to_excel('Desafio.xlsx', index=False)

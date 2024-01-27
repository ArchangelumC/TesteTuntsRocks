import pandas as pd

# Reading of sheet data
sheet_location = 'Desafio.xlsx'
dados = pd.read_excel(sheet_location)

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

dados.to_excel('Desafio.xlsx', index=False)
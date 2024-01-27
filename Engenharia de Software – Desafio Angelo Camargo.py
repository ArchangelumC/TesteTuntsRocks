import pandas as pd

# Leitura dos dados da planilha
caminho_planilha = 'Desafio.xlsx'
dados = pd.read_excel(caminho_planilha)

# Cálculo da média e verificação de aprovação para cada aluno
for indice, linha in dados.iterrows():
    faltas = linha['Faltas']
    nome = linha['Aluno']
    P1 = linha['P1']
    P2 = linha['P2']
    P3 = linha['P3']

    media = (P1 + P2 + P3) / 3

    if faltas > 15:
        dados.at[indice, 'Situação'] = "Reprovado por Falta"
        dados.at[indice, 'Nota para Aprovação Final'] = 0
    elif media >= 70:
        dados.at[indice, 'Situação'] = "Aprovado"
        dados.at[indice, 'Nota para Aprovação Final'] = 0
    elif media >= 50 and media <70:
        dados.at[indice, 'Situação'] = "Exame Final"
        dados.at[indice, 'Nota para Aprovação Final'] = round(100 - media,2) #It cames from the equation 5 <= (m + naf)/2
    else:
        dados.at[indice, 'Situação'] = "Reprovado"
        dados.at[indice, 'Nota para Aprovação Final'] = 0



    print(f"Aluno: {nome} - Média: {media:.2f} - Situação: {linha['Situação']} - Nota para Aprovação Final:{linha['Nota para Aprovação Final']}")

dados.to_excel('Novo.xlsx', index=False)
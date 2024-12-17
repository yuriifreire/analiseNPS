import pandas as pd

#Puxar a planilha
planilha = "analiseDados.xlsx"
df = pd.read_excel(planilha)

print(df.head())

#Os Promotores tem notas 9 e 10, os Neutros tem notas 7 e 8 e os Detratores tem notas entre 0 e 6
def classificar_nps(nota):
    if nota >= 9:
        return "Promotor"
    elif 7 <= nota <= 8:
        return "Neutro"
    else:
        return "Detrator"

#Os Satisfeitos tem notas 4 e 5, os Neutros tem notas 3 e os Instatisfeitos tem notas 1 e 2
def classificar_csat(nota):
    if nota >= 4:
        return "Satisifeito"
    elif nota == 3:
        return "Neutro"
    else:
        return "Insatisfeito"

#Passar as funções dentro da planilha
df['Classificação NPS'] = df['NPS'].apply(classificar_nps)
df['Classificação CSAT'] = df['CSAT'].apply(classificar_csat)

#exibir planilha classificada
print(df)
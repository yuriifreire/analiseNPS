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
        return "Satisfeito"
    elif nota == 3:
        return "Neutro"
    else:
        return "Insatisfeito"

#Passar as funções dentro da planilha
df['Classificação NPS'] = df['NPS'].apply(classificar_nps)
df['Classificação CSAT'] = df['CSAT'].apply(classificar_csat)

#Exibir planilha classificada
print(df)

#Fazendo a separação de NPS por empresa
nps_por_empresa = df.groupby('EMPRESA')['Classificação NPS'].value_counts()
print(nps_por_empresa)

#Fazendo a separação de CSAT por empresa
csat_por_empresa = df.groupby('EMPRESA')['Classificação CSAT'].value_counts()
print(csat_por_empresa)

#Calcular o NPS por Empresa
def calcular_nps(grupo):
    total = len(grupo)
    promotores = len(grupo[grupo['Classificação NPS'] == "Promotor"])
    detratores = len(grupo[grupo['Classificação NPS'] == "Detrator"])
    nps = ((promotores - detratores) / total) * 100 #multiplicamos por 100 por se tratar de um valor em porcentagem
    return nps

#Calcular o CSAT por Empresa
def calcular_csat(grupo):
    total = len(grupo)
    satisfeitos = len(grupo[grupo['Classificação CSAT'] == "Satisfeito"])
    csat = (satisfeitos / total) * 100
    return csat


#Passando a função do cálculo de NPS por empresa
nota_nps_empresa = df.groupby('EMPRESA').apply(calcular_nps).reset_index()
nota_nps_empresa.columns = ['EMPRESA', 'NPS']
print(nota_nps_empresa)

#Passando a função do cálculo de CSAT por empresa
nota_csat_empresa = df.groupby('EMPRESA').apply(calcular_csat).reset_index()
nota_csat_empresa.columns = ['EMPRESA','CSAT']
print(csat_por_empresa)
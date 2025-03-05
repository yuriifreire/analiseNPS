import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

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

#Gerar planilha com resultado geral
df.to_excel("classificacao_geral_nps_csat.xlsx", index=False)

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

#Gerar planilha para NPS por empresa
nota_nps_empresa.to_excel("nps_por_empresa.xlsx", index=False)

#Passando a função do cálculo de CSAT por empresa
nota_csat_empresa = df.groupby('EMPRESA').apply(calcular_csat).reset_index()
nota_csat_empresa.columns = ['EMPRESA','CSAT']
print(csat_por_empresa)

#Gerar Planilha para CSAT por empresa
nota_csat_empresa.to_excel("csat_por_empresa.xlsx", index=False)


#Determinar a previsão com base no Indicador Combinado
def previsao(indicador):
    if indicador >= 75:
        return "Previsão de se manter ou se tornar promotor"
    elif 50 < indicador < 75:
        return "Previsão de se manter neutro"
    elif 0 < indicador <= 50:
        return "Previsão de se tornar detrator"
    else:
        return "Previsão de ocorrer um futuro cancelamento"


#Agrupar por cliente calculando o NPS, CSAT e o Indicador Combinado
indicadores_por_empresa = df.groupby('EMPRESA').apply(
    lambda grupo: pd.Series({
        'NPS': calcular_nps(grupo),
        'CSAT': calcular_csat(grupo),
        'Indicador Combinado': (calcular_nps(grupo) + calcular_csat(grupo)) / 2,
    })
).reset_index()

#Adicionando a coluna de previsão na planilha
indicadores_por_empresa['Previsão'] = indicadores_por_empresa['Indicador Combinado'].apply(previsao)

print(indicadores_por_empresa)

#Gerar planilha com a previsão
indicadores_por_empresa.to_excel("Previsao.xlsx", index=False)
print("Arquivo Salvo: Previsao.xlsx")


#Puxar os dados da planilha de previsao
planilha_previsao = "Previsao.xlsx"
df2 = pd.read_excel(planilha_previsao)

#Ordenar o DF pelo "Indicador Combinado" em ordem decrescente
df2 = df2.sort_values(by="Indicador Combinado", ascending=False)

#Configurando o estilo do seaborn
sns.set_style("whitegrid")

#Tamanho do gráfico
plt.figure(figsize=(12, 8))

#Criar o gráfico de barras
sns.barplot(
    x="Indicador Combinado",
    y="EMPRESA",
    hue="Previsão",
    data=df2,
    dodge=False, #Remover a separação das barras por hue
    palette="magma" #Definição da paleta de cores
)

#Configurando os rótulos e títulos
plt.title("Indicador Combinado por Empresa e Previsão", fontsize=16)
plt.xlabel("Indicador Combinado", fontsize=12)
plt.ylabel("Empresas", fontsize=12)


#Ajustando a legenda
plt.legend(title="Previsão", bbox_to_anchor=(1.05, 1), loc="upper left")


#Exibindo o gráfico
plt.tight_layout()
plt.show()


#Salvando o gráfico de barras em imagem .png
plt.savefig("previsao.png", dpi=300, bbox_inches='tight')


#Criar gráfico de pontos
pontos = sns.scatterplot(
    y="Indicador Combinado",
    x="EMPRESA",
    hue="Previsão",
    size="Indicador Combinado",
    sizes=(50,300), #Definindo o tamanho dos pontos baseado no Indicador Combinado
    palette="magma",
    data=df2
)

#Adicionando faixas específicas
plt.axhline(y=75, color='green', linestyle='--', label="75 (Promotor)")
plt.axhline(y=50, color='orange', linestyle='--', label="50 (Neutro)")
plt.axhline(y=0, color='red', linestyle='--', label="0 (Churn)")


#Configurando os rótulos e títulos
plt.title("Indicador Combinado por Empresa e Previsão", fontsize=16)
plt.ylabel("Indicador Combinado", fontsize=12)
plt.xlabel("Empresas", fontsize=12)


#Ajustando a legenda
plt.legend(title="Previsão", bbox_to_anchor=(1.05, 1), loc="upper left")


#Exibindo o gráfico
plt.tight_layout()
plt.show()

#Salvar o gráfico de pontos em imagem .png
plt.savefig("previsaoPontos.png", dpi=300, bbox_inches='tight')

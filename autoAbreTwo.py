import pandas as pd
import textwrap

# Função para dividir texto sem quebrar palavras e redistribuir entre colunas
def dividir_texto_sem_quebrar_palavras(texto, limite=45):
    # Usar textwrap para garantir que as palavras não sejam quebradas
    partes = textwrap.wrap(texto, width=limite, break_long_words=False)
    return partes

# Ler a planilha
df = pd.read_excel('C:/Users/ana.pinheiro/OneDrive - Telecall/Área de Trabalho/AutoAbreviacao/arquivoEntrada/cadastros_campos_concatenados.xlsx')

# Definir as colunas para abreviar
colunas_para_abreviar = ['billingStreet', 'billingDescription']

# Processar cada linha para dividir e redistribuir texto
for i, row in df.iterrows():
    # Concatenar os textos das colunas atuais com vírgulas
    texto_completo = ', '.join([str(row[coluna]) for coluna in colunas_para_abreviar if pd.notna(row[coluna])])
    partes = dividir_texto_sem_quebrar_palavras(texto_completo, limite=45)
    
    # Atualizar as colunas com as partes do texto
    for j in range(len(colunas_para_abreviar)):
        if j < len(partes):
            df.at[i, colunas_para_abreviar[j]] = partes[j]
        else:
            df.at[i, colunas_para_abreviar[j]] = ''

# Salvar a planilha atualizada
df.to_excel('C:/Users/ana.pinheiro/OneDrive - Telecall/Área de Trabalho/AutoAbreviacao/ArquivoSaida/cadastros_campos_concatenados_geral.xlsx', index=False)

print("Processo concluído e planilha salva com sucesso.")

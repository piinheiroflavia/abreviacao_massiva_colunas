import pandas as pd
import textwrap

# Função para abreviar texto
def abreviar_texto(texto, limite=45):
    if len(texto) > limite:
        return textwrap.shorten(texto, width=limite, placeholder="...")
    return texto

# Ler a planilha
df = pd.read_excel('C:/caminha_da_pasta_testEndereco.xlsx')


# Abreviar os valores das colunas desejadas
colunas_para_abreviar = ['nomeColuna1', 'nomeColuna2' ,'nomeColuna3', 'nomeColuna4']
for coluna in colunas_para_abreviar:
    df[coluna] = df[coluna].apply(lambda x: abreviar_texto(str(x), limite=45))

# Salvar a planilha atualizada
df.to_excel('C:/caminha_da_pasta_testEndereco.xlsx/testeEndereco.xlsx', index=False)
df.to_excel('', index=False)

print("Processo concluído e planilha salva com sucesso.")

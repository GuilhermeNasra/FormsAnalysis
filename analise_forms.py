import pandas as pd
import matplotlib.pyplot as plt
from operator import itemgetter


def input_palavras_chave():
    lista_palavras_chave = []
    print('Aperte <Enter> para interromper!\nEscreva as palavras-chave desejadas:')
    print('Para remover a última palavra-chave digitada, escreva <remover> e aperte <Enter>\n')

    palavras = input()
    while palavras != '':
        if palavras == 'remover':
            removido = lista_palavras_chave.pop()
            print(f'A palavra-chave <{removido}> foi removida com sucesso.\n')
        else:
            lista_palavras_chave.append(palavras)
        palavras = input()
    return lista_palavras_chave


def quantidade_respostas(lista_palavras_chave, op_fisica):
    i = 0
    respostas = []
    phrase = []
    df = pd.DataFrame()
    while i < len(lista_palavras_chave):
        contador = 0
        for indice, linha in op_fisica.iteritems():
            if lista_palavras_chave[i] in linha:
                contador += 1
                phrase.append(linha)
            else:
                pass
        df.insert(i, lista_palavras_chave[i], pd.Series(phrase), allow_duplicates=True)
        phrase.clear()
        respostas.append((lista_palavras_chave[i], contador))
        i += 1
    df.to_excel(r'keywords_dataframe.xlsx')

    # Ordenando os números de forma decrescente
    numeros_ordenados = sorted(respostas, key=itemgetter(1), reverse=True)
    for i in numeros_ordenados:
        print('Quantidade de respostas com {}: {}'.format(i[0], i[1]))
    return numeros_ordenados


def plot_graph(numeros_ordenados):
    # Plotando o gráfico
    list_x = []
    list_y = []
    for x, y in numeros_ordenados:
        list_x.append(x)
        list_y.append(y)

    # Gráfico de barras
    plt.figure(figsize=(15, 8))
    plt.title('Análise de palavras-chave')
    plt.bar(list_x, list_y, color='red')
    for i in range(len(list_x)):
        plt.text(i, list_y[i], list_y[i], ha='center')
    plt.xticks(rotation=45)
    plt.xlabel('Palavras-chave')
    plt.ylabel('Quantidade')

    plt.show()


def main():
    # Importando planilha
    tabela_df = pd.read_excel('base_teste.xlsx')
    op_fisica = tabela_df['base']
    op_fisica = op_fisica.dropna(how='all', axis=0)

    # Analisando dados
    lista = input_palavras_chave()

    # Organizando palavras em ordem decrescente
    list_sorted = quantidade_respostas(lista, op_fisica)

    # Plotando gráfico
    plot_graph(list_sorted)

    input('\nPressione <Enter> para encerrar!')


main()

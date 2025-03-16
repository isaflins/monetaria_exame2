import pandas as pd
import matplotlib.pyplot as plt
import os

# Caminho completo do arquivo Excel
caminho_excel = r"C:\Users\maria\Documents\1_Graduação\6_Economia Monetaria\Trabalho\Bovespa_ipeadata.xlsx"

# Carregar os dados da Sheet1 do arquivo Excel
df = pd.read_excel(caminho_excel, sheet_name='Sheet1')

# Exibir os nomes das colunas para conferência
print("Colunas do DataFrame:", df.columns.tolist())

# Diretório para salvar os gráficos
diretorio = r"C:\Users\maria\Documents\1_Graduação\6_Economia Monetaria\Trabalho"
os.makedirs(diretorio, exist_ok=True)

# Suposições:
# - Coluna B será o eixo X (df.iloc[:,1])
# - Coluna C será plotada no eixo Y esquerdo (df.iloc[:,2])
# - Coluna D será plotada no eixo Y direito (df.iloc[:,3])

# Criar a figura e o primeiro eixo (para Coluna C)
fig, ax1 = plt.subplots(figsize=(10,6))
color1 = 'tab:red'  # Linha da Coluna C em vermelho
ax1.set_xlabel(df.columns[1])  # Eixo X: Coluna B
ax1.set_ylabel(df.columns[2])  # Eixo Y esquerdo: Coluna C (rótulos neutros)
line1, = ax1.plot(df.iloc[:,1], df.iloc[:,2], linestyle='-', color=color1, label=df.columns[2])

# Criar o segundo eixo (para Coluna D) compartilhando o mesmo eixo X
ax2 = ax1.twinx()
color2 = 'tab:blue'  # Linha da Coluna D em azul
ax2.set_ylabel(df.columns[3])  # Eixo Y direito: Coluna D (rótulos neutros)
line2, = ax2.plot(df.iloc[:,1], df.iloc[:,3], linestyle='-', color=color2, label=df.columns[3])

# Combinar as legendas de ambos os eixos
lines = [line1, line2]
labels = [line.get_label() for line in lines]
ax1.legend(lines, labels, loc='best')

# Título e grid
plt.title("Compração Selic x Bovespa")
ax1.grid(True)

# Ajustar layout para não cortar os rótulos
fig.tight_layout()

# Salvar o gráfico como PNG
nome_arquivo = os.path.join(diretorio, "Comparacao_Composto_Selic_Bovespa.png")
plt.savefig(nome_arquivo, format='png')
plt.close()

print("Gráfico composto salvo com sucesso!")

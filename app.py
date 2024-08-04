import streamlit as st
import pandas as pd
from datetime import datetime
import io
import seaborn as sns
import matplotlib.pyplot as plt

# Configurar o layout da página para ser expansível
st.set_page_config(layout="wide")

# Função para aplicar as transformações
def process_data(df):
    df.columns = [x.strip() for x in df.columns]
    df = df[['Data/Hora Empenho','Código Requisição', 'Id Volante','Volante','Código do Produto','Descrição do Produto','Qtde Atendida','Qtde Empenhada','Qtde Requisitada','Descrição Segmento Destino','Estoque Físico']]
    d1 = pd.to_datetime(datetime.today())
    df['dia'] = [d1 - x for x in pd.to_datetime(df['Data/Hora Empenho'])]
    df['dia'] = df['dia'].astype(str).str.extract('(.*?) days')
    df['Estoque Físico'] = [x.replace(' - ',' ')[4:] for x in df['Estoque Físico']]
    df['Estoque Físico'] = df['Estoque Físico'].str.strip()
    
    def repl(x):
        mapping = {
            'PROPRIO GERAL': 'PROPRIO',
            'O&M PROPRIO GERAL TIM FATURA B': 'GERAL TIM',
            'SPEEDY/FTTX CLIENTE': 'CASA CLIENTE',
            'MANUTENCAO RPO CLIENTE': 'MANUTENCAO',
            'EXEC SEGREGADO IMPLANTACAO RPO CLIENTE': 'EXEC SEGREGADO',
            'BOL IMPLANTACAO RPO CLIENTE': 'IMPLANTACAO',
            'DADOS CLIENTE': 'DADOS',
            'FERRAMENTAL': 'FERRAMENTAL',
            'UNIFORME': 'UNIFORME',
            'EPI-EPC': 'EPI',
            'BRINDES': 'BRINDES',
            'RPO MANUTENCAO FIBRA OPTICA BACKBONE': 'BACKBONE',
            'CLASSE D': 'CLASSE D',
            'MANUTENCAO RPO MATERIAL REUTILIZACAO': 'MANUTENCAO REUT',
            'EQUIPAMENTOS': 'EQUIPAMENTOS',
            'PROPRIO GERAL TIM': 'PROPRIO TIM',
            'KIT FERRAMENTAL CONTRATACOES': 'KIT CONTRATACOES',
            'MATERIAL DE ESCRITORIO': 'MATERIAL DE ESCRITORIO',
            'CELULARES': 'CELULARES',
            'MANUTENCAO RPO-MATERIAL REUTILIZACAO': 'MANUTENCAO-REUTILIZACAO',
            'MATERIAIS ENTREGUES EM OBRA': 'MATERIAIS ENTREGUES EM OBRA',
            'EQUIPAMENTOS TI': 'EQUIPAMENTOS TI',
            'LVUT IMPLANTACAO RPO CLIENTE': 'IMPLANTACAO',
            'PLANTA INTERNA FIXA': 'PLANTA INTERNA'
        }
        return mapping.get(x, x)
    
    df['Estoque'] = df['Estoque Físico'].apply(repl)
    df['Status'] = ["A Separar" if x <= 0 else "Separado" for x in df["Qtde Atendida"]]
    return df

# Função para criar gráficos com seaborn e matplotlib
def create_seaborn_charts(df):
    # Gráfico 1: Número de Valores Únicos (Id Volante) a Separar por Dia
    # Filtrar as requisições a separar
    requisicoes_a_separar = df[df['Status'] == 'A Separar']

    # Contar o número de valores únicos na coluna 'Id Volante' por dia
    valores_unicos_por_dia = requisicoes_a_separar.groupby(requisicoes_a_separar['Data/Hora Empenho'].dt.date)['Id Volante'].nunique()

    # Configurar o estilo do Seaborn
    sns.set_theme(style='whitegrid')

    # Criar subplots lado a lado
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))

    # Plotar o gráfico de barras do número de valores únicos por dia
    sns.barplot(x=valores_unicos_por_dia.index, y=valores_unicos_por_dia.values, color='skyblue', ax=ax1)
    ax1.set_title('Número de Valores Únicos (Id Volante) a Separar por Dia')
    ax1.set_ylabel('Número de Valores Únicos')
    ax1.tick_params(axis='x', rotation=45)

    # Adicionar valores nas barras
    for p in ax1.patches:
        ax1.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width() / 2., p.get_height()), ha='center', va='center', xytext=(0, 10), textcoords='offset points')

    # Plotar o gráfico de pizza da porcentagem
    ax2.pie(valores_unicos_por_dia, labels=valores_unicos_por_dia.index, autopct='%1.1f%%', colors=sns.color_palette('pastel'), startangle=90)
    ax2.set_title('Porcentagem de Valores Únicos (Id Volante) a Separar por Dia')

    plt.tight_layout()
    return fig

# Gráfico 2: Número de Requisições a Separar por Estoque
def create_bar_chart(df):
    requisicoes_a_separar = df[df['Status'] == 'A Separar']

    # Criar a tabela pivot contando valores únicos de 'Id Volante'
    pivot_table = pd.pivot_table(requisicoes_a_separar, values='Id Volante', index='Estoque', aggfunc=pd.Series.nunique).fillna(0)

    # Configurar o estilo do Seaborn
    sns.set_theme(style='whitegrid', palette='muted', font_scale=1.2)

    # Plotar o gráfico de barras empilhadas
    fig, ax = plt.subplots(figsize=(16, 8))
    ax = pivot_table.plot(kind='bar', stacked=True, ax=ax, colormap='viridis', edgecolor='white')

    # Adicionar rótulos e título
    ax.set_ylabel('Número de Requisições a Separar', labelpad=15)
    ax.set_xlabel('Estoque', labelpad=15)
    ax.set_title('Número de Requisições a Separar por Estoque', pad=20)

    # Adicionar legenda
    ax.legend(title='Status', bbox_to_anchor=(1.05, 1), loc='upper left')

    # Adicionar valores nas barras
    for p in ax.patches:
        ax.annotate(f'{p.get_height():.0f}', (p.get_x() + p.get_width() / 2., p.get_height()),
                    ha='center', va='center', fontsize=12, color='black', xytext=(0, 5),
                    textcoords='offset points', weight='bold')

    # Remover bordas desnecessárias
    sns.despine()
    plt.grid(False)
    plt.xticks(rotation=60)
    plt.tight_layout()
    return fig

# Gráfico 3: Distribuição entre Separado e A Separar por Estoque
def create_stacked_bar_chart(df):
    # Filtrar dados para 'Separado' e 'A Separar'
    separado = df[df['Status'] == 'Separado']
    a_separar = df[df['Status'] == 'A Separar']

    separado = separado.drop_duplicates(subset=['Código Requisição','Estoque'], keep='first')
    a_separar = a_separar.drop_duplicates(subset=['Código Requisição','Estoque'], keep='first')

    # Criar tabelas pivot para cada status contando valores únicos
    pivot_separado = pd.pivot_table(separado, values='Código Requisição', index='Estoque', aggfunc='nunique').fillna(0)
    pivot_a_separar = pd.pivot_table(a_separar, values='Código Requisição', index='Estoque', aggfunc='nunique').fillna(0)

    # Unir os DataFrames para garantir que todos os estoques estejam presentes em ambos
    merged_pivot = pd.merge(pivot_separado, pivot_a_separar, how='outer', left_index=True, right_index=True, suffixes=('_Separado', '_A_Separar')).fillna(0)

    # Configurar o estilo do Seaborn
    sns.set_theme(style='whitegrid', palette='muted', font_scale=1.2)

    # Plotar o gráfico de barras empilhadas
    fig, ax = plt.subplots(figsize=(16, 6))

    # Barra 'Separado'
    ax.bar(merged_pivot.index, merged_pivot['Código Requisição_Separado'], label='Separado', color='green', alpha=0.7)

    # Barra 'A Separar' empilhada sobre 'Separado'
    ax.bar(merged_pivot.index, merged_pivot['Código Requisição_A_Separar'], bottom=merged_pivot['Código Requisição_Separado'], label='A Separar', color='orange', alpha=0.7)

    # Adicionar rótulos e título
    ax.set_ylabel('Número de Requisições Únicas', labelpad=15)
    ax.set_xlabel('Estoque', labelpad=15)
    ax.set_title('Distribuição entre Separado e A Separar por Estoque', pad=20)

    # Adicionar legenda
    ax.legend(title='Status', bbox_to_anchor=(1.05, 1), loc='upper left')

    # Adicionar valores nas barras
    for p, qtd_separado, qtd_a_separar in zip(ax.patches, merged_pivot['Código Requisição_A_Separar'], merged_pivot['Código Requisição_Separado']):
        height = p.get_height() + p.get_y()
        ax.text(p.get_x() + p.get_width() / 2., height, f'{int(qtd_separado):.0f}', ha='center', va='bottom', fontsize=10, color='black', weight='bold')
        ax.text(p.get_x() + p.get_width() / 2., height, f'{int(qtd_a_separar):.0f}', ha='center', va='top', fontsize=10, color='black', weight='bold')

    # Remover bordas desnecessárias
    sns.despine()
    plt.xticks(rotation=60)
    plt.grid(False)
    plt.tight_layout()
    return fig

# Interface Streamlit
st.title("Processamento de Dados de Atendimento")

uploaded_file = st.file_uploader("Carregar arquivo Excel", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("Dados brutos carregados:")
    st.write(df.head())

    if st.button("Processar dados"):
        df_processed = process_data(df)
        st.write("Dados processados:")
        st.write(df_processed.head())

        # Gerar o arquivo Excel para download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_processed.to_excel(writer, index=False, sheet_name='Processed Data')
        processed_file = output.getvalue()

        st.download_button(
            label="Baixar arquivo Excel",
            data=processed_file,
            file_name="dados_processados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Criar e exibir os gráficos do Seaborn e Matplotlib
        st.subheader("Gráficos de Análise por Id Volante")
        fig1 = create_seaborn_charts(df_processed)
        st.pyplot(fig1)

        st.subheader("Número de Requisições a Separar por Estoque")
        fig2 = create_bar_chart(df_processed)
        st.pyplot(fig2)

        st.subheader("Distribuição entre Separado e A Separar por Estoque")
        fig3 = create_stacked_bar_chart(df_processed)
        st.pyplot(fig3)


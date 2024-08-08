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
    df['dia'] = (d1 - pd.to_datetime(df['Data/Hora Empenho'])).dt.days
    df['Estoque Físico'] = df['Estoque Físico'].str.replace(' - ', ' ').str[4:].str.strip()
    
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
            'PLANTA INTERNA FIXA': 'PLANTA INTERNA',
            'PROJETO SANTANDER':'SANTANDER',
            'PROPRIO TIM VANDALISMO':'TIM VANDALISMO'
        }
        return mapping.get(x, x)
    
    df['Estoque'] = df['Estoque Físico'].apply(repl)

    def determine_status(df):
        requisicoes = df['Código Requisição'].unique()
        status_list = []

        for requisicao in requisicoes:
            subset = df[df['Código Requisição'] == requisicao]
            if (subset['Qtde Atendida'] > 0).any():
                status = "Separado"
            else:
                status = "A Separar"
            status_list.extend([status] * len(subset))

        return status_list

    # Aplique a função ao DataFrame
    df['Status'] = determine_status(df)

    return df

# Gráficos com seaborn e matplotlib
def create_seaborn_charts(df):
    requisicoes_a_separar = df[df['Status'] == 'A Separar']
    valores_unicos_por_dia = requisicoes_a_separar.groupby(requisicoes_a_separar['Data/Hora Empenho'].dt.date)['Id Volante'].nunique()
    
    # Configurando o estilo do Seaborn e a paleta de cores
    sns.set_theme(style='darkgrid', palette='pastel')

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
    
    sns.barplot(x=valores_unicos_por_dia.index, y=valores_unicos_por_dia.values, color='#3498db', ax=ax1)
    ax1.set_title('Número de Valores Únicos (Id Volante) a Separar por Dia', fontsize=14, weight='bold')
    ax1.set_ylabel('Número de Valores Únicos', fontsize=12)
    ax1.set_xlabel('Data', fontsize=12)
    ax1.tick_params(axis='x', rotation=45)

    for p in ax1.patches:
        ax1.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width() / 2., p.get_height()), 
                     ha='center', va='center', xytext=(0, 10), textcoords='offset points', fontsize=10, weight='bold')

    ax2.pie(valores_unicos_por_dia, labels=valores_unicos_por_dia.index, autopct='%1.1f%%', 
            colors=sns.color_palette('pastel'), startangle=90)
    ax2.set_title('Porcentagem de Valores Únicos a Separar por Dia', fontsize=14, weight='bold')

    plt.tight_layout()
    return fig

def create_bar_chart(df):
    requisicoes_a_separar = df[df['Status'] == 'A Separar']
    pivot_table = pd.pivot_table(requisicoes_a_separar, values='Id Volante', index='Estoque', aggfunc=pd.Series.nunique).fillna(0)

    sns.set_theme(style='darkgrid', palette='deep', font_scale=1.1)

    fig, ax = plt.subplots(figsize=(16, 8))
    ax = pivot_table.plot(kind='bar', stacked=True, ax=ax, colormap='viridis', edgecolor='white')

    ax.set_ylabel('Número de Requisições a Separar', labelpad=15, fontsize=12)
    ax.set_xlabel('Estoque', labelpad=15, fontsize=12)
    ax.set_title('Número de Requisições a Separar por Estoque', pad=20, fontsize=14, weight='bold')

    ax.legend(title='Status', bbox_to_anchor=(1.05, 1), loc='upper left')

    for p in ax.patches:
        ax.annotate(f'{p.get_height():.0f}', (p.get_x() + p.get_width() / 2., p.get_height()), 
                    ha='center', va='center', fontsize=10, color='white', xytext=(0, 5), 
                    textcoords='offset points', weight='bold')

    sns.despine()
    plt.grid(False)
    plt.xticks(rotation=60, fontsize=10)
    plt.tight_layout()
    return fig

def create_stacked_bar_chart(df):
    separado = df[df['Status'] == 'Separado']
    a_separar = df[df['Status'] == 'A Separar']

    separado = separado.drop_duplicates(subset=['Código Requisição','Estoque'], keep='first')
    a_separar = a_separar.drop_duplicates(subset=['Código Requisição','Estoque'], keep='first')

    pivot_separado = pd.pivot_table(separado, values='Código Requisição', index='Estoque', aggfunc='nunique').fillna(0)
    pivot_a_separar = pd.pivot_table(a_separar, values='Código Requisição', index='Estoque', aggfunc='nunique').fillna(0)

    merged_pivot = pd.merge(pivot_separado, pivot_a_separar, how='outer', left_index=True, right_index=True, suffixes=('_Separado', '_A_Separar')).fillna(0)
    
    sns.set_theme(style='darkgrid', palette='coolwarm', font_scale=1.1)

    fig, ax = plt.subplots(figsize=(16, 6))
    ax.bar(merged_pivot.index, merged_pivot['Código Requisição_Separado'], label='Separado', color='#2ecc71', alpha=0.8)
    ax.bar(merged_pivot.index, merged_pivot['Código Requisição_A_Separar'], bottom=merged_pivot['Código Requisição_Separado'], label='A Separar', color='#e74c3c', alpha=0.8)

    ax.set_ylabel('Número de Requisições Únicas', labelpad=15, fontsize=12)
    ax.set_xlabel('Estoque', labelpad=15, fontsize=12)
    ax.set_title('Distribuição entre Separado e A Separar por Estoque', pad=20, fontsize=14, weight='bold')

    ax.legend(title='Status', bbox_to_anchor=(1.05, 1), loc='upper left')

    for p, qtd_separado, qtd_a_separar in zip(ax.patches, merged_pivot['Código Requisição_A_Separar'], merged_pivot['Código Requisição_Separado']):
        height = p.get_height() + p.get_y()
        ax.text(p.get_x() + p.get_width() / 2., height, f'{int(qtd_separado):.0f}', ha='center', va='bottom', fontsize=10, color='black', weight='bold')
        ax.text(p.get_x() + p.get_width() / 2., height, f'{int(qtd_a_separar):.0f}', ha='center', va='top', fontsize=10, color='black', weight='bold')

    sns.despine()
    plt.xticks(rotation=60, fontsize=10)
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

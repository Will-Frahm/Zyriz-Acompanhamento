import pandas as pd
import sqlite3
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# ========== CONFIGURAÇÕES ========== #
colunas_excel = [
    "Data", "Marketplace", "Cod. Produto", "Produto", "ID Anúncio", "Vendedor",
    "Preço", "Preço Sugerido", "% Dif. Preço Sugerido", "Diferença Preço Sugerido",
    "Link", "Cidade", "Estado", "Grupo Nome", "Grupo CNPJ", "Razão Social CNPJ",
    "Catálogo Mercado Livre"
]

planilha = "IRREGULARES"
tabela_dados = "dados_marketplace"
tabela_irregularidades = "irregularidades"

# ========== FUNÇÕES ========== #
def filtrar_produtos(df):
    """Remove produtos restritos/não monitorados."""
    filtro = ~df["Produto"].str.startswith(("(RESTRITOS)", "(NÃO MONITORADOS)", "(NÂO MONITORADOS)","(NÃO MONITORADO)"))
    return df[filtro].copy()

def atualizar_irregularidades(conexao, df):
    cursor = conexao.cursor()    
    try:
        df['diff_convertida'] = (
            df['diferenca_preco_sugerido']
            .astype(str)
            .str.replace('[^\d,-]', '', regex=True)
            .str.replace(',', '.')
            .replace('', '0')
            .astype(float)
        )
        
        df['data'] = pd.to_datetime(df['data'], dayfirst=True).dt.strftime('%d/%m/%Y')

        df_irreg = df[df['diff_convertida'] < 0]
        print(f"\n🚨 Anúncios irregulares detectados: {len(df_irreg)}")

        for _, row in df_irreg.iterrows():
            id_anuncio = row['id_anuncio']
            data_atual = row['data']            

            cursor.execute("""
            SELECT data_irregular, dias_irregular 
            FROM irregularidades 
            WHERE id_anuncio = ? AND status = 'ATIVO'
            """, (id_anuncio,))
            resultado = cursor.fetchone()

            if resultado:  
                data_inicial = resultado[0]
                dias_atual = resultado[1]
                                
                if data_atual != data_inicial:
                    cursor.execute("""
                    UPDATE irregularidades 
                    SET dias_irregular = ?,
                        data_irregular = ?  
                    WHERE id_anuncio = ? AND status = 'ATIVO'
                    """, (dias_atual + 1, data_inicial, id_anuncio))
                    print(f"🔄 Atualizado: {id_anuncio} (+1 dia) | Data Inicial: {data_inicial}")
            else:
                cursor.execute("""
                INSERT INTO irregularidades 
                (id_anuncio, data_irregular, status, dias_irregular)
                VALUES (?, ?, 'ATIVO', 1)
                """, (id_anuncio, data_atual))  
                print(f"✅ Novo irregular: {id_anuncio} | Data Inicial: {data_atual}")
        
        anuncios_irreg_atual = set(df_irreg['id_anuncio'])
        if anuncios_irreg_atual:
            cursor.execute(f"""
            DELETE FROM irregularidades
            WHERE status = 'ATIVO'
            AND id_anuncio NOT IN ({','.join(['?']*len(anuncios_irreg_atual))})
            """, list(anuncios_irreg_atual))
            print(f"\n♻️ Anúncios regularizados: {cursor.rowcount}")

        conexao.commit()

    except Exception as e:
        print(f"❌ ERRO: {str(e)}")
        conexao.rollback()
        raise
    finally:
        cursor.close()

Tk().withdraw()
arquivo_excel = askopenfilename(
    title="Selecione o arquivo Excel",
    filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")]
)

if not arquivo_excel:
    print("Nenhum arquivo selecionado. Saindo.")
    exit()

df = pd.read_excel(arquivo_excel, sheet_name=planilha, dtype=str)
df = filtrar_produtos(df)
df = df[colunas_excel]

conexao = sqlite3.connect(r"\\wts\instala\Zyriz_relatorio_base\bd_zyriz.db")

mapeamento_colunas = {
    "Data": "data",
    "Marketplace": "marketplace",
    "Cod. Produto": "cod_produto",
    "Produto": "produto",
    "ID Anúncio": "id_anuncio",
    "Vendedor": "vendedor",
    "Preço": "preco",
    "Preço Sugerido": "preco_sugerido",
    "% Dif. Preço Sugerido": "percentual_diferenca_preco_sugerido",
    "Diferença Preço Sugerido": "diferenca_preco_sugerido",
    "Link": "link",
    "Cidade": "cidade",
    "Estado": "estado",
    "Grupo Nome": "grupo_nome",
    "Grupo CNPJ": "grupo_cnpj",
    "Razão Social CNPJ": "razao_social_cnpj",
    "Catálogo Mercado Livre": "catalogo_mercado_livre"
}

df.rename(columns=mapeamento_colunas, inplace=True)
df['data'] = pd.to_datetime(df['data'], dayfirst=True).dt.strftime('%d/%m/%Y')

cursor = conexao.cursor()
cursor.execute(f"""
CREATE TABLE IF NOT EXISTS {tabela_irregularidades} (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    id_anuncio TEXT,
    data_irregular TEXT,
    status TEXT CHECK(status IN ('ATIVO', 'REGULARIZADO')),
    dias_irregular INTEGER DEFAULT 1,
    FOREIGN KEY (id_anuncio) REFERENCES {tabela_dados}(id_anuncio)
)
""")

conexao.commit()
cursor.close()
df.to_sql(tabela_dados, conexao, if_exists="append", index=False)
atualizar_irregularidades(conexao, df)
conexao.close()
print("\n✔️ Processo concluído!")
from flask import Flask, render_template, request, jsonify, send_file
from sqlalchemy import create_engine, text
import pandas as pd
from datetime import datetime
import os
import tempfile

app = Flask(__name__)

# =====================================================
# CONFIGURAÇÃO BANCO DE DADOS
# =====================================================
usuario = "postgres"
senha = "postgres"
ip = "192.168.0.250"
porta = "5432"
banco = "DB-CONTAS A RECEBER"

engine = create_engine(
    f"postgresql+psycopg2://{usuario}:{senha}@{ip}:{porta}/{banco}",
    pool_pre_ping=True,
    pool_size=5,
    max_overflow=10
)

# =====================================================
# HOME
# =====================================================
@app.route("/")
def index():
    return render_template("index.html")

# =====================================================
# CONSULTA INICIAL
# =====================================================
@app.route("/consultar", methods=["POST"])
def consultar():
    dados = request.get_json()
    venc_ini = dados.get("venc_ini")
    venc_fim = dados.get("venc_fim")

    if not venc_ini or not venc_fim:
        return jsonify({"erro": "Informe o período de vencimento"})

    print("\n==============================")
    print("Consultando Inadimplência e Faturamento...")
    print("==============================")

    # 1. SQL Inadimplência
    sql_inad = text("""
        SELECT *
        FROM marcelo."Relatório de inadiplencia detalhado - (0205)"
        WHERE CAST("Data de vencimento" AS DATE) BETWEEN :inicio AND :fim
        ORDER BY "Data de vencimento"
    """)

    # 2. SQL Faturamento (Usando Data de vencimento no filtro)
    sql_fat = text("""
        SELECT 
            CAST("Data de vencimento" AS DATE) AS data_venc,
            COALESCE("DE PARA CR", 'SEM REDE') AS rede,
            COALESCE(CAST("Total do documento" AS NUMERIC), 0) AS valor_faturamento
        FROM marcelo."FATURAMENTO DASH"
        WHERE CAST("Data de vencimento" AS DATE) BETWEEN :inicio AND :fim
    """)

    params = {"inicio": venc_ini, "fim": venc_fim}

    df_inad = pd.read_sql(sql_inad, engine, params=params)
    df_fat = pd.read_sql(sql_fat, engine, params=params)

    # Tratamento da Inadimplência
    df_inad = df_inad.fillna("")
    for coluna in df_inad.columns:
        df_inad[coluna] = df_inad[coluna].astype(str)

    # Tratamento do Faturamento
    df_fat['valor_faturamento'] = pd.to_numeric(df_fat['valor_faturamento'], errors='coerce').fillna(0)
    df_fat['mes_ano'] = pd.to_datetime(df_fat['data_venc']).dt.strftime('%m/%Y')

    # Totalizador Geral de Faturamento
    faturamento_total = float(df_fat['valor_faturamento'].sum())
    
    # Faturamento por Mês e Faturamento por Rede (DE PARA CR)
    fat_por_mes = df_fat.groupby('mes_ano')['valor_faturamento'].sum().to_dict()
    fat_por_rede = df_fat.groupby('rede')['valor_faturamento'].sum().to_dict()

    retorno = {
        "total": len(df_inad),
        "dados": df_inad.to_dict(orient="records"),
        "faturamento": {
            "total": faturamento_total,
            "por_mes": fat_por_mes,
            "por_rede": fat_por_rede
        }
    }

    return jsonify(retorno)

# =====================================================
# EXPORTAR EXCEL
# =====================================================
@app.route("/exportar", methods=["POST"])
def exportar():
    dados = request.get_json()
    df = pd.DataFrame(dados)

    nome = f"INADIMPLENCIA_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    caminho = os.path.join(tempfile.gettempdir(), nome)

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Inadimplência", index=False)

    return send_file(caminho, as_attachment=True, download_name=nome)

# =====================================================
# START
# =====================================================
if __name__ == "__main__":
    app.run(debug=True, threaded=True)
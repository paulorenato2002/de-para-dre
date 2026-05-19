import streamlit as st
import pandas as pd
from supabase import create_client, Client
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# --- 1. SETUP E CONEXÃO ---
st.set_page_config(page_title="Conciliador Contábil", page_icon="📊", layout="wide")


@st.cache_resource
def init_connection() -> Client:
    if "SUPABASE_URL" not in st.secrets:
        st.error("Faltou configurar `SUPABASE_URL` nos secrets do Streamlit Cloud.")
        st.stop()

    if "SUPABASE_SERVICE_ROLE_KEY" in st.secrets:
        key = st.secrets["SUPABASE_SERVICE_ROLE_KEY"]
    elif "SUPABASE_KEY" in st.secrets:
        key = st.secrets["SUPABASE_KEY"]
    else:
        st.error("Faltou configurar `SUPABASE_KEY` ou `SUPABASE_SERVICE_ROLE_KEY` nos secrets do Streamlit Cloud.")
        st.stop()

    url = st.secrets["SUPABASE_URL"]
    return create_client(url, key)


supabase = init_connection()


# Lista padrão de empresas, usada como fallback quando a tabela no banco ainda não existir.
EMPRESAS_PADRAO = {
    "401": "401 - SST",
    "370": "370 - STI",
    "536": "536 - SCD",
    "570": "570 - SSD",
    "556": "556 - SAC",
}


@st.cache_data(ttl=300)
def carregar_empresas():
    try:
        response = supabase.table("empresas").select("id, nome, ativo").eq("ativo", True).execute()
        dados = response.data or []
        if dados:
            return {str(item["id"]): str(item["nome"]) for item in dados}
    except Exception:
        pass
    return EMPRESAS_PADRAO


EMPRESAS = carregar_empresas()


# --- 2. FUNÇÕES ÚTEIS ---
def tratar_valor(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return abs(float(val))
    val = str(val).replace("R$", "").replace(".", "").replace(",", ".").strip()
    try:
        return abs(float(val))
    except Exception:
        return 0.0


def normalizar_texto(val):
    if pd.isna(val):
        return ""
    return str(val).strip()


def normalizar_chave(val):
    texto = normalizar_texto(val)
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto


def limpar_colunas(df):
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df


def obter_coluna(df, candidatos):
    for coluna in candidatos:
        if coluna in df.columns:
            return coluna
    return None


def colunas_existentes(df, colunas):
    return [coluna for coluna in colunas if coluna in df.columns]


def formatar_moeda(valor):
    return f"R$ {abs(float(valor)):,.2f}"


def classificar_diferenca(diff):
    if abs(diff) <= 0.01:
        return "🟢 Bateu", "Os valores fecham certinho entre contábil e cliente."
    if diff > 0:
        return "🔴 Divergente", f"Faltam {formatar_moeda(diff)} no cliente para bater com o contábil."
    return "🔴 Divergente", f"Sobram {formatar_moeda(diff)} no cliente em relação ao contábil."


def status_linha(valor, total_ok):
    if pd.isna(valor) or abs(float(valor)) <= 0.0:
        return "Não encontrado"
    return "Encontrado" if total_ok else "Revisar"


def preparar_detalhe_contabil(df_cont_detalhe, col_conta_contabil, col_grupo_contabil, col_historico_contabil, total_ok):
    if df_cont_detalhe.empty:
        return pd.DataFrame(columns=["GRUPO DRE", "HISTORICO", "CONTA DEB", "VALOR", "STATUS"])

    df = df_cont_detalhe.copy()
    df["GRUPO DRE"] = df["_Grupo_Contabil"].apply(normalizar_texto)
    if col_historico_contabil:
        df["HISTORICO"] = df[col_historico_contabil].apply(normalizar_texto)
    else:
        df["HISTORICO"] = ""
    df["CONTA DEB"] = df[col_conta_contabil].apply(normalizar_chave)
    df["VALOR"] = df["Valor_Tratado"].apply(float)
    df["STATUS"] = "Encontrado" if total_ok else "Revisar"

    df = (
        df.groupby(["GRUPO DRE", "HISTORICO", "CONTA DEB"], dropna=False, as_index=False)
        .agg({"VALOR": "sum", "STATUS": "first"})
        .sort_values(["VALOR", "GRUPO DRE", "HISTORICO"], ascending=[False, True, True])
    )
    return df


def preparar_detalhe_cliente(df_cli_detalhe, col_fornecedor_cliente, total_ok):
    if df_cli_detalhe.empty:
        return pd.DataFrame(columns=["GRUPO DRE", "FORNECEDOR", "VALOR", "STATUS"])

    df = df_cli_detalhe.copy()
    df["GRUPO DRE"] = df["_Nivel_3"].apply(normalizar_texto)
    if col_fornecedor_cliente and col_fornecedor_cliente in df.columns:
        df["FORNECEDOR"] = df[col_fornecedor_cliente].apply(normalizar_texto)
    else:
        df["FORNECEDOR"] = df["_Nivel_4"].apply(normalizar_texto)
    df["VALOR"] = df["Valor_Tratado"].apply(float)
    df["STATUS"] = "Encontrado" if total_ok else "Revisar"

    df = (
        df.groupby(["GRUPO DRE", "FORNECEDOR"], dropna=False, as_index=False)
        .agg({"VALOR": "sum", "STATUS": "first"})
        .sort_values(["VALOR", "GRUPO DRE", "FORNECEDOR"], ascending=[False, True, True])
    )
    return df


def montar_quadro_lado_a_lado(df_cont, df_cli):
    max_len = max(len(df_cont), len(df_cli), 1)
    df_cont = df_cont.reindex(range(max_len)).reset_index(drop=True)
    df_cli = df_cli.reindex(range(max_len)).reset_index(drop=True)

    cont = pd.DataFrame(
        {
            ("CONTABIL", "GRUPO DRE"): df_cont["GRUPO DRE"] if "GRUPO DRE" in df_cont.columns else [""] * max_len,
            ("CONTABIL", "HISTORICO"): df_cont["HISTORICO"] if "HISTORICO" in df_cont.columns else [""] * max_len,
            ("CONTABIL", "CONTA DEB"): df_cont["CONTA DEB"] if "CONTA DEB" in df_cont.columns else [""] * max_len,
            ("CONTABIL", "VALOR"): df_cont["VALOR"] if "VALOR" in df_cont.columns else [0.0] * max_len,
            ("CONTABIL", "STATUS"): df_cont["STATUS"] if "STATUS" in df_cont.columns else [""] * max_len,
        }
    )

    cli = pd.DataFrame(
        {
            ("CLIENTE", "GRUPO DRE"): df_cli["GRUPO DRE"] if "GRUPO DRE" in df_cli.columns else [""] * max_len,
            ("CLIENTE", "FORNECEDOR"): df_cli["FORNECEDOR"] if "FORNECEDOR" in df_cli.columns else [""] * max_len,
            ("CLIENTE", "VALOR"): df_cli["VALOR"] if "VALOR" in df_cli.columns else [0.0] * max_len,
            ("CLIENTE", "STATUS"): df_cli["STATUS"] if "STATUS" in df_cli.columns else [""] * max_len,
        }
    )

    quadro = pd.concat([cont, cli], axis=1)
    quadro.columns = pd.MultiIndex.from_tuples(quadro.columns)
    quadro_export = pd.DataFrame(
        {
            "CONTABIL - GRUPO DRE": df_cont["GRUPO DRE"] if "GRUPO DRE" in df_cont.columns else [""] * max_len,
            "CONTABIL - HISTORICO": df_cont["HISTORICO"] if "HISTORICO" in df_cont.columns else [""] * max_len,
            "CONTABIL - CONTA DEB": df_cont["CONTA DEB"] if "CONTA DEB" in df_cont.columns else [""] * max_len,
            "CONTABIL - VALOR": df_cont["VALOR"] if "VALOR" in df_cont.columns else [0.0] * max_len,
            "CONTABIL - STATUS": df_cont["STATUS"] if "STATUS" in df_cont.columns else [""] * max_len,
            "CLIENTE - GRUPO DRE": df_cli["GRUPO DRE"] if "GRUPO DRE" in df_cli.columns else [""] * max_len,
            "CLIENTE - FORNECEDOR": df_cli["FORNECEDOR"] if "FORNECEDOR" in df_cli.columns else [""] * max_len,
            "CLIENTE - VALOR": df_cli["VALOR"] if "VALOR" in df_cli.columns else [0.0] * max_len,
            "CLIENTE - STATUS": df_cli["STATUS"] if "STATUS" in df_cli.columns else [""] * max_len,
        }
    )
    return quadro, quadro_export


def carregar_parametrizacao_empresa(empresa_id):
    response = (
        supabase.table("parametrizacao_contas")
        .select("id, conta_contabil, fornecedor_cliente")
        .eq("empresa_id", str(empresa_id))
        .execute()
    )
    df_banco = pd.DataFrame(response.data)
    if df_banco.empty:
        df_banco = pd.DataFrame(columns=["id", "conta_contabil", "fornecedor_cliente"])
    return df_banco


def preparar_pendencias_parametrizacao(df_cliente, df_parametros, col_nivel3_cliente, col_conta_cliente, col_fornecedor_cliente=None):
    if df_cliente is None or df_cliente.empty:
        return pd.DataFrame(columns=["GRUPO DRE", "CÓD. PLANO 04", "FORNECEDOR", "VALOR", "QTD LANCAMENTOS", "STATUS"])

    df = df_cliente.copy()
    if "_Nivel_3" not in df.columns and col_nivel3_cliente in df.columns:
        df["_Nivel_3"] = df[col_nivel3_cliente].apply(normalizar_texto)
    if "_Nivel_4" not in df.columns and col_conta_cliente in df.columns:
        df["_Nivel_4"] = df[col_conta_cliente].apply(normalizar_texto)
    if "_Fornecedor" not in df.columns and col_fornecedor_cliente and col_fornecedor_cliente in df.columns:
        df["_Fornecedor"] = df[col_fornecedor_cliente].apply(normalizar_texto)

    parametros_existentes = set()
    if df_parametros is not None and not df_parametros.empty and "fornecedor_cliente" in df_parametros.columns:
        parametros_existentes = set(df_parametros["fornecedor_cliente"].dropna().astype(str).map(normalizar_texto))

    df["__parametrizado"] = df["_Nivel_4"].isin(parametros_existentes)
    df_pend = df[~df["__parametrizado"]].copy()

    if df_pend.empty:
        return pd.DataFrame(columns=["GRUPO DRE", "CÓD. PLANO 04", "FORNECEDOR", "VALOR", "QTD LANCAMENTOS", "STATUS"])

    df_pend["FORNECEDOR"] = df_pend["_Fornecedor"] if "_Fornecedor" in df_pend.columns else df_pend["_Nivel_4"]
    df_pend["VALOR"] = df_pend["Valor_Tratado"].apply(float)
    df_pend["STATUS"] = "Sem parametrização"

    df_pend = (
        df_pend.groupby(["_Nivel_3", "_Nivel_4", "FORNECEDOR"], dropna=False, as_index=False)
        .agg(QTD_LANCAMENTOS=("FORNECEDOR", "size"), VALOR=("VALOR", "sum"))
        .rename(columns={"_Nivel_3": "GRUPO DRE", "_Nivel_4": "CÓD. PLANO 04"})
        .sort_values(["VALOR", "QTD_LANCAMENTOS"], ascending=False)
    )

    df_pend["STATUS"] = "Sem parametrização"
    return df_pend


def exportar_excel_relatorio(nome_cliente, df_resumo, quadro_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumo_export = df_resumo.copy()
        resumo_export.to_excel(writer, sheet_name="Resumo", index=False)

        quadro_export.to_excel(writer, sheet_name="Detalhe", index=False, startrow=2)

        wb = writer.book

        # Estilos leves para leitura.
        azul = PatternFill("solid", fgColor="DCE6F1")
        verde = PatternFill("solid", fgColor="E2F0D9")
        cinza = PatternFill("solid", fgColor="F3F6F9")
        titulo = PatternFill("solid", fgColor="2F5597")
        branco = Font(color="FFFFFF", bold=True)
        bold = Font(bold=True)
        thin = Side(style="thin", color="D9E2F3")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        ws = wb["Resumo"]
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = bold
            cell.fill = cinza
            cell.border = border

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = border

        for col in ws.columns:
            largura = 0
            letra = get_column_letter(col[0].column)
            for cell in col:
                try:
                    largura = max(largura, len(str(cell.value)) if cell.value is not None else 0)
                except Exception:
                    pass
            ws.column_dimensions[letra].width = min(max(largura + 2, 12), 28)

        wd = wb["Detalhe"]
        wd.freeze_panes = "A4"
        wd.merge_cells("A1:I1")
        wd["A1"] = f"NOME DO CLIENTE: {nome_cliente}"
        wd["A1"].fill = titulo
        wd["A1"].font = branco
        wd["A1"].alignment = Alignment(horizontal="center", vertical="center")

        wd.merge_cells("A2:E2")
        wd.merge_cells("F2:I2")
        wd["A2"] = "CONTABIL"
        wd["F2"] = "CLIENTE"
        wd["A2"].fill = azul
        wd["F2"].fill = verde
        wd["A2"].font = bold
        wd["F2"].font = bold
        wd["A2"].alignment = Alignment(horizontal="center")
        wd["F2"].alignment = Alignment(horizontal="center")

        for cell in wd[3]:
            cell.font = bold
            cell.fill = cinza
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in wd.iter_rows(min_row=4):
            for cell in row:
                cell.border = border
                if cell.column in [2, 3, 4, 5]:
                    cell.fill = azul
                elif cell.column in [6, 7, 8, 9]:
                    cell.fill = verde

        for row in wd.iter_rows(min_row=4):
            for cell in row:
                if cell.column in [5, 9]:
                    valor = str(cell.value).lower() if cell.value is not None else ""
                    if "bateu" in valor or "encontrado" in valor:
                        cell.fill = PatternFill("solid", fgColor="E2F0D9")
                    elif "diverg" in valor or "revis" in valor:
                        cell.fill = PatternFill("solid", fgColor="FCE4D6")
                    elif "não" in valor or "nao" in valor:
                        cell.fill = PatternFill("solid", fgColor="FFF2CC")

        for col in wd.columns:
            largura = 0
            letra = get_column_letter(col[0].column)
            for cell in col:
                try:
                    largura = max(largura, len(str(cell.value)) if cell.value is not None else 0)
                except Exception:
                    pass
            wd.column_dimensions[letra].width = min(max(largura + 2, 14), 35)

        # Formatação numérica
        for row in wd.iter_rows(min_row=4):
            for cell in row:
                if cell.column in [4, 8]:
                    cell.number_format = 'R$ #,##0.00'

    output.seek(0)
    return output.getvalue()


def assinatura_arquivos(*arquivos):
    return tuple(
        (
            getattr(arquivo, "name", None),
            getattr(arquivo, "size", None),
        )
        if arquivo is not None
        else (None, None)
        for arquivo in arquivos
    )


def renderizar_analise_conferencia(analise):
    df_resultados = analise["df_resultados"].copy()
    detalhes_por_conta = analise["detalhes_por_conta"]
    empresa_id = analise["empresa_id"]
    empresa_nome = analise["empresa_nome"]
    col_conta_contabil = analise["col_conta_contabil"]
    col_grupo_contabil = analise["col_grupo_contabil"]
    col_historico_contabil = analise["col_historico_contabil"]
    col_fornecedor_cliente = analise["col_fornecedor_cliente"]

    if df_resultados.empty:
        st.info("Nenhuma conta com valor diferente de zero foi encontrada para comparar.")
        return

    st.markdown("---")
    tot_batidos = len(df_resultados[df_resultados["STATUS"] == "🟢 Bateu"])
    tot_divergentes = len(df_resultados[df_resultados["STATUS"] == "🔴 Divergente"])
    tot_sem_map = len(df_resultados[df_resultados["STATUS"] == "⚠️ Sem parametrização"])

    c1, c2, c3 = st.columns(3)
    c1.metric("🟢 Bateu", tot_batidos)
    c2.metric("🔴 Divergências", tot_divergentes)
    c3.metric("⚠️ Sem parametrização", tot_sem_map)

    filtro_key = f"filtro_visao_{empresa_id}"
    if filtro_key not in st.session_state:
        st.session_state[filtro_key] = "Todos"

    filtro = st.radio(
        "Filtrar Visão:",
        ["Todos", "🔴 Divergente", "⚠️ Sem parametrização", "🟢 Bateu"],
        horizontal=True,
        key=filtro_key,
    )
    df_filtrado = df_resultados
    if filtro != "Todos":
        df_filtrado = df_filtrado[df_filtrado["STATUS"] == filtro]

    if df_filtrado.empty:
        st.info("Nenhuma conta corresponde ao filtro selecionado.")
        return

    colunas_visiveis = [
        "Conta Débito",
        "Grupo DRE (Contábil)",
        "Valor Contábil",
        "Valor Cliente",
        "Diferença",
        "STATUS",
        "MOTIVO",
    ]
    st.dataframe(
        df_filtrado[colunas_visiveis].style.format(
            {
                "Valor Contábil": "R$ {:,.2f}",
                "Valor Cliente": "R$ {:,.2f}",
                "Diferença": "R$ {:,.2f}",
            }
        ),
        use_container_width=True,
    )

    # --- DETALHAMENTO ANALÍTICO ---
    st.markdown("---")
    st.subheader("Quadro dos Lançamentos")
    st.caption("Contábil e cliente lado a lado, com o total da conta selecionada.")

    contas_visiveis = df_filtrado["Conta Débito"].tolist()
    conta_key = f"conta_detalhe_{empresa_id}"
    if conta_key not in st.session_state or st.session_state[conta_key] not in contas_visiveis:
        st.session_state[conta_key] = contas_visiveis[0]

    conta_detalhe = st.selectbox("Conta para detalhar", contas_visiveis, key=conta_key)

    detalhe_atual = detalhes_por_conta.get(
        conta_detalhe, {"contabil": pd.DataFrame(), "cliente": pd.DataFrame()}
    )
    df_cont_detalhe = detalhe_atual["contabil"]
    df_cli_detalhe = detalhe_atual["cliente"]
    resumo_linha = df_filtrado[df_filtrado["Conta Débito"] == conta_detalhe].iloc[0]
    total_ok = resumo_linha["STATUS"] == "🟢 Bateu"

    st.markdown(f"### NOME DO CLIENTE: {empresa_nome}")
    st.markdown("#### CONTABIL | CLIENTE")

    df_cont_rel = preparar_detalhe_contabil(
        df_cont_detalhe, col_conta_contabil, col_grupo_contabil, col_historico_contabil, total_ok
    )
    df_cli_rel = preparar_detalhe_cliente(df_cli_detalhe, col_fornecedor_cliente, total_ok)
    quadro_display, quadro_export = montar_quadro_lado_a_lado(df_cont_rel, df_cli_rel)

    st.dataframe(
        quadro_display.style.format(
            {
                ("CONTABIL", "VALOR"): "R$ {:,.2f}",
                ("CLIENTE", "VALOR"): "R$ {:,.2f}",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    resumo_export = df_resultados[
        [
            "Conta Débito",
            "Grupo DRE (Contábil)",
            "Valor Contábil",
            "Valor Cliente",
            "Diferença",
            "STATUS",
            "MOTIVO",
            "QTD CONTÁBIL",
            "QTD CLIENTE",
            "DIF. QTD",
        ]
    ].copy()

    arquivo_excel = exportar_excel_relatorio(empresa_nome, resumo_export, quadro_export)

    st.download_button(
        "📥 Exportar relatório em Excel",
        data=arquivo_excel,
        file_name=f"conciliacao_{empresa_id}_{conta_detalhe}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


# --- 3. MENU LATERAL ---
st.sidebar.title("Navegação")
empresa_selecionada = st.sidebar.selectbox(
    "🏢 Selecione a Empresa:",
    list(EMPRESAS.keys()),
    format_func=lambda x: EMPRESAS[x],
)
menu = st.sidebar.radio("Ir para", ["Conciliação", "Parametrização"])


# ==========================================
# TELA 1: CONCILIAÇÃO (DUPLA CONFERÊNCIA)
# ==========================================
if menu == "Conciliação":
    st.title(f"📊 Conciliação Simples - {EMPRESAS[empresa_selecionada]}")
    st.markdown("Conferência direta entre o contábil e o cliente, sem firula: valor, lançamento e diferença.")

    col1, col2 = st.columns(2)
    with col1:
        file_contabil = st.file_uploader("📂 Upload CONTÁBIL (.xlsx)", type=["xlsx", "xls"], key="contabil")
    with col2:
        file_cliente = st.file_uploader("📂 Upload CLIENTE / DRE (.xlsx)", type=["xlsx", "xls"], key="cliente")

    if file_contabil and file_cliente:
        if st.button("🚀 Iniciar Conferência", use_container_width=True, type="primary"):
            with st.spinner("Lendo arquivos e cruzando com o banco de dados..."):
                try:
                    # Lendo os arquivos
                    df_contabil = limpar_colunas(pd.read_excel(file_contabil, engine="openpyxl"))
                    df_cliente = limpar_colunas(pd.read_excel(file_cliente, engine="openpyxl"))

                    # Colunas esperadas
                    col_conta_contabil = obter_coluna(df_contabil, ["codic"])
                    col_grupo_contabil = obter_coluna(df_contabil, ["nomec"])
                    col_valor_contabil = obter_coluna(df_contabil, ["valdeb"])
                    col_data_contabil = obter_coluna(df_contabil, ["datalan"])
                    col_historico_contabil = obter_coluna(df_contabil, ["historico_excel", "historico"])

                    col_conta_cliente = obter_coluna(df_cliente, ["Cód. Plano de conta nº. 04"])
                    col_nivel3_cliente = obter_coluna(df_cliente, ["Plano de conta nº. 03"])
                    col_valor_cliente = obter_coluna(df_cliente, ["Débito"])
                    col_data_cliente = obter_coluna(df_cliente, ["Dt. Movimento"])
                    col_descricao_cliente = obter_coluna(df_cliente, ["Descrição"])
                    col_fornecedor_cliente = obter_coluna(df_cliente, ["Cliente/Fornecedor"])

                    faltantes = []
                    if not col_conta_contabil:
                        faltantes.append("codic")
                    if not col_grupo_contabil:
                        faltantes.append("nomec")
                    if not col_valor_contabil:
                        faltantes.append("valdeb")
                    if not col_conta_cliente:
                        faltantes.append("Cód. Plano de conta nº. 04")
                    if not col_nivel3_cliente:
                        faltantes.append("Plano de conta nº. 03")
                    if not col_valor_cliente:
                        faltantes.append("Débito")

                    if faltantes:
                        raise ValueError(
                            "As planilhas enviadas não têm as colunas esperadas: " + ", ".join(faltantes)
                        )

                    # Tratamento de valores e chaves
                    df_contabil["Valor_Tratado"] = df_contabil[col_valor_contabil].apply(tratar_valor)
                    df_contabil["_Conta_Contabil"] = df_contabil[col_conta_contabil].apply(normalizar_chave)
                    df_contabil["_Grupo_Contabil"] = df_contabil[col_grupo_contabil].apply(normalizar_texto)

                    df_cliente["Valor_Tratado"] = df_cliente[col_valor_cliente].apply(tratar_valor)
                    df_cliente["_Nivel_3"] = df_cliente[col_nivel3_cliente].apply(normalizar_texto)
                    df_cliente["_Nivel_4"] = df_cliente[col_conta_cliente].apply(normalizar_texto)
                    if col_fornecedor_cliente:
                        df_cliente["_Fornecedor"] = df_cliente[col_fornecedor_cliente].apply(normalizar_texto)

                    # Puxar Parametrização da empresa selecionada
                    response = (
                        supabase.table("parametrizacao_contas")
                        .select("*")
                        .eq("empresa_id", str(empresa_selecionada))
                        .execute()
                    )
                    df_parametros = pd.DataFrame(response.data)

                    if df_parametros.empty:
                        st.warning(
                            f"⚠️ Nenhuma regra encontrada no banco para a empresa {EMPRESAS[empresa_selecionada]}. Vá na aba de Parametrização!"
                        )
                    else:
                        df_parametros = limpar_colunas(df_parametros)
                        if "conta_contabil" not in df_parametros.columns or "fornecedor_cliente" not in df_parametros.columns:
                            raise ValueError("A tabela parametrizacao_contas precisa ter as colunas conta_contabil e fornecedor_cliente.")

                        df_parametros["conta_contabil"] = df_parametros["conta_contabil"].apply(normalizar_chave)
                        df_parametros["fornecedor_cliente"] = df_parametros["fornecedor_cliente"].apply(normalizar_texto)

                        # Agrupar Contábil por Conta Contábil
                        df_cont_grp = (
                            df_contabil.groupby("_Conta_Contabil")
                            .agg(
                                Grupo_Contabil=("_Grupo_Contabil", "first"),
                                Valor_Contabil=("Valor_Tratado", "sum"),
                                Qtde_Contabil=("_Conta_Contabil", "size"),
                            )
                            .reset_index()
                            .rename(columns={"_Conta_Contabil": "Conta Débito"})
                        )

                        resultados = []
                        detalhes_por_conta = {}

                        # Conferência linha a linha por Conta Contábil
                        for _, row in df_cont_grp.iterrows():
                            conta_contabil = str(row["Conta Débito"]).strip()
                            valor_contabil = round(float(row["Valor_Contabil"]), 2)
                            grupo_contabil = str(row["Grupo_Contabil"]).strip()
                            qtde_contabil = int(row["Qtde_Contabil"])

                            if valor_contabil == 0:
                                continue

                            regras = df_parametros[df_parametros["conta_contabil"] == conta_contabil]

                            df_cont_detalhe = df_contabil[df_contabil["_Conta_Contabil"] == conta_contabil].copy()

                            if regras.empty:
                                resultados.append(
                                    {
                                        "Conta Débito": conta_contabil,
                                        "Grupo DRE (Contábil)": grupo_contabil,
                                        "Valor Contábil": valor_contabil,
                                        "Valor Cliente": 0.0,
                                        "Diferença": valor_contabil,
                                        "QTD CONTÁBIL": qtde_contabil,
                                        "QTD CLIENTE": 0,
                                        "DIF. QTD": qtde_contabil,
                                        "STATUS": "⚠️ Sem parametrização",
                                        "MOTIVO": "Conta sem regra cadastrada no mini banco.",
                                    }
                                )
                                detalhes_por_conta[conta_contabil] = {
                                    "contabil": df_cont_detalhe,
                                    "cliente": pd.DataFrame(),
                                }
                            else:
                                lista_fornecedores = regras["fornecedor_cliente"].tolist()
                                filtro_cli = df_cliente["_Nivel_4"].isin(lista_fornecedores)
                                df_cli_match = df_cliente[filtro_cli].copy()

                                valor_cliente = round(df_cli_match["Valor_Tratado"].sum(), 2)
                                qtde_cliente = int(len(df_cli_match))
                                diff = round(valor_contabil - valor_cliente, 2)
                                diff_qtd = qtde_contabil - qtde_cliente
                                status, motivo = classificar_diferenca(diff)

                                resultados.append(
                                    {
                                        "Conta Débito": conta_contabil,
                                        "Grupo DRE (Contábil)": grupo_contabil,
                                        "Valor Contábil": valor_contabil,
                                        "Valor Cliente": valor_cliente,
                                        "Diferença": diff,
                                        "QTD CONTÁBIL": qtde_contabil,
                                        "QTD CLIENTE": qtde_cliente,
                                        "DIF. QTD": diff_qtd,
                                        "STATUS": status,
                                        "MOTIVO": motivo,
                                    }
                                )

                                detalhes_por_conta[conta_contabil] = {
                                    "contabil": df_cont_detalhe,
                                    "cliente": df_cli_match,
                                }

                        df_resultados = pd.DataFrame(resultados)
                        assinatura_atual = assinatura_arquivos(file_contabil, file_cliente)
                        st.session_state["analise_conciliacao"] = {
                            "df_resultados": df_resultados,
                            "detalhes_por_conta": detalhes_por_conta,
                            "empresa_id": empresa_selecionada,
                            "empresa_nome": EMPRESAS[empresa_selecionada],
                            "df_cliente": df_cliente,
                            "df_contabil": df_contabil,
                            "col_conta_contabil": col_conta_contabil,
                            "col_grupo_contabil": col_grupo_contabil,
                            "col_historico_contabil": col_historico_contabil,
                            "col_nivel3_cliente": col_nivel3_cliente,
                            "col_conta_cliente": col_conta_cliente,
                            "col_fornecedor_cliente": col_fornecedor_cliente,
                        }
                        st.session_state["analise_conciliacao_assinatura"] = assinatura_atual
                        st.session_state["analise_conciliacao_empresa"] = empresa_selecionada

                except Exception as e:
                    st.error(f"❌ Erro ao ler planilhas ou cruzar dados. Detalhe: {e}")

    analise_armazenada = st.session_state.get("analise_conciliacao")
    assinatura_atual = assinatura_arquivos(file_contabil, file_cliente) if file_contabil and file_cliente else None
    if (
        analise_armazenada
        and st.session_state.get("analise_conciliacao_empresa") == empresa_selecionada
        and st.session_state.get("analise_conciliacao_assinatura") == assinatura_atual
    ):
        renderizar_analise_conferencia(analise_armazenada)
    elif file_contabil and file_cliente:
        st.info("Clique em Iniciar Conferência para gerar a análise.")


# ==========================================
# TELA 2: GESTÃO DO BANCO DE DADOS (CRUD)
# ==========================================
elif menu == "Parametrização":
    st.title("⚙️ Parametrização da Empresa")
    st.markdown(f"Gerenciando as regras da empresa: **{EMPRESAS[empresa_selecionada]}**")
    st.caption("Aqui você consulta o que já foi parametrizado e identifica o que ainda falta vincular no cliente.")

    df_banco = carregar_parametrizacao_empresa(empresa_selecionada)

    st.subheader("Consulta de Parametrização")
    st.caption("Conta Débito do contábil x Cód. Plano de conta nº. 04 do cliente.")

    filtro_consulta = st.text_input("Buscar na parametrização", placeholder="Digite conta débito ou código do cliente...")
    if filtro_consulta:
        termo = filtro_consulta.strip().lower()
        mascara = (
            df_banco["conta_contabil"].astype(str).str.lower().str.contains(termo, na=False)
            | df_banco["fornecedor_cliente"].astype(str).str.lower().str.contains(termo, na=False)
        )
        df_banco_visivel = df_banco[mascara].copy()
    else:
        df_banco_visivel = df_banco.copy()

    if df_banco_visivel.empty:
        st.info("Nenhuma regra encontrada para o filtro atual.")
    else:
        st.dataframe(
            df_banco_visivel.rename(
                columns={
                    "conta_contabil": "CONTA DÉBITO",
                    "fornecedor_cliente": "CÓD. PLANO 04",
                }
            )[["CONTA DÉBITO", "CÓD. PLANO 04"]]
            .sort_values(["CONTA DÉBITO", "CÓD. PLANO 04"]),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("---")
    st.subheader("Fornecedores sem parametrização")
    st.caption("Lista automática dos lançamentos do cliente que ainda não têm de/para cadastrado.")

    fonte_cliente = st.session_state.get("analise_conciliacao", {}).get("df_cliente")
    col_nivel3_cliente = st.session_state.get("analise_conciliacao", {}).get("col_nivel3_cliente")
    col_conta_cliente = st.session_state.get("analise_conciliacao", {}).get("col_conta_cliente")
    col_fornecedor_cliente = st.session_state.get("analise_conciliacao", {}).get("col_fornecedor_cliente")

    if fonte_cliente is None or fonte_cliente.empty:
        st.info(
            "Para listar os fornecedores sem parametrização, carregue uma base do cliente abaixo ou rode uma conciliação primeiro."
        )
        arquivo_cliente_consulta = st.file_uploader(
            "📂 Upload CLIENTE / DRE para consulta",
            type=["xlsx", "xls"],
            key="cliente_consulta_parametrizacao",
        )
        if arquivo_cliente_consulta:
            fonte_cliente = limpar_colunas(pd.read_excel(arquivo_cliente_consulta, engine="openpyxl"))
            col_nivel3_cliente = obter_coluna(fonte_cliente, ["Plano de conta nº. 03"])
            col_conta_cliente = obter_coluna(fonte_cliente, ["Cód. Plano de conta nº. 04"])
            col_fornecedor_cliente = obter_coluna(fonte_cliente, ["Cliente/Fornecedor"])
            col_valor_cliente = obter_coluna(fonte_cliente, ["Débito"])
            if col_nivel3_cliente and col_conta_cliente and col_valor_cliente:
                fonte_cliente["Valor_Tratado"] = fonte_cliente[col_valor_cliente].apply(tratar_valor)
                fonte_cliente["_Nivel_3"] = fonte_cliente[col_nivel3_cliente].apply(normalizar_texto)
                fonte_cliente["_Nivel_4"] = fonte_cliente[col_conta_cliente].apply(normalizar_texto)
                if col_fornecedor_cliente:
                    fonte_cliente["_Fornecedor"] = fonte_cliente[col_fornecedor_cliente].apply(normalizar_texto)
            else:
                st.warning("A base enviada não tem as colunas mínimas esperadas para consulta.")

    if fonte_cliente is not None and not fonte_cliente.empty and col_nivel3_cliente and col_conta_cliente:
        df_pendencias = preparar_pendencias_parametrizacao(
            fonte_cliente,
            df_banco,
            col_nivel3_cliente,
            col_conta_cliente,
            col_fornecedor_cliente,
        )

        if df_pendencias.empty:
            st.success("Todas as contas/fornecedores dessa base já estão parametrizados.")
        else:
            st.dataframe(
                df_pendencias.rename(
                    columns={
                        "GRUPO DRE": "GRUPO DRE",
                        "CÓD. PLANO 04": "CÓD. PLANO 04",
                        "FORNECEDOR": "FORNECEDOR",
                        "VALOR": "VALOR",
                        "QTD LANCAMENTOS": "QTD LANCAMENTOS",
                        "STATUS": "STATUS",
                    }
                ).style.format({"VALOR": "R$ {:,.2f}"}),
                use_container_width=True,
                hide_index=True,
            )
    elif fonte_cliente is not None and fonte_cliente.empty:
        st.warning("Não foi possível montar a lista de pendências porque a base do cliente está vazia.")

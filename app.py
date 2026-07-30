"""
DRE Gerencial Automatizado — Conector Nibo + Streamlit
=======================================================

Estrutura: Grupo (linha da DRE) > Subgrupo (opcional) > Categoria real do
Nibo. Cada linha soma as categorias/subgrupos listados nela; os expansores
abaixo da tabela mostram a composição detalhada, e toda tabela tem botão
de exportar para Excel.

Segurança: o token NUNCA fica no código. Ele é lido de st.secrets, que no
Streamlit Cloud é configurado em Settings > Secrets (nunca vai pro GitHub).
"""

import io
import json
import os
import re
import streamlit as st
import pandas as pd
import requests
import plotly.graph_objects as go
from datetime import datetime, timedelta

# =========================================================================
# 1. ESTRUTURA DA DRE — Grupo/Subgrupo/Categoria (plano de contas real da
#    GC - Marketing & Gestão Estratégica, exportado do Nibo)
# =========================================================================
#
# Cada linha da DRE_STRUCTURE tem "sinal" (+1 entrada / -1 saída) e OU:
#   - "categorias": lista simples de categorias do Nibo, OU
#   - "subgrupos": dict {nome_subgrupo: [categorias...]} quando a linha tem
#     mais de um nível de composição (ex.: Custo Fixo)
#
# Mudanças estruturais desta versão (confirmadas com o usuário):
#   - "Custo Fixo" passa a ser um grupo único com 2 subgrupos (Despesas
#     Administrativas + Folha/CMO), substituindo a separação anterior
#     entre CSP e Despesas Operacionais. Por isso não há mais uma linha
#     "Lucro Bruto" distinta — vai direto de Receita Líquida para Lucro
#     Operacional.
#   - "Financeira - tarifas bancárias" saiu de Tributos e entrou no novo
#     grupo "Despesas Financeiras".
#   - "Despesas administrativas - pró-labore" agora está incluída no
#     subgrupo administrativo do Custo Fixo.

DRE_STRUCTURE = {
    "Receita Bruta": {
        "sinal": 1,
        "categorias": [
            "Receita com vendas - Boleto",
            "Receita com vendas - cartão de crédito",
            "Financeira - receita financeira",
            "Receita com vendas - pix",
        ],
    },
    "Tributos": {  # redutor de receita / impostos e custos sobre venda
        "sinal": -1,
        "categorias": [
            "Tributos - simples nacional",
            "Custo com cobrança - boleto",
            "Custo meio de emissão NF",
            "Custo meio de pagamento - máquina crédito",
            "Custo meio de pagamento - máquina crédito à vista",
            "Custo meio de pagamento - máquina débito",
            "Custo meio de pagamento - máquina pix",
        ],
    },
    "Custo Fixo": {
        "sinal": -1,
        "subgrupos": {
            "Despesa Operacional Administrativa": [
                "Custos operacionais - a identificar",
                "Despesas Administrativas - Cartão de Crédito",
                "Despesas administrativas - certificado digital",
                "Despesas administrativas - Confraternização Equipe",
                "Despesas administrativas - contabilidade",
                "Despesas administrativas - copa e cozinha",
                "Despesas administrativas - escritório de advocacia",
                "Despesas administrativas - eventos",
                "Despesas administrativas - farmáxia",
                "Despesas administrativas - Gráfica",
                "Despesas administrativas - informática",
                "Despesas administrativas - jurídico",
                "Despesas administrativas - licenças e software",
                "Despesas administrativas - material de escritório",
                "Despesas administrativas - Medicina do Trabalho",
                "Despesas administrativas - Patrocinios",
                "Despesas administrativas - pró-labore",
                "Despesas administrativas - serviço de limpeza",
                "Despesas administrativas - serviço de terceiros",
                "Despesas administrativas - telefonia e internet",
                "Despesas administrativas - transporte urbano",
                "Despesas administrativas - uniforme",
                "Despesas administrativas - viagem",
                "Despesas com instalação - água e esgoto",
                "Despesas com instalação - aluguel",
                "Despesas com instalação - energia elétrica",
                "Despesas com instalação - iptu",
                "Despesas com instalação - manutenção e conservação",
                "Despesas com instalação - segurança e monitorament",
            ],
            "Despesas com Folha - CMO": [
                "Folha - adiantamento de salários",
                "Folha - cursos e treinamentos",
                "Folha - fgts",
                "Folha - inss",
                "Folha - plano de saúde",
                "Folha - rescisões",
                "Folha - salários",
                "Folha - vale alimentação",
                "Folha - vale transporte",
            ],
        },
    },
    "Investimentos": {  # vem após o Lucro Operacional
        "sinal": -1,
        "categorias": [
            "Investimento - Participação em Eventos",
            "Investimento em Sistema Próprio",
            "Máquinas e equipamentos",
            "Marketing e publicidade - facebook ads",
            "Marketing e publicidade - google ads",
            "Móveis, utensílios e instalações",
            "Compra de ativo fixo",
        ],
    },
    "Despesas Financeiras": {
        "sinal": -1,
        "categorias": [
            "Taxas e contribuições",
            "Tributos - IOF",
            "Financeira - despesas financeiras",
            "Financeira - estornos",
            "Financeira - juros fornecedores",
            "Financeira - negociação de divida",
            "Financeira - tarifas bancárias",
        ],
    },
    "Atividade de Financiamento": {  # espelha o relatório nativo do Nibo
        "sinal": -1,
        "categorias": [
            "Retirada de capital",
            "Distribuição de lucros",
            "Juros sobre empréstimo bnds",
            "Pagamento empréstimo bnds",
            "Pagamento de empréstimos a terceiros",
            "Pagamento de empréstimo de sócios",
            "Pagamento de empréstimos bancários",
            "Juros sobre empréstimos bancários",
            "Multas sobre empréstimos bancários",
        ],
    },
}

# Ordem de exibição da DRE, incluindo as linhas calculadas (subtotais)
DRE_LINES_ORDER = [
    ("Receita Bruta", "detalhe"),
    ("Tributos", "detalhe"),
    ("Receita Líquida", "subtotal"),
    ("Custo Fixo", "detalhe"),
    ("Lucro Operacional", "subtotal"),
    ("Investimentos", "detalhe"),
    ("Despesas Financeiras", "detalhe"),
    ("Atividade de Financiamento", "detalhe"),
    ("Não Classificado", "detalhe"),
    ("Geração de Caixa Realizada", "subtotal"),
]

NIBO_BASE_URL = "https://api.nibo.com.br/empresas/v1"

# Paleta oficial da marca Breakr (Manual de Marca, nov/2024) — usada só nos
# gráficos, conforme solicitado.
BREAKR_AMARELO = "#FF9406"   # Amarelo Fagulha
BREAKR_VERMELHO = "#CA3F17"  # Vermelho Brasa
BREAKR_CINZA = "#F3F4F7"     # Cinza Vapor
BREAKR_PRETO = "#0F0D05"     # Preto Fumaça


def _norm(texto: str) -> str:
    """Normaliza texto para comparação: remove espaços não separáveis (nbsp),
    espaços duplicados e diferenças de maiúscula/minúscula."""
    if texto is None:
        return ""
    texto = texto.replace("\xa0", " ")
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto.lower()


# Mapas invertidos: categoria normalizada -> (linha da DRE, subgrupo)
_CATEGORIA_PARA_LINHA = {}
_CATEGORIA_PARA_SUBGRUPO = {}
for _linha, _info in DRE_STRUCTURE.items():
    if "subgrupos" in _info:
        for _sub, _cats in _info["subgrupos"].items():
            for _cat in _cats:
                _CATEGORIA_PARA_LINHA[_norm(_cat)] = _linha
                _CATEGORIA_PARA_SUBGRUPO[_norm(_cat)] = _sub
    else:
        for _cat in _info["categorias"]:
            _CATEGORIA_PARA_LINHA[_norm(_cat)] = _linha
            _CATEGORIA_PARA_SUBGRUPO[_norm(_cat)] = _linha  # sem subgrupo real: usa o próprio nome da linha


def classify_dre_line(categoria: str):
    return _CATEGORIA_PARA_LINHA.get(_norm(categoria))


def classify_subgrupo(categoria: str):
    return _CATEGORIA_PARA_SUBGRUPO.get(_norm(categoria))


# =========================================================================
# 2. CONEXÃO COM A API DO NIBO
# =========================================================================

def get_credentials():
    """Lê token e organization_id de st.secrets, com mensagens claras de erro."""
    try:
        token = st.secrets["NIBO_API_TOKEN"]
    except (KeyError, FileNotFoundError):
        st.error(
            "❌ Token do Nibo não encontrado em st.secrets. "
            "Configure `NIBO_API_TOKEN` em Settings > Secrets no Streamlit Cloud "
            "(ou em `.streamlit/secrets.toml` localmente)."
        )
        st.stop()
    org_id = st.secrets.get("NIBO_ORGANIZATION_ID", None)
    return token, org_id


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_schedules(token: str, org_id: str, date_from: str, date_to: str) -> pd.DataFrame:
    """
    Busca lançamentos (contas a pagar/receber) do Nibo entre duas datas,
    paginando via $top/$skip (padrão OData usado pela API do Nibo).
    """
    url = f"{NIBO_BASE_URL}/schedules"
    headers = {"apitoken": token}
    if org_id:
        headers["organization_id"] = org_id

    all_items = []
    top = 500
    skip = 0

    while True:
        params = {
            "$top": top,
            "$skip": skip,
            "$filter": f"dueDate ge {date_from} and dueDate le {date_to}",
            "$orderby": "dueDate asc",
        }
        try:
            resp = requests.get(url, headers=headers, params=params, timeout=30)
        except requests.exceptions.RequestException as e:
            st.error(f"❌ Falha de conexão com a API do Nibo: {e}")
            st.stop()

        if resp.status_code == 401:
            st.error("❌ Token do Nibo inválido ou expirado (401 Unauthorized).")
            st.stop()
        if resp.status_code != 200:
            st.error(f"❌ Erro na API do Nibo (HTTP {resp.status_code}): {resp.text[:300]}")
            st.stop()

        payload = resp.json()
        items = payload.get("items", payload if isinstance(payload, list) else [])
        if not items:
            break

        all_items.extend(items)
        if len(items) < top:
            break
        skip += top

    if not all_items:
        return pd.DataFrame()

    # --- normalização dos campos ---
    # Calculamos a FRAÇÃO efetivamente paga/recebida de cada lançamento
    # (value vs paidValue), não um booleano isPaid puro — isso captura
    # pagamentos parciais corretamente.
    rows = []
    for it in all_items:
        categories = it.get("categories") or [{
            "categoryName": (it.get("category") or {}).get("name", "Não categorizado"),
            "value": it.get("value", 0),
        }]
        data_competencia = it.get("accrualDate") or it.get("dueDate") or it.get("scheduleDate")
        # Regime de Caixa: usa Vencimento (dueDate). Só funciona corretamente
        # se, ao dar baixa em cada conta no Nibo, o Vencimento for ajustado
        # para bater com a data real do pagamento (prática adotada pelo
        # usuário de jan/2026 em diante).
        data_pagamento = it.get("dueDate") or data_competencia
        if not data_competencia:
            continue

        valor_categorias = [abs(float(cat.get("value", 0) or 0)) for cat in categories]
        soma_categorias = sum(valor_categorias)
        valor_schedule = abs(float(it.get("value", 0) or 0))
        denom = soma_categorias if soma_categorias > 0 else valor_schedule

        paid_value = it.get("paidValue")
        if paid_value is None:
            paid_value = valor_schedule if it.get("isPaid") else 0.0
        paid_value = abs(float(paid_value or 0))

        if denom > 0:
            fracao_paga = min(paid_value / denom, 1.0)
        else:
            fracao_paga = 1.0 if it.get("isPaid") else 0.0

        if fracao_paga <= 0:
            continue

        if soma_categorias > 0:
            # Caso normal: cada categoria tem seu próprio valor de rateio.
            for cat, valor_cat_base in zip(categories, valor_categorias):
                nome_cat = cat.get("categoryName") or cat.get("name") or "Não categorizado"
                valor_cat = valor_cat_base * fracao_paga
                if valor_cat == 0:
                    continue
                rows.append({
                    "data_competencia": data_competencia,
                    "data_pagamento": data_pagamento,
                    "categoria": nome_cat,
                    "valor": valor_cat,
                    "tipo_cat": cat.get("type", "out"),
                })
        else:
            # Lançamento sem valor discriminado por categoria (comum em
            # entradas manuais já criadas como pagas) — usa o valor
            # efetivamente pago, dividido entre as categorias existentes,
            # em vez de descartar o lançamento inteiro.
            valor_realizado = paid_value if paid_value > 0 else valor_schedule * fracao_paga
            if valor_realizado > 0 and categories:
                valor_por_cat = valor_realizado / len(categories)
                for cat in categories:
                    nome_cat = cat.get("categoryName") or cat.get("name") or "Não categorizado"
                    rows.append({
                        "data_competencia": data_competencia,
                        "data_pagamento": data_pagamento,
                        "categoria": nome_cat,
                        "valor": valor_por_cat,
                        "tipo_cat": cat.get("type", "out"),
                    })

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df["data_competencia"] = pd.to_datetime(df["data_competencia"], errors="coerce")
    df["data_pagamento"] = pd.to_datetime(df["data_pagamento"], errors="coerce")
    df = df.dropna(subset=["data_competencia"])
    df["data_pagamento"] = df["data_pagamento"].fillna(df["data_competencia"])
    return df


# =========================================================================
# 3. MONTAGEM DA DRE
# =========================================================================

def build_pivot_por_categoria(df: pd.DataFrame) -> pd.DataFrame:
    """Pivot (linha_dre, subgrupo, categoria) x mês, com sinal já aplicado.
    Categorias que não batem com nenhuma linha caem em "Não Classificado"
    usando o tipo in/out do próprio lançamento para o sinal."""
    if df.empty:
        return pd.DataFrame()

    df = df.copy()
    df["linha_dre"] = df["categoria"].apply(classify_dre_line)
    df["subgrupo"] = df["categoria"].apply(classify_subgrupo)

    def calc_valor_sinal(row):
        linha = row["linha_dre"]
        if linha is not None and linha in DRE_STRUCTURE:
            return row["valor"] * DRE_STRUCTURE[linha]["sinal"]
        return row["valor"] * (1 if row["tipo_cat"] == "in" else -1)

    df["valor_sinal"] = df.apply(calc_valor_sinal, axis=1)
    df["linha_dre"] = df["linha_dre"].fillna("Não Classificado")
    df["subgrupo"] = df["subgrupo"].fillna("Não Classificado")

    pivot = df.pivot_table(
        index=["linha_dre", "subgrupo", "categoria"], columns="mes",
        values="valor_sinal", aggfunc="sum", fill_value=0.0,
    )
    return pivot


def build_dre(pivot_categoria: pd.DataFrame) -> pd.DataFrame:
    """Soma o pivot por linha_dre e monta a DRE completa com subtotais."""
    if pivot_categoria.empty:
        return pd.DataFrame()

    meses = sorted(pivot_categoria.columns)
    por_linha = pivot_categoria.groupby(level="linha_dre").sum()
    por_linha = por_linha.reindex(columns=meses, fill_value=0.0)

    for linha in list(DRE_STRUCTURE) + ["Não Classificado"]:
        if linha not in por_linha.index:
            por_linha.loc[linha] = 0.0

    dre = pd.DataFrame(index=[l for l, _ in DRE_LINES_ORDER], columns=meses, dtype=float)

    for linha in DRE_STRUCTURE:
        dre.loc[linha] = por_linha.loc[linha]
    dre.loc["Não Classificado"] = por_linha.loc["Não Classificado"]

    dre.loc["Receita Líquida"] = dre.loc["Receita Bruta"] + dre.loc["Tributos"]
    dre.loc["Lucro Operacional"] = dre.loc["Receita Líquida"] + dre.loc["Custo Fixo"]
    dre.loc["Geração de Caixa Realizada"] = (
        dre.loc["Lucro Operacional"] + dre.loc["Investimentos"]
        + dre.loc["Despesas Financeiras"] + dre.loc["Atividade de Financiamento"]
        + dre.loc["Não Classificado"]
    )

    return dre


# =========================================================================
# 4. EXPORTAÇÃO PARA EXCEL
# =========================================================================

def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Dados") -> bytes:
    """Converte um DataFrame (com o índice preservado) em bytes de .xlsx."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name[:31])
    return buffer.getvalue()


def botao_exportar(df: pd.DataFrame, nome_arquivo: str, label: str = "⬇️ Exportar para Excel"):
    st.download_button(
        label=label,
        data=to_excel_bytes(df, sheet_name=nome_arquivo[:31]),
        file_name=f"{nome_arquivo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"download_{nome_arquivo}_{id(df)}",
    )


# =========================================================================
# 5. METAS POR LINHA DA DRE (mês a mês)
# =========================================================================
# Persistidas em um arquivo JSON local. ⚠️ No Streamlit Cloud, o disco é
# efêmero — se o app reiniciar (redeploy, hibernação por inatividade), esse
# arquivo pode ser perdido. Por isso o botão "Exportar metas" serve como
# backup: se isso acontecer, é só importar de novo.
METAS_PATH = "metas.json"

# Linhas em que bater a meta significa "atingir ou superar" o valor (metas
# de receita/lucro/caixa). Nas demais linhas de custo, bater a meta
# significa "não gastar mais do que o valor definido" (a meta é um teto).
LINHAS_META_SUPERAR = {"Receita Bruta", "Receita Líquida", "Lucro Operacional", "Geração de Caixa Realizada"}


def carregar_metas() -> dict:
    if os.path.exists(METAS_PATH):
        try:
            with open(METAS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def salvar_metas(metas: dict):
    try:
        with open(METAS_PATH, "w", encoding="utf-8") as f:
            json.dump(metas, f, ensure_ascii=False, indent=2)
        return True
    except OSError:
        return False


def status_meta(linha: str, valor_real: float, meta: float):
    """Retorna (emoji, texto) comparando o realizado com a meta."""
    if meta is None or meta == 0:
        return "", ""
    if linha in LINHAS_META_SUPERAR:
        pct = (valor_real / meta * 100) if meta else 0
        if valor_real >= meta:
            return "✅", f"{pct:.0f}% da meta"
        elif pct >= 90:
            return "⚠️", f"{pct:.0f}% da meta"
        else:
            return "❌", f"{pct:.0f}% da meta"
    else:
        # Linha de custo: meta é um teto de gasto (valor positivo).
        gasto_real = abs(valor_real)
        pct = (gasto_real / meta * 100) if meta else 0
        if gasto_real <= meta:
            return "✅", f"{pct:.0f}% do teto"
        elif pct <= 110:
            return "⚠️", f"{pct:.0f}% do teto"
        else:
            return "❌", f"{pct:.0f}% do teto"


def gerar_insights(dre: pd.DataFrame, metas: dict, fmt_fn) -> dict:
    """Analisa a DRE e devolve insights automáticos agrupados em
    Destaques, Atenção e Sugestões — tudo calculado por regras simples
    (variação mês a mês, metas, tendência de margem), sem IA externa."""
    meses_local = list(dre.columns)
    destaques, atencao, sugestoes = [], [], []

    if len(meses_local) < 2:
        return {"Destaques": [], "Atenção": [], "Sugestões": []}

    linhas_detalhe = [l for l, t in DRE_LINES_ORDER if t == "detalhe"]

    # 1) Variações bruscas mês a mês em qualquer linha de detalhe
    variacoes_detectadas = []
    for linha in linhas_detalhe:
        serie = dre.loc[linha]
        for i in range(1, len(meses_local)):
            atual, anterior = serie.iloc[i], serie.iloc[i - 1]
            mes_atual, mes_ant = meses_local[i], meses_local[i - 1]
            if abs(anterior) < 1000:  # ignora bases muito pequenas (% explode sem sentido)
                continue
            variacao_pct = (atual - anterior) / abs(anterior) * 100
            if abs(variacao_pct) >= 50:
                variacoes_detectadas.append((abs(atual - anterior), linha, mes_atual, mes_ant, atual, anterior, variacao_pct))
    # Prioriza pelas variações de maior impacto em R$, não só em %.
    variacoes_detectadas.sort(key=lambda x: x[0], reverse=True)
    for _, linha, mes_atual, mes_ant, atual, anterior, variacao_pct in variacoes_detectadas[:5]:
        direcao = "aumentou" if abs(atual) > abs(anterior) else "caiu"
        destaques.append(
            f"📌 **{linha}** {direcao} {abs(variacao_pct):.0f}% em {mes_atual} "
            f"({fmt_fn(atual)}) em relação a {mes_ant} ({fmt_fn(anterior)})."
        )

    # 2) Metas não atingidas com frequência
    for linha, metas_linha in metas.items():
        if not metas_linha or linha not in dre.index:
            continue
        falhas = []
        for mes in meses_local:
            meta_val = metas_linha.get(mes, 0.0)
            if not meta_val:
                continue
            emoji, _ = status_meta(linha, dre.loc[linha, mes], meta_val)
            if emoji == "❌":
                falhas.append(mes)
        meses_com_meta = [m for m in meses_local if metas_linha.get(m, 0.0)]
        if meses_com_meta and len(falhas) >= max(2, len(meses_com_meta) // 2 + 1):
            atencao.append(
                f"⚠️ **{linha}** ficou fora da meta em {len(falhas)} de {len(meses_com_meta)} "
                f"meses com meta definida ({', '.join(falhas)}) — vale revisar o teto/objetivo ou os gastos dessa linha."
            )

    # 3) Tendência de margem operacional (primeiro vs último mês)
    receita_serie = dre.loc["Receita Bruta"]
    lucro_serie = dre.loc["Lucro Operacional"]
    margem_inicio = (lucro_serie.iloc[0] / receita_serie.iloc[0] * 100) if receita_serie.iloc[0] else None
    margem_fim = (lucro_serie.iloc[-1] / receita_serie.iloc[-1] * 100) if receita_serie.iloc[-1] else None
    if margem_inicio is not None and margem_fim is not None:
        delta_margem = margem_fim - margem_inicio
        if delta_margem >= 5:
            destaques.append(
                f"📈 A margem operacional melhorou de {margem_inicio:.1f}% ({meses_local[0]}) "
                f"para {margem_fim:.1f}% ({meses_local[-1]})."
            )
        elif delta_margem <= -5:
            atencao.append(
                f"📉 A margem operacional caiu de {margem_inicio:.1f}% ({meses_local[0]}) "
                f"para {margem_fim:.1f}% ({meses_local[-1]}) — vale investigar o motivo."
            )

    # 4) "Não Classificado" relevante (problema de dados, não do negócio)
    if "Não Classificado" in dre.index:
        for mes in meses_local:
            nc = dre.loc["Não Classificado", mes]
            receita_mes = dre.loc["Receita Bruta", mes]
            if receita_mes and abs(nc) / abs(receita_mes) >= 0.03 and abs(nc) >= 500:
                sugestoes.append(
                    f"💡 Em {mes}, {fmt_fn(nc)} ficou em 'Não Classificado' "
                    f"({abs(nc)/abs(receita_mes)*100:.1f}% da receita do mês) — mapeie essas "
                    f"categorias no DRE_STRUCTURE para uma visão mais precisa."
                )

    # 5) Melhor e piores meses de caixa
    caixa_serie = dre.loc["Geração de Caixa Realizada"]
    mes_melhor = caixa_serie.idxmax()
    mes_pior = caixa_serie.idxmin()
    if mes_melhor != mes_pior:
        destaques.append(f"🏆 Melhor geração de caixa: **{mes_melhor}** ({fmt_fn(caixa_serie[mes_melhor])}).")
        if caixa_serie[mes_pior] < 0:
            atencao.append(f"🔻 Pior geração de caixa: **{mes_pior}** ({fmt_fn(caixa_serie[mes_pior])}).")

    # 6) Sugestão: linha de custo que mais cresceu em proporção à receita
    if len(meses_local) >= 2:
        maior_crescimento, linha_maior = 0, None
        for linha in ["Custo Fixo", "Investimentos", "Despesas Financeiras", "Atividade de Financiamento"]:
            if linha not in dre.index:
                continue
            pct_inicio = abs(dre.loc[linha, meses_local[0]]) / receita_serie.iloc[0] * 100 if receita_serie.iloc[0] else 0
            pct_fim = abs(dre.loc[linha, meses_local[-1]]) / receita_serie.iloc[-1] * 100 if receita_serie.iloc[-1] else 0
            crescimento = pct_fim - pct_inicio
            if crescimento > maior_crescimento:
                maior_crescimento, linha_maior = crescimento, linha
        if linha_maior and maior_crescimento >= 5:
            sugestoes.append(
                f"💡 **{linha_maior}** cresceu {maior_crescimento:.1f} pontos percentuais como "
                f"proporção da receita entre {meses_local[0]} e {meses_local[-1]} — é o maior "
                f"candidato a revisão se o objetivo for melhorar a margem."
            )

    if not sugestoes:
        sugestoes.append("✅ Nenhum ponto crítico de estrutura de custo identificado no período — continue monitorando.")

    return {"Destaques": destaques, "Atenção": atencao, "Sugestões": sugestoes}


def simular_metas(dre: pd.DataFrame, pivot_categoria: pd.DataFrame, meta_margem_pct: float,
                   meta_caixa: float, num_socios: int = 1) -> dict:
    """Calcula 'de trás para frente' o faturamento e os custos necessários
    para atingir uma margem operacional mínima E uma geração de caixa alvo,
    mantendo a folha de salários fixa (no patamar médio atual), os demais
    custos administrativos fixos na média, e isolando o pró-labore como a
    variável de ajuste dentro do orçamento administrativo — dividido pelo
    número de sócios."""
    meses_local = list(dre.columns)
    n = len(meses_local)

    avg_investimentos = dre.loc["Investimentos"].mean()
    avg_despesas_fin = dre.loc["Despesas Financeiras"].mean()
    avg_atividade_fin = dre.loc["Atividade de Financiamento"].mean()

    receita_total = dre.loc["Receita Bruta"].sum()
    tributos_total = dre.loc["Tributos"].sum()
    taxa_tributos = abs(tributos_total) / receita_total if receita_total else 0

    # Folha de salários (subgrupo dentro de Custo Fixo) — média histórica,
    # mantida fixa na simulação, conforme pedido.
    try:
        folha_media = pivot_categoria.xs(
            ("Custo Fixo", "Despesas com Folha - CMO"), level=("linha_dre", "subgrupo")
        ).sum().mean()
    except KeyError:
        folha_media = 0.0

    try:
        admin_media_atual = pivot_categoria.xs(
            ("Custo Fixo", "Despesa Operacional Administrativa"), level=("linha_dre", "subgrupo")
        ).sum().mean()
    except KeyError:
        admin_media_atual = 0.0

    # Pró-labore especificamente, dentro do subgrupo administrativo — é a
    # variável de ajuste; o resto do administrativo fica fixo na média.
    try:
        prolabore_media_atual = pivot_categoria.xs(
            "Despesas administrativas - pró-labore", level="categoria"
        ).sum().mean()
    except KeyError:
        prolabore_media_atual = 0.0
    outros_admin_media = admin_media_atual - prolabore_media_atual

    if meta_margem_pct <= 0:
        return {"erro": "A margem mínima precisa ser maior que 0%."}

    margem_frac = meta_margem_pct / 100
    lucro_operacional_necessario = meta_caixa - (avg_investimentos + avg_despesas_fin + avg_atividade_fin)
    receita_sugerida = lucro_operacional_necessario / margem_frac
    tributos_sugeridos = -taxa_tributos * receita_sugerida
    receita_liquida_sugerida = receita_sugerida + tributos_sugeridos
    custo_fixo_total_sugerido = lucro_operacional_necessario - receita_liquida_sugerida
    admin_sugerido = custo_fixo_total_sugerido - folha_media
    prolabore_sugerido_total = admin_sugerido - outros_admin_media
    num_socios = max(1, num_socios)
    prolabore_por_socio = prolabore_sugerido_total / num_socios

    return {
        "receita_sugerida": receita_sugerida,
        "taxa_tributos": taxa_tributos,
        "tributos_sugeridos": tributos_sugeridos,
        "receita_liquida_sugerida": receita_liquida_sugerida,
        "folha_media": folha_media,
        "admin_sugerido": admin_sugerido,
        "admin_media_atual": admin_media_atual,
        "outros_admin_media": outros_admin_media,
        "prolabore_media_atual": prolabore_media_atual,
        "prolabore_sugerido_total": prolabore_sugerido_total,
        "prolabore_por_socio": prolabore_por_socio,
        "num_socios": num_socios,
        "custo_fixo_total_sugerido": custo_fixo_total_sugerido,
        "lucro_operacional_necessario": lucro_operacional_necessario,
        "avg_investimentos": avg_investimentos,
        "avg_despesas_fin": avg_despesas_fin,
        "avg_atividade_fin": avg_atividade_fin,
        "meta_caixa": meta_caixa,
    }


# =========================================================================
# 6. INTERFACE — ESTILO EXECUTIVO
# =========================================================================

st.set_page_config(page_title="DRE Gerencial", layout="wide", page_icon="📊")

st.markdown("""
<style>
    .block-container {padding-top: 2rem;}
    div[data-testid="stMetric"] {
        background-color: #F7F9FB;
        border: 1px solid #E5E9EF;
        border-radius: 10px;
        padding: 16px 18px;
    }
    div[data-testid="stMetricLabel"] {font-size: 0.85rem; color: #5A6472;}
    thead tr th {background-color: #1F2A44 !important; color: white !important;}
    h1, h2, h3 {color: #1F2A44;}
</style>
""", unsafe_allow_html=True)

st.title("📊 DRE Gerencial — Breakr Assessoria")
st.caption("Dados sincronizados automaticamente com o Nibo")

token, org_id = get_credentials()

with st.sidebar:
    st.header("Filtros")
    hoje = datetime.today()
    data_inicio = st.date_input("Data inicial", value=hoje.replace(day=1) - timedelta(days=180))
    data_fim = st.date_input("Data final", value=hoje)
    regime = st.radio(
        "Regime de apresentação",
        options=["Competência", "Caixa (data de pagamento)"],
        help="Competência: agrupa pelo mês de competência do lançamento "
             "(igual ao Painel de acompanhamento nativo do Nibo). "
             "Caixa: agrupa pelo mês do Vencimento no Nibo — funciona bem "
             "se, ao dar baixa em cada conta, o Vencimento for ajustado "
             "para bater com a data real do pagamento.",
    )
    if st.button("🔄 Atualizar dados", use_container_width=True):
        st.cache_data.clear()
    st.divider()
    st.caption("Fonte: API Nibo · Atualização automática a cada 1h (cache)")

buffer_dias = timedelta(days=45)
date_from_str = (data_inicio - buffer_dias).strftime("%Y-%m-%dT00:00:00Z")
date_to_str = (data_fim + buffer_dias).strftime("%Y-%m-%dT23:59:59Z")

with st.spinner("Buscando lançamentos no Nibo..."):
    df_raw = fetch_schedules(token, org_id, date_from_str, date_to_str)

st.caption(f"📐 Regime: **{regime}**")

if df_raw.empty:
    st.warning("Nenhum lançamento retornado para o período selecionado.")
    st.stop()

# Calcula o "mês" de cada lançamento de acordo com o regime escolhido —
# não precisa buscar de novo na API, só reagrupar os dados já carregados.
data_ref = df_raw["data_competencia"] if regime == "Competência" else df_raw["data_pagamento"]
df_raw = df_raw.copy()
df_raw["mes"] = data_ref.dt.to_period("M").astype(str)

# Descarta a margem de segurança: mantém só os meses dentro do período que
# o usuário realmente selecionou nos filtros.
mes_min_real = data_inicio.strftime("%Y-%m")
mes_max_real = data_fim.strftime("%Y-%m")
df_raw = df_raw[(df_raw["mes"] >= mes_min_real) & (df_raw["mes"] <= mes_max_real)]

if df_raw.empty:
    st.warning("Nenhum lançamento no período selecionado (após ajuste de competência/caixa).")
    st.stop()


def fmt_moeda(v):
    if pd.isna(v):
        v = 0.0
    sinal = "-" if v < 0 else ""
    return f"{sinal}R$ {abs(v):,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")


def md(texto: str) -> str:
    """Escapa o cifrão para exibição segura em st.markdown/warning/success/
    error — sem isso, o Streamlit interpreta pares de '$' como fórmula
    matemática (LaTeX) e quebra o texto quando há mais de um valor em R$
    na mesma frase."""
    return texto.replace("$", "\\$")


with st.expander("🔬 Diagnóstico — total bruto por mês (para comparar com o relatório do Nibo)"):
    st.caption(
        "Esta tabela soma TUDO que a API retornou, com o sinal do campo "
        "'tipo_cat' (entrada/saída), antes de qualquer classificação de "
        "DRE. Se o total de algum mês aqui já não bater com o total de "
        "movimentações que você vê no próprio Nibo para o mesmo mês, o "
        "problema é na busca de dados (API/filtro), não na classificação."
    )
    df_diag = df_raw.copy()
    df_diag["valor_sinal"] = df_diag.apply(
        lambda r: r["valor"] * (1 if r["tipo_cat"] == "in" else -1), axis=1
    )
    diag = df_diag.groupby("mes").agg(
        total_liquido=("valor_sinal", "sum"),
        qtd_lancamentos=("valor_sinal", "count"),
        total_entradas=("valor_sinal", lambda s: s[s > 0].sum()),
        total_saidas=("valor_sinal", lambda s: s[s < 0].sum()),
    ).reset_index()
    diag_fmt = diag.copy()
    for col in ["total_liquido", "total_entradas", "total_saidas"]:
        diag_fmt[col] = diag_fmt[col].apply(fmt_moeda)
    st.dataframe(diag_fmt, use_container_width=True)
    botao_exportar(diag, "diagnostico_totais_por_mes")

pivot_categoria = build_pivot_por_categoria(df_raw)
dre = build_dre(pivot_categoria)

if dre.empty or len(dre.columns) == 0:
    st.warning(
        "Lançamentos encontrados, mas nenhuma categoria bateu com a DRE_STRUCTURE. "
        "Confira os nomes exatos das categorias abaixo e ajuste o dicionário "
        "DRE_STRUCTURE no topo do app.py."
    )
    resumo_categorias = (
        df_raw.groupby("categoria")["valor"]
        .agg(qtd_lancamentos="count", valor_total="sum")
        .sort_values("valor_total", ascending=False)
        .reset_index()
    )
    st.dataframe(resumo_categorias, use_container_width=True)
    botao_exportar(resumo_categorias, "categorias_nao_mapeadas")
    st.stop()

meses = list(dre.columns)
ultimo_mes = meses[-1]
mes_anterior = meses[-2] if len(meses) > 1 else None

metas_atuais = carregar_metas()

with st.expander("🧠 Insights automáticos do período", expanded=False):
    insights = gerar_insights(dre, metas_atuais, fmt_moeda)
    cols_insight = st.columns(3)
    titulos = ["📈 Destaques", "⚠️ Atenção", "💡 Sugestões"]
    chaves = ["Destaques", "Atenção", "Sugestões"]
    for col, titulo, chave in zip(cols_insight, titulos, chaves):
        with col:
            st.markdown(f"**{titulo}**")
            itens = insights[chave]
            if not itens:
                st.caption("Nada relevante identificado.")
            else:
                for item in itens:
                    st.markdown(f"- {md(item)}")

with st.expander("🎯 Simulador de metas — quanto preciso faturar e gastar?", expanded=False):
    st.caption(
        "Define a margem operacional mínima e a geração de caixa que você quer atingir "
        "por mês. O app calcula, de trás para frente, o faturamento necessário e o "
        "orçamento sugerido para cada linha — mantendo a **folha de salários** no "
        "patamar médio atual e os **Tributos** proporcionais à receita (calculados "
        "automaticamente, não são um valor fixo a definir)."
    )
    sc1, sc2, sc3 = st.columns(3)
    with sc1:
        meta_margem_input = st.number_input(
            "Margem Operacional mínima desejada (%)", min_value=0.1, max_value=95.0,
            value=40.0, step=1.0,
        )
    with sc2:
        meta_caixa_input = st.number_input(
            "Geração de Caixa desejada por mês (R$)", min_value=0.0, value=20000.0, step=1000.0,
        )
    with sc3:
        num_socios_input = st.number_input(
            "Número de sócios (para dividir o pró-labore)", min_value=1, max_value=20,
            value=1, step=1,
        )

    sim = simular_metas(dre, pivot_categoria, meta_margem_input, meta_caixa_input, int(num_socios_input))

    if "erro" in sim:
        st.error(sim["erro"])
    else:
        st.markdown("#### 📋 Plano sugerido (valor por mês)")
        p1, p2, p3, p4 = st.columns(4)
        p1.metric("Receita Bruta necessária", fmt_moeda(sim["receita_sugerida"]))
        p2.metric("Tributos (estimado)", fmt_moeda(sim["tributos_sugeridos"]),
                   f"{sim['taxa_tributos']*100:.1f}% da receita")
        p3.metric("Receita Líquida", fmt_moeda(sim["receita_liquida_sugerida"]))
        p4.metric("Lucro Operacional", fmt_moeda(sim["lucro_operacional_necessario"]))

        st.markdown("#### 💰 Orçamento de Custo Fixo")
        cf1, cf2, cf3 = st.columns(3)
        cf1.metric("Folha de Salários (fixo — média atual)", fmt_moeda(sim["folha_media"]))
        delta_admin = sim["admin_sugerido"] - sim["admin_media_atual"]
        cf2.metric(
            "Despesas Administrativas sugerido", fmt_moeda(sim["admin_sugerido"]),
            f"{fmt_moeda(delta_admin)} vs. média atual ({fmt_moeda(sim['admin_media_atual'])})",
            delta_color="inverse",
        )
        cf3.metric("Custo Fixo total permitido", fmt_moeda(sim["custo_fixo_total_sugerido"]))

        st.markdown("#### 👤 Pró-labore sugerido")
        st.caption(
            "Isola o pró-labore dentro do orçamento administrativo, mantendo os demais "
            "custos administrativos (contabilidade, jurídico, licenças etc.) fixos na "
            "média histórica — o pró-labore é quem absorve o ajuste."
        )
        pl1, pl2, pl3 = st.columns(3)
        pl1.metric("Outras despesas administrativas (fixo)", fmt_moeda(sim["outros_admin_media"]))
        delta_prolabore = sim["prolabore_sugerido_total"] - sim["prolabore_media_atual"]
        pl2.metric(
            "Pró-labore total sugerido", fmt_moeda(sim["prolabore_sugerido_total"]),
            f"{fmt_moeda(delta_prolabore)} vs. atual ({fmt_moeda(sim['prolabore_media_atual'])})",
            delta_color="inverse",
        )
        pl3.metric(
            f"Pró-labore por sócio ({sim['num_socios']}x)",
            fmt_moeda(abs(sim["prolabore_por_socio"])),
        )
        if sim["prolabore_sugerido_total"] > 0:
            st.error(
                "❌ Esse cenário exigiria pró-labore **positivo** (ou seja, negativo como "
                "custo não sobra orçamento nenhum) — a meta de margem/caixa não é viável "
                "sem cortar outras despesas administrativas ou aumentar mais o faturamento."
            )

        st.markdown("#### 📦 Orçamento sugerido (mantido no patamar médio histórico)")
        b1, b2, b3 = st.columns(3)
        b1.metric("Investimentos", fmt_moeda(sim["avg_investimentos"]))
        b2.metric("Despesas Financeiras", fmt_moeda(sim["avg_despesas_fin"]))
        b3.metric("Atividade de Financiamento", fmt_moeda(sim["avg_atividade_fin"]))

        if delta_admin < 0:
            st.warning(md(
                f"⚠️ Para bater essa meta mantendo o faturamento sugerido, as despesas "
                f"administrativas (incluindo pró-labore) precisam **cair {fmt_moeda(abs(delta_admin))}** "
                f"em relação à média atual. Se não for viável cortar, o caminho é aumentar "
                f"o faturamento além do valor sugerido."
            ))
        else:
            st.success(md(
                f"✅ Nesse cenário, as despesas administrativas têm folga de "
                f"{fmt_moeda(delta_admin)} em relação à média atual."
            ))

        plano_df = pd.DataFrame({
            "Linha": ["Receita Bruta", "Tributos", "Receita Líquida", "Folha de Salários",
                      "Outras Despesas Administrativas", "Pró-labore Total Sugerido",
                      f"Pró-labore por sócio ({sim['num_socios']}x)",
                      "Despesas Administrativas (Total)", "Custo Fixo Total", "Lucro Operacional",
                      "Investimentos", "Despesas Financeiras", "Atividade de Financiamento",
                      "Geração de Caixa (meta)"],
            "Valor Sugerido": [
                sim["receita_sugerida"], sim["tributos_sugeridos"], sim["receita_liquida_sugerida"],
                sim["folha_media"], sim["outros_admin_media"], sim["prolabore_sugerido_total"],
                sim["prolabore_por_socio"], sim["admin_sugerido"], sim["custo_fixo_total_sugerido"],
                sim["lucro_operacional_necessario"], sim["avg_investimentos"], sim["avg_despesas_fin"],
                sim["avg_atividade_fin"], sim["meta_caixa"],
            ],
        })
        botao_exportar(plano_df, "plano_de_metas", label="⬇️ Exportar plano sugerido")

        st.divider()
        st.caption(
            "O botão abaixo preenche a tabela de metas (seção '🎯 Definir metas por "
            "linha') com este plano, repetido em todos os meses do período — assim os "
            "✅⚠️❌ na DRE já refletem se cada mês está no caminho para bater essa meta."
        )
        if st.button("✅ Usar este plano como meta para todos os meses", use_container_width=True):
            novas_metas = {
                "Receita Bruta": {m: sim["receita_sugerida"] for m in meses},
                "Tributos": {m: abs(sim["tributos_sugeridos"]) for m in meses},
                "Receita Líquida": {m: sim["receita_liquida_sugerida"] for m in meses},
                "Custo Fixo": {m: abs(sim["custo_fixo_total_sugerido"]) for m in meses},
                "Lucro Operacional": {m: sim["lucro_operacional_necessario"] for m in meses},
                "Investimentos": {m: abs(sim["avg_investimentos"]) for m in meses},
                "Despesas Financeiras": {m: abs(sim["avg_despesas_fin"]) for m in meses},
                "Atividade de Financiamento": {m: abs(sim["avg_atividade_fin"]) for m in meses},
                "Não Classificado": {m: 0.0 for m in meses},
                "Geração de Caixa Realizada": {m: sim["meta_caixa"] for m in meses},
            }
            if salvar_metas(novas_metas):
                st.success("Metas aplicadas a todos os meses! Role até a tabela DRE para ver os ✅⚠️❌.")
                st.rerun()
            else:
                st.error("Não consegui salvar as metas em disco.")


def variacao(atual, anterior):
    if anterior in (0, None) or pd.isna(anterior):
        return None
    return (atual - anterior) / abs(anterior) * 100


# ---- KPIs ----
st.subheader(f"Indicadores — {ultimo_mes}")
k1, k2, k3, k4, k5 = st.columns(5)

receita = dre.loc["Receita Bruta", ultimo_mes]
lucro_op = dre.loc["Lucro Operacional", ultimo_mes]
caixa = dre.loc["Geração de Caixa Realizada", ultimo_mes]
margem_op = (lucro_op / receita * 100) if receita else 0
margem_caixa = (caixa / receita * 100) if receita else 0

receita_ant = dre.loc["Receita Bruta", mes_anterior] if mes_anterior else None
caixa_ant = dre.loc["Geração de Caixa Realizada", mes_anterior] if mes_anterior else None
lucro_op_ant = dre.loc["Lucro Operacional", mes_anterior] if mes_anterior else None

k1.metric("Receita Bruta", fmt_moeda(receita),
          f"{variacao(receita, receita_ant):.1f}%" if variacao(receita, receita_ant) is not None else None)
k2.metric("Lucro Operacional", fmt_moeda(lucro_op), f"{margem_op:.1f}% margem")
k3.metric("Variação Lucro Op.",
          f"{variacao(lucro_op, lucro_op_ant):.1f}%" if variacao(lucro_op, lucro_op_ant) is not None else "—")
k4.metric("Caixa Gerado", fmt_moeda(caixa),
          f"{variacao(caixa, caixa_ant):.1f}%" if variacao(caixa, caixa_ant) is not None else None)
k5.metric("Meses no período", f"{len(meses)}")

st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)


def gauge_kpi(titulo, valor_pct, valor_abs_fmt, faixa_max=60):
    """Gauge circular estilo Power BI/dashboard executivo."""
    cor = BREAKR_AMARELO if valor_pct >= 0 else BREAKR_VERMELHO
    limite = max(abs(valor_pct) * 1.3, faixa_max)
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=valor_pct,
        number={"suffix": "%", "font": {"size": 44, "color": cor}},
        gauge={
            "axis": {"range": [-limite, limite], "tickcolor": "#5A6472", "tickfont": {"size": 10}},
            "bar": {"color": cor, "thickness": 0.28},
            "bgcolor": "white",
            "borderwidth": 0,
            "threshold": {"line": {"color": BREAKR_PRETO, "width": 3}, "thickness": 0.75, "value": valor_pct},
        },
        domain={"x": [0, 1], "y": [0, 1]},
    ))
    fig.update_layout(height=200, margin=dict(t=20, b=0, l=30, r=30))
    st.markdown(f"<p style='text-align:center; font-weight:700; font-size:1.05rem; "
                f"color:#1F2A44; margin-bottom:0;'>{titulo}</p>", unsafe_allow_html=True)
    st.plotly_chart(fig, use_container_width=True)
    st.markdown(f"<p style='text-align:center; font-size:0.85rem; color:#5A6472; "
                f"margin-top:-10px;'>{md(valor_abs_fmt)}</p>", unsafe_allow_html=True)


kc1, kc2 = st.columns(2)
with kc1:
    gauge_kpi("Margem Operacional sobre Faturamento", margem_op,
              f"Lucro Op. {fmt_moeda(lucro_op)} ÷ Receita {fmt_moeda(receita)}")
with kc2:
    gauge_kpi("Caixa Gerado sobre Faturamento", margem_caixa,
              f"Caixa {fmt_moeda(caixa)} ÷ Receita {fmt_moeda(receita)}")


st.divider()

# ---- Metas por linha (mês a mês) ----
with st.expander("🎯 Definir metas por linha (mês a mês)"):
    st.caption(
        "Para linhas de Receita/Lucro/Caixa, a meta é um valor a **superar**. "
        "Para linhas de custo (Tributos, Custo Fixo, Investimentos, "
        "Despesas Financeiras, Atividade de Financiamento), a meta é um "
        "**teto de gasto** — digite sempre um número positivo, mesmo para "
        "linhas de custo."
    )
    metas_salvas = carregar_metas()
    linhas_todas = [l for l, _ in DRE_LINES_ORDER]
    metas_df = pd.DataFrame(
        {mes: [metas_salvas.get(linha, {}).get(mes, 0.0) for linha in linhas_todas] for mes in meses},
        index=linhas_todas,
    )
    metas_editadas = st.data_editor(
        metas_df, use_container_width=True,
        column_config={mes: st.column_config.NumberColumn(mes, format="%.0f") for mes in meses},
        key="editor_metas",
    )
    col_salvar, col_export, col_import = st.columns([1, 1, 2])
    with col_salvar:
        if st.button("💾 Salvar metas"):
            novas_metas = {
                linha: {mes: float(metas_editadas.loc[linha, mes]) for mes in meses}
                for linha in linhas_todas
            }
            if salvar_metas(novas_metas):
                st.success("Metas salvas!")
            else:
                st.error("Não consegui salvar em disco — exporta para Excel como backup.")
    with col_export:
        botao_exportar(metas_editadas, "metas_dre", label="⬇️ Exportar metas")
    with col_import:
        arquivo_metas = st.file_uploader("Importar metas (.xlsx)", type=["xlsx"], key="upload_metas")
        if arquivo_metas is not None:
            df_importado = pd.read_excel(arquivo_metas, index_col=0)
            metas_importadas = {
                linha: {mes: float(df_importado.loc[linha, mes]) for mes in df_importado.columns if linha in df_importado.index}
                for linha in linhas_todas
            }
            if salvar_metas(metas_importadas):
                st.success("Metas importadas! Atualize a página para ver refletido na tabela.")

metas_atuais = carregar_metas()

st.divider()

# ---- Tabela DRE comparativa ----
st.subheader("DRE Comparativa Mês a Mês")

dre_display_fmt = dre.copy().map(fmt_moeda)

margem_op_row = (dre.loc["Lucro Operacional"] / dre.loc["Receita Bruta"].replace(0, pd.NA) * 100).round(1)
for mes in meses:
    if pd.notna(margem_op_row[mes]):
        dre_display_fmt.loc["Lucro Operacional", mes] += f"  ({margem_op_row[mes]:.1f}%)"

# Anexa o ícone de status (✅/⚠️/❌) em cada célula que tiver meta definida.
for linha in dre.index:
    metas_linha = metas_atuais.get(linha, {})
    for mes in meses:
        meta_val = metas_linha.get(mes, 0.0)
        if meta_val:
            emoji, _texto = status_meta(linha, dre.loc[linha, mes], meta_val)
            if emoji:
                dre_display_fmt.loc[linha, mes] += f" {emoji}"

linhas_subtotal = {"Receita Líquida", "Lucro Operacional", "Geração de Caixa Realizada"}
index_labels = []
for linha, tipo in DRE_LINES_ORDER:
    prefixo = "🔹 " if tipo == "subtotal" else "   "
    index_labels.append(prefixo + linha)
dre_display_fmt.index = index_labels


def destacar_totalizadores(row):
    nome_linha = row.name.replace("🔹 ", "").strip()
    if nome_linha in linhas_subtotal:
        return ["background-color: #EAF3FF; color: #1F2A44; font-weight: 700;"] * len(row)
    return [""] * len(row)


styler = dre_display_fmt.style.apply(destacar_totalizadores, axis=1)
st.dataframe(styler, use_container_width=True)
st.caption("✅ Meta atingida · ⚠️ Perto da meta (dentro de 10%) · ❌ Fora da meta — defina metas no expansor acima.")
botao_exportar(dre, "dre_gerencial", label="⬇️ Exportar DRE completa para Excel")

st.divider()

# ---- Drill-down: composição de cada linha ----
st.subheader("🔍 Composição de cada linha (clique para expandir)")
st.caption("Mostra quanto cada subgrupo/categoria real do Nibo contribuiu, mês a mês.")

for linha, _tipo in DRE_LINES_ORDER:
    if linha not in DRE_STRUCTURE and linha != "Não Classificado":
        continue  # subtotais não têm composição própria

    titulo = f"{linha} — composição"
    if linha == "Não Classificado":
        titulo = "⚠️ Não Classificado — categorias que ainda não estão em nenhum grupo da DRE"

    with st.expander(titulo):
        try:
            sub = pivot_categoria.xs(linha, level="linha_dre")
        except KeyError:
            sub = None
        if sub is None or sub.empty:
            st.caption("Nenhum lançamento nessa linha no período selecionado.")
            continue

        sub = sub.reindex(columns=meses, fill_value=0.0)
        subgrupos_distintos = sub.index.get_level_values("subgrupo").unique()

        # Se a linha tem mais de um subgrupo real (não é só o nome da própria
        # linha repetido), mostra também o resumo por subgrupo antes do detalhe.
        tem_subgrupos_reais = len(subgrupos_distintos) > 1 or subgrupos_distintos[0] != linha
        if tem_subgrupos_reais:
            st.markdown("**Resumo por subgrupo:**")
            por_sub = sub.groupby(level="subgrupo").sum()
            por_sub = por_sub.loc[por_sub.sum(axis=1).sort_values(ascending=False).index]
            por_sub_fmt = por_sub.map(fmt_moeda)
            por_sub_fmt.insert(0, "Total no período", por_sub.sum(axis=1).apply(fmt_moeda))
            st.dataframe(por_sub_fmt, use_container_width=True)
            botao_exportar(por_sub, f"{linha}_subgrupos".replace(" ", "_"), label="⬇️ Exportar subgrupos")
            st.markdown("**Detalhe por categoria:**")

        detalhe = sub.droplevel("subgrupo") if "subgrupo" in sub.index.names else sub
        detalhe = detalhe.loc[detalhe.sum(axis=1).sort_values(ascending=False).index]
        detalhe_fmt = detalhe.map(fmt_moeda)
        detalhe_fmt.insert(0, "Total no período", detalhe.sum(axis=1).apply(fmt_moeda))
        st.dataframe(detalhe_fmt, use_container_width=True)
        botao_exportar(detalhe, f"{linha}_categorias".replace(" ", "_"), label="⬇️ Exportar categorias")

st.divider()

# ---- Gráficos ----
c1, c2 = st.columns(2)

with c1:
    st.subheader("Receita x Lucro Operacional")
    fig1 = go.Figure()
    fig1.add_trace(go.Bar(x=meses, y=dre.loc["Receita Bruta"], name="Receita Bruta", marker_color=BREAKR_PRETO))
    fig1.add_trace(go.Scatter(x=meses, y=dre.loc["Lucro Operacional"], name="Lucro Operacional",
                               mode="lines+markers", line=dict(color=BREAKR_AMARELO, width=3)))
    fig1.update_layout(height=380, margin=dict(t=20, b=20, l=10, r=10),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(fig1, use_container_width=True)

with c2:
    st.subheader("Tendência de Caixa")
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(x=meses, y=dre.loc["Geração de Caixa Realizada"], name="Caixa Gerado",
                               mode="lines+markers", fill="tozeroy",
                               line=dict(color=BREAKR_AMARELO, width=3)))
    fig2.update_layout(height=380, margin=dict(t=20, b=20, l=10, r=10))
    st.plotly_chart(fig2, use_container_width=True)

c3, c4 = st.columns(2)

with c3:
    st.subheader("Custo Fixo x Investimentos x Despesas Financeiras")
    fig3 = go.Figure()
    fig3.add_trace(go.Bar(x=meses, y=-dre.loc["Custo Fixo"], name="Custo Fixo", marker_color=BREAKR_VERMELHO))
    fig3.add_trace(go.Bar(x=meses, y=-dre.loc["Investimentos"], name="Investimentos", marker_color=BREAKR_AMARELO))
    fig3.add_trace(go.Bar(x=meses, y=-dre.loc["Despesas Financeiras"], name="Despesas Financeiras",
                           marker_color=BREAKR_VERMELHO))
    fig3.update_layout(height=380, barmode="group", margin=dict(t=20, b=20, l=10, r=10),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(fig3, use_container_width=True)

with c4:
    st.subheader("Margem Operacional (%)")
    fig4 = go.Figure()
    fig4.add_trace(go.Scatter(x=meses, y=margem_op_row, name="Margem Operacional %",
                               mode="lines+markers", line=dict(color=BREAKR_AMARELO, width=3)))
    fig4.update_layout(height=380, margin=dict(t=20, b=20, l=10, r=10),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(fig4, use_container_width=True)

st.caption("⚠️ Valores em regime de caixa (proporcional ao que já foi efetivamente pago/recebido no Nibo). "
           "Revise DRE_STRUCTURE no início do app.py para ajustar o mapeamento de categorias.")

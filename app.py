"""
Sistema de Onboarding de Prestadores de Serviço
------------------------------------------------
App Streamlit (arquivo único) que conduz o prestador pelas 7 etapas:

  1. Entrevistas                      (RH)
  2. Formalização e validação de CNPJ (Candidato)
  3. Validação das documentações      (RH)
  4. Assinatura do contrato           (RH)
  5. Provisionamento de acessos       (RH/TI)
  6. Reunião de integração            (RH)
  7. Início das operações             (RH)

Armazenamento:
  * Sem configuração nenhuma, grava em CSV local (data/prestadores.csv).
  * Se existir [gcp_service_account] e sheet_url em st.secrets, grava no Google Sheets.
  * Se existir drive_folder_id, os cartões CNPJ vão para uma pasta do Google Drive.

Validação de CNPJ: BrasilAPI (gratuita) com fallback para publica.cnpj.ws e,
se ambas falharem, declaração manual do candidato marcada para conferência do RH.
"""

import json
import re
import secrets as pysecrets
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

# =============================================================================
# CONFIGURAÇÕES GERAIS
# =============================================================================
st.set_page_config(page_title="Onboarding de Prestadores", page_icon="🤝", layout="wide")

CNAE_EXIGIDO = "7319002"  # 7319-0/02 Promoção de vendas
VAGAS = ["Operador EGV", "Gerente de Negócios", "Gestor Comercial"]
VAGAS_ACEITAM_MEI = {"Operador EGV"}  # nas demais vagas o CNPJ não pode ser MEI
STATUS_ENTREVISTA = ["Em Entrevista", "Aprovado", "Banco de Espera", "Reprovado"]
UFS = ["", "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR",
       "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"]
ESTADOS_CIVIS = ["Solteiro(a)", "Casado(a)", "União estável", "Divorciado(a)", "Viúvo(a)"]
STATUS_ATIVO = "Ativo / Operação Iniciada"
DOC_APROVADA = "Aprovada"
DOC_APROVADA_PEND = "Aprovada (CNPJ pendente)"
DOCS_LIBERADAS = (DOC_APROVADA, DOC_APROVADA_PEND)
SIM = "Sim"

DATA_DIR = Path(__file__).parent / "data"
CSV_PATH = DATA_DIR / "prestadores.csv"
UPLOAD_DIR = DATA_DIR / "uploads"

COLUNAS = [
    "id", "codigo_acesso", "criado_em", "atualizado_em",
    # Etapa 1
    "nome", "email", "telefone", "vaga", "status_entrevista", "obs_entrevista", "curriculo",
    "avaliacao_nota", "avaliacao_resumo", "avaliacao",
    # Etapa 2
    "nome_completo", "cpf", "rg", "data_nascimento", "endereco", "estado_civil",
    "cep", "uf", "cidade", "bairro", "logradouro", "numero", "complemento",
    "possui_cnpj", "cnpj", "razao_social", "natureza_juridica", "situacao_cnpj",
    "is_mei", "tem_cnae", "cnae_principal", "cnaes_secundarios", "fonte_validacao",
    "cnpj_apto", "cartao_cnpj", "form_enviado_em",
    # Etapa 3
    "doc_status", "doc_obs", "doc_validado_em",
    # Etapa 4
    "contrato_enviado", "contrato_assinado", "contrato_data", "contrato_arquivo",
    # Etapa 5
    "ti_email", "ti_looker", "ti_propostas", "ti_email_usuario", "ti_looker_usuario", "ti_propostas_usuario",
    # Etapa 6
    "onb_material", "onb_metas", "onb_cultura", "data_integracao", "padrinho",
    # Etapa 7
    "status_final", "operacao_iniciada_em",
]

ETAPAS = {
    1: "1. Entrevistas",
    2: "2. Formalização / CNPJ",
    3: "3. Validação de documentos",
    4: "4. Assinatura do contrato",
    5: "5. Acessos (TI)",
    6: "6. Integração / Onboarding",
    7: "7. Início das operações",
}


def secret(chave, padrao=None):
    """Lê st.secrets sem quebrar quando não existe secrets.toml."""
    try:
        return st.secrets[chave]
    except Exception:
        return padrao


def agora():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def data_br(txt):
    try:
        return datetime.strptime(txt, "%Y-%m-%d %H:%M").strftime("%d/%m/%Y às %H:%M")
    except ValueError:
        return txt


def so_digitos(txt):
    return re.sub(r"\D", "", txt or "")


def fmt_cnae(c):
    c = so_digitos(c).zfill(7)
    return f"{c[:4]}-{c[4]}/{c[5:]}"


def fmt_cnpj(c):
    c = so_digitos(c)
    if len(c) != 14:
        return c
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"


# =============================================================================
# ARMAZENAMENTO (CSV local ou Google Sheets, troca transparente)
# =============================================================================
def usa_sheets():
    return bool(secret("gcp_service_account")) and bool(secret("sheet_url"))


@st.cache_resource(show_spinner=False)
def _worksheet():
    import gspread

    gc = gspread.service_account_from_dict(dict(secret("gcp_service_account")))
    sh = gc.open_by_url(secret("sheet_url"))
    nome_aba = secret("sheet_tab", "prestadores")
    try:
        ws = sh.worksheet(nome_aba)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=nome_aba, rows=1000, cols=len(COLUNAS))
    return ws


@st.cache_data(ttl=20, show_spinner=False)
def _ler_sheets():
    valores = _worksheet().get_all_values()
    if not valores:
        return pd.DataFrame(columns=COLUNAS)
    return pd.DataFrame(valores[1:], columns=valores[0])


def carregar():
    if usa_sheets():
        df = _ler_sheets().copy()
    elif CSV_PATH.exists():
        df = pd.read_csv(CSV_PATH, dtype=str, keep_default_na=False)
    else:
        df = pd.DataFrame(columns=COLUNAS)
    for c in COLUNAS:
        if c not in df.columns:
            df[c] = ""
    return df[COLUNAS].fillna("").astype(str)


def salvar(df):
    df = df[COLUNAS].fillna("").astype(str)
    if usa_sheets():
        ws = _worksheet()
        ws.clear()
        ws.update(range_name="A1", values=[COLUNAS] + df.values.tolist())
        _ler_sheets.clear()
    else:
        DATA_DIR.mkdir(exist_ok=True)
        df.to_csv(CSV_PATH, index=False)


def buscar(pid):
    df = carregar()
    linha = df[df["id"] == pid]
    return None if linha.empty else linha.iloc[0].to_dict()


def atualizar(pid, **campos):
    df = carregar()
    idx = df.index[df["id"] == pid]
    if len(idx) == 0:
        st.error("Registro não encontrado.")
        return
    for k, v in campos.items():
        df.loc[idx, k] = "" if v is None else str(v)
    df.loc[idx, "atualizado_em"] = agora()
    salvar(df)


def criar(nome, email, telefone, vaga, status):
    df = carregar()
    novo = {c: "" for c in COLUNAS}
    novo.update(
        id=pysecrets.token_hex(4),
        codigo_acesso=pysecrets.token_hex(3).upper(),
        criado_em=agora(), atualizado_em=agora(),
        nome=nome.strip(), email=email.strip().lower(), telefone=telefone.strip(),
        vaga=vaga.strip(), status_entrevista=status,
    )
    df = pd.concat([df, pd.DataFrame([novo])], ignore_index=True)
    salvar(df)
    return novo


def guardar_arquivo(pid, arquivo):
    """Salva um anexo (cartão CNPJ, currículo, contrato) no Google Drive (se configurado) ou em disco local."""
    nome = f"{pid}_{re.sub(r'[^A-Za-z0-9._-]', '_', arquivo.name)}"
    conteudo = arquivo.getvalue()
    pasta = secret("drive_folder_id")
    if pasta and secret("gcp_service_account"):
        try:
            from google.auth.transport.requests import AuthorizedSession
            from google.oauth2.service_account import Credentials

            creds = Credentials.from_service_account_info(
                dict(secret("gcp_service_account")),
                scopes=["https://www.googleapis.com/auth/drive"],
            )
            meta = json.dumps({"name": nome, "parents": [pasta]})
            limite = "onboarding_boundary"
            corpo = (
                f"--{limite}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n{meta}\r\n"
                f"--{limite}\r\nContent-Type: {arquivo.type or 'application/octet-stream'}\r\n\r\n"
            ).encode() + conteudo + f"\r\n--{limite}--".encode()
            r = AuthorizedSession(creds).post(
                "https://www.googleapis.com/upload/drive/v3/files"
                "?uploadType=multipart&supportsAllDrives=true&fields=id,webViewLink",
                data=corpo,
                headers={"Content-Type": f"multipart/related; boundary={limite}"},
                timeout=60,
            )
            r.raise_for_status()
            return r.json()["webViewLink"]
        except Exception as e:  # cai para o disco local
            st.warning(f"Não foi possível enviar ao Google Drive ({e}). Arquivo salvo localmente.")
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    (UPLOAD_DIR / nome).write_bytes(conteudo)
    return f"local:{nome}"


# =============================================================================
# VALIDAÇÕES (CPF, CNPJ, MEI, CNAE)
# =============================================================================
def cpf_valido(cpf):
    c = so_digitos(cpf)
    if len(c) != 11 or len(set(c)) == 1:
        return False
    for i in (9, 10):
        soma = sum(int(c[j]) * (i + 1 - j) for j in range(i))
        if (soma * 10 % 11) % 10 != int(c[i]):
            return False
    return True


def cnpj_valido(cnpj):
    c = so_digitos(cnpj)
    if len(c) != 14 or len(set(c)) == 1:
        return False

    def dv(base, pesos):
        resto = sum(int(a) * b for a, b in zip(base, pesos)) % 11
        return "0" if resto < 2 else str(11 - resto)

    p1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    return c[12] == dv(c[:12], p1) and c[13] == dv(c[:13], [6] + p1)


def _parse_brasilapi(d):
    secundarios = [str(x.get("codigo") or "") for x in (d.get("cnaes_secundarios") or [])]
    return {
        "razao_social": d.get("razao_social", ""),
        "natureza_juridica": d.get("natureza_juridica", ""),
        "situacao": (d.get("descricao_situacao_cadastral") or "").upper(),
        "is_mei": bool(d.get("opcao_pelo_mei")) and not d.get("data_exclusao_do_mei"),
        "cnae_principal": str(d.get("cnae_fiscal") or ""),
        "cnaes_secundarios": [s for s in secundarios if s and s != "0"],
        "fonte": "BrasilAPI",
    }


def _parse_cnpjws(d):
    est = d.get("estabelecimento") or {}
    simples = d.get("simples") or {}
    return {
        "razao_social": d.get("razao_social", ""),
        "natureza_juridica": (d.get("natureza_juridica") or {}).get("descricao", ""),
        "situacao": (est.get("situacao_cadastral") or "").upper(),
        "is_mei": (simples.get("mei") or "").lower() == "sim",
        "cnae_principal": str((est.get("atividade_principal") or {}).get("id") or ""),
        "cnaes_secundarios": [str(x.get("id")) for x in (est.get("atividades_secundarias") or [])],
        "fonte": "CNPJ.ws",
    }


def _parse_receitaws(d):
    if d.get("status") == "ERROR":
        raise LookupError(d.get("message", ""))
    simei = d.get("simei") or {}
    principais = [so_digitos(x.get("code", "")) for x in (d.get("atividade_principal") or [])]
    secundarios = [so_digitos(x.get("code", "")) for x in (d.get("atividades_secundarias") or [])]
    return {
        "razao_social": d.get("nome", ""),
        "natureza_juridica": d.get("natureza_juridica", ""),
        "situacao": (d.get("situacao") or "").upper(),
        # O ReceitaWS gratuito nem sempre informa o SIMEI; sem essa informação o candidato declara.
        "is_mei": simei.get("optante") if "optante" in simei else None,
        "cnae_principal": principais[0] if principais else "",
        "cnaes_secundarios": [x for x in principais[1:] + secundarios if x and set(x) != {"0"}],
        "fonte": "ReceitaWS",
    }


@st.cache_data(ttl=600, show_spinner=False)
def consultar_cnpj(cnpj):
    """Retorna dict com status: ok | nao_encontrado | indisponivel.

    BrasilAPI e CNPJ.ws usam a base pública da Receita, atualizada cerca de uma vez por mês:
    CNPJs abertos há poucas semanas ainda não aparecem nelas. O ReceitaWS consulta a Receita
    de forma mais direta e costuma achar CNPJs novos (limite gratuito de 3 consultas/minuto).
    """
    c = so_digitos(cnpj)
    fontes = [
        (f"https://brasilapi.com.br/api/cnpj/v1/{c}", _parse_brasilapi),
        (f"https://publica.cnpj.ws/cnpj/{c}", _parse_cnpjws),
        (f"https://receitaws.com.br/v1/cnpj/{c}", _parse_receitaws),
    ]
    nao_encontrado = False
    for url, parser in fontes:
        try:
            r = requests.get(url, timeout=15, headers={"User-Agent": "onboarding-prestadores/1.0"})
            if r.status_code == 404:
                nao_encontrado = True
                continue
            r.raise_for_status()
            dados = parser(r.json())
            dados["status"] = "ok"
            return dados
        except LookupError:
            nao_encontrado = True
        except Exception:
            continue
    return {"status": "nao_encontrado" if nao_encontrado else "indisponivel"}


@st.cache_data(ttl=86400, show_spinner=False)
def consultar_cep(cep):
    """ViaCEP com reserva na BrasilAPI. Retorna dict com uf, cidade, bairro, logradouro ou None."""
    c = so_digitos(cep)
    if len(c) != 8:
        return None
    fontes = [
        (f"https://viacep.com.br/ws/{c}/json/",
         lambda d: None if d.get("erro") else {"uf": d.get("uf", ""), "cidade": d.get("localidade", ""),
                                               "bairro": d.get("bairro", ""), "logradouro": d.get("logradouro", "")}),
        (f"https://brasilapi.com.br/api/cep/v2/{c}",
         lambda d: {"uf": d.get("state", ""), "cidade": d.get("city", ""),
                    "bairro": d.get("neighborhood", ""), "logradouro": d.get("street", "")}),
    ]
    for url, parser in fontes:
        try:
            r = requests.get(url, timeout=10, headers={"User-Agent": "onboarding-prestadores/1.0"})
            if r.status_code == 200:
                dados = parser(r.json())
                if dados and dados.get("cidade"):
                    return dados
        except Exception:
            continue
    return None


def fmt_cep(c):
    c = so_digitos(c)
    return f"{c[:5]}-{c[5:]}" if len(c) == 8 else c


def montar_endereco(d):
    linha = ", ".join(x for x in [d["logradouro"], d["numero"]] if x)
    if d["complemento"]:
        linha += f" - {d['complemento']}"
    partes = [linha, d["bairro"], "/".join(x for x in [d["cidade"], d["uf"]] if x)]
    if d["cep"]:
        partes.append(f"CEP {fmt_cep(d['cep'])}")
    return " · ".join(x for x in partes if x)


def aceita_mei(vaga):
    return vaga in VAGAS_ACEITAM_MEI


def avaliar_regras(is_mei, tem_cnae, situacao="ATIVA", vaga=""):
    """Aplica as regras de negócio. Retorna (apto, lista_de_pendências)."""
    pendencias = []
    if situacao and situacao != "ATIVA":
        pendencias.append(f"O CNPJ está com situação cadastral **{situacao}**. Ele precisa estar **ATIVO**.")
    if is_mei and not aceita_mei(vaga):
        pendencias.append(
            "Sua empresa está enquadrada como **MEI**, o que não é aceito para esta parceria. "
            "Solicite o desenquadramento do MEI e a migração para ME (Empresário Individual), "
            "SLU ou LTDA junto ao seu contador ou ao Portal do Simples Nacional."
        )
    if not tem_cnae:
        pendencias.append(
            f"O CNAE **{fmt_cnae(CNAE_EXIGIDO)} (Promoção de vendas)** não consta no seu CNPJ. "
            "Solicite ao seu contador a inclusão desse CNAE (principal ou secundário) "
            "por meio de alteração cadastral na Junta Comercial / Redesim."
        )
    return (len(pendencias) == 0, pendencias)


AVISO_SEM_MEI = """
**Importante: a parceria NÃO aceita MEI (Microempreendedor Individual).**
O MEI não permite a atividade de promoção de vendas nos moldes da parceria e tem limite de faturamento baixo.
"""

AVISO_COM_MEI = """
**Para esta vaga o MEI é aceito.** Se preferir, abra como MEI com o CNAE abaixo; os modelos a seguir também servem.
"""

ORIENTACOES_SEM_CNPJ = """
**Modelos jurídicos recomendados**
- **ME – Empresário Individual (EI):** abertura simples, sem sócios; o patrimônio pessoal não é separado do da empresa.
- **SLU – Sociedade Limitada Unipessoal:** um único dono com patrimônio pessoal protegido e sem capital mínimo. *Costuma ser a opção mais indicada.*
- **LTDA – Sociedade Limitada:** quando houver dois ou mais sócios.
- **EPP – Empresa de Pequeno Porte:** é um porte (faturamento de R$ 360 mil a R$ 4,8 milhões/ano), aplicável a qualquer um dos modelos acima.

**Regime tributário**
- ✅ **Simples Nacional (recomendado):** guia única (DAS), menor burocracia e alíquotas iniciais reduzidas.
- **Lucro Presumido:** pode compensar com faturamento mais alto ou margens elevadas; avalie com o contador.
- **Lucro Real:** obrigatório acima de R$ 78 milhões/ano; raramente indicado para prestadores.

**Atividade obrigatória:** inclua o CNAE **7319-0/02 – Promoção de vendas** (principal ou secundário).

Depois de abrir a empresa, volte a este formulário com o mesmo código de acesso e informe o CNPJ.
"""


# =============================================================================
# ETAPA ATUAL / PROGRESSO
# =============================================================================
def cnpj_declarado(r):
    """CNPJ que não estava nas bases públicas e ainda não foi conferido pelo RH no cartão."""
    return r["possui_cnpj"] == SIM and "declara" in r["fonte_validacao"].lower()


def cnpj_pendente(r):
    return r["cnpj_apto"] != SIM or cnpj_declarado(r)


def etapas_concluidas(r):
    """Situação de cada etapa de onboarding. As etapas correm em paralelo, em qualquer ordem."""
    return {
        2: bool(r["form_enviado_em"]) and not cnpj_pendente(r),
        3: r["doc_status"] in DOCS_LIBERADAS,
        4: r["contrato_assinado"] == SIM,
        5: all(r[k] == SIM for k in ("ti_email", "ti_looker", "ti_propostas")),
        6: bool(r["padrinho"]) and all(r[k] == SIM for k in ("onb_material", "onb_metas", "onb_cultura")),
    }


NOMES_CURTOS = {2: "CNPJ/Formulário", 3: "Documentos", 4: "Contrato", 5: "Acessos", 6: "Integração"}


def pendencias(r):
    return [NOMES_CURTOS[k] for k, ok in etapas_concluidas(r).items() if not ok]


def fase(r):
    """1 = entrevista, 2 = onboarding, 8 = ativo."""
    if r["status_final"] == STATUS_ATIVO:
        return 8
    if r["status_entrevista"] != "Aprovado":
        return 1
    return 2


def rotulo_etapa(r):
    f = fase(r)
    if f == 8:
        return STATUS_ATIVO
    if f == 1:
        return f"1. Entrevistas ({r['status_entrevista'] or '-'})"
    pend = pendencias(r)
    if not pend:
        return "Onboarding concluído · aguardando início"
    return "Onboarding · falta: " + ", ".join(pend)




# =============================================================================
# VISUAL (CSS, chips, progresso)
# =============================================================================
CSS = """
<style>
.block-container {padding-top: 3.2rem; max-width: 1600px;}
.stTabs [role="tablist"] {gap: 6px; flex-wrap: wrap; border-bottom: none;}
.stTabs [data-testid="stTab"] {background: #fff; border: 1px solid #E5E7EB; border-radius: 10px;
                            padding: 6px 14px; height: auto; margin: 0;}
.stTabs [data-testid="stTab"][aria-selected="true"] {background: #4F46E5; border-color: #4F46E5;}
.stTabs [data-testid="stTab"][aria-selected="true"] p {color: #fff;}
.stTabs .react-aria-SelectionIndicator {display: none;}
h1, h2, h3 {letter-spacing: -0.01em;}
.hero {background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 100%); color: #fff;
       padding: 20px 26px; border-radius: 16px; margin-bottom: 18px;}
.hero h2 {color: #fff; margin: 0; font-size: 1.55rem; font-weight: 700; padding: 0;}
.hero p {margin: 6px 0 0; opacity: .92; font-size: .92rem;}
.hero code {background: rgba(255,255,255,.18); color: #fff; padding: 1px 8px; border-radius: 6px;}
.kpis {display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; margin-bottom: 18px;}
.kpi {background: #fff; border: 1px solid #E5E7EB; border-radius: 14px; padding: 12px 16px;}
.kpi b {font-size: 1.6rem; display: block; line-height: 1.2; color: #111827;}
.kpi b.txt {font-size: 1.15rem; padding: 6px 0 3px;}
.kpi span {color: #6B7280; font-size: .8rem;}
.kb-head {display: flex; justify-content: space-between; align-items: center; padding: 8px 12px;
          border-radius: 10px; font-weight: 600; font-size: .86rem; color: #fff; margin-bottom: 6px;}
.kb-count {background: rgba(255,255,255,.28); border-radius: 999px; padding: 0 9px; font-size: .78rem;}
.kb-vazio {color: #9CA3AF; font-size: .8rem; text-align: center; padding: 14px 0;}
.card-nome {font-weight: 650; font-size: .93rem; color: #111827; margin-bottom: 1px;}
.card-sub {color: #6B7280; font-size: .76rem; margin-bottom: 6px;}
.chip {display: inline-block; padding: 1px 8px; border-radius: 999px; font-size: .69rem;
       font-weight: 600; margin: 0 4px 4px 0; white-space: nowrap;}
.chip-ok {background: #DCFCE7; color: #166534;}
.chip-pend {background: #F3F4F6; color: #4B5563;}
.chip-warn {background: #FEF3C7; color: #92400E;}
.chip-bad {background: #FEE2E2; color: #991B1B;}
.chip-info {background: #E0E7FF; color: #3730A3;}
.dots {display: flex; gap: 3px; margin: 4px 0 2px;}
.dots span {flex: 1; height: 5px; border-radius: 3px; background: #E5E7EB;}
.dots span.on {background: #22C55E;}
.passos {display: flex; flex-wrap: wrap; gap: 6px; margin: 0 0 12px;}
.passo {padding: 5px 12px; border-radius: 999px; font-size: .8rem; font-weight: 600;
        background: #F3F4F6; color: #4B5563; border: 1px solid #E5E7EB;}
.passo.ok {background: #DCFCE7; color: #166534; border-color: #BBF7D0;}
.campos {display: grid; grid-template-columns: 170px 1fr; gap: 4px 12px; font-size: .9rem; margin-bottom: 8px;}
.campos .k {color: #6B7280;}
.campos .v {color: #111827; word-break: break-word;}
</style>
"""

# Tela do prestador: sem barra lateral e com o login em destaque.
CSS_CANDIDATO = """
<style>
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"], [data-testid="stExpandSidebarButton"] {display: none;}
.st-key-btn_acesso_rh, .st-key-btn_voltar_cand {position: fixed; top: 58px; right: 22px; z-index: 999990; width: auto;}
.st-key-btn_acesso_rh button, .st-key-btn_voltar_cand button {color: #9CA3AF; font-size: .78rem; min-height: 0; padding: 2px 6px;}
.st-key-btn_acesso_rh button p, .st-key-btn_voltar_cand button p {font-size: .78rem;}
.st-key-btn_acesso_rh button:hover, .st-key-btn_voltar_cand button:hover {color: #4F46E5;}
.login-cand {text-align: center; margin: 7vh 0 22px;}
.login-icone {font-size: 3.6rem; line-height: 1; margin-bottom: 10px;}
.login-cand h1 {font-size: 2.5rem; font-weight: 800; color: #1F2937; padding: 0; margin: 0 0 10px; letter-spacing: -0.02em;}
.login-cand p {font-size: 1.15rem; color: #6B7280; margin: 0;}
.st-key-login_cand {background: #fff; border: 1px solid #E5E7EB; border-radius: 20px; padding: 28px 30px 30px;
                    box-shadow: 0 12px 32px rgba(79, 70, 229, .10);}
.st-key-login_cand label p {font-size: 1.05rem; font-weight: 600; color: #374151;}
.st-key-login_cand input {font-size: 1.25rem; padding: 14px 16px;}
.st-key-login_cand [data-testid="stFormSubmitButton"] button {min-height: 3.4rem; border-radius: 12px; margin-top: 8px;}
.st-key-login_cand [data-testid="stFormSubmitButton"] button p {font-size: 1.2rem; font-weight: 700;}
</style>
"""

KANBAN = [
    ("entrevista", "🗣️ Entrevista", "#6366F1"),
    (2, "📝 Formalização", "#0EA5E9"),
    (3, "📂 Documentos", "#14B8A6"),
    (4, "✍️ Contrato", "#F59E0B"),
    (5, "💻 Acessos TI", "#8B5CF6"),
    (6, "🤝 Integração", "#EC4899"),
    (7, "🚀 Início / Ativos", "#22C55E"),
]


def avisar(msg, icone="✅"):
    """Guarda um aviso para aparecer como notificação no canto da tela depois do st.rerun()."""
    st.session_state["_aviso"] = (msg, icone)


def mostrar_aviso():
    aviso = st.session_state.pop("_aviso", None)
    if aviso:
        st.toast(aviso[0], icon=aviso[1])


def html(txt):
    st.markdown(txt, unsafe_allow_html=True)


def esc(txt):
    return (str(txt or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def chip(txt, tipo="pend"):
    return f'<span class="chip chip-{tipo}">{esc(txt)}</span>'


def coluna_kanban(r):
    """O card fica na primeira etapa ainda pendente; as demais aparecem nos indicadores."""
    if r["status_final"] == STATUS_ATIVO:
        return 7
    if r["status_entrevista"] != "Aprovado":
        return "entrevista"
    for k, ok in etapas_concluidas(r).items():
        if not ok:
            return k
    return 7


def passos_html(r):
    feitas = etapas_concluidas(r)
    itens = [f'<span class="passo ok">1 · Entrevista ✓</span>' if r["status_entrevista"] == "Aprovado"
             else f'<span class="passo">1 · Entrevista ({esc(r["status_entrevista"])})</span>']
    for k, ok in feitas.items():
        itens.append(f'<span class="passo{" ok" if ok else ""}">{k} · {NOMES_CURTOS[k]}{" ✓" if ok else ""}</span>')
    ativo = r["status_final"] == STATUS_ATIVO
    itens.append(f'<span class="passo{" ok" if ativo else ""}">7 · Início{" ✓" if ativo else ""}</span>')
    return '<div class="passos">' + "".join(itens) + "</div>"


def campos_html(d):
    linhas = "".join(f'<div class="k">{esc(k)}</div><div class="v">{esc(v) or "—"}</div>' for k, v in d.items())
    return f'<div class="campos">{linhas}</div>'


def barra_progresso(r):
    """Versão para o candidato."""
    html(passos_html(r))


# =============================================================================
# VISÃO DO CANDIDATO
# =============================================================================
@st.dialog("Dados enviados ✅")
def dialogo_envio(info):
    if info["apto"]:
        st.markdown("Recebemos seus **dados pessoais e o Cartão CNPJ**. 🎉")
        st.markdown("O RH vai conferir a documentação e falar com você sobre o contrato e os próximos passos.")
    elif info["possui"] == "Não":
        st.markdown("Recebemos seus **dados pessoais**.")
        st.markdown("Assim que abrir o CNPJ (com o CNAE 7319-0/02, seguindo as orientações do formulário), volte aqui com o mesmo "
                    "código de acesso e informe o número.")
    else:
        st.markdown("Recebemos seus **dados pessoais**.")
        st.markdown("Seu CNPJ ainda precisa de ajustes. Quando regularizar, volte aqui com o mesmo código e "
                    "valide de novo.")
    if st.button("Entendi", type="primary", width="stretch"):
        st.rerun()


def pendencias_cnpj(r):
    if r["possui_cnpj"] == "Não":
        tipos = "MEI, ME, SLU ou LTDA" if aceita_mei(r["vaga"]) else "ME, SLU ou LTDA (não pode ser MEI)"
        return [f"Você informou que ainda não tem CNPJ. Abra a empresa como {tipos}, "
                f"com o CNAE {fmt_cnae(CNAE_EXIGIDO)} (Promoção de vendas)."]
    _, pend = avaliar_regras(r["is_mei"] == SIM, r["tem_cnae"] == SIM, r["situacao_cnpj"], r["vaga"])
    return pend or ["Confira o CNPJ informado e valide de novo."]


def botao_acesso_rh():
    """Botão discreto no canto superior direito que leva ao login do RH."""
    if st.button("Acesso RH", key="btn_acesso_rh", type="tertiary"):
        st.session_state["modo"] = "rh"
        st.rerun()


def tela_candidato():
    pid = st.session_state.get("cand_id")
    if not pid:
        botao_acesso_rh()
        _, meio, _ = st.columns([1, 1.6, 1])
        with meio:
            html('<div class="login-cand"><div class="login-icone">🤝</div>'
                 '<h1>Bem-vindo ao seu onboarding</h1>'
                 '<p>Digite o código de acesso que você recebeu do RH e o seu e-mail para continuar.</p></div>')
        with meio, st.container(key="login_cand"), st.form("form_login_cand", border=False):
            codigo = st.text_input("Código de acesso", placeholder="Ex.: A1B2C3").strip().upper()
            email = st.text_input("E-mail", placeholder="seu@email.com").strip().lower()
            if st.form_submit_button("Acessar meu onboarding", type="primary", width="stretch"):
                df = carregar()
                achou = df[(df["codigo_acesso"] == codigo) & (df["email"] == email)]
                if achou.empty:
                    st.error("Código ou e-mail não conferem. Verifique com o RH.")
                else:
                    st.session_state["cand_id"] = achou.iloc[0]["id"]
                    st.rerun()
        return

    r = buscar(pid)
    if r is None:
        st.session_state.pop("cand_id", None)
        st.rerun()

    info_envio = st.session_state.pop("_envio_cand", None)
    if info_envio:
        dialogo_envio(info_envio)

    topo1, topo2 = st.columns([5, 1])
    with topo1:
        html(f'<div class="hero"><h2>Olá, {esc(r["nome"].split(" ")[0])}!</h2>'
             f'<p>{esc(r["vaga"])} · acompanhe aqui o seu onboarding.</p></div>')
    if topo2.button("Sair", width="stretch"):
        st.session_state.pop("cand_id", None)
        st.rerun()
    barra_progresso(r)

    if r["status_entrevista"] != "Aprovado":
        msgs = {
            "Em Entrevista": "Seu processo ainda está na fase de entrevistas. Aguarde o contato do RH.",
            "Banco de Espera": "Você está no nosso banco de espera. Entraremos em contato assim que houver oportunidade.",
            "Reprovado": "Agradecemos sua participação. Neste momento não seguiremos com o seu processo.",
        }
        st.info(msgs.get(r["status_entrevista"], "Aguarde o contato do RH."))
        return

    if r["status_final"] == STATUS_ATIVO:
        st.success("🎉 Seu onboarding foi concluído. Boas-vindas e boas vendas!")
        return

    bloqueado = r["form_enviado_em"] and r["cnpj_apto"] == SIM and r["doc_status"] != "Reprovada"
    if bloqueado:
        st.success("Recebemos seus dados e documentos. O RH vai conduzir as próximas etapas com você.")
        if r["padrinho"]:
            st.write(f"**Seu padrinho no primeiro mês:** {r['padrinho']}")
        with st.expander("Ver dados enviados"):
            html(campos_html({"Nome": r["nome_completo"], "CPF": r["cpf"], "E-mail": r["email"],
                              "Telefone": r["telefone"], "CNPJ": fmt_cnpj(r["cnpj"]),
                              "Razão social": r["razao_social"]}))
        return

    if r["doc_status"] != "Reprovada" and r["form_enviado_em"]:
        st.success(f"✅ Recebemos seus dados pessoais em {data_br(r['form_enviado_em'])}.")
        with st.container(border=True):
            st.markdown("##### ⚠️ Falta regularizar o CNPJ")
            for p in pendencias_cnpj(r):
                st.markdown(f"- {p}")
            st.caption("Quando estiver tudo certo, abra a seção abaixo, informe o CNPJ e envie de novo.")
        with st.expander("✏️ Atualizar meus dados ou informar o CNPJ"):
            formulario_candidato(r)
        return

    if r["doc_status"] == "Reprovada":
        st.error(f"O RH pediu ajustes na sua documentação: {r['doc_obs'] or 'veja com o RH.'}")
    with st.container(border=True):
        formulario_candidato(r)


def campos_endereco(r, k):
    """Campos de endereço. Ao informar o CEP (Enter ou sair do campo), preenche estado, cidade, bairro e rua."""
    campos = ["cep", "uf", "cidade", "bairro", "logradouro", "numero", "complemento"]
    for c in campos:  # valores iniciais vêm do cadastro; depois o session_state manda
        st.session_state.setdefault(k(c), fmt_cep(r[c]) if c == "cep" else r[c])
    if st.session_state[k("uf")] not in UFS:
        st.session_state[k("uf")] = ""

    def ao_mudar_cep():
        dados = consultar_cep(st.session_state[k("cep")])
        st.session_state[k("cep_msg")] = "ok" if dados else "erro"
        if dados:
            st.session_state[k("cep")] = fmt_cep(st.session_state[k("cep")])
            for c, v in dados.items():
                if v:
                    st.session_state[k(c)] = v if c != "uf" or v in UFS else ""

    st.markdown("##### Endereço")
    if r["endereco"] and not r["logradouro"]:
        st.caption(f"Endereço informado antes: {r['endereco']}")
    c1, c2, c3 = st.columns([1, 1, 2])
    c1.text_input("CEP *", key=k("cep"), placeholder="00000-000", on_change=ao_mudar_cep,
                  help="Digite o CEP e aperte Enter: estado, cidade, bairro e rua são preenchidos sozinhos.")
    c2.selectbox("Estado *", UFS, key=k("uf"))
    c3.text_input("Cidade *", key=k("cidade"))
    msg = st.session_state.get(k("cep_msg"))
    if msg == "ok":
        c1.caption("✅ Endereço encontrado")
    elif msg == "erro":
        c1.caption("⚠️ CEP não encontrado. Preencha à mão.")
    c1, c2 = st.columns([1, 2])
    c1.text_input("Bairro *", key=k("bairro"))
    c2.text_input("Endereço (rua, avenida...) *", key=k("logradouro"))
    c1, c2 = st.columns([1, 2])
    c1.text_input("Número *", key=k("numero"))
    c2.text_input("Complemento", key=k("complemento"), placeholder="Apto, bloco, sala...")
    return {c: st.session_state[k(c)].strip() if isinstance(st.session_state[k(c)], str) else st.session_state[k(c)]
            for c in campos}


def formulario_candidato(r, prefixo="cand"):
    """Formulário da Etapa 2. Usado pelo candidato e, no painel, pelo RH em nome dele."""
    pid = r["id"]

    def k(nome):
        return f"{prefixo}_{nome}_{pid}"

    st.markdown("##### Dados pessoais")
    c1, c2 = st.columns(2)
    nome_completo = c1.text_input("Nome completo *", r["nome_completo"] or r["nome"], key=k("nome"))
    cpf = c2.text_input("CPF *", r["cpf"], placeholder="000.000.000-00", key=k("cpf"))
    rg = c1.text_input("RG *", r["rg"], key=k("rg"))
    try:
        dn_padrao = date.fromisoformat(r["data_nascimento"]) if r["data_nascimento"] else date(1990, 1, 1)
    except ValueError:
        dn_padrao = date(1990, 1, 1)
    nasc = c2.date_input("Data de nascimento *", dn_padrao, min_value=date(1940, 1, 1),
                         max_value=date.today(), format="DD/MM/YYYY", key=k("nasc"))
    telefone = c1.text_input("Telefone *", r["telefone"], key=k("tel"))
    email = c2.text_input("E-mail *", r["email"], key=k("email"))
    ec_idx = ESTADOS_CIVIS.index(r["estado_civil"]) if r["estado_civil"] in ESTADOS_CIVIS else 0
    estado_civil = c1.selectbox("Estado civil *", ESTADOS_CIVIS, index=ec_idx, key=k("ec"))
    end = campos_endereco(r, k)

    st.markdown("##### Pessoa jurídica")
    possui = st.radio("Já possui CNPJ?", ["Sim", "Não"], horizontal=True,
                      index=0 if r["possui_cnpj"] != "Não" else 1, key=k("possui"))

    cnpj_info = {}
    arquivo = None
    apto = False

    if possui == "Não":
        st.info((AVISO_COM_MEI if aceita_mei(r["vaga"]) else AVISO_SEM_MEI) + ORIENTACOES_SEM_CNPJ)
        cnpj_info = {"possui_cnpj": "Não", "cnpj_apto": ""}
    else:
        cnpj = st.text_input("CNPJ *", fmt_cnpj(r["cnpj"]), placeholder="00.000.000/0000-00", key=k("cnpj"))
        chave = f"cnpj_res_{pid}"
        if st.button("🔎 Validar CNPJ", key=k("validar")):
            if not cnpj_valido(cnpj):
                st.session_state[chave] = {"status": "invalido"}
            else:
                with st.spinner("Consultando a Receita Federal..."):
                    st.session_state[chave] = {**consultar_cnpj(so_digitos(cnpj)), "cnpj": so_digitos(cnpj)}

        res = st.session_state.get(chave)
        if res and res.get("cnpj") and res["cnpj"] != so_digitos(cnpj):
            res = None  # o CNPJ digitado mudou, precisa validar de novo

        if not res:
            st.caption("Clique em **Validar CNPJ** para conferirmos MEI e CNAE automaticamente.")
        elif res["status"] == "invalido":
            st.error("CNPJ inválido. Confira os números digitados.")
        else:
            declarado = False
            if res["status"] == "ok":
                cnaes = [res["cnae_principal"]] + res["cnaes_secundarios"]
                tem_cnae = CNAE_EXIGIDO in [so_digitos(x).zfill(7) for x in cnaes if x]
                situacao, fonte = res["situacao"], f"Automática ({res['fonte']})"
                st.write(f"**{res['razao_social']}** · {res['natureza_juridica']} · Situação: {situacao}")
                col_a, col_b = st.columns(2)
                col_b.metric(f"CNAE {fmt_cnae(CNAE_EXIGIDO)}?", "Sim ✅" if tem_cnae else "Não ❌")
                if res["is_mei"] is None:  # a fonte não informou o SIMEI
                    is_mei = col_a.radio("A empresa é MEI?", ["Sim", "Não"], horizontal=True, index=1,
                                         key=k("mei_ws")) == "Sim"
                    fonte += " · MEI declarado"
                    declarado = True
                else:
                    is_mei = res["is_mei"]
                    col_a.metric("Enquadrado como MEI?", ("Sim ✅" if aceita_mei(r["vaga"]) else "Sim ❌")
                                 if is_mei else "Não ✅")
            else:  # não encontrado ou fontes fora do ar -> declaração conferida pelo RH no cartão
                if res["status"] == "nao_encontrado":
                    st.warning(
                        "Não encontramos este CNPJ nas bases públicas da Receita. **Isso é normal para CNPJs abertos "
                        "há poucas semanas**: essas bases são atualizadas cerca de uma vez por mês. "
                        "Responda abaixo e anexe o Cartão CNPJ; o RH vai conferir."
                    )
                    fonte = "Declaração do candidato (CNPJ ainda fora das bases públicas)"
                else:
                    st.warning("Não conseguimos consultar a Receita agora. Responda abaixo e anexe o Cartão CNPJ; "
                               "o RH vai conferir.")
                    fonte = "Declaração do candidato (consulta indisponível)"
                is_mei = st.radio("A empresa é MEI?", ["Sim", "Não"], horizontal=True, index=1, key=k("mei")) == "Sim"
                tem_cnae = st.radio(f"O CNPJ possui o CNAE {fmt_cnae(CNAE_EXIGIDO)} (Promoção de vendas)?",
                                    ["Sim", "Não"], horizontal=True, key=k("cnae")) == "Sim"
                situacao = ""
                declarado = True
                res = {**res, "razao_social": "", "natureza_juridica": "", "cnae_principal": "", "cnaes_secundarios": []}

            apto, pend = avaliar_regras(is_mei, tem_cnae, situacao, r["vaga"])
            confirmou = True
            if apto and declarado:
                confirmou = st.checkbox("Declaro que essas informações conferem com o Cartão CNPJ que vou anexar.",
                                        key=k("declaro"))
            if apto and confirmou:
                st.success("O CNPJ atende aos requisitos. Anexe o Cartão CNPJ para concluir.")
                arquivo = st.file_uploader("Cartão CNPJ (PDF ou imagem) *", type=["pdf", "png", "jpg", "jpeg"],
                                           key=k("cartao"))
            elif apto:
                st.info("Marque a declaração acima para anexar o Cartão CNPJ.")
            else:
                st.error("O CNPJ precisa de regularização antes de seguir:")
                for p in pend:
                    st.markdown(f"- {p}")
                st.caption("Dá para enviar os dados pessoais agora e voltar com o CNPJ regularizado.")
            apto = apto and confirmou

            cnpj_info = {
                "possui_cnpj": "Sim", "cnpj": so_digitos(cnpj),
                "razao_social": res.get("razao_social", ""), "natureza_juridica": res.get("natureza_juridica", ""),
                "situacao_cnpj": situacao, "is_mei": SIM if is_mei else "Não",
                "tem_cnae": SIM if tem_cnae else "Não",
                "cnae_principal": fmt_cnae(res["cnae_principal"]) if res.get("cnae_principal") else "",
                "cnaes_secundarios": ", ".join(fmt_cnae(x) for x in res.get("cnaes_secundarios", [])),
                "fonte_validacao": fonte, "cnpj_apto": SIM if apto else "Não",
            }

    st.divider()
    rotulo = "Enviar dados e documentos" if apto else "Enviar dados pessoais"
    if prefixo != "cand":
        rotulo = "Salvar Etapa 2"
    if st.button(rotulo, type="primary", key=k("enviar")):
        erros = []
        if not all([nome_completo.strip(), rg.strip(), telefone.strip(), email.strip()]):
            erros.append("Preencha todos os campos obrigatórios (*).")
        if len(so_digitos(end["cep"])) != 8 or not all(end[x] for x in ("uf", "cidade", "bairro", "logradouro", "numero")):
            erros.append("Preencha o endereço completo: CEP, estado, cidade, bairro, endereço e número.")
        if not cpf_valido(cpf):
            erros.append("CPF inválido.")
        if possui == "Sim" and not cnpj_info:
            erros.append("Valide o CNPJ antes de enviar.")
        if apto and arquivo is None and not r["cartao_cnpj"]:
            erros.append("Anexe o Cartão CNPJ.")
        if erros:
            for e in erros:
                st.error(e)
            return

        campos = dict(
            nome_completo=nome_completo.strip(), cpf=so_digitos(cpf), rg=rg.strip(),
            data_nascimento=nasc.isoformat(), telefone=telefone.strip(), email=email.strip().lower(),
            estado_civil=estado_civil, endereco=montar_endereco(end), form_enviado_em=agora(), **end, **cnpj_info,
        )
        if arquivo is not None:
            campos["cartao_cnpj"] = guardar_arquivo(pid, arquivo)
        if apto and r["doc_status"] != DOC_APROVADA_PEND:
            campos.update(doc_status="Pendente", doc_obs="")
        atualizar(pid, **campos)
        if prefixo == "cand":
            st.session_state["_envio_cand"] = {"apto": apto, "possui": possui}
        else:
            avisar("Etapa 2 salva." + ir_para_proxima(pid, 1, True))
        st.rerun()


# =============================================================================
# PAINEL DO RH
# =============================================================================
def login_rh():
    if st.session_state.get("rh_ok"):
        return True
    senha_cfg = secret("rh_password", "rh123")
    if st.button("← Acesso do prestador", key="btn_voltar_cand", type="tertiary"):
        st.session_state["modo"] = "candidato"
        st.rerun()
    html('<div class="hero"><h2>🔒 Painel de Controle · RH</h2><p>Acesso restrito à equipe de RH.</p></div>')
    _, meio, _ = st.columns([1, 2, 1])
    with meio:
        if senha_cfg == "rh123":
            st.caption("Senha de demonstração: **rh123** (defina `rh_password` em secrets para trocar).")
        with st.form("login_rh"):
            senha = st.text_input("Senha", type="password")
            if st.form_submit_button("Entrar", type="primary", width="stretch"):
                if senha == senha_cfg:
                    st.session_state["rh_ok"] = True
                    st.rerun()
                st.error("Senha incorreta.")
    return False


def tela_rh():
    if not login_rh():
        return
    pid = st.session_state.get("ficha")
    if pid and buscar(pid):
        ficha_prestador(buscar(pid))
        return
    st.session_state.pop("ficha", None)

    html('<div class="hero"><h2>📋 Painel de Controle · RH</h2>'
         '<p>Acompanhe cada prestador da entrevista ao início das operações.</p></div>')
    aba_quadro, aba_novo, aba_comp, aba_base = st.tabs(
        ["🗂️ Quadro", "➕ Novo candidato", "📈 Comparativo", "📋 Base completa"])
    with aba_quadro:
        quadro_kanban()
    with aba_novo:
        novo_candidato()
    with aba_comp:
        comparativo()
    with aba_base:
        base_completa()


def kpis(df):
    ativos = int((df["status_final"] == STATUS_ATIVO).sum()) if not df.empty else 0
    em_ent = int((df["status_entrevista"] == "Em Entrevista").sum()) if not df.empty else 0
    onb = int(df.apply(lambda r: fase(r) == 2, axis=1).sum()) if not df.empty else 0
    pend_cnpj = int(df.apply(lambda r: fase(r) == 2 and bool(r["form_enviado_em"]) and cnpj_pendente(r),
                                  axis=1).sum()) if not df.empty else 0
    itens = [("Candidatos", len(df)), ("Em entrevista", em_ent), ("Em onboarding", onb),
             ("CNPJ pendente", pend_cnpj), ("Ativos", ativos)]
    html('<div class="kpis">' + "".join(f'<div class="kpi"><b>{v}</b><span>{t}</span></div>' for t, v in itens)
         + "</div>")


def card_prestador(r):
    feitas = etapas_concluidas(r)
    chips = [chip(r["vaga"], "info")]
    if r["status_entrevista"] != "Aprovado":
        tipo = {"Reprovado": "bad", "Banco de Espera": "warn"}.get(r["status_entrevista"], "pend")
        chips.append(chip(r["status_entrevista"], tipo))
    if r["avaliacao_nota"]:
        chips.append(chip(f"Nota {r['avaliacao_nota']}", "info"))
    if fase(r) == 2 and r["form_enviado_em"] and cnpj_pendente(r):
        chips.append(chip("CNPJ pendente", "warn"))
    if r["doc_status"] == "Reprovada":
        chips.append(chip("Docs reprovados", "bad"))
    if r["status_final"] == STATUS_ATIVO:
        chips.append(chip("Ativo", "ok"))
    elif fase(r) == 2 and not pendencias(r):
        chips.append(chip("Pronto para iniciar", "ok"))
    dots = ""
    if fase(r) != 1:
        dots = '<div class="dots">' + "".join(
            f'<span class="{"on" if ok else ""}" title="{NOMES_CURTOS[k]}"></span>' for k, ok in feitas.items()) + "</div>"
    with st.container(border=True):
        html(f'<div class="card-nome">{esc(r["nome"])}</div>'
             f'<div class="card-sub">Código {esc(r["codigo_acesso"])}</div>{"".join(chips)}{dots}')
        if st.button("Abrir ficha", key=f"abrir_{r['id']}", width="stretch"):
            st.session_state["ficha"] = r["id"]
            st.rerun()


def quadro_kanban():
    df = carregar()
    kpis(df)
    if df.empty:
        st.info("Nenhum candidato ainda. Cadastre o primeiro na aba **➕ Novo candidato**.")
        return
    f1, f2, f3 = st.columns([2, 2, 1.4])
    busca = f1.text_input("Buscar por nome ou código", key="kb_busca", placeholder="Digite para filtrar")
    vagas = ["Todas as vagas"] + sorted(v for v in df["vaga"].unique() if v)
    vaga = f2.selectbox("Vaga", vagas, key="kb_vaga")
    f3.write("")
    mostrar_fora = f3.toggle("Banco de espera e reprovados", key="kb_fora")

    if busca:
        b = busca.strip().lower()
        df = df[df["nome"].str.lower().str.contains(b, regex=False) | df["codigo_acesso"].str.lower().str.contains(b)]
    if vaga != "Todas as vagas":
        df = df[df["vaga"] == vaga]
    if not mostrar_fora:
        df = df[~df["status_entrevista"].isin(["Banco de Espera", "Reprovado"])]

    df = df.assign(_col=df.apply(coluna_kanban, axis=1) if not df.empty else [])
    st.caption("Cada card fica na primeira etapa pendente. As barrinhas mostram as etapas 2 a 6: "
               "verde é concluída. As etapas podem ser feitas em qualquer ordem dentro da ficha.")
    colunas = st.columns(len(KANBAN), gap="small")
    for (chave, titulo, cor), col in zip(KANBAN, colunas):
        itens = df[df["_col"] == chave] if not df.empty else df
        with col:
            html(f'<div class="kb-head" style="background:{cor}"><span>{titulo}</span>'
                 f'<span class="kb-count">{len(itens)}</span></div>')
            if itens.empty:
                html('<div class="kb-vazio">Nenhum</div>')
            for _, r in itens.sort_values("atualizado_em", ascending=False).iterrows():
                card_prestador(r.to_dict())


def novo_candidato():
    _, meio, _ = st.columns([1, 3, 1])
    with meio, st.form("novo_cand", clear_on_submit=True):
        st.markdown("##### Cadastrar candidato")
        c1, c2 = st.columns(2)
        nome = c1.text_input("Nome *")
        email = c2.text_input("E-mail *")
        tel = c1.text_input("Telefone *")
        vaga = c2.selectbox("Vaga / Cargo *", VAGAS, index=None, placeholder="Selecione a vaga")
        status = st.selectbox("Status da entrevista", STATUS_ENTREVISTA)
        cv = st.file_uploader("Currículo (PDF ou Word, opcional)", type=["pdf", "doc", "docx"])
        if st.form_submit_button("Cadastrar", type="primary", width="stretch"):
            if not all([nome.strip(), email.strip(), tel.strip(), vaga]):
                st.error("Preencha todos os campos.")
            elif "@" not in email:
                st.error("E-mail inválido.")
            elif (carregar()["email"] == email.strip().lower()).any():
                st.error("Já existe um candidato com esse e-mail.")
            else:
                novo = criar(nome, email, tel, vaga, status)
                if cv is not None:
                    atualizar(novo["id"], curriculo=guardar_arquivo(f"{novo['id']}_cv", cv))
                st.success(f"{novo['nome']} cadastrado. Código de acesso: **{novo['codigo_acesso']}**")


def base_completa():
    df = carregar()
    if df.empty:
        st.info("Nenhum candidato cadastrado ainda.")
        return
    df["situacao"] = df.apply(rotulo_etapa, axis=1)
    df["cnpj_sit"] = df.apply(
        lambda r: "" if not r["form_enviado_em"] else ("Regular" if not cnpj_pendente(r) else "Pendente"), axis=1)
    st.dataframe(
        df[["nome", "codigo_acesso", "vaga", "email", "telefone", "status_entrevista", "avaliacao_nota",
            "situacao", "cnpj_sit", "padrinho", "atualizado_em"]]
        .rename(columns={"nome": "Nome", "codigo_acesso": "Código de acesso", "vaga": "Vaga", "email": "E-mail",
                         "telefone": "Telefone", "status_entrevista": "Entrevista", "avaliacao_nota": "Nota",
                         "situacao": "Situação", "cnpj_sit": "CNPJ", "padrinho": "Padrinho",
                         "atualizado_em": "Atualizado em"}),
        hide_index=True, width="stretch",
    )
    st.download_button("⬇️ Exportar base completa (CSV)", df[COLUNAS].to_csv(index=False).encode("utf-8-sig"),
                       "prestadores.csv", "text/csv")


def caminho_local(valor):
    return UPLOAD_DIR / valor[6:] if valor.startswith("local:") else None


def mostrar_arquivo(valor, rotulo, key):
    if not valor:
        st.caption(f"{rotulo}: não anexado")
    elif valor.startswith("local:"):
        caminho = UPLOAD_DIR / valor[6:]
        if caminho.exists():
            st.download_button(f"📎 Baixar {rotulo}", caminho.read_bytes(), caminho.name, key=key)
        else:
            st.warning(f"{rotulo}: arquivo não encontrado (o servidor pode ter sido reiniciado). Anexe de novo.")
    else:
        st.link_button(f"📎 Abrir {rotulo} no Drive", valor)


def ck(r, campo, rotulo, key):
    # A chave muda a cada salvamento para a caixa refletir o que foi gravado (ex.: marcada ao anexar o contrato).
    return st.checkbox(rotulo, value=r[campo] == SIM, key=f"{key}_{r['id']}_{r['atualizado_em']}")


def sn(v):
    return SIM if v else ""


def mensagem_acesso(r):
    return (f"Olá, {r['nome'].split(' ')[0]}! Para seguir com o seu onboarding, acesse "
            f"{secret('app_url', 'https://ehsonboard.streamlit.app')} e entre com:\n"
            f"Código de acesso: {r['codigo_acesso']}\nE-mail: {r['email']}")


# ---------------------------------------------------------------------------
# Ficha do prestador: uma aba por etapa
# ---------------------------------------------------------------------------
def situacao_ficha(r):
    if r["status_final"] == STATUS_ATIVO:
        return f"✅ {STATUS_ATIVO} desde {r['operacao_iniciada_em']}"
    if r["status_entrevista"] != "Aprovado":
        return f"Entrevista: {r['status_entrevista']}"
    feitas = etapas_concluidas(r)
    pend = [NOMES_CURTOS[k] for k, ok in feitas.items() if not ok]
    txt = f"{sum(feitas.values())} de {len(feitas)} etapas de onboarding concluídas"
    return txt + (" · falta: " + ", ".join(pend) if pend else " · pronto para iniciar")


ABA_RESUMO = "📋 Resumo"
ABAS_FICHA = ["1 · Entrevista", "2 · Formalização / CNPJ", "3 · Documentos", "4 · Contrato",
              "5 · Acessos TI", "6 · Integração", "7 · Início"]
# status: (ícone, cor da borda, cor de fundo, legenda)
CORES_STATUS = {
    "ok": ("✅", "#16A34A", "#ECFDF3", "Concluída"),
    "andamento": ("🟡", "#D97706", "#FFFBEB", "Em andamento"),
    "erro": ("🔴", "#DC2626", "#FEF2F2", "Reprovada / precisa de ajuste"),
    "todo": ("⚪", "#D1D5DB", "#FFFFFF", "Não iniciada"),
}


def status_etapas(r):
    """Status visual de cada etapa da ficha: ok | andamento | erro | todo."""
    feitas = etapas_concluidas(r)
    form = bool(r["form_enviado_em"])
    st1 = {"Aprovado": "ok", "Reprovado": "erro"}.get(r["status_entrevista"], "andamento")
    if feitas[2]:
        st2 = "ok"
    elif form and r["possui_cnpj"] == SIM and r["cnpj_apto"] != SIM:
        st2 = "erro"  # CNPJ fora das regras (MEI, CNAE ou situação)
    else:
        st2 = "andamento" if form else "todo"
    if r["doc_status"] == DOC_APROVADA:
        st3 = "ok"
    elif r["doc_status"] == "Reprovada":
        st3 = "erro"
    else:
        st3 = "andamento" if form or r["doc_status"] == DOC_APROVADA_PEND else "todo"
    st4 = "ok" if feitas[4] else ("andamento" if r["contrato_enviado"] == SIM else "todo")
    ti = [r[k] == SIM for k in ("ti_email", "ti_looker", "ti_propostas")]
    st5 = "ok" if all(ti) else ("andamento" if any(ti) else "todo")
    onb = [r[k] == SIM for k in ("onb_material", "onb_metas", "onb_cultura")] + [bool(r["padrinho"])]
    st6 = "ok" if all(onb) else ("andamento" if any(onb) else "todo")
    if r["status_final"] == STATUS_ATIVO:
        st7 = "ok"
    else:
        st7 = "andamento" if r["status_entrevista"] == "Aprovado" and all(feitas.values()) else "todo"
    return [st1, st2, st3, st4, st5, st6, st7]


def rotulos_abas(r):
    return [f"{CORES_STATUS[s][0]} {nome}" for s, nome in zip(status_etapas(r), ABAS_FICHA)]


def css_abas(r):
    regras = []
    for i, s in enumerate(status_etapas(r), start=1):
        _, borda, fundo, _ = CORES_STATUS[s]
        sel = f'.st-key-ficha_abas [data-testid="stTab"]:nth-of-type({i + 1})'
        regras.append(f"{sel} {{background: {fundo}; border-color: {borda}; box-shadow: inset 0 -3px 0 {borda};}}")
        regras.append(f'{sel}[aria-selected="true"] {{background: #4F46E5; border-color: {borda}; '
                      f"box-shadow: inset 0 -4px 0 {borda};}}")
    legenda = " · ".join(f"{ic} {txt}" for ic, _, _, txt in CORES_STATUS.values())
    return f"<style>{' '.join(regras)}</style>", legenda


def ir_para_proxima(pid, atual, avancar=None):
    """Depois de salvar, abre a próxima etapa pendente, ou volta ao quadro quando o candidato não está
    aprovado, já está ativo ou não sobrou nada.

    atual é o índice da aba (0 a 6). Por padrão só avança se a etapa atual ficou concluída;
    avancar=True força o avanço, avancar=False mantém na mesma aba.
    """
    r = buscar(pid)
    status = status_etapas(r)
    if avancar is None:
        avancar = status[atual] == "ok"
    if not avancar:
        st.session_state["ficha_ir"] = atual
        return ""
    if r["status_final"] == STATUS_ATIVO or r["status_entrevista"] != "Aprovado":
        st.session_state.pop("ficha", None)
        return " De volta ao quadro."
    ordem = list(range(atual + 1, 6)) + list(range(1, atual))  # etapas 2 a 6, a partir da seguinte
    for i in ordem:
        if status[i] != "ok":
            st.session_state["ficha_ir"] = i
            return f" Próxima: {ABAS_FICHA[i]}."
    st.session_state["ficha_ir"] = 6
    return f" Tudo pronto: {ABAS_FICHA[6]}."


def ficha_prestador(r):
    if st.button("← Voltar ao quadro"):
        st.session_state.pop("ficha", None)
        st.rerun()
    html(f'<div class="hero"><h2>{esc(r["nome"])}</h2>'
         f'<p>{esc(r["vaga"])} · {esc(r["email"])} · {esc(r["telefone"])} · '
         f'Código de acesso <code>{esc(r["codigo_acesso"])}</code></p>'
         f'<p><b>{situacao_ficha(r)}</b></p></div>')

    rotulos = [ABA_RESUMO] + rotulos_abas(r)
    # A aba aberta é guardada pela posição, porque o ícone do rótulo muda quando o status muda.
    # Posição 0 é o resumo; a etapa N (índice N-1 em ABAS_FICHA) fica na posição N.
    if st.session_state.get("ficha_abas_de") != r["id"]:  # ficha recém-aberta: começa no resumo
        idx = 0
    else:
        atual = st.session_state.get("ficha_abas") or ABA_RESUMO
        idx = next((i + 1 for i, nome in enumerate(ABAS_FICHA) if atual.endswith(nome)), 0)
    if "ficha_ir" in st.session_state:
        idx = st.session_state.pop("ficha_ir") + 1
    st.session_state["ficha_abas_de"] = r["id"]
    st.session_state["ficha_abas"] = rotulos[idx]
    estilo, legenda = css_abas(r)
    html(estilo)
    st.caption(legenda)
    abas = st.tabs(rotulos, key="ficha_abas", on_change="rerun")
    with abas[0]:
        aba_resumo(r)
    with abas[1]:
        aba_entrevista(r)
    with abas[2]:
        aba_formalizacao(r)
    with abas[3]:
        aba_documentos(r)
    with abas[4]:
        aba_contrato(r)
    with abas[5]:
        aba_acessos(r)
    with abas[6]:
        aba_integracao(r)
    with abas[7]:
        aba_inicio(r)


def checklist_etapas(r):
    """Itens de cada etapa para o resumo: (concluído, item, detalhe)."""
    def arq(v):
        return "anexado" if v else ""
    acessos = [("ti_email", "E-mail corporativo"), ("ti_looker", "Looker"), ("ti_propostas", "Sistema de propostas")]
    materiais = [("onb_material", "Material de onboarding"), ("onb_metas", "Metas e comissionamento"),
                 ("onb_cultura", "Cultura e código de conduta")]
    cnpj_txt = fmt_cnpj(r["cnpj"]) if r["possui_cnpj"] == SIM else ("não possui" if r["possui_cnpj"] == "Não" else "")
    return [
        [(r["status_entrevista"] == "Aprovado", "Entrevista aprovada", r["status_entrevista"]),
         (bool(r["curriculo"]), "Currículo", arq(r["curriculo"])),
         (bool(r["avaliacao_nota"]), "Avaliação de perfil", f"nota {r['avaliacao_nota']}" if r["avaliacao_nota"] else "")],
        [(bool(r["form_enviado_em"]), "Formulário enviado", data_br(r["form_enviado_em"])),
         (bool(r["cpf"]) and bool(r["endereco"]), "Dados pessoais e endereço", r["cidade"] and f"{r['cidade']}/{r['uf']}"),
         (r["possui_cnpj"] == SIM, "CNPJ", cnpj_txt),
         (r["possui_cnpj"] == SIM and not cnpj_pendente(r), "CNPJ dentro das regras",
          "MEI: " + (r["is_mei"] or "—") + f" · CNAE {fmt_cnae(CNAE_EXIGIDO)}: " + (r["tem_cnae"] or "—")
          if r["possui_cnpj"] == SIM else ""),
         (bool(r["cartao_cnpj"]), "Cartão CNPJ", arq(r["cartao_cnpj"]))],
        [(r["doc_status"] in DOCS_LIBERADAS, "Documentação conferida pelo RH", r["doc_status"] or "Pendente"),
         (r["doc_status"] == DOC_APROVADA, "Aprovação definitiva", r["doc_obs"])],
        [(r["contrato_enviado"] == SIM, "Contrato enviado", ""),
         (r["contrato_assinado"] == SIM, "Contrato assinado", data_br(r["contrato_data"])),
         (bool(r["contrato_arquivo"]), "Contrato assinado anexado", arq(r["contrato_arquivo"]))],
        [(r[c] == SIM, nome, r[f"{c}_usuario"]) for c, nome in acessos],
        [(r[c] == SIM, nome, "apresentado" if r[c] == SIM else "") for c, nome in materiais]
        + [(bool(r["data_integracao"]), "Reunião de integração",
            date.fromisoformat(r["data_integracao"]).strftime("%d/%m/%Y") if r["data_integracao"] else ""),
           (bool(r["padrinho"]), "Padrinho", r["padrinho"])],
        [(r["status_final"] == STATUS_ATIVO, STATUS_ATIVO, data_br(r["operacao_iniciada_em"]))],
    ]


def onepage_html(r):
    """Resumo do candidato em uma página, com estilos inline para colar ou anexar em e-mail."""
    status = status_etapas(r)
    itens = checklist_etapas(r)
    feitos = sum(ok for etapa in itens for ok, _, _ in etapa)
    total = sum(len(etapa) for etapa in itens)
    etapas_ok = sum(s == "ok" for s in status)
    pct = round(100 * feitos / total)
    fonte = "font-family:Arial,Helvetica,sans-serif;"
    blocos = []
    for nome, s, lista in zip(ABAS_FICHA, status, itens):
        icone, borda, fundo, legenda = CORES_STATUS[s]
        linhas = "".join(
            f'<tr><td style="width:26px;padding:5px 0;vertical-align:top;font-size:15px">{"✅" if ok else "⬜"}</td>'
            f'<td style="padding:5px 0;font-size:14px;color:{"#111827" if ok else "#6B7280"}">{esc(item)}</td>'
            f'<td style="padding:5px 0;font-size:13px;color:#6B7280;text-align:right">{esc(det)}</td></tr>'
            for ok, item, det in lista)
        blocos.append(
            f'<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:0 0 12px;'
            f'border:1px solid {borda};border-left:6px solid {borda};border-radius:8px">'
            f'<tr><td style="background:{fundo};padding:10px 14px">'
            f'<span style="font-size:15px;font-weight:bold;color:#111827">{icone} {esc(nome)}</span>'
            f'<span style="float:right;font-size:12px;color:#374151;padding-top:2px">{legenda}</span></td></tr>'
            f'<tr><td style="padding:4px 14px 8px"><table width="100%" cellpadding="0" cellspacing="0">{linhas}'
            f"</table></td></tr></table>")
    dados = [("Vaga", r["vaga"]), ("E-mail", r["email"]), ("Telefone", r["telefone"]),
             ("Código de acesso", r["codigo_acesso"]), ("Situação", r["status_final"] or "Em onboarding")]
    cab = "".join(f'<td style="padding:0 18px 0 0;font-size:12px;color:#E0E7FF">{esc(k)}<br>'
                  f'<b style="font-size:14px;color:#fff">{esc(v) or "—"}</b></td>' for k, v in dados)
    return (
        f'<div style="{fonte}background:#F3F4F6;padding:20px 0">'
        f'<table width="680" align="center" cellpadding="0" cellspacing="0" style="max-width:680px;width:100%;'
        f'background:#fff;border-radius:12px;border-collapse:separate">'
        f'<tr><td style="background:#4F46E5;padding:22px 24px;border-radius:12px 12px 0 0">'
        f'<div style="font-size:12px;color:#C7D2FE;letter-spacing:.06em">RESUMO DE ONBOARDING</div>'
        f'<div style="font-size:24px;font-weight:bold;color:#fff;margin:4px 0 14px">{esc(r["nome"])}</div>'
        f'<table cellpadding="0" cellspacing="0"><tr>{cab}</tr></table></td></tr>'
        f'<tr><td style="padding:18px 24px 6px">'
        f'<div style="font-size:14px;color:#111827"><b>{etapas_ok} de 7 etapas concluídas</b> · '
        f"{feitos} de {total} itens do checklist ({pct}%)</div>"
        f'<div style="background:#E5E7EB;height:8px;border-radius:4px;margin:8px 0 6px">'
        f'<div style="background:#16A34A;height:8px;width:{pct}%;border-radius:4px"></div></div>'
        f'<div style="font-size:12px;color:#6B7280">'
        + " · ".join(f"{ic} {txt}" for ic, _, _, txt in CORES_STATUS.values()) +
        f'</div></td></tr><tr><td style="padding:12px 24px 8px">{"".join(blocos)}</td></tr>'
        f'<tr><td style="padding:4px 24px 20px;font-size:11px;color:#9CA3AF">'
        f"Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')} pelo sistema de Onboarding de Prestadores."
        f"</td></tr></table></div>"
    )


def aba_resumo(r):
    pagina = onepage_html(r)
    c1, c2 = st.columns([1, 2], vertical_alignment="center")
    arquivo = (f'<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>Resumo · {esc(r["nome"])}'
               f"</title></head><body style=\"margin:0\">{pagina}</body></html>")
    c1.download_button("📧 Baixar one page para e-mail", arquivo.encode("utf-8"),
                       f"resumo_{slug(r['nome'])}.html", "text/html", key=f"dl_resumo_{r['id']}", type="primary")
    c2.caption("Abra o arquivo no navegador, selecione tudo (Ctrl+A), copie e cole no corpo do e-mail. "
               "Também dá para anexar o arquivo ou salvar como PDF (Ctrl+P).")
    st.html(pagina)


def aba_entrevista(r):
    c0, c1, c2 = st.columns([1, 1, 2])
    opcoes_vaga = VAGAS + ([r["vaga"]] if r["vaga"] and r["vaga"] not in VAGAS else [])
    nova_vaga = c0.selectbox("Vaga / Cargo", opcoes_vaga, index=opcoes_vaga.index(r["vaga"]) if r["vaga"] else 0,
                             key=f"vaga_{r['id']}")
    novo_status = c1.selectbox("Status da entrevista", STATUS_ENTREVISTA,
                               index=STATUS_ENTREVISTA.index(r["status_entrevista"])
                               if r["status_entrevista"] in STATUS_ENTREVISTA else 0, key=f"st_{r['id']}")
    obs = c2.text_input("Observações da entrevista", r["obs_entrevista"], key=f"obs_{r['id']}")
    if st.button("Salvar", key=f"salva_st_{r['id']}", type="primary"):
        atualizar(r["id"], vaga=nova_vaga, status_entrevista=novo_status, obs_entrevista=obs)
        avisar("Dados da entrevista salvos." + ir_para_proxima(r["id"], 0, True))
        st.rerun()

    st.markdown("##### Currículo")
    c1, c2 = st.columns([1, 2])
    with c1:
        mostrar_arquivo(r["curriculo"], "Currículo", key=f"dl_cv_{r['id']}")
    with c2:
        novo_cv = st.file_uploader("Anexar ou trocar currículo", type=["pdf", "doc", "docx"], key=f"up_cv_{r['id']}")
        if novo_cv is not None and st.button("Salvar currículo", key=f"salva_cv_{r['id']}"):
            atualizar(r["id"], curriculo=guardar_arquivo(f"{r['id']}_cv", novo_cv))
            avisar("Currículo salvo." + ir_para_proxima(r["id"], 0, False))
            st.rerun()

    avaliacao_candidato(r)


def avaliacao_candidato(r):
    st.markdown("##### Avaliação de perfil")
    st.caption("Cole a nota e o resumo da avaliação feita pelo Claude no projeto e anexe o arquivo completo.")
    c1, c2 = st.columns([1, 3])
    try:
        nota_atual = float(r["avaliacao_nota"].replace(",", ".")) if r["avaliacao_nota"] else None
    except ValueError:
        nota_atual = None
    nota = c1.number_input("Nota geral (0 a 10)", 0.0, 10.0, value=nota_atual, step=0.1, format="%.1f",
                           key=f"nota_{r['id']}")
    resumo = c2.text_area("Resumo executivo", r["avaliacao_resumo"], height=110, key=f"res_{r['id']}")
    arq = st.file_uploader("Avaliação completa (.md, .pdf ou .docx)", type=["md", "pdf", "docx"],
                           key=f"up_av_{r['id']}")
    if st.button("Salvar avaliação", key=f"salva_av_{r['id']}", type="primary"):
        campos = dict(avaliacao_nota="" if nota is None else f"{nota:.1f}", avaliacao_resumo=resumo.strip())
        if arq is not None:
            campos["avaliacao"] = guardar_arquivo(f"{r['id']}_avaliacao", arq)
        atualizar(r["id"], **campos)
        avisar("Avaliação salva." + ir_para_proxima(r["id"], 0, False))
        st.rerun()
    mostrar_arquivo(r["avaliacao"], "Avaliação completa", key=f"dl_av_{r['id']}")
    local = caminho_local(r["avaliacao"])
    if local and local.suffix == ".md" and local.exists():
        with st.expander("Ler avaliação completa"):
            st.markdown(local.read_text(encoding="utf-8", errors="ignore"))


def aba_formalizacao(r):
    if r["status_entrevista"] != "Aprovado":
        st.info("O candidato ainda não foi aprovado na entrevista. Você já pode adiantar a formalização, se quiser.")
    if r["form_enviado_em"]:
        situ = "✅ CNPJ regular" if not cnpj_pendente(r) else (
            "⚠️ Sem CNPJ" if r["possui_cnpj"] == "Não" else
            "🔎 CNPJ declarado, conferir na Etapa 3" if cnpj_declarado(r) else "⚠️ CNPJ fora das regras")
        (st.warning if cnpj_pendente(r) else st.success)(f"Formulário enviado em {data_br(r['form_enviado_em'])} · {situ}")
    else:
        st.warning("O candidato ainda não preencheu o formulário.")

    st.markdown("##### Enviar acesso ao candidato")
    st.caption("Copie a mensagem abaixo e mande por WhatsApp ou e-mail.")
    st.code(mensagem_acesso(r), language=None)

    with st.expander("✏️ Preencher ou corrigir a Etapa 2 pelo candidato", expanded=not r["form_enviado_em"]):
        formulario_candidato(r, prefixo="rh")


def validacao_manual_cnpj(r):
    """CNPJs recém-abertos não aparecem nas bases públicas: o RH confere no cartão e valida aqui."""
    with st.container(border=True):
        st.markdown("**🔎 Validar CNPJ pelo cartão**")
        st.caption("As bases públicas ainda não tinham este CNPJ (comum em empresas abertas há poucas semanas), "
                   "então o candidato declarou os dados. Abra o cartão CNPJ acima, ou consulte no site da "
                   "Receita Federal, e confirme.")
        st.link_button("Abrir consulta na Receita Federal",
                       "https://solucoes.receita.fazenda.gov.br/servicos/cnpjreva/cnpjreva_solicitacao.asp")
        ativa = st.checkbox("Situação cadastral ATIVA", key=f"mv_ativa_{r['id']}")
        if aceita_mei(r["vaga"]):
            nao_mei = True
            st.caption(f"Para a vaga {r['vaga']} o MEI é aceito, então não é preciso conferir o enquadramento.")
        else:
            nao_mei = st.checkbox("Não é MEI (confira no Portal do Simples Nacional: "
                                  "o cartão CNPJ de MEI e de ME podem ser iguais)", key=f"mv_mei_{r['id']}")
        cnae = st.checkbox(f"Possui o CNAE {fmt_cnae(CNAE_EXIGIDO)} (principal ou secundário)", key=f"mv_cnae_{r['id']}")
        if st.button("✅ Confirmar validação do CNPJ", key=f"mv_ok_{r['id']}",
                     disabled=not (ativa and nao_mei and cnae)):
            atualizar(r["id"], is_mei=r["is_mei"] if aceita_mei(r["vaga"]) else "Não", tem_cnae=SIM, situacao_cnpj="ATIVA", cnpj_apto=SIM,
                      fonte_validacao=f"Conferido pelo RH no cartão CNPJ em {data_br(agora())}")
            avisar("CNPJ validado pelo RH." + ir_para_proxima(r["id"], 2, False))
            st.rerun()


def aba_documentos(r):
    if not r["form_enviado_em"]:
        st.warning("Ainda não há dados da Etapa 2 para conferir.")
        return
    c1, c2 = st.columns(2)
    with c1, st.container(border=True):
        st.markdown("**Dados pessoais**")
        html(campos_html({
            "Nome completo": r["nome_completo"], "CPF": r["cpf"], "RG": r["rg"],
            "Nascimento": r["data_nascimento"], "Telefone": r["telefone"], "E-mail": r["email"],
            "Estado civil": r["estado_civil"], "Endereço": r["endereco"],
        }))
    with c2, st.container(border=True):
        st.markdown("**Pessoa jurídica**")
        html(campos_html({
            "Possui CNPJ": r["possui_cnpj"], "CNPJ": fmt_cnpj(r["cnpj"]), "Razão social": r["razao_social"],
            "Natureza jurídica": r["natureza_juridica"], "Situação": r["situacao_cnpj"],
            "MEI": r["is_mei"], f"CNAE {fmt_cnae(CNAE_EXIGIDO)}": r["tem_cnae"],
            "CNAE principal": r["cnae_principal"], "CNAEs secundários": r["cnaes_secundarios"],
            "Validação": r["fonte_validacao"], "Apto": r["cnpj_apto"],
        }))
        mostrar_arquivo(r["cartao_cnpj"], "Cartão CNPJ", key=f"dl_cartao_{r['id']}")
    if cnpj_declarado(r):
        validacao_manual_cnpj(r)
    st.write(f"Status da documentação: **{r['doc_status'] or 'Pendente'}**")
    obs = st.text_input("Observação para o candidato (obrigatória ao reprovar)", r["doc_obs"], key=f"docobs_{r['id']}")
    b1, b2 = st.columns(2)
    if cnpj_pendente(r):
        if r["possui_cnpj"] == "Não":
            motivo = "não tem CNPJ"
        elif r["cnpj_apto"] == SIM:
            motivo = "tem CNPJ ainda não conferido pelo RH (use o bloco acima)"
        else:
            motivo = "tem CNPJ fora das regras (MEI, CNAE ou situação)"
        st.warning(f"O candidato {motivo}. Você pode **aprovar com o CNPJ pendente** e seguir com as outras etapas. "
                   "Quando o CNPJ for regularizado, volte aqui para a aprovação definitiva.")
        if r["doc_status"] != DOC_APROVADA_PEND and b1.button(
                "⏭️ Aprovar com CNPJ pendente", key=f"aprova_pend_{r['id']}", type="primary"):
            atualizar(r["id"], doc_status=DOC_APROVADA_PEND, doc_obs=obs, doc_validado_em=agora())
            avisar("Documentação aprovada com CNPJ pendente." + ir_para_proxima(r["id"], 2, True))
            st.rerun()
    else:
        if r["doc_status"] == DOC_APROVADA_PEND:
            st.info("O CNPJ foi regularizado. Confira o cartão e aprove a documentação definitiva.")
        if r["doc_status"] != DOC_APROVADA and b1.button(
                "✅ Aprovar documentação", key=f"aprova_{r['id']}", type="primary"):
            atualizar(r["id"], doc_status=DOC_APROVADA, doc_obs=obs, doc_validado_em=agora())
            avisar("Documentação aprovada." + ir_para_proxima(r["id"], 2, True))
            st.rerun()
    if r["doc_status"] not in DOCS_LIBERADAS and b2.button("❌ Reprovar / pedir ajuste", key=f"reprova_{r['id']}"):
        if not obs.strip():
            st.error("Descreva o ajuste necessário.")
        else:
            atualizar(r["id"], doc_status="Reprovada", doc_obs=obs, doc_validado_em=agora())
            avisar("Ajuste solicitado ao candidato." + ir_para_proxima(r["id"], 2, True), "📝")
            st.rerun()


def aba_contrato(r):
    enviado = ck(r, "contrato_enviado", "Contrato de parceria enviado para assinatura", "ce")
    assinado = ck(r, "contrato_assinado", "Contrato assinado por ambas as partes", "ca")
    c1, c2 = st.columns([2, 1])
    arq = c1.file_uploader("Contrato assinado (PDF ou imagem)", type=["pdf", "png", "jpg", "jpeg"],
                           key=f"up_ct_{r['id']}_{r['atualizado_em']}")
    with c2:
        mostrar_arquivo(r["contrato_arquivo"], "Contrato assinado", key=f"dl_ct_{r['id']}")
    if arq is not None:  # anexar o contrato assinado já marca envio e assinatura
        enviado = assinado = True
    if st.button("Salvar contrato", key=f"salva_ct_{r['id']}", type="primary"):
        if assinado and not enviado:
            st.error("Marque o envio antes da assinatura.")
        elif assinado and arq is None and not r["contrato_arquivo"]:
            st.error("Anexe o contrato assinado para registrar a assinatura.")
        else:
            campos = dict(contrato_enviado=sn(enviado), contrato_assinado=sn(assinado),
                          contrato_data=agora() if assinado and not r["contrato_data"] else r["contrato_data"])
            if arq is not None:
                campos["contrato_arquivo"] = guardar_arquivo(f"{r['id']}_contrato", arq)
            atualizar(r["id"], **campos)
            avisar("Contrato salvo." + ir_para_proxima(r["id"], 3, True))
            st.rerun()
    if r["contrato_data"]:
        st.caption(f"Assinatura registrada em {data_br(r['contrato_data'])}.")


def aba_acessos(r):
    acessos = [
        ("ti_email", "Criação de e-mail corporativo", "E-mail corporativo", "nome@empresa.com.br"),
        ("ti_looker", "Acesso ao sistema de gestão Looker", "Usuário Looker", ""),
        ("ti_propostas", "Acesso ao sistema de digitação de propostas", "Usuário do sistema de propostas", ""),
    ]
    campos, faltando = {}, []
    for campo, rotulo, rotulo_usuario, exemplo in acessos:
        c1, c2 = st.columns([1, 1], vertical_alignment="bottom")
        with c1:
            liberado = ck(r, campo, rotulo, campo)
        usuario = c2.text_input(rotulo_usuario, r[f"{campo}_usuario"], placeholder=exemplo,
                                key=f"{campo}_usuario_{r['id']}").strip()
        liberado = liberado or bool(usuario)  # informar o usuário já marca o acesso como liberado
        if liberado and not usuario:
            faltando.append(rotulo_usuario)
        campos[campo], campos[f"{campo}_usuario"] = sn(liberado), usuario
    if st.button("Salvar acessos", key=f"salva_ti_{r['id']}", type="primary"):
        if faltando:
            st.error("Informe o usuário de: " + ", ".join(faltando) + ".")
            return
        atualizar(r["id"], **campos)
        avisar("Acessos salvos." + ir_para_proxima(r["id"], 4, True))
        st.rerun()


def aba_integracao(r):
    links = secret("links", {}) or {}
    materiais = [
        ("onb_material", "Material de Onboarding (manuais e processos)", "material_onboarding"),
        ("onb_metas", "Detalhamento de metas e comissionamento", "metas_comissionamento"),
        ("onb_cultura", "Cultura organizacional e código de conduta", "cultura_conduta"),
    ]
    marcados = {}
    for campo, rotulo, chave_link in materiais:
        c1, c2 = st.columns([3, 1])
        with c1:
            marcados[campo] = ck(r, campo, f"{rotulo} — apresentado", campo)
        url = links.get(chave_link) if hasattr(links, "get") else None
        if url:
            c2.link_button("Abrir", url)
        else:
            c2.caption(f"Configure `links.{chave_link}` em secrets")
    try:
        di = date.fromisoformat(r["data_integracao"]) if r["data_integracao"] else date.today()
    except ValueError:
        di = date.today()
    c1, c2 = st.columns(2)
    data_int = c1.date_input("Data da reunião de integração", di, format="DD/MM/YYYY", key=f"di_{r['id']}")
    padrinho = c2.text_input("Padrinho responsável pelo primeiro mês *", r["padrinho"], key=f"pad_{r['id']}")
    if st.button("Salvar integração", key=f"salva_onb_{r['id']}", type="primary"):
        atualizar(r["id"], data_integracao=data_int.isoformat(), padrinho=padrinho.strip(),
                  **{k: sn(v) for k, v in marcados.items()})
        avisar("Integração salva." + ir_para_proxima(r["id"], 5, True))
        st.rerun()


def aba_inicio(r):
    if r["status_final"] == STATUS_ATIVO:
        st.success(f"🎉 {STATUS_ATIVO} desde {r['operacao_iniciada_em']}.")
        return
    pend = pendencias(r)
    if not cnpj_pendente(r) and r["doc_status"] == DOC_APROVADA_PEND:
        pend.append("Aprovação definitiva dos documentos")
    if pend:
        st.warning("Ainda há pendências: **" + ", ".join(pend) + "**.")
        forcar = st.checkbox("Iniciar a operação mesmo com essas pendências", key=f"forca_{r['id']}")
    else:
        st.write("Todas as etapas foram concluídas. Confirme o início das operações.")
        forcar = True
    if st.button("🚀 Iniciar operação", key=f"ativa_{r['id']}", type="primary", disabled=not forcar):
        atualizar(r["id"], status_final=STATUS_ATIVO, operacao_iniciada_em=agora())
        avisar("Operação iniciada! Prestador ativo." + ir_para_proxima(r["id"], 6, True), "🚀")
        st.rerun()


# ---------------------------------------------------------------------------
# Comparativo por vaga
# ---------------------------------------------------------------------------
def slug(txt):
    return re.sub(r"[^a-z0-9]+", "-", txt.lower()).strip("-") or "vaga"


def comparativo():
    df = carregar()
    if df.empty:
        st.info("Nenhum candidato cadastrado ainda.")
        return
    vagas = sorted(v for v in df["vaga"].unique() if v)
    vaga = st.selectbox("Vaga", vagas, key="comp_vaga")
    d = df[df["vaga"] == vaga].copy()
    d["nota"] = pd.to_numeric(d["avaliacao_nota"].str.replace(",", "."), errors="coerce")
    d = d.sort_values("nota", ascending=False, na_position="last")
    d.insert(0, "posicao", range(1, len(d) + 1))
    d["Currículo"] = d["curriculo"].map(lambda v: "Sim" if v else "")

    st.markdown(f"##### Ranking · {vaga}")
    st.caption(f"{d['nota'].notna().sum()} de {len(d)} candidatos com avaliação.")
    st.dataframe(
        d[["posicao", "nome", "nota", "status_entrevista", "avaliacao_resumo", "Currículo"]].rename(columns={
            "posicao": "#", "nome": "Nome", "nota": "Nota", "status_entrevista": "Status",
            "avaliacao_resumo": "Resumo executivo"}),
        hide_index=True, width="stretch",
        column_config={"Nota": st.column_config.ProgressColumn("Nota", min_value=0, max_value=10, format="%.1f")},
    )

    st.markdown("##### Comparativo detalhado")
    base = f"comparativo_{slug(vaga)}"
    existentes = sorted(UPLOAD_DIR.glob(base + ".*")) if UPLOAD_DIR.exists() else []
    if existentes:
        arq = existentes[-1]
        st.download_button(f"📎 Baixar {arq.name}", arq.read_bytes(), arq.name, key="dl_comp")
        try:
            if arq.suffix == ".xlsx":
                st.dataframe(pd.read_excel(arq), hide_index=True, width="stretch")
            elif arq.suffix == ".csv":
                st.dataframe(pd.read_csv(arq), hide_index=True, width="stretch")
            elif arq.suffix == ".md":
                with st.expander("Ler comparativo", expanded=True):
                    st.markdown(arq.read_text(encoding="utf-8", errors="ignore"))
        except Exception as e:
            st.warning(f"Não consegui exibir o arquivo ({e}). Use o botão de download.")
    else:
        st.caption("Nenhum comparativo detalhado anexado para esta vaga.")
    novo = st.file_uploader("Anexar ou trocar comparativo (.xlsx, .csv ou .md)", type=["xlsx", "csv", "md"],
                            key=f"up_comp_{slug(vaga)}")
    if novo is not None and st.button("Salvar comparativo", key="salva_comp"):
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        for antigo in existentes:
            antigo.unlink()
        (UPLOAD_DIR / (base + Path(novo.name).suffix.lower())).write_bytes(novo.getvalue())
        avisar("Comparativo salvo.")
        st.rerun()


# =============================================================================
# NAVEGAÇÃO
# =============================================================================
html(CSS)
mostrar_aviso()
# O acesso padrão é o do prestador; o RH entra pelo botão discreto do canto ou pelo link ?rh=1.
if "modo" not in st.session_state:
    st.session_state["modo"] = "rh" if st.query_params.get("rh") else "candidato"

if st.session_state["modo"] == "candidato":
    html(CSS_CANDIDATO)
    tela_candidato()
else:
    if st.session_state.get("rh_ok"):
        with st.sidebar:
            st.markdown("### 🤝 Onboarding de Prestadores")
            st.caption("Armazenamento: " + ("Google Sheets ✅" if usa_sheets() else "CSV local (demonstração)"))
            if st.button("Sair do painel RH", width="stretch"):
                for chave in ("rh_ok", "ficha"):
                    st.session_state.pop(chave, None)
                st.session_state["modo"] = "candidato"
                st.rerun()
    else:
        html(CSS_CANDIDATO)
    tela_rh()

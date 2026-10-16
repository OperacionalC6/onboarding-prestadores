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
STATUS_ENTREVISTA = ["Em Entrevista", "Aprovado", "Banco de Espera", "Reprovado"]
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
    "possui_cnpj", "cnpj", "razao_social", "natureza_juridica", "situacao_cnpj",
    "is_mei", "tem_cnae", "cnae_principal", "cnaes_secundarios", "fonte_validacao",
    "cnpj_apto", "cartao_cnpj", "form_enviado_em",
    # Etapa 3
    "doc_status", "doc_obs", "doc_validado_em",
    # Etapa 4
    "contrato_enviado", "contrato_assinado", "contrato_data",
    # Etapa 5
    "ti_email", "ti_looker", "ti_propostas",
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
    """Salva um anexo (cartão CNPJ, currículo) no Google Drive (se configurado) ou em disco local."""
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


@st.cache_data(ttl=3600, show_spinner=False)
def consultar_cnpj(cnpj):
    """Retorna dict com status: ok | nao_encontrado | indisponivel."""
    c = so_digitos(cnpj)
    fontes = [
        (f"https://brasilapi.com.br/api/cnpj/v1/{c}", _parse_brasilapi),
        (f"https://publica.cnpj.ws/cnpj/{c}", _parse_cnpjws),
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
        except Exception:
            continue
    return {"status": "nao_encontrado" if nao_encontrado else "indisponivel"}


def avaliar_regras(is_mei, tem_cnae, situacao="ATIVA"):
    """Aplica as regras de negócio. Retorna (apto, lista_de_pendências)."""
    pendencias = []
    if situacao and situacao != "ATIVA":
        pendencias.append(f"O CNPJ está com situação cadastral **{situacao}**. Ele precisa estar **ATIVO**.")
    if is_mei:
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


ORIENTACOES_SEM_CNPJ = """
**Importante: a parceria NÃO aceita MEI (Microempreendedor Individual).**
O MEI não permite a atividade de promoção de vendas nos moldes da parceria e tem limite de faturamento baixo.

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
def cnpj_pendente(r):
    return r["cnpj_apto"] != SIM


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


def barra_progresso(r):
    if fase(r) == 8:
        st.progress(1.0, text=f"**{STATUS_ATIVO}**")
        return
    feitas = etapas_concluidas(r)
    n = sum(feitas.values())
    marcas = " · ".join(f"{'✅' if ok else '⬜'} {NOMES_CURTOS[k]}" for k, ok in feitas.items())
    st.progress(n / len(feitas), text=f"**{n} de {len(feitas)} etapas concluídas**  ·  {marcas}")


# =============================================================================
# VISÃO DO CANDIDATO
# =============================================================================
def tela_candidato():
    st.title("🤝 Formulário do Prestador")

    pid = st.session_state.get("cand_id")
    if not pid:
        st.write("Informe o **código de acesso** que você recebeu do RH e o seu e-mail.")
        with st.form("login_cand"):
            codigo = st.text_input("Código de acesso").strip().upper()
            email = st.text_input("E-mail").strip().lower()
            if st.form_submit_button("Acessar", type="primary"):
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

    topo1, topo2 = st.columns([4, 1])
    topo1.subheader(f"Olá, {r['nome'].split(' ')[0]}!")
    if topo2.button("Sair"):
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
            st.write({k: r[k] for k in ("nome_completo", "cpf", "email", "telefone", "cnpj", "razao_social")})
        return

    if r["doc_status"] == "Reprovada":
        st.error(f"O RH pediu ajustes na sua documentação: {r['doc_obs'] or 'veja com o RH.'}")
    elif r["form_enviado_em"]:
        st.warning("Recebemos seus dados pessoais. Falta regularizar o CNPJ para seguir. Atualize abaixo quando estiver pronto.")

    formulario_candidato(r)


def formulario_candidato(r):
    pid = r["id"]
    st.markdown("### Etapa 2 · Dados pessoais")
    c1, c2 = st.columns(2)
    nome_completo = c1.text_input("Nome completo *", r["nome_completo"] or r["nome"])
    cpf = c2.text_input("CPF *", r["cpf"], placeholder="000.000.000-00")
    rg = c1.text_input("RG *", r["rg"])
    try:
        dn_padrao = date.fromisoformat(r["data_nascimento"]) if r["data_nascimento"] else date(1990, 1, 1)
    except ValueError:
        dn_padrao = date(1990, 1, 1)
    nasc = c2.date_input("Data de nascimento *", dn_padrao, min_value=date(1940, 1, 1),
                         max_value=date.today(), format="DD/MM/YYYY")
    telefone = c1.text_input("Telefone *", r["telefone"])
    email = c2.text_input("E-mail *", r["email"])
    ec_idx = ESTADOS_CIVIS.index(r["estado_civil"]) if r["estado_civil"] in ESTADOS_CIVIS else 0
    estado_civil = c1.selectbox("Estado civil *", ESTADOS_CIVIS, index=ec_idx)
    endereco = st.text_area("Endereço completo (rua, nº, complemento, bairro, cidade/UF, CEP) *", r["endereco"])

    st.markdown("### Etapa 2 · Pessoa jurídica")
    possui = st.radio("Você já possui CNPJ?", ["Sim", "Não"], horizontal=True,
                      index=0 if r["possui_cnpj"] != "Não" else 1)

    cnpj_info = {}
    arquivo = None
    apto = False

    if possui == "Não":
        st.info(ORIENTACOES_SEM_CNPJ)
        cnpj_info = {"possui_cnpj": "Não", "cnpj_apto": ""}
    else:
        cnpj = st.text_input("CNPJ *", fmt_cnpj(r["cnpj"]), placeholder="00.000.000/0000-00")
        chave = f"cnpj_res_{pid}"
        if st.button("🔎 Validar CNPJ"):
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
        elif res["status"] == "nao_encontrado":
            st.error("CNPJ não encontrado na base da Receita Federal.")
        else:
            if res["status"] == "ok":
                cnaes = [res["cnae_principal"]] + res["cnaes_secundarios"]
                tem_cnae = CNAE_EXIGIDO in [so_digitos(x).zfill(7) for x in cnaes if x]
                is_mei, situacao, fonte = res["is_mei"], res["situacao"], f"Automática ({res['fonte']})"
                st.write(f"**{res['razao_social']}** · {res['natureza_juridica']} · Situação: {situacao}")
                col_a, col_b = st.columns(2)
                col_a.metric("Enquadrado como MEI?", "Sim ❌" if is_mei else "Não ✅")
                col_b.metric(f"CNAE {fmt_cnae(CNAE_EXIGIDO)}?", "Sim ✅" if tem_cnae else "Não ❌")
            else:  # indisponivel -> declaração manual
                st.warning("Não conseguimos consultar a Receita agora. Responda abaixo; o RH vai conferir no cartão CNPJ.")
                is_mei = st.radio("Sua empresa é MEI?", ["Sim", "Não"], horizontal=True, index=1) == "Sim"
                tem_cnae = st.radio(f"Seu CNPJ possui o CNAE {fmt_cnae(CNAE_EXIGIDO)} (Promoção de vendas)?",
                                    ["Sim", "Não"], horizontal=True) == "Sim"
                situacao, fonte = "", "Declaração do candidato (API indisponível)"
                res = {**res, "razao_social": "", "natureza_juridica": "", "cnae_principal": "", "cnaes_secundarios": []}

            apto, pendencias = avaliar_regras(is_mei, tem_cnae, situacao)
            if apto:
                st.success("Seu CNPJ atende aos requisitos. Anexe o Cartão CNPJ para concluir.")
                arquivo = st.file_uploader("Cartão CNPJ (PDF ou imagem) *", type=["pdf", "png", "jpg", "jpeg"])
            else:
                st.error("Seu CNPJ precisa de regularização antes de seguir:")
                for p in pendencias:
                    st.markdown(f"- {p}")
                st.caption("Você pode enviar seus dados pessoais agora e voltar com o CNPJ regularizado.")

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
    if st.button(rotulo, type="primary"):
        erros = []
        if not all([nome_completo.strip(), rg.strip(), telefone.strip(), email.strip(), endereco.strip()]):
            erros.append("Preencha todos os campos obrigatórios (*).")
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
            estado_civil=estado_civil, endereco=endereco.strip(), form_enviado_em=agora(), **cnpj_info,
        )
        if arquivo is not None:
            campos["cartao_cnpj"] = guardar_arquivo(pid, arquivo)
        if apto and r["doc_status"] != DOC_APROVADA_PEND:
            campos.update(doc_status="Pendente", doc_obs="")
        atualizar(pid, **campos)
        st.success("Dados enviados com sucesso!")
        st.rerun()


# =============================================================================
# PAINEL DO RH
# =============================================================================
def login_rh():
    if st.session_state.get("rh_ok"):
        return True
    senha_cfg = secret("rh_password", "rh123")
    st.title("🔒 Painel de Controle · RH")
    if senha_cfg == "rh123":
        st.caption("Senha de demonstração: **rh123** (defina `rh_password` em secrets para trocar).")
    with st.form("login_rh"):
        senha = st.text_input("Senha", type="password")
        if st.form_submit_button("Entrar", type="primary"):
            if senha == senha_cfg:
                st.session_state["rh_ok"] = True
                st.rerun()
            st.error("Senha incorreta.")
    return False


def tela_rh():
    if not login_rh():
        return
    st.title("📋 Painel de Controle · RH")
    aba_geral, aba_entrevistas, aba_comp, aba_prestador = st.tabs(
        ["📊 Visão geral", "🗣️ Etapa 1 · Entrevistas", "📈 Comparativo", "🧾 Etapas 3 a 7 · Prestador"]
    )
    with aba_geral:
        visao_geral()
    with aba_entrevistas:
        etapa_entrevistas()
    with aba_comp:
        comparativo()
    with aba_prestador:
        gestao_prestador()


def visao_geral():
    df = carregar()
    if df.empty:
        st.info("Nenhum candidato cadastrado ainda. Comece pela aba **Etapa 1 · Entrevistas**.")
        return
    df["etapa"] = df.apply(rotulo_etapa, axis=1)
    df["n_etapa"] = df.apply(fase, axis=1)
    df["situacao_cnpj_pj"] = df.apply(
        lambda r: "" if not r["form_enviado_em"] else ("Regular" if r["cnpj_apto"] == SIM else "Pendente"), axis=1)
    cols = st.columns(5)
    cols[0].metric("Candidatos", len(df))
    cols[1].metric("Em entrevista", int((df["status_entrevista"] == "Em Entrevista").sum()))
    cols[2].metric("Em onboarding", int((df["n_etapa"] == 2).sum()))
    cols[3].metric("CNPJ pendente", int(((df["situacao_cnpj_pj"] == "Pendente") & (df["n_etapa"] < 8)).sum()))
    cols[4].metric("Ativos", int((df["status_final"] == STATUS_ATIVO).sum()))

    df["pendentes"] = df.apply(lambda r: pendencias(r) if fase(r) == 2 else [], axis=1)
    filtro = st.multiselect("Filtrar prestadores com pendência em", list(NOMES_CURTOS.values()))
    vis = df[df["pendentes"].apply(lambda p: all(f in p for f in filtro))] if filtro else df
    st.dataframe(
        vis[["nome", "codigo_acesso", "vaga", "email", "telefone", "status_entrevista", "etapa", "situacao_cnpj_pj", "padrinho",
             "atualizado_em"]]
        .rename(columns={"nome": "Nome", "codigo_acesso": "Código de acesso", "vaga": "Vaga", "email": "E-mail", "telefone": "Telefone",
                         "status_entrevista": "Entrevista", "etapa": "Situação", "situacao_cnpj_pj": "CNPJ",
                         "padrinho": "Padrinho", "atualizado_em": "Atualizado em"}),
        hide_index=True, width="stretch",
    )
    st.download_button("⬇️ Exportar base completa (CSV)", df[COLUNAS].to_csv(index=False).encode("utf-8-sig"),
                       "prestadores.csv", "text/csv")


def etapa_entrevistas():
    st.markdown("#### Cadastrar candidato")
    with st.form("novo_cand", clear_on_submit=True):
        c1, c2 = st.columns(2)
        nome = c1.text_input("Nome *")
        email = c2.text_input("E-mail *")
        tel = c1.text_input("Telefone *")
        vaga = c2.text_input("Vaga / Cargo *")
        status = st.selectbox("Status", STATUS_ENTREVISTA)
        cv = st.file_uploader("Currículo (PDF ou Word, opcional)", type=["pdf", "doc", "docx"])
        if st.form_submit_button("Cadastrar", type="primary"):
            if not all([nome.strip(), email.strip(), tel.strip(), vaga.strip()]):
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

    df = carregar()
    if df.empty:
        return
    st.markdown("#### Atualizar status da entrevista")
    nomes = {r["id"]: f"{r['nome']} · {r['vaga']} ({r['status_entrevista']})" for _, r in df.iterrows()}
    r = buscar(st.selectbox("Candidato", list(nomes), format_func=nomes.get, key="sel_entrevista"))
    c1, c2 = st.columns([1, 2])
    novo_status = c1.selectbox("Status", STATUS_ENTREVISTA, index=STATUS_ENTREVISTA.index(r["status_entrevista"])
                               if r["status_entrevista"] in STATUS_ENTREVISTA else 0, key=f"st_{r['id']}")
    obs = c2.text_input("Observações da entrevista", r["obs_entrevista"], key=f"obs_{r['id']}")
    if st.button("Salvar status", key=f"salva_st_{r['id']}"):
        atualizar(r["id"], status_entrevista=novo_status, obs_entrevista=obs)
        st.success("Status atualizado.")
        st.rerun()

    if r["status_entrevista"] == "Aprovado":
        st.info(
            f"Envie ao candidato o link deste app, o **código de acesso `{r['codigo_acesso']}`** e o e-mail "
            f"cadastrado (`{r['email']}`) para ele preencher a Etapa 2."
        )

    c1, c2 = st.columns([1, 2])
    with c1:
        mostrar_arquivo(r["curriculo"], "Currículo", key=f"dl_cv_{r['id']}")
    with c2:
        novo_cv = st.file_uploader("Anexar ou trocar currículo", type=["pdf", "doc", "docx"], key=f"up_cv_{r['id']}")
        if novo_cv is not None and st.button("Salvar currículo", key=f"salva_cv_{r['id']}"):
            atualizar(r["id"], curriculo=guardar_arquivo(f"{r['id']}_cv", novo_cv))
            st.success("Currículo salvo.")
            st.rerun()

    avaliacao_candidato(r)


def caminho_local(valor):
    return UPLOAD_DIR / valor[6:] if valor.startswith("local:") else None


def avaliacao_candidato(r):
    st.markdown("#### Avaliação de perfil")
    st.caption("Cole aqui a nota e o resumo da avaliação feita pelo Claude no projeto e anexe o arquivo completo.")
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
        st.success("Avaliação salva.")
        st.rerun()
    mostrar_arquivo(r["avaliacao"], "Avaliação completa", key=f"dl_av_{r['id']}")
    local = caminho_local(r["avaliacao"])
    if local and local.suffix == ".md" and local.exists():
        with st.expander("Ler avaliação completa"):
            st.markdown(local.read_text(encoding="utf-8", errors="ignore"))


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

    st.markdown(f"#### Ranking · {vaga}")
    st.caption(f"{d['nota'].notna().sum()} de {len(d)} candidatos com avaliação.")
    st.dataframe(
        d[["posicao", "nome", "nota", "status_entrevista", "avaliacao_resumo", "Currículo"]].rename(columns={
            "posicao": "#", "nome": "Nome", "nota": "Nota", "status_entrevista": "Status",
            "avaliacao_resumo": "Resumo executivo"}),
        hide_index=True, width="stretch",
        column_config={"Nota": st.column_config.ProgressColumn("Nota", min_value=0, max_value=10, format="%.1f")},
    )

    st.markdown("#### Comparativo detalhado")
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
        st.success("Comparativo salvo.")
        st.rerun()


def mostrar_arquivo(valor, rotulo, key):
    if not valor:
        st.write(f"{rotulo}: *não anexado*")
    elif valor.startswith("local:"):
        caminho = UPLOAD_DIR / valor[6:]
        if caminho.exists():
            st.download_button(f"📎 Baixar {rotulo}", caminho.read_bytes(), caminho.name, key=key)
        else:
            st.warning(f"{rotulo}: arquivo não encontrado (o servidor pode ter sido reiniciado). Anexe de novo.")
    else:
        st.link_button(f"📎 Abrir {rotulo} no Drive", valor)


def mostrar_cartao(valor):
    mostrar_arquivo(valor, "Cartão CNPJ", key="dl_cartao")


def ck(r, campo, rotulo, key):
    return st.checkbox(rotulo, value=r[campo] == SIM, key=f"{key}_{r['id']}")


def sn(v):
    return SIM if v else ""


def gestao_prestador():
    df = carregar()
    df = df[df["status_entrevista"] == "Aprovado"]
    if df.empty:
        st.info("Nenhum candidato aprovado na entrevista ainda.")
        return
    nomes = {r["id"]: f"{r['nome']} · {rotulo_etapa(r)}" for _, r in df.iterrows()}
    r = buscar(st.selectbox("Prestador", list(nomes), format_func=nomes.get, key="sel_prestador"))
    barra_progresso(r)
    st.caption("As etapas são independentes: preencha cada aba na ordem que for mais conveniente.")

    t3, t4, t5, t6, t7 = st.tabs([ETAPAS[3], ETAPAS[4], ETAPAS[5], ETAPAS[6], ETAPAS[7]])

    # ---------------- Etapa 3 ----------------
    with t3:
        if not r["form_enviado_em"]:
            st.warning("O candidato ainda não enviou o formulário da Etapa 2.")
        else:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Dados pessoais**")
                st.write({
                    "Nome completo": r["nome_completo"], "CPF": r["cpf"], "RG": r["rg"],
                    "Nascimento": r["data_nascimento"], "Telefone": r["telefone"], "E-mail": r["email"],
                    "Estado civil": r["estado_civil"], "Endereço": r["endereco"],
                })
            with c2:
                st.markdown("**Pessoa jurídica**")
                st.write({
                    "Possui CNPJ": r["possui_cnpj"], "CNPJ": fmt_cnpj(r["cnpj"]), "Razão social": r["razao_social"],
                    "Natureza jurídica": r["natureza_juridica"], "Situação": r["situacao_cnpj"],
                    "MEI": r["is_mei"], f"CNAE {fmt_cnae(CNAE_EXIGIDO)}": r["tem_cnae"],
                    "CNAE principal": r["cnae_principal"], "CNAEs secundários": r["cnaes_secundarios"],
                    "Validação": r["fonte_validacao"], "Apto": r["cnpj_apto"],
                })
                mostrar_cartao(r["cartao_cnpj"])
            if r["fonte_validacao"].startswith("Declaração"):
                st.warning("A validação foi declarada pelo candidato. Confira MEI e CNAE no cartão CNPJ.")
            st.write(f"Status atual da documentação: **{r['doc_status'] or 'Pendente'}**")
            obs = st.text_input("Observação para o candidato (obrigatória ao reprovar)", r["doc_obs"],
                                key=f"docobs_{r['id']}")
            if cnpj_pendente(r):
                motivo = "não tem CNPJ" if r["possui_cnpj"] == "Não" else "tem CNPJ fora das regras (MEI, CNAE ou situação)"
                st.warning(
                    f"O candidato {motivo}. Você pode **aprovar com o CNPJ pendente** para seguir com contrato, "
                    "acessos e integração. O início das operações (Etapa 7) só libera depois que o CNPJ "
                    "for regularizado: o candidato volta ao formulário com o mesmo código e valida o novo CNPJ."
                )
                b1, b2 = st.columns(2)
                if r["doc_status"] != DOC_APROVADA_PEND and b1.button(
                        "⏭️ Aprovar com CNPJ pendente", key=f"aprova_pend_{r['id']}", type="primary"):
                    atualizar(r["id"], doc_status=DOC_APROVADA_PEND, doc_obs=obs, doc_validado_em=agora())
                    st.rerun()
            else:
                if r["doc_status"] == DOC_APROVADA_PEND:
                    st.info("O candidato regularizou o CNPJ. Confira o cartão acima e aprove a documentação definitiva.")
                b1, b2 = st.columns(2)
                if r["doc_status"] != DOC_APROVADA and b1.button(
                        "✅ Aprovar documentação", key=f"aprova_{r['id']}", type="primary"):
                    atualizar(r["id"], doc_status=DOC_APROVADA, doc_obs=obs, doc_validado_em=agora())
                    st.rerun()
            if r["doc_status"] not in DOCS_LIBERADAS:
                if b2.button("❌ Reprovar / pedir ajuste", key=f"reprova_{r['id']}"):
                    if not obs.strip():
                        st.error("Descreva o ajuste necessário.")
                    else:
                        atualizar(r["id"], doc_status="Reprovada", doc_obs=obs, doc_validado_em=agora())
                        st.rerun()

    # ---------------- Etapa 4 ----------------
    with t4:
        enviado = ck(r, "contrato_enviado", "Contrato de parceria enviado para assinatura", "ce")
        assinado = ck(r, "contrato_assinado", "Contrato assinado por ambas as partes", "ca")
        if st.button("Salvar contrato", key=f"salva_ct_{r['id']}", type="primary"):
            if assinado and not enviado:
                st.error("Marque o envio antes da assinatura.")
            else:
                atualizar(r["id"], contrato_enviado=sn(enviado), contrato_assinado=sn(assinado),
                          contrato_data=agora() if assinado and not r["contrato_data"] else r["contrato_data"])
                st.rerun()
        if r["contrato_data"]:
            st.caption(f"Assinatura registrada em {r['contrato_data']}.")

    # ---------------- Etapa 5 ----------------
    with t5:
        a = ck(r, "ti_email", "Criação de e-mail corporativo", "ti1")
        b = ck(r, "ti_looker", "Acesso ao sistema de gestão Looker", "ti2")
        c = ck(r, "ti_propostas", "Acesso ao sistema de digitação de propostas", "ti3")
        if st.button("Salvar acessos", key=f"salva_ti_{r['id']}", type="primary"):
            atualizar(r["id"], ti_email=sn(a), ti_looker=sn(b), ti_propostas=sn(c))
            st.rerun()

    # ---------------- Etapa 6 ----------------
    with t6:
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
        data_int = st.date_input("Data da reunião de integração", di, format="DD/MM/YYYY", key=f"di_{r['id']}")
        padrinho = st.text_input("Padrinho responsável pelo primeiro mês *", r["padrinho"], key=f"pad_{r['id']}")
        if st.button("Salvar integração", key=f"salva_onb_{r['id']}", type="primary"):
            atualizar(r["id"], data_integracao=data_int.isoformat(), padrinho=padrinho.strip(),
                      **{k: sn(v) for k, v in marcados.items()})
            st.rerun()

    # ---------------- Etapa 7 ----------------
    with t7:
        if r["status_final"] == STATUS_ATIVO:
            st.success(f"🎉 {STATUS_ATIVO} desde {r['operacao_iniciada_em']}.")
        else:
            pend = pendencias(r)
            if cnpj_pendente(r) is False and r["doc_status"] == DOC_APROVADA_PEND:
                pend.append("Aprovação definitiva dos documentos")
            if pend:
                st.warning("Ainda há pendências: **" + ", ".join(pend) + "**.")
                forcar = st.checkbox("Iniciar a operação mesmo com essas pendências", key=f"forca_{r['id']}")
            else:
                st.write("Todas as etapas foram concluídas. Confirme o início das operações.")
                forcar = True
            if st.button("🚀 Iniciar operação", key=f"ativa_{r['id']}", type="primary", disabled=not forcar):
                atualizar(r["id"], status_final=STATUS_ATIVO, operacao_iniciada_em=agora())
                st.balloons()
                st.rerun()


# =============================================================================
# NAVEGAÇÃO
# =============================================================================
with st.sidebar:
    st.header("Onboarding de Prestadores")
    tela = st.radio("Menu", ["Visão do Candidato / Formulário", "Painel de Controle / RH"])
    st.divider()
    st.caption("Armazenamento: " + ("Google Sheets ✅" if usa_sheets() else "CSV local (demonstração)"))
    if st.session_state.get("rh_ok") and st.button("Sair do painel RH"):
        st.session_state.pop("rh_ok")
        st.rerun()

if tela.startswith("Visão do Candidato"):
    tela_candidato()
else:
    tela_rh()

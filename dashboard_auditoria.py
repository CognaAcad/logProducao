import re
import json
import hashlib
import hmac
import sqlite3
import base64
import requests
from io import BytesIO
from pathlib import Path
from datetime import timedelta, datetime, date

import pandas as pd
import streamlit as st
import plotly.express as px


# ==========================================================
# CONFIGURAÇÃO
# ==========================================================

st.set_page_config(
    page_title="Dashboard de Auditoria SharePoint",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Dashboard de Auditoria SharePoint")

st.caption(
    "Auditoria das entregas, controle de produção "
    "e acompanhamento do cronograma."
)


# ==========================================================
# ARQUIVOS LOCAIS
# ==========================================================

ARQUIVO_BANCO = Path(
    "dashboard_auditoria.db"
)

ARQUIVO_CONTROLE = Path(
    "controle_producao.xlsx"
)

ARQUIVO_MAPEAMENTO = Path(
    "mapeamento_bibliotecas.json"
)


ARQUIVO_RESPONSAVEIS = Path(
    "responsaveis_disciplinas.json"
)


# Link direto para download do arquivo de backup no GitHub.
URL_BACKUP = (
    "https://raw.githubusercontent.com/"
    "CognaAcad/logProducao/main/"
    "auditoria_sharepoint_seleniumv4.xlsx"
)


# ==========================================================
# ACESSO ADMINISTRATIVO
# ==========================================================

ADMIN_USERNAME = "allan.c.souza@cogna.com.br"

# SHA-256 da senha administrativa informada.
# A senha em texto puro não fica gravada no código.
ADMIN_PASSWORD_SHA256 = (
    "d41ca9b3ff93b24da439c32ab28c24fd"
    "03220fbee13d3c4650f20125172ae72d"
)


# ==========================================================
# CONFIGURAÇÕES DO CONTROLE
# ==========================================================

STATUS_PRODUCAO = [
    "",
    "Não iniciado",
    "Em produção",
    "Em revisão",
    "Ajuste solicitado",
    "Aguardando",
    "Validado",
    "Não se aplica",
]


COLUNAS_ETAPAS = [
    "Texto",
    "Relatório Antiplágio",
    "PPT",
    "Questões",
    "Roteiro Podcast",
    "RAP",
]


COLUNAS_CONTROLE = [
    "Biblioteca Original",
    "Área",
    "Consultor",
    "Revisor",
    "Revisor e-mail",
    "Autor",
    "Autor e-mail",
    "Responsável",
    "Disciplina",
    "Unidade",
    "Aula",
    "Prazo",
    "Texto",
    "Relatório Antiplágio",
    "PPT",
    "Questões",
    "Roteiro Podcast",
    "RAP",
    "Última Atualização",
    "Status da Produção",
]


# ==========================================================
# BANCO DE DADOS - PERSISTÊNCIA DA AUDITORIA
# ==========================================================

def conectar_banco():

    return sqlite3.connect(
        ARQUIVO_BANCO
    )


def inicializar_banco():

    conexao = conectar_banco()

    try:

        cursor = conexao.cursor()

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS importacoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                arquivo TEXT NOT NULL,
                hash_arquivo TEXT NOT NULL,
                data_importacao TEXT NOT NULL
            )
            """
        )

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS cronograma_disciplinas (
                biblioteca_original TEXT PRIMARY KEY,
                disciplina TEXT NOT NULL,
                data_inicio TEXT NOT NULL,
                data_atualizacao TEXT NOT NULL
            )
            """
        )

        conexao.commit()

    finally:

        conexao.close()


def carregar_data_inicio_cronograma(biblioteca_original):

    conexao = conectar_banco()

    try:

        cursor = conexao.cursor()

        cursor.execute(
            """
            SELECT data_inicio
            FROM cronograma_disciplinas
            WHERE biblioteca_original = ?
            LIMIT 1
            """,
            (str(biblioteca_original).strip(),)
        )

        resultado = cursor.fetchone()

        if not resultado:

            return None

        return datetime.strptime(
            resultado[0],
            "%Y-%m-%d"
        ).date()

    except (ValueError, TypeError, sqlite3.Error):

        return None

    finally:

        conexao.close()


def salvar_data_inicio_cronograma(
    biblioteca_original,
    disciplina,
    data_inicio
):

    if data_inicio is None:

        return

    biblioteca_original = str(
        biblioteca_original
    ).strip()

    disciplina = str(
        disciplina
    ).strip()

    data_inicio_iso = pd.Timestamp(
        data_inicio
    ).strftime(
        "%Y-%m-%d"
    )

    data_atualizacao = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    conexao = conectar_banco()

    try:

        cursor = conexao.cursor()

        cursor.execute(
            """
            INSERT INTO cronograma_disciplinas (
                biblioteca_original,
                disciplina,
                data_inicio,
                data_atualizacao
            )
            VALUES (?, ?, ?, ?)
            ON CONFLICT(biblioteca_original)
            DO UPDATE SET
                disciplina = excluded.disciplina,
                data_inicio = excluded.data_inicio,
                data_atualizacao = excluded.data_atualizacao
            """,
            (
                biblioteca_original,
                disciplina,
                data_inicio_iso,
                data_atualizacao,
            )
        )

        conexao.commit()

    finally:

        conexao.close()


def carregar_cronograma_banco():

    conexao = conectar_banco()

    try:

        return pd.read_sql_query(
            """
            SELECT
                biblioteca_original,
                disciplina,
                data_inicio,
                data_atualizacao
            FROM cronograma_disciplinas
            ORDER BY disciplina
            """,
            conexao
        )

    except Exception:

        return pd.DataFrame(
            columns=[
                "biblioteca_original",
                "disciplina",
                "data_inicio",
                "data_atualizacao",
            ]
        )

    finally:

        conexao.close()


def tabela_auditoria_existe():

    if not ARQUIVO_BANCO.exists():

        return False

    conexao = conectar_banco()

    try:

        cursor = conexao.cursor()

        cursor.execute(
            """
            SELECT name
            FROM sqlite_master
            WHERE type = 'table'
              AND name = 'auditoria'
            """
        )

        return cursor.fetchone() is not None

    finally:

        conexao.close()


def salvar_auditoria_banco(df_auditoria):

    conexao = conectar_banco()

    try:

        df_auditoria.to_sql(
            "auditoria",
            conexao,
            if_exists="replace",
            index=False
        )

        conexao.commit()

    finally:

        conexao.close()


def carregar_auditoria_banco():

    if not tabela_auditoria_existe():

        return None

    conexao = conectar_banco()

    try:

        df_salvo = pd.read_sql_query(
            "SELECT * FROM auditoria",
            conexao
        )

        if df_salvo.empty:

            return None

        return df_salvo.fillna("")

    except Exception:

        return None

    finally:

        conexao.close()


def registrar_importacao(
    nome_arquivo,
    hash_arquivo
):

    conexao = conectar_banco()

    try:

        cursor = conexao.cursor()

        cursor.execute(
            """
            INSERT INTO importacoes (
                arquivo,
                hash_arquivo,
                data_importacao
            )
            VALUES (?, ?, ?)
            """,
            (
                nome_arquivo,
                hash_arquivo,
                datetime.now().strftime(
                    "%d/%m/%Y %H:%M:%S"
                )
            )
        )

        conexao.commit()

    finally:

        conexao.close()


def obter_ultima_importacao():

    if not ARQUIVO_BANCO.exists():

        return None

    conexao = conectar_banco()

    try:

        cursor = conexao.cursor()

        cursor.execute(
            """
            SELECT
                arquivo,
                hash_arquivo,
                data_importacao
            FROM importacoes
            ORDER BY id DESC
            LIMIT 1
            """
        )

        resultado = cursor.fetchone()

        if not resultado:

            return None

        return {
            "arquivo": resultado[0],
            "hash_arquivo": resultado[1],
            "data_importacao": resultado[2],
        }

    except Exception:

        return None

    finally:

        conexao.close()


def calcular_hash_arquivo(conteudo):

    return hashlib.sha256(
        conteudo
    ).hexdigest()


inicializar_banco()


# ==========================================================
# AUTENTICAÇÃO
# ==========================================================

def hash_senha(senha):
    return hashlib.sha256(
        str(senha).encode("utf-8")
    ).hexdigest()


def esta_autenticado():
    return bool(
        st.session_state.get(
            "admin_autenticado",
            False
        )
    )


def verificar_credenciais(usuario, senha):
    usuario_ok = hmac.compare_digest(
        str(usuario).strip(),
        ADMIN_USERNAME
    )

    senha_ok = hmac.compare_digest(
        hash_senha(senha),
        ADMIN_PASSWORD_SHA256
    )

    return usuario_ok and senha_ok


def formulario_login(local="pagina"):
    st.markdown("### 🔐 Área restrita")
    st.caption(
        "Informe o usuário e a senha administrativos para continuar."
    )

    with st.form(
        key=f"login_admin_{local}",
        clear_on_submit=False
    ):
        usuario = st.text_input(
            "Usuário",
            key=f"login_usuario_{local}"
        )

        senha = st.text_input(
            "Senha",
            type="password",
            key=f"login_senha_{local}"
        )

        entrar = st.form_submit_button(
            "🔓 Entrar",
            type="primary",
            use_container_width=True
        )

    if entrar:
        if verificar_credenciais(usuario, senha):
            st.session_state[
                "admin_autenticado"
            ] = True

            st.session_state[
                "admin_usuario"
            ] = ADMIN_USERNAME

            st.success(
                "✅ Acesso autorizado."
            )
            st.rerun()

        else:
            st.error(
                "Usuário ou senha inválidos."
            )


def formulario_login_sidebar():
    st.sidebar.warning(
        "🔒 Cronograma protegido"
    )

    with st.sidebar.form(
        key="login_admin_sidebar",
        clear_on_submit=False
    ):
        usuario = st.text_input(
            "Usuário",
            key="login_usuario_sidebar"
        )

        senha = st.text_input(
            "Senha",
            type="password",
            key="login_senha_sidebar"
        )

        entrar = st.form_submit_button(
            "🔓 Entrar",
            use_container_width=True
        )

    if entrar:
        if verificar_credenciais(usuario, senha):
            st.session_state[
                "admin_autenticado"
            ] = True

            st.session_state[
                "admin_usuario"
            ] = ADMIN_USERNAME

            st.rerun()

        else:
            st.sidebar.error(
                "Usuário ou senha inválidos."
            )


def botao_logout(local="pagina"):
    if not esta_autenticado():
        return

    st.caption(
        f"🔓 Acesso administrativo: {ADMIN_USERNAME}"
    )

    if st.button(
        "🚪 Sair da área administrativa",
        key=f"logout_admin_{local}"
    ):
        st.session_state[
            "admin_autenticado"
        ] = False

        st.session_state.pop(
            "admin_usuario",
            None
        )

        st.rerun()


# ==========================================================
# FUNÇÕES
# ==========================================================

def gerar_key(texto):

    return hashlib.md5(
        str(texto).encode("utf-8")
    ).hexdigest()[:10]


# ==========================================================
# DISCIPLINA
# ==========================================================

def limpar_nome_disciplina(nome):

    nome = str(nome)

    nome = nome.rsplit(".", 1)[0]

    nome = nome.replace(
        "_",
        " "
    )

    nome = nome.replace(
        "-",
        " "
    )

    nome = re.sub(
        r"\s+",
        " ",
        nome
    )

    return nome.strip().upper()


def extrair_disciplina_dos_arquivos(
    arquivos
):

    arquivos = str(
        arquivos
    )

    padrao = (
        r"U\d+A\d{2,3}_([^;]+)"
    )

    achados = re.findall(
        padrao,
        arquivos,
        flags=re.IGNORECASE
    )

    if achados:

        return limpar_nome_disciplina(
            achados[0]
        )

    return ""


def descobrir_disciplina_automatica(
    grupo
):
    """
    Prioridade para identificar a disciplina:
    1. coluna Disciplina gerada pelo robô Selenium/SharePoint;
    2. nome extraído dos arquivos encontrados;
    3. vazio, para permitir os fallbacks antigos.
    """

    # ------------------------------------------------------
    # 1. NOME VINDO DIRETAMENTE DO SHAREPOINT
    # ------------------------------------------------------
    if "Disciplina" in grupo.columns:
        candidatas_sharepoint = []

        for valor in grupo["Disciplina"]:
            nome = str(valor).strip()

            if not nome:
                continue

            # Ignora placeholders antigos como Biblioteca 1.
            if re.fullmatch(r"Biblioteca\s*\d+", nome, flags=re.IGNORECASE):
                continue

            candidatas_sharepoint.append(nome.upper())

        if candidatas_sharepoint:
            return pd.Series(candidatas_sharepoint).mode().iloc[0]

    # ------------------------------------------------------
    # 2. FALLBACK: EXTRAI DO NOME DOS ARQUIVOS
    # ------------------------------------------------------
    disciplinas = []

    if "Arquivos encontrados" in grupo.columns:
        for arquivos in grupo["Arquivos encontrados"]:
            disciplina = extrair_disciplina_dos_arquivos(arquivos)

            if disciplina:
                disciplinas.append(disciplina)

    if disciplinas:
        return pd.Series(disciplinas).mode().iloc[0]

    return ""


# ==========================================================
# MAPEAMENTO
# ==========================================================

def configuracao_github_disponivel():

    try:
        return (
            "github" in st.secrets
            and bool(st.secrets["github"].get("token"))
            and bool(st.secrets["github"].get("owner"))
            and bool(st.secrets["github"].get("repo"))
        )
    except Exception:
        return False


def configuracao_github():

    if not configuracao_github_disponivel():
        return None

    return {
        "token": st.secrets["github"]["token"],
        "owner": st.secrets["github"]["owner"],
        "repo": st.secrets["github"]["repo"],
        "branch": st.secrets["github"].get(
            "branch",
            "main"
        ),
        "arquivo": st.secrets["github"].get(
            "arquivo_mapeamento",
            "mapeamento_bibliotecas.json"
        ),
    }


def obter_headers_github():

    config = configuracao_github()

    if config is None:
        return {}

    return {
        "Authorization": f"Bearer {config['token']}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def carregar_mapeamento_local():

    if not ARQUIVO_MAPEAMENTO.exists():
        return {}

    try:
        with open(
            ARQUIVO_MAPEAMENTO,
            "r",
            encoding="utf-8"
        ) as arquivo:
            return json.load(arquivo)
    except Exception:
        return {}


def salvar_mapeamento_local(mapa):

    with open(
        ARQUIVO_MAPEAMENTO,
        "w",
        encoding="utf-8"
    ) as arquivo:
        json.dump(
            mapa,
            arquivo,
            ensure_ascii=False,
            indent=4
        )


def carregar_mapeamento_github():
    """Lê mapeamento_bibliotecas.json diretamente do GitHub."""

    config = configuracao_github()

    if config is None:
        return carregar_mapeamento_local()

    url = (
        f"https://api.github.com/repos/"
        f"{config['owner']}/"
        f"{config['repo']}/contents/"
        f"{config['arquivo']}"
    )

    resposta = requests.get(
        url,
        headers=obter_headers_github(),
        params={"ref": config["branch"]},
        timeout=20
    )

    if resposta.status_code == 404:
        return carregar_mapeamento_local()

    resposta.raise_for_status()

    dados = resposta.json()

    conteudo_base64 = dados.get("content", "")

    if not conteudo_base64:
        return {}

    conteudo = base64.b64decode(
        conteudo_base64
    ).decode("utf-8")

    return json.loads(conteudo)


def salvar_mapeamento_github(mapa):
    """Cria/atualiza mapeamento_bibliotecas.json no GitHub."""

    config = configuracao_github()

    # Execução local sem Secrets: mantém o comportamento antigo.
    if config is None:
        salvar_mapeamento_local(mapa)
        return True, "local"

    url = (
        f"https://api.github.com/repos/"
        f"{config['owner']}/"
        f"{config['repo']}/contents/"
        f"{config['arquivo']}"
    )

    headers = obter_headers_github()

    resposta_atual = requests.get(
        url,
        headers=headers,
        params={"ref": config["branch"]},
        timeout=20
    )

    sha_atual = None

    if resposta_atual.status_code == 200:
        sha_atual = resposta_atual.json().get("sha")
    elif resposta_atual.status_code != 404:
        resposta_atual.raise_for_status()

    conteudo_json = json.dumps(
        mapa,
        ensure_ascii=False,
        indent=4
    )

    conteudo_base64 = base64.b64encode(
        conteudo_json.encode("utf-8")
    ).decode("utf-8")

    payload = {
        "message": "Atualiza nomes das disciplinas pelo Streamlit",
        "content": conteudo_base64,
        "branch": config["branch"],
    }

    if sha_atual:
        payload["sha"] = sha_atual

    resposta = requests.put(
        url,
        headers=headers,
        json=payload,
        timeout=20
    )

    resposta.raise_for_status()

    # Mantém também uma cópia local durante a sessão atual.
    salvar_mapeamento_local(mapa)

    return True, "github"


def carregar_mapeamento():

    try:
        return carregar_mapeamento_github()
    except Exception as erro:
        st.warning(
            "Não foi possível carregar o mapeamento do GitHub. "
            f"Será usada a cópia local, se existir. Erro: {erro}"
        )
        return carregar_mapeamento_local()


def salvar_mapeamento(mapa):

    try:
        return salvar_mapeamento_github(mapa)
    except Exception as erro:
        st.error(
            "Não foi possível salvar o mapeamento no GitHub. "
            f"Erro: {erro}"
        )
        return False, "erro"


def _config_arquivo_github(nome_secret, padrao):

    config = configuracao_github()

    if config is None:
        return None

    try:
        arquivo = st.secrets["github"].get(nome_secret, padrao)
    except Exception:
        arquivo = padrao

    novo = dict(config)
    novo["arquivo"] = arquivo
    return novo


def carregar_json_local(caminho):

    caminho = Path(caminho)

    if not caminho.exists():
        return {}

    try:
        with open(caminho, "r", encoding="utf-8") as arquivo:
            dados = json.load(arquivo)
            return dados if isinstance(dados, dict) else {}
    except Exception:
        return {}


def salvar_json_local(caminho, dados):

    with open(caminho, "w", encoding="utf-8") as arquivo:
        json.dump(dados, arquivo, ensure_ascii=False, indent=4)


def carregar_json_github(nome_secret, padrao, caminho_local):

    config = _config_arquivo_github(nome_secret, padrao)

    if config is None:
        return carregar_json_local(caminho_local)

    url = (
        f"https://api.github.com/repos/"
        f"{config['owner']}/"
        f"{config['repo']}/contents/"
        f"{config['arquivo']}"
    )

    resposta = requests.get(
        url,
        headers=obter_headers_github(),
        params={"ref": config["branch"]},
        timeout=20,
    )

    if resposta.status_code == 404:
        return carregar_json_local(caminho_local)

    resposta.raise_for_status()
    dados = resposta.json()
    conteudo_base64 = dados.get("content", "")

    if not conteudo_base64:
        return {}

    conteudo = base64.b64decode(conteudo_base64).decode("utf-8")
    resultado = json.loads(conteudo)
    return resultado if isinstance(resultado, dict) else {}


def salvar_json_github(nome_secret, padrao, caminho_local, dados, mensagem):

    config = _config_arquivo_github(nome_secret, padrao)

    if config is None:
        salvar_json_local(caminho_local, dados)
        return True, "local"

    url = (
        f"https://api.github.com/repos/"
        f"{config['owner']}/"
        f"{config['repo']}/contents/"
        f"{config['arquivo']}"
    )

    headers = obter_headers_github()
    resposta_atual = requests.get(
        url,
        headers=headers,
        params={"ref": config["branch"]},
        timeout=20,
    )

    sha_atual = None
    if resposta_atual.status_code == 200:
        sha_atual = resposta_atual.json().get("sha")
    elif resposta_atual.status_code != 404:
        resposta_atual.raise_for_status()

    conteudo_json = json.dumps(dados, ensure_ascii=False, indent=4)
    conteudo_base64 = base64.b64encode(
        conteudo_json.encode("utf-8")
    ).decode("utf-8")

    payload = {
        "message": mensagem,
        "content": conteudo_base64,
        "branch": config["branch"],
    }

    if sha_atual:
        payload["sha"] = sha_atual

    resposta = requests.put(
        url,
        headers=headers,
        json=payload,
        timeout=20,
    )
    resposta.raise_for_status()
    salvar_json_local(caminho_local, dados)
    return True, "github"


def carregar_responsaveis():

    try:
        return carregar_json_github(
            "arquivo_responsaveis",
            "responsaveis_disciplinas.json",
            ARQUIVO_RESPONSAVEIS,
        )
    except Exception as erro:
        st.warning(
            "Não foi possível carregar Autor/Responsável do GitHub. "
            f"Será usada a cópia local, se existir. Erro: {erro}"
        )
        return carregar_json_local(ARQUIVO_RESPONSAVEIS)


def salvar_responsaveis(dados):

    try:
        return salvar_json_github(
            "arquivo_responsaveis",
            "responsaveis_disciplinas.json",
            ARQUIVO_RESPONSAVEIS,
            dados,
            "Atualiza autores e responsáveis pelo Streamlit",
        )
    except Exception as erro:
        st.error(
            "Não foi possível salvar Autor/Responsável no GitHub. "
            f"Erro: {erro}"
        )
        return False, "erro"


def normalizar_chave_disciplina(valor):

    return re.sub(
        r"\s+",
        " ",
        str(valor).strip().upper(),
    )


def aplicar_responsaveis_controle(controle, cadastro):

    if controle is None or controle.empty:
        return controle

    controle = controle.copy()

    for indice, row in controle.iterrows():
        chave = normalizar_chave_disciplina(row.get("Disciplina", ""))
        dados = cadastro.get(chave, {})

        if not isinstance(dados, dict):
            continue

        autor = str(dados.get("autor", "")).strip()
        responsavel = str(dados.get("responsavel", "")).strip()

        if autor:
            controle.at[indice, "Autor"] = autor

        if responsavel:
            controle.at[indice, "Responsável"] = responsavel

    return controle


def criar_tabela_mapeamento(
    df
):

    mapa_salvo = (
        carregar_mapeamento()
    )

    dados = []

    for biblioteca, grupo in df.groupby(
        "Biblioteca Original"
    ):

        biblioteca = str(
            biblioteca
        ).strip()

        automatica = (
            descobrir_disciplina_automatica(
                grupo
            )
        )

        disciplina_salva = (
            mapa_salvo.get(
                biblioteca,
                ""
            )
        )

        # O nome detectado diretamente pelo SharePoint tem prioridade.
        # O JSON fica apenas como fallback para relatórios antigos.
        disciplina_final = (
            automatica
            or disciplina_salva
            or biblioteca
        )

        dados.append(
            {
                "Biblioteca":
                    biblioteca,

                "Disciplina automática":
                    automatica,

                "Nome da disciplina":
                    disciplina_final,
            }
        )

    return pd.DataFrame(
        dados
    )


def aplicar_mapeamento(
    df,
    tabela_mapeamento
):

    mapa = dict(
        zip(
            tabela_mapeamento[
                "Biblioteca"
            ],
            tabela_mapeamento[
                "Nome da disciplina"
            ]
        )
    )

    df[
        "Biblioteca"
    ] = (
        df[
            "Biblioteca Original"
        ]
        .map(
            mapa
        )
    )

    df[
        "Biblioteca"
    ] = df[
        "Biblioteca"
    ].fillna(
        df[
            "Biblioteca Original"
        ]
    )

    return df


# ==========================================================
# AULAS
# ==========================================================

def extrair_numero_aula(
    aula
):

    aula = str(
        aula
    ).upper().strip()

    # Aceita:
    # A01
    # A1
    # Aula 01
    # Aula 1
    # Aula 16

    match = re.search(
        r"(?:AULA\s*|A)(\d{1,3})",
        aula
    )

    if match:

        return int(
            match.group(1)
        )

    return None


def extrair_numero_unidade_encerramento(
    aula
):

    aula = str(
        aula
    ).upper().strip()

    # Aceita:
    # Encerramento Unidade 01
    # Encerramento Unidade 1
    # ENC 01
    # ENC1

    match = re.search(
        r"(?:ENCERRAMENTO\s*UNIDADE|ENC)\s*0?(\d+)",
        aula
    )

    if match:

        return int(
            match.group(1)
        )

    return None


# ==========================================================
# PRAZOS
# ==========================================================

def calcular_prazo(
    aula,
    data_inicio
):

    if not data_inicio:

        return ""

    data_inicio = pd.Timestamp(
        data_inicio
    )

    # ------------------------------------------------------
    # AULA NORMAL
    # ------------------------------------------------------

    numero_aula = (
        extrair_numero_aula(
            aula
        )
    )

    if numero_aula is not None:

        data_entrega = (
            data_inicio
            + timedelta(
                days=(
                    numero_aula - 1
                ) * 7
            )
        )

        return data_entrega.strftime(
            "%d/%m/%Y"
        )

    # ------------------------------------------------------
    # ENCERRAMENTO
    # ------------------------------------------------------

    numero_unidade = (
        extrair_numero_unidade_encerramento(
            aula
        )
    )

    if numero_unidade is not None:

        # Unidade 1 -> depois da aula 04
        # Unidade 2 -> depois da aula 08
        # Unidade 3 -> depois da aula 12
        # Unidade 4 -> depois da aula 16

        ultima_aula_unidade = (
            numero_unidade * 4
        )

        data_ultima_aula = (
            data_inicio
            + timedelta(
                days=(
                    ultima_aula_unidade
                    - 1
                ) * 7
            )
        )

        data_encerramento = (
            data_ultima_aula
            + timedelta(
                days=7
            )
        )

        return (
            data_encerramento
            .strftime(
                "%d/%m/%Y"
            )
        )

    return ""


def converter_data(
    valor
):

    if valor is None:

        return pd.NaT

    if isinstance(
        valor,
        pd.Timestamp
    ):

        return valor

    if isinstance(
        valor,
        datetime
    ):

        return pd.Timestamp(
            valor
        )

    if isinstance(
        valor,
        date
    ):

        return pd.Timestamp(
            valor
        )

    texto = str(
        valor
    ).strip()

    if not texto:

        return pd.NaT

    return pd.to_datetime(
        texto,
        dayfirst=True,
        errors="coerce"
    )


# ==========================================================
# STATUS DA PRODUÇÃO
# ==========================================================

def status_linha(
    row
):

    etapas = []

    for coluna in COLUNAS_ETAPAS:
        if coluna not in row:
            continue

        valor = str(row[coluna]).strip()

        if valor and valor != "Não se aplica":
            etapas.append(valor)

    # Tudo aplicável validado.
    if etapas and all(etapa == "Validado" for etapa in etapas):
        return "✅ Concluído"

    tem_validado = any(etapa == "Validado" for etapa in etapas)
    status_andamento = [
        "Em produção",
        "Em revisão",
        "Ajuste solicitado",
        "Aguardando",
    ]
    tem_andamento = any(etapa in status_andamento for etapa in etapas)

    prazo = converter_data(row.get("Prazo", ""))
    hoje = pd.Timestamp(date.today())

    if pd.notna(prazo):
        diferenca = (prazo.normalize() - hoje).days

        # Se já existe evidência de entrega/andamento, não chama de atrasado.
        if diferenca < 0:
            if tem_validado or tem_andamento:
                return "🔵 Em andamento"
            return "🔴 Atrasado"

        if diferenca <= 2 and not (tem_validado or tem_andamento):
            return "🟡 Próximo do prazo"

    if tem_andamento or tem_validado:
        return "🔵 Em andamento"

    return "⚪ Não iniciado"


# ==========================================================
# CORES
# ==========================================================

def destacar_status(
    valor
):

    valor = str(
        valor
    ).lower()

    if (
        valor == "ok"
        or "validado" in valor
        or "aprovado" in valor
        or "concluído" in valor
    ):

        return (
            "background-color:#d1fae5;"
            "color:#065f46;"
            "font-weight:bold;"
        )

    if (
        "faltando" in valor
        or "erro" in valor
        or "incompleto" in valor
        or "atrasado" in valor
    ):

        return (
            "background-color:#fee2e2;"
            "color:#991b1b;"
            "font-weight:bold;"
        )

    if (
        "aguardando" in valor
        or "pendente" in valor
        or "próximo" in valor
    ):

        return (
            "background-color:#fef3c7;"
            "color:#92400e;"
            "font-weight:bold;"
        )

    if "andamento" in valor:

        return (
            "background-color:#dbeafe;"
            "color:#1e40af;"
            "font-weight:bold;"
        )

    return ""


# ==========================================================
# CONTROLE DE PRODUÇÃO
# ==========================================================

def criar_controle_inicial(
    df
):

    controle = pd.DataFrame()

    controle[
        "Biblioteca Original"
    ] = df[
        "Biblioteca Original"
    ]

    controle[
        "Disciplina"
    ] = df[
        "Biblioteca"
    ]

    controle[
        "Unidade"
    ] = df[
        "Unidade"
    ]

    controle[
        "Aula"
    ] = df[
        "Aula"
    ]

    controle[
        "Prazo"
    ] = df[
        "Prazo"
    ]

    controle[
        "Área"
    ] = ""

    controle[
        "Consultor"
    ] = ""

    controle[
        "Revisor"
    ] = ""

    controle[
        "Revisor e-mail"
    ] = ""

    controle[
        "Autor"
    ] = ""

    controle[
        "Autor e-mail"
    ] = ""

    controle[
        "Responsável"
    ] = ""

    for coluna in COLUNAS_ETAPAS:

        controle[
            coluna
        ] = "Não iniciado"

    controle[
        "Última Atualização"
    ] = ""

    controle[
        "Status da Produção"
    ] = controle.apply(
        status_linha,
        axis=1
    )

    return controle[
        COLUNAS_CONTROLE
    ]


def atualizar_prazos_controle(
    controle,
    df_base
):

    mapa_prazos = {}

    for _, row in (
        df_base.iterrows()
    ):

        chave = (
            str(
                row[
                    "Biblioteca Original"
                ]
            ).strip(),

            str(
                row[
                    "Unidade"
                ]
            ).strip(),

            str(
                row[
                    "Aula"
                ]
            ).strip(),
        )

        mapa_prazos[
            chave
        ] = row[
            "Prazo"
        ]

    for indice, row in (
        controle.iterrows()
    ):

        chave = (
            str(
                row[
                    "Biblioteca Original"
                ]
            ).strip(),

            str(
                row[
                    "Unidade"
                ]
            ).strip(),

            str(
                row[
                    "Aula"
                ]
            ).strip(),
        )

        if chave in mapa_prazos:

            controle.at[
                indice,
                "Prazo"
            ] = mapa_prazos[
                chave
            ]

    return controle


def atualizar_disciplinas_controle(
    controle,
    df_base
):

    mapa_disciplina = (
        df_base[
            [
                "Biblioteca Original",
                "Biblioteca"
            ]
        ]
        .drop_duplicates()
        .set_index(
            "Biblioteca Original"
        )[
            "Biblioteca"
        ]
        .to_dict()
    )

    controle[
        "Disciplina"
    ] = controle.apply(
        lambda row:
        mapa_disciplina.get(
            row[
                "Biblioteca Original"
            ],
            row[
                "Disciplina"
            ]
        ),
        axis=1
    )

    return controle


def _valor_ok(valor):

    return str(valor).strip().upper() == "OK"


def sincronizar_auditoria_controle(controle, df_base):
    """Promove etapas do controle para Validado quando a auditoria comprova o arquivo."""

    if controle is None:
        controle = pd.DataFrame(columns=COLUNAS_CONTROLE)

    controle = controle.copy()

    # Garante que novas linhas da auditoria também entrem no controle.
    chaves_existentes = set()
    for _, row in controle.iterrows():
        chaves_existentes.add((
            normalizar_chave_disciplina(row.get("Disciplina", "")),
            str(row.get("Unidade", "")).strip().upper(),
            str(row.get("Aula", "")).strip().upper(),
        ))

    novas_linhas = []
    for _, row in df_base.iterrows():
        chave = (
            normalizar_chave_disciplina(row.get("Biblioteca", "")),
            str(row.get("Unidade", "")).strip().upper(),
            str(row.get("Aula", "")).strip().upper(),
        )
        if chave not in chaves_existentes:
            base = {coluna: "" for coluna in COLUNAS_CONTROLE}
            base["Biblioteca Original"] = row.get("Biblioteca Original", "")
            base["Disciplina"] = row.get("Biblioteca", "")
            base["Unidade"] = row.get("Unidade", "")
            base["Aula"] = row.get("Aula", "")
            base["Prazo"] = row.get("Prazo", "")
            for etapa in COLUNAS_ETAPAS:
                base[etapa] = "Não iniciado"
            novas_linhas.append(base)
            chaves_existentes.add(chave)

    if novas_linhas:
        controle = pd.concat(
            [controle, pd.DataFrame(novas_linhas)],
            ignore_index=True,
        )

    auditoria_por_chave = {}
    for _, row in df_base.iterrows():
        chave = (
            normalizar_chave_disciplina(row.get("Biblioteca", "")),
            str(row.get("Unidade", "")).strip().upper(),
            str(row.get("Aula", "")).strip().upper(),
        )
        auditoria_por_chave[chave] = row

    for indice, row in controle.iterrows():
        chave = (
            normalizar_chave_disciplina(row.get("Disciplina", "")),
            str(row.get("Unidade", "")).strip().upper(),
            str(row.get("Aula", "")).strip().upper(),
        )
        audit = auditoria_por_chave.get(chave)

        if audit is None:
            continue

        # Não rebaixa status manual; apenas promove para Validado.
        if _valor_ok(audit.get("DOCX", "")):
            controle.at[indice, "Texto"] = "Validado"

        if _valor_ok(audit.get("PDF", "")):
            controle.at[indice, "Relatório Antiplágio"] = "Validado"

        if _valor_ok(audit.get("PPT", "")):
            controle.at[indice, "PPT"] = "Validado"

        if _valor_ok(audit.get("Questões", "")):
            controle.at[indice, "Questões"] = "Validado"

    controle["Status da Produção"] = controle.apply(status_linha, axis=1)
    return controle[COLUNAS_CONTROLE]


def carregar_controle(
    df_base
):

    if not ARQUIVO_CONTROLE.exists():

        controle = criar_controle_inicial(
            df_base
        )
        controle = sincronizar_auditoria_controle(
            controle,
            df_base
        )
        controle = aplicar_responsaveis_controle(
            controle,
            carregar_responsaveis()
        )
        return controle[COLUNAS_CONTROLE]

    try:

        controle = pd.read_excel(
            ARQUIVO_CONTROLE,
            engine="openpyxl"
        ).fillna("")

    except Exception as erro:

        st.warning(
            f"Erro ao abrir "
            f"{ARQUIVO_CONTROLE}: "
            f"{erro}"
        )

        return criar_controle_inicial(
            df_base
        )

    # ------------------------------------------------------
    # GARANTE COLUNAS
    # ------------------------------------------------------

    for coluna in COLUNAS_CONTROLE:

        if coluna not in (
            controle.columns
        ):

            controle[
                coluna
            ] = ""

    # ------------------------------------------------------
    # ATUALIZA DISCIPLINAS
    # ------------------------------------------------------

    controle = (
        atualizar_disciplinas_controle(
            controle,
            df_base
        )
    )

    # ------------------------------------------------------
    # ATUALIZA PRAZOS
    # ------------------------------------------------------

    controle = (
        atualizar_prazos_controle(
            controle,
            df_base
        )
    )

    # ------------------------------------------------------
    # SINCRONIZA AUDITORIA -> CONTROLE
    # ------------------------------------------------------

    controle = sincronizar_auditoria_controle(
        controle,
        df_base
    )

    # ------------------------------------------------------
    # AUTOR / RESPONSÁVEL CADASTRADOS UMA ÚNICA VEZ
    # ------------------------------------------------------

    controle = aplicar_responsaveis_controle(
        controle,
        carregar_responsaveis()
    )

    controle[
        "Status da Produção"
    ] = controle.apply(
        status_linha,
        axis=1
    )

    return controle[
        COLUNAS_CONTROLE
    ]


def salvar_controle(
    controle
):

    controle = (
        controle.copy()
    )

    controle[
        "Última Atualização"
    ] = datetime.now().strftime(
        "%d/%m/%Y %H:%M"
    )

    controle[
        "Status da Produção"
    ] = controle.apply(
        status_linha,
        axis=1
    )

    controle.to_excel(
        ARQUIVO_CONTROLE,
        index=False
    )

    return controle


# ==========================================================
# UPLOAD / CARREGAMENTO AUTOMÁTICO
# ==========================================================

st.subheader(
    "📂 Fonte dos dados"
)

st.caption(
    "Ao enviar um novo Excel, os dados da aba Auditoria "
    "são salvos automaticamente no SQLite. Se nenhum "
    "arquivo for enviado, o sistema usa o último relatório salvo."
)

arquivo = st.file_uploader(
    "📂 Envie um novo relatório Excel para atualizar os dados",
    type=[
        "xlsx"
    ],
    key=(
        "upload_relatorio_"
        "auditoria"
    )
)


colunas_esperadas = [
    "Biblioteca",
    "Unidade",
    "Aula",
    "DOCX",
    "PPT",
    "PDF",
    "Questões",
    "REV1",
    "AJ1",
    "Aprovado",
    "Validado",
    "Status Geral",
    "Pendências",
    "Arquivos encontrados",
]


df = None


if arquivo is not None:

    try:

        conteudo_arquivo = arquivo.getvalue()

        hash_atual = calcular_hash_arquivo(
            conteudo_arquivo
        )

        df_excel = pd.read_excel(
            BytesIO(
                conteudo_arquivo
            ),
            sheet_name="Auditoria",
            engine="openpyxl"
        )

        df_excel = df_excel.fillna("")

        faltantes = [
            coluna
            for coluna in colunas_esperadas
            if coluna not in df_excel.columns
        ]

        if faltantes:

            st.error(
                "Colunas ausentes no relatório: "
                f"{faltantes}"
            )

            st.stop()

        ultima_importacao = (
            obter_ultima_importacao()
        )

        hash_anterior = (
            ultima_importacao.get(
                "hash_arquivo"
            )
            if ultima_importacao
            else None
        )

        precisa_salvar = (
            hash_atual != hash_anterior
            or not tabela_auditoria_existe()
        )

        if precisa_salvar:

            salvar_auditoria_banco(
                df_excel
            )

            registrar_importacao(
                arquivo.name,
                hash_atual
            )

            if (
                "controle_producao"
                in st.session_state
            ):

                del st.session_state[
                    "controle_producao"
                ]

            st.success(
                "✅ Relatório importado e salvo automaticamente."
            )

        else:

            st.info(
                "ℹ️ Este relatório já é o último arquivo salvo."
            )

        df = df_excel.copy()

    except Exception as erro:

        st.error(
            "Não foi possível processar o Excel. "
            f"Erro: {erro}"
        )

        st.stop()


else:

    df_salvo = carregar_auditoria_banco()

    if df_salvo is None:

        st.info(
            "Ainda não existe um relatório salvo. "
            "Envie o arquivo auditoria_sharepoint_selenium.xlsx "
            "para iniciar."
        )

        st.stop()

    df = df_salvo.copy()

    st.info(
        "📦 Utilizando o último relatório salvo no banco."
    )


df = df.fillna("")


# ==========================================================
# VALIDAÇÃO FINAL
# ==========================================================

faltantes = [
    coluna
    for coluna in colunas_esperadas
    if coluna not in df.columns
]


if faltantes:

    st.error(
        "Os dados salvos estão incompatíveis. "
        "Colunas ausentes: "
        f"{faltantes}"
    )

    st.stop()


# ==========================================================
# ÚLTIMA IMPORTAÇÃO
# ==========================================================

ultima_importacao = (
    obter_ultima_importacao()
)


if ultima_importacao:

    info1, info2, info3 = st.columns(
        [2.5, 2.0, 1.0]
    )

    info1.caption(
        "📄 Último arquivo salvo: "
        f"{ultima_importacao['arquivo']}"
    )

    info2.caption(
        "🕒 Última importação: "
        f"{ultima_importacao['data_importacao']}"
    )

    with info3:
        st.link_button(
            "⬇️ Backup",
            URL_BACKUP,
            use_container_width=True,
            help=(
                "Baixar o arquivo de backup da auditoria "
                "diretamente do GitHub."
            ),
        )


# ==========================================================
# BIBLIOTECA ORIGINAL
# ==========================================================

# Usa o nome real da disciplina, quando o novo robô já o trouxe.
# Isso elimina a dependência de Biblioteca 1, Biblioteca 2 etc.
def _chave_biblioteca_original(row):
    disciplina = str(row.get("Disciplina", "")).strip()

    if (
        disciplina
        and not re.fullmatch(
            r"Biblioteca\s*\d+",
            disciplina,
            flags=re.IGNORECASE
        )
    ):
        return disciplina.upper()

    biblioteca = str(row.get("Biblioteca", "")).strip()

    if biblioteca:
        return biblioteca

    referencia = str(row.get("Biblioteca Ref", "")).strip()
    return referencia


df["Biblioteca Original"] = df.apply(
    _chave_biblioteca_original,
    axis=1
)


# ==========================================================
# MAPEAMENTO DAS DISCIPLINAS
# ==========================================================

tabela_mapeamento = (
    criar_tabela_mapeamento(
        df
    )
)


df = aplicar_mapeamento(
    df,
    tabela_mapeamento
)


# ==========================================================
# CRONOGRAMA
# ==========================================================

st.sidebar.header(
    "📅 Cronograma"
)


disciplinas_unicas = sorted(
    [
        disciplina
        for disciplina
        in df[
            "Biblioteca"
        ].unique()
        if str(
            disciplina
        ).strip()
    ]
)


pares_cronograma = (
    df[
        [
            "Biblioteca Original",
            "Biblioteca",
        ]
    ]
    .drop_duplicates()
    .sort_values(
        "Biblioteca"
    )
)


datas_inicio = {}


# ==========================================================
# PROTEÇÃO DO CRONOGRAMA LATERAL
# ==========================================================
# Sem login, as datas já salvas continuam sendo usadas para
# calcular os prazos, mas não podem ser alteradas.
# O mesmo login administrativo libera também a aba de
# Configuração do Cronograma.

if esta_autenticado():
    st.sidebar.success(
        "🔓 Cronograma liberado para edição"
    )

    st.sidebar.caption(
        f"Administrador: {ADMIN_USERNAME}"
    )

    if st.sidebar.button(
        "🚪 Sair da área administrativa",
        key="logout_admin_sidebar"
    ):
        st.session_state[
            "admin_autenticado"
        ] = False

        st.session_state.pop(
            "admin_usuario",
            None
        )

        st.rerun()

    for _, item_cronograma in pares_cronograma.iterrows():
        biblioteca_original = str(
            item_cronograma[
                "Biblioteca Original"
            ]
        ).strip()

        disciplina = str(
            item_cronograma[
                "Biblioteca"
            ]
        ).strip()

        data_salva = carregar_data_inicio_cronograma(
            biblioteca_original
        )

        if data_salva is None:
            data_salva = date.today()

        chave = gerar_key(
            biblioteca_original
        )

        with st.sidebar.expander(
            disciplina
        ):
            data_selecionada = st.date_input(
                "Data de início",
                value=data_salva,
                key=(
                    "data_inicio_"
                    f"{chave}"
                ),
                format="DD/MM/YYYY"
            )

            datas_inicio[
                disciplina
            ] = data_selecionada

            if data_selecionada != data_salva:
                salvar_data_inicio_cronograma(
                    biblioteca_original,
                    disciplina,
                    data_selecionada
                )

            elif carregar_data_inicio_cronograma(
                biblioteca_original
            ) is None:
                salvar_data_inicio_cronograma(
                    biblioteca_original,
                    disciplina,
                    data_selecionada
                )

else:
    formulario_login_sidebar()

    st.sidebar.caption(
        "As datas abaixo são somente para consulta. "
        "Faça login para alterá-las."
    )

    for _, item_cronograma in pares_cronograma.iterrows():
        biblioteca_original = str(
            item_cronograma[
                "Biblioteca Original"
            ]
        ).strip()

        disciplina = str(
            item_cronograma[
                "Biblioteca"
            ]
        ).strip()

        data_salva = carregar_data_inicio_cronograma(
            biblioteca_original
        )

        # Se ainda não houver uma data cadastrada, o sistema
        # usa a data atual apenas para manter o cálculo funcional.
        # Essa data não é gravada no banco até um administrador
        # autenticar e realizar a edição.
        data_calculo = (
            data_salva
            if data_salva is not None
            else date.today()
        )

        datas_inicio[
            disciplina
        ] = data_calculo

        with st.sidebar.expander(
            disciplina
        ):
            if data_salva is not None:
                st.write(
                    "**Data de início:** "
                    f"{data_salva.strftime('%d/%m/%Y')}"
                )
            else:
                st.warning(
                    "Data de início ainda não cadastrada."
                )

# ==========================================================
# CALCULA PRAZOS
# ==========================================================

df[
    "Prazo"
] = df.apply(
    lambda row:
    calcular_prazo(
        row[
            "Aula"
        ],
        datas_inicio.get(
            row[
                "Biblioteca"
            ]
        )
    ),
    axis=1
)


# ==========================================================
# ABAS
# ==========================================================

tab_auditoria, tab_producao, tab_config = (
    st.tabs(
        [
            "📊 Auditoria SharePoint",
            "📅 Controle de Produção",
            "⚙️ Configuração do Cronograma",
        ]
    )
)


# ==========================================================
# ABA 1 - AUDITORIA
# ==========================================================

with tab_auditoria:

    st.subheader(
        "📊 Auditoria das entregas"
    )


    # ======================================================
    # FILTROS
    # ======================================================

    st.markdown(
        "### 🔎 Filtros"
    )


    f1, f2, f3, f4 = (
        st.columns(
            4
        )
    )


    lista_disciplinas = sorted(
        df[
            "Biblioteca"
        ].unique()
    )


    lista_status = sorted(
        df[
            "Status Geral"
        ].unique()
    )


    lista_unidades = sorted(
        df[
            "Unidade"
        ].unique()
    )


    with f1:

        bibliotecas = st.multiselect(
            "Disciplina",
            lista_disciplinas,
            default=lista_disciplinas,
            key=(
                "auditoria_"
                "filtro_disciplina"
            )
        )


    with f2:

        status = st.multiselect(
            "Status Geral",
            lista_status,
            default=lista_status,
            key=(
                "auditoria_"
                "filtro_status"
            )
        )


    with f3:

        unidades = st.multiselect(
            "Unidade",
            lista_unidades,
            default=lista_unidades,
            key=(
                "auditoria_"
                "filtro_unidade"
            )
        )


    with f4:

        busca = st.text_input(
            "Buscar",
            key=(
                "auditoria_"
                "campo_busca"
            )
        )


    df_filtrado = df[
        df[
            "Biblioteca"
        ].isin(
            bibliotecas
        )
        &
        df[
            "Status Geral"
        ].isin(
            status
        )
        &
        df[
            "Unidade"
        ].isin(
            unidades
        )
    ]


    if busca:

        busca_lower = (
            busca.lower()
        )

        df_filtrado = df_filtrado[
            df_filtrado.apply(
                lambda row:
                busca_lower
                in " ".join(
                    row.astype(
                        str
                    )
                ).lower(),
                axis=1
            )
        ]


    # ======================================================
    # MÉTRICAS
    # ======================================================

    total = len(
        df_filtrado
    )


    completos = len(
        df_filtrado[
            df_filtrado[
                "Status Geral"
            ].astype(
                str
            ).str.contains(
                "completa|aprovado|validado",
                case=False,
                na=False
            )
        ]
    )


    incompletos = len(
        df_filtrado[
            df_filtrado[
                "Status Geral"
            ].astype(
                str
            ).str.contains(
                "incompleto|erro|sem acesso",
                case=False,
                na=False
            )
        ]
    )


    pendentes = len(
        df_filtrado[
            df_filtrado[
                "Pendências"
            ].astype(
                str
            ).str.strip()
            != ""
        ]
    )


    c1, c2, c3, c4 = (
        st.columns(
            4
        )
    )


    c1.metric(
        "Total de registros",
        total
    )


    c2.metric(
        "Completos / Aprovados",
        completos
    )


    c3.metric(
        "Incompletos / Erros",
        incompletos
    )


    c4.metric(
        "Com pendências",
        pendentes
    )


    st.divider()


    # ======================================================
    # GRÁFICOS
    # ======================================================

    g1, g2 = st.columns(
        2
    )


    with g1:

        resumo_status = (
            df_filtrado[
                "Status Geral"
            ]
            .value_counts()
            .reset_index()
        )

        resumo_status.columns = [
            "Status",
            "Quantidade"
        ]

        if not (
            resumo_status.empty
        ):

            fig_status = px.pie(
                resumo_status,
                names="Status",
                values="Quantidade",
                title=(
                    "Distribuição por Status"
                )
            )

            st.plotly_chart(
                fig_status,
                use_container_width=True,
                key=(
                    "grafico_status_"
                    "auditoria"
                )
            )


    with g2:

        resumo_unidade = (
            df_filtrado
            .groupby(
                [
                    "Unidade",
                    "Status Geral"
                ]
            )
            .size()
            .reset_index(
                name="Quantidade"
            )
        )

        if not (
            resumo_unidade.empty
        ):

            fig_unidade = px.bar(
                resumo_unidade,
                x="Unidade",
                y="Quantidade",
                color="Status Geral",
                barmode="group",
                title=(
                    "Status por Unidade"
                )
            )

            st.plotly_chart(
                fig_unidade,
                use_container_width=True,
                key=(
                    "grafico_unidade_"
                    "auditoria"
                )
            )


    st.divider()


    # ======================================================
    # VISÃO RÁPIDA
    # ======================================================

    st.subheader(
        "✅ Visão rápida das entregas"
    )


    colunas_visao = [
        "Biblioteca",
        "Unidade",
        "Aula",
        "Prazo",
        "DOCX",
        "PPT",
        "PDF",
        "Questões",
        "REV1",
        "AJ1",
        "Aprovado",
        "Validado",
        "Status Geral",
        "Pendências",
    ]


    st.dataframe(
        df_filtrado[
            colunas_visao
        ].style.map(
            destacar_status
        ),
        use_container_width=True,
        height=450
    )


    # ======================================================
    # DOWNLOAD
    # ======================================================

    st.divider()


    st.download_button(
        "⬇️ Baixar auditoria filtrada",
        data=(
            df_filtrado
            .to_csv(
                index=False
            )
            .encode(
                "utf-8-sig"
            )
        ),
        file_name=(
            "auditoria_filtrada.csv"
        ),
        mime="text/csv",
        key=(
            "download_"
            "auditoria"
        )
    )


# ==========================================================
# ABA 2 - CONTROLE DE PRODUÇÃO
# ==========================================================

with tab_producao:

    st.subheader(
        "📅 Controle de Produção"
    )


    # ======================================================
    # CARREGA CONTROLE
    # ======================================================

    if (
        "controle_producao"
        not in st.session_state
    ):

        st.session_state[
            "controle_producao"
        ] = carregar_controle(
            df
        )


    controle = (
        st.session_state[
            "controle_producao"
        ].copy()
    )


    # ======================================================
    # SEMPRE ATUALIZA PRAZOS
    # ======================================================

    controle = (
        atualizar_prazos_controle(
            controle,
            df
        )
    )


    controle = (
        atualizar_disciplinas_controle(
            controle,
            df
        )
    )

    # Sincroniza automaticamente os OK encontrados pela auditoria.
    controle = sincronizar_auditoria_controle(
        controle,
        df
    )

    # Preenche Autor e Responsável cadastrados uma única vez.
    controle = aplicar_responsaveis_controle(
        controle,
        carregar_responsaveis()
    )

    controle[
        "Status da Produção"
    ] = controle.apply(
        status_linha,
        axis=1
    )


    # ======================================================
    # MÉTRICAS
    # ======================================================

    qtd_total = len(
        controle
    )


    qtd_concluidos = (
        controle[
            "Status da Produção"
        ]
        .astype(
            str
        )
        .str.contains(
            "Concluído",
            na=False
        )
        .sum()
    )


    qtd_andamento = (
        controle[
            "Status da Produção"
        ]
        .astype(
            str
        )
        .str.contains(
            "andamento",
            case=False,
            na=False
        )
        .sum()
    )


    qtd_proximos = (
        controle[
            "Status da Produção"
        ]
        .astype(
            str
        )
        .str.contains(
            "Próximo",
            case=False,
            na=False
        )
        .sum()
    )


    qtd_atrasados = (
        controle[
            "Status da Produção"
        ]
        .astype(
            str
        )
        .str.contains(
            "Atrasado",
            case=False,
            na=False
        )
        .sum()
    )


    m1, m2, m3, m4, m5 = (
        st.columns(
            5
        )
    )


    m1.metric(
        "📚 Total",
        qtd_total
    )


    m2.metric(
        "✅ Concluídos",
        int(
            qtd_concluidos
        )
    )


    m3.metric(
        "🔵 Em andamento",
        int(
            qtd_andamento
        )
    )


    m4.metric(
        "🟡 Próximos do prazo",
        int(
            qtd_proximos
        )
    )


    m5.metric(
        "🔴 Atrasados",
        int(
            qtd_atrasados
        )
    )


    st.divider()


    # ======================================================
    # FILTROS
    # ======================================================

    f1, f2, f3, f4 = (
        st.columns(
            4
        )
    )


    disciplinas_controle = sorted(
        [
            item
            for item in controle[
                "Disciplina"
            ].unique()
            if str(
                item
            ).strip()
        ]
    )


    autores = sorted(
        [
            item
            for item in controle[
                "Autor"
            ].unique()
            if str(
                item
            ).strip()
        ]
    )


    unidades_controle = sorted(
        [
            item
            for item in controle[
                "Unidade"
            ].unique()
            if str(
                item
            ).strip()
        ]
    )


    status_controle = sorted(
        [
            item
            for item in controle[
                "Status da Produção"
            ].unique()
            if str(
                item
            ).strip()
        ]
    )


    with f1:

        filtro_disciplina = (
            st.multiselect(
                "Disciplina",
                disciplinas_controle,
                default=(
                    disciplinas_controle
                ),
                key=(
                    "controle_"
                    "filtro_disciplina"
                )
            )
        )


    with f2:

        filtro_autor = (
            st.multiselect(
                "Autor",
                autores,
                key=(
                    "controle_"
                    "filtro_autor"
                )
            )
        )


    with f3:

        filtro_unidade = (
            st.multiselect(
                "Unidade",
                unidades_controle,
                default=(
                    unidades_controle
                ),
                key=(
                    "controle_"
                    "filtro_unidade"
                )
            )
        )


    with f4:

        filtro_status = (
            st.multiselect(
                "Status",
                status_controle,
                default=(
                    status_controle
                ),
                key=(
                    "controle_"
                    "filtro_status"
                )
            )
        )


    mascara = pd.Series(
        True,
        index=controle.index
    )


    if filtro_disciplina:

        mascara &= controle[
            "Disciplina"
        ].isin(
            filtro_disciplina
        )


    if filtro_autor:

        mascara &= controle[
            "Autor"
        ].isin(
            filtro_autor
        )


    if filtro_unidade:

        mascara &= controle[
            "Unidade"
        ].isin(
            filtro_unidade
        )


    if filtro_status:

        mascara &= controle[
            "Status da Produção"
        ].isin(
            filtro_status
        )


    controle_filtrado = (
        controle.loc[
            mascara
        ].copy()
    )


    controle_filtrado[
        "_indice_original"
    ] = controle_filtrado.index


    # ======================================================
    # EDITOR
    # ======================================================

    st.markdown(
        "### 📝 Cronograma de Produção"
    )


    config_colunas = {

        "Biblioteca Original":
            None,

        "_indice_original":
            None,

        "Autor":
            st.column_config.TextColumn(
                "Autor",
                disabled=True
            ),

        "Responsável":
            st.column_config.TextColumn(
                "Responsável",
                disabled=True
            ),

        "Prazo":
            st.column_config.TextColumn(
                "Prazo",
                disabled=True
            ),

        "Texto":
            st.column_config.SelectboxColumn(
                "Texto",
                options=STATUS_PRODUCAO
            ),

        "Relatório Antiplágio":
            st.column_config.SelectboxColumn(
                "Relatório Antiplágio",
                options=STATUS_PRODUCAO
            ),

        "PPT":
            st.column_config.SelectboxColumn(
                "PPT",
                options=STATUS_PRODUCAO
            ),

        "Questões":
            st.column_config.SelectboxColumn(
                "Questões",
                options=STATUS_PRODUCAO
            ),

        "Roteiro Podcast":
            st.column_config.SelectboxColumn(
                "Roteiro Podcast",
                options=STATUS_PRODUCAO
            ),

        "RAP":
            st.column_config.SelectboxColumn(
                "RAP",
                options=STATUS_PRODUCAO
            ),

        "Status da Produção":
            st.column_config.TextColumn(
                "Status da Produção",
                disabled=True
            ),

        "Última Atualização":
            st.column_config.TextColumn(
                "Última Atualização",
                disabled=True
            ),
    }


    controle_editado = (
        st.data_editor(
            controle_filtrado,
            column_config=(
                config_colunas
            ),
            use_container_width=True,
            hide_index=True,
            height=600,
            num_rows="fixed",
            key=(
                "editor_"
                "controle_producao"
            )
        )
    )


    b1, b2, b3 = st.columns(
        [
            1.5,
            1.5,
            5
        ]
    )


    with b1:

        salvar = st.button(
            "💾 Salvar alterações",
            type="primary",
            use_container_width=True,
            key=(
                "salvar_"
                "controle"
            )
        )


    with b2:

        atualizar = st.button(
            "🔄 Atualizar status",
            use_container_width=True,
            key=(
                "atualizar_"
                "controle"
            )
        )


    # ======================================================
    # SALVAR
    # ======================================================

    if salvar:

        for _, linha in (
            controle_editado
            .iterrows()
        ):

            indice_original = int(
                linha[
                    "_indice_original"
                ]
            )

            for coluna in (
                COLUNAS_CONTROLE
            ):

                if coluna in (
                    linha.index
                ):

                    controle.at[
                        indice_original,
                        coluna
                    ] = linha[
                        coluna
                    ]

        controle[
            "Status da Produção"
        ] = controle.apply(
            status_linha,
            axis=1
        )

        controle = (
            salvar_controle(
                controle
            )
        )

        st.session_state[
            "controle_producao"
        ] = controle

        st.success(
            "✅ Alterações salvas."
        )

        st.rerun()


    # ======================================================
    # ATUALIZAR STATUS
    # ======================================================

    if atualizar:

        controle[
            "Status da Produção"
        ] = controle.apply(
            status_linha,
            axis=1
        )

        st.session_state[
            "controle_producao"
        ] = controle

        st.rerun()


    # ======================================================
    # ATRASADOS
    # ======================================================

    atrasados = controle[
        controle[
            "Status da Produção"
        ]
        .astype(
            str
        )
        .str.contains(
            "Atrasado",
            case=False,
            na=False
        )
    ]


    if not (
        atrasados.empty
    ):

        st.divider()

        st.subheader(
            "🚨 Entregas atrasadas"
        )

        st.dataframe(
            atrasados[
                [
                    "Disciplina",
                    "Autor",
                    "Responsável",
                    "Revisor",
                    "Unidade",
                    "Aula",
                    "Prazo",
                    "Status da Produção",
                ]
            ],
            use_container_width=True,
            hide_index=True
        )


# ==========================================================
# ABA 3 - CONFIGURAÇÃO
# ==========================================================

with tab_config:

    if not esta_autenticado():
        st.subheader(
            "⚙️ Configuração do Cronograma"
        )
        st.warning(
            "🔒 Esta área é restrita aos administradores."
        )
        formulario_login(
            "configuracao_cronograma"
        )

    else:
        botao_logout(
            "configuracao_cronograma"
        )

        st.subheader(
            "⚙️ Configuração do Cronograma"
        )


        # ======================================================
        # NOMES DAS DISCIPLINAS
        # ======================================================

        st.markdown(
            "## 🏷️ Nomes das disciplinas"
        )


        st.caption(
            "O dashboard prioriza automaticamente o nome da disciplina "
            "recebido do SharePoint. O mapeamento manual permanece apenas "
            "como fallback para relatórios antigos."
        )


        editor_mapeamento = (
            st.data_editor(
                tabela_mapeamento,
                column_config={

                    "Biblioteca":
                        st.column_config.TextColumn(
                            "Biblioteca",
                            disabled=True
                        ),

                    "Disciplina automática":
                        st.column_config.TextColumn(
                            "Disciplina automática",
                            disabled=True
                        ),

                    "Nome da disciplina":
                        st.column_config.TextColumn(
                            "Nome da disciplina",
                            required=True
                        ),
                },
                hide_index=True,
                use_container_width=True,
                key=(
                    "editor_"
                    "mapeamento"
                )
            )
        )


        if st.button(
            "💾 Salvar nomes das disciplinas",
            type="primary",
            key=(
                "salvar_"
                "nomes_disciplinas"
            )
        ):

            novo_mapa = {}

            for _, row in (
                editor_mapeamento
                .iterrows()
            ):

                biblioteca = str(
                    row[
                        "Biblioteca"
                    ]
                ).strip()

                disciplina = str(
                    row[
                        "Nome da disciplina"
                    ]
                ).strip()

                if not disciplina:

                    disciplina = (
                        biblioteca
                    )

                novo_mapa[
                    biblioteca
                ] = disciplina

            salvou, destino = salvar_mapeamento(
                novo_mapa
            )

            if salvou:

                if (
                    "controle_producao"
                    in st.session_state
                ):

                    del st.session_state[
                        "controle_producao"
                    ]

                if destino == "github":
                    st.success(
                        "✅ Nomes salvos no GitHub com sucesso."
                    )
                else:
                    st.success(
                        "✅ Nomes salvos localmente. "
                        "Configure os Secrets do GitHub para persistência no Streamlit Cloud."
                    )

                st.rerun()


        st.divider()

        # ======================================================
        # AUTOR E RESPONSÁVEL POR DISCIPLINA
        # ======================================================

        st.markdown(
            "## 👥 Autor e responsável por disciplina"
        )

        st.caption(
            "Cadastre uma única vez. O dashboard reaplica automaticamente "
            "Autor e Responsável a todas as aulas da disciplina."
        )

        cadastro_responsaveis = carregar_responsaveis()
        linhas_responsaveis = []

        for disciplina in disciplinas_unicas:
            chave = normalizar_chave_disciplina(disciplina)
            dados_resp = cadastro_responsaveis.get(chave, {})

            if not isinstance(dados_resp, dict):
                dados_resp = {}

            linhas_responsaveis.append({
                "Disciplina": disciplina,
                "Autor/Revisor": str(dados_resp.get("autor", "")),
                "Responsável": str(dados_resp.get("responsavel", "")),
            })

        df_responsaveis = pd.DataFrame(linhas_responsaveis)

        editor_responsaveis = st.data_editor(
            df_responsaveis,
            column_config={
                "Disciplina": st.column_config.TextColumn(
                    "Disciplina",
                    disabled=True,
                ),
                "Autor": st.column_config.TextColumn(
                    "Autor",
                ),
                "Responsável": st.column_config.TextColumn(
                    "Responsável",
                ),
            },
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            key="editor_responsaveis_disciplinas",
        )

        if st.button(
            "💾 Salvar autor e responsável",
            type="primary",
            key="salvar_responsaveis_disciplinas",
        ):
            novo_cadastro = {}

            for _, row_resp in editor_responsaveis.iterrows():
                disciplina = str(row_resp.get("Disciplina", "")).strip()
                if not disciplina:
                    continue

                chave = normalizar_chave_disciplina(disciplina)
                novo_cadastro[chave] = {
                    "autor": str(row_resp.get("Autor", "")).strip(),
                    "responsavel": str(row_resp.get("Responsável", "")).strip(),
                }

            salvou_resp, destino_resp = salvar_responsaveis(novo_cadastro)

            if salvou_resp:
                if "controle_producao" in st.session_state:
                    del st.session_state["controle_producao"]

                if destino_resp == "github":
                    st.success(
                        "✅ Autor e responsável salvos no GitHub com sucesso."
                    )
                else:
                    st.success(
                        "✅ Autor e responsável salvos localmente."
                    )

                st.rerun()

        st.divider()


        # ======================================================
        # DATAS INICIAIS
        # ======================================================

        st.markdown(
            "## 🗓️ Datas iniciais"
        )


        dados_datas = []


        for disciplina in (
            disciplinas_unicas
        ):

            dados_datas.append(
                {
                    "Disciplina":
                        disciplina,

                    "Data de início":
                        datas_inicio[
                            disciplina
                        ].strftime(
                            "%d/%m/%Y"
                        )
                }
            )


        df_datas = pd.DataFrame(
            dados_datas
        )


        st.dataframe(
            df_datas,
            use_container_width=True,
            hide_index=True
        )


        st.divider()


        # ======================================================
        # PRAZOS CALCULADOS
        # ======================================================

        st.markdown(
            "## 🗓️ Prazos calculados"
        )


        cronograma_completo = df[
            [
                "Biblioteca",
                "Unidade",
                "Aula",
                "Prazo"
            ]
        ].copy()


        cronograma_completo.columns = [
            "Disciplina",
            "Unidade",
            "Aula",
            "Prazo"
        ]


        st.dataframe(
            cronograma_completo,
            use_container_width=True,
            hide_index=True,
            height=600
        )

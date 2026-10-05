"""
Script 1 - Geração da newsletter diária (versão RSS, 100% gratuita)

Fluxo:
1. Coleta manchetes e resumos recentes de feeds RSS, por seção.
2. Lê o log dos últimos 7 dias na planilha (antirrepetição).
3. Chamada 1 ao Gemini (editor): escolhe, por seção, a matéria principal e até 2 complementares.
4. O script abre e lê o texto completo das matérias escolhidas (se o site bloquear ou for
   paywall, usa só manchete e resumo, e avisa o Gemini do nível de detalhe disponível).
5. Chamada 2 ao Gemini (redator): escreve as 9 seções a partir do texto lido.
6. Valida o JSON, salva a edição na aba "edicoes" e os temas na aba "log".

Variáveis de ambiente (GitHub Secrets):
- GEMINI_API_KEY
- GOOGLE_SERVICE_ACCOUNT_JSON  (conteúdo inteiro do .json da service account)
- SPREADSHEET_ID

Abas da planilha:
- "log":     data | secao | tema
- "edicoes": data | conteudo_json

Para trocar ou acrescentar fontes, edite só o dicionário FEEDS abaixo.
Feeds que falharem são pulados e listados no log da execução.
"""

import os
import re
import json
import time
import html
import calendar
import datetime

import requests
import feedparser
import trafilatura
import gspread
from google import genai
from google.genai import types
from google.oauth2.service_account import Credentials

MODELO = "gemini-3.1-flash-lite"
DIAS_DE_HISTORICO = 7
JANELA_HORAS = 48          # só entram notícias das últimas 48h
MAX_POR_FEED = 8
MAX_POR_SECAO = 30
TIMEOUT_FEED = 15
USER_AGENT = "Mozilla/5.0 (compatible; DebateNewsletterBot/1.0)"
TIMEOUT_MATERIA = 20
MAX_CHARS_PRINCIPAL = 7000       # texto lido da matéria principal
MAX_CHARS_COMPLEMENTAR = 3500    # texto lido de cada complementar
MIN_CHARS_INTEGRAL = 700         # abaixo disso, considera que não leu a matéria (paywall/bloqueio)

SECOES = [
    "ia",
    "relacoes_internacionais",
    "portugal",
    "politica_brasileira",
    "meio_ambiente",
    "economia",
    "cultura",
    "movimentos_sociais",
    "ciencia",
]

# (nome do veículo, url do feed). Só entram veículos com RSS aberto.
FEEDS = {
    "ia": [
        ("Tecnoblog", "https://tecnoblog.net/feed/"),
        ("MIT Technology Review", "https://www.technologyreview.com/feed/"),
        ("The Verge", "https://www.theverge.com/rss/index.xml"),
        ("Wired", "https://www.wired.com/feed/rss"),
        ("Ars Technica", "https://feeds.arstechnica.com/arstechnica/technology-lab"),
        ("Rest of World", "https://restofworld.org/feed/latest"),
        ("Folha (Tec)", "https://feeds.folha.uol.com.br/tec/rss091.xml"),
    ],
    "relacoes_internacionais": [
        ("Folha (Mundo)", "https://feeds.folha.uol.com.br/mundo/rss091.xml"),
        ("BBC News Brasil", "https://feeds.bbci.co.uk/portuguese/rss.xml"),
        ("BBC World", "https://feeds.bbci.co.uk/news/world/rss.xml"),
        ("Al Jazeera", "https://www.aljazeera.com/xml/rss/all.xml"),
        ("The Guardian (World)", "https://www.theguardian.com/world/rss"),
        ("Le Monde (International)", "https://www.lemonde.fr/en/international/rss_full.xml"),
        ("The Diplomat", "https://thediplomat.com/feed/"),
        ("Foreign Policy", "https://foreignpolicy.com/feed/"),
        ("Nexo", "https://www.nexojornal.com.br/rss.xml"),
    ],
    "portugal": [
        ("Público", "https://feeds.feedburner.com/PublicoRSS"),
        ("Observador", "https://observador.pt/feed/"),
        ("RTP Notícias", "https://www.rtp.pt/noticias/rss"),
        ("ECO", "https://eco.sapo.pt/feed/"),
        ("Politico Europe", "https://www.politico.eu/feed/"),
    ],
    "politica_brasileira": [
        ("Agência Brasil (Política)", "https://agenciabrasil.ebc.com.br/rss/politica/feed.xml"),
        ("Folha (Poder)", "https://feeds.folha.uol.com.br/poder/rss091.xml"),
        ("Poder360", "https://www.poder360.com.br/feed/"),
        ("G1 (Política)", "https://g1.globo.com/rss/g1/politica/"),
        ("Agência Pública", "https://apublica.org/feed/"),
        ("Congresso em Foco", "https://congressoemfoco.uol.com.br/feed/"),
        ("piauí", "https://piaui.folha.uol.com.br/feed/"),
        ("Intercept Brasil", "https://www.intercept.com.br/feed/"),
    ],
    "meio_ambiente": [
        ("((o))eco", "https://oeco.org.br/feed/"),
        ("Mongabay", "https://news.mongabay.com/feed/"),
        ("Carbon Brief", "https://www.carbonbrief.org/feed/"),
        ("The Guardian (Environment)", "https://www.theguardian.com/environment/rss"),
        ("Grist", "https://grist.org/feed/"),
        ("Climate Home News", "https://www.climatechangenews.com/feed/"),
        ("Observatório do Clima", "https://oc.eco.br/feed/"),
        ("Folha (Ambiente)", "https://feeds.folha.uol.com.br/ambiente/rss091.xml"),
        ("Inside Climate News", "https://insideclimatenews.org/feed/"),
        ("InfoAmazonia", "https://infoamazonia.org/feed/"),
    ],
    "economia": [
        ("Folha (Mercado)", "https://feeds.folha.uol.com.br/mercado/rss091.xml"),
        ("Agência Brasil (Economia)", "https://agenciabrasil.ebc.com.br/rss/economia/feed.xml"),
        ("InfoMoney", "https://www.infomoney.com.br/feed/"),
        ("Brazil Journal", "https://braziljournal.com/feed/"),
        ("Exame", "https://exame.com/feed/"),
        ("Valor Econômico", "https://pox.globo.com/rss/valor"),
        ("Project Syndicate", "https://www.project-syndicate.org/rss"),
        ("The Guardian (Business)", "https://www.theguardian.com/business/rss"),
        ("BBC Business", "https://feeds.bbci.co.uk/news/business/rss.xml"),
        ("Nikkei Asia", "https://asia.nikkei.com/rss/feed/nar"),
    ],
    "cultura": [
        ("Folha (Ilustrada)", "https://feeds.folha.uol.com.br/ilustrada/rss091.xml"),
        ("Agência Brasil (Cultura)", "https://agenciabrasil.ebc.com.br/rss/cultura/feed.xml"),
        ("The Guardian (Culture)", "https://www.theguardian.com/culture/rss"),
        ("The New York Times (Arts)", "https://rss.nytimes.com/services/xml/rss/nyt/Arts.xml"),
        ("The New Yorker", "https://www.newyorker.com/feed/everything"),
        ("The Atlantic", "https://www.theatlantic.com/feed/all/"),
        ("Variety", "https://variety.com/feed/"),
        ("Pitchfork", "https://pitchfork.com/feed/feed-news/rss"),
        ("ARTnews", "https://www.artnews.com/feed/"),
    ],
    "movimentos_sociais": [
        ("Brasil de Fato", "https://www.brasildefato.com.br/rss2.xml"),
        ("Agência Pública", "https://apublica.org/feed/"),
        ("Repórter Brasil", "https://reporterbrasil.org.br/feed/"),
        ("Ponte Jornalismo", "https://ponte.org/feed/"),
        ("AzMina", "https://azmina.com.br/feed/"),
        ("Agência Mural", "https://www.agenciamural.org.br/feed/"),
        ("Amazônia Real", "https://amazoniareal.com.br/feed/"),
        ("Global Voices", "https://globalvoices.org/feed/"),
        ("Waging Nonviolence", "https://wagingnonviolence.org/feed/"),
        ("openDemocracy", "https://www.opendemocracy.net/en/rss.xml"),
        ("Democracy Now!", "https://www.democracynow.org/democracynow.rss"),
        ("The Guardian (Global Development)", "https://www.theguardian.com/global-development/rss"),
        ("Jacobin", "https://jacobin.com/feed"),
    ],
    "ciencia": [
        ("Nature (News)", "https://www.nature.com/nature.rss"),
        ("Science (News)", "https://www.science.org/rss/news_current.xml"),
        ("Scientific American", "https://rss.sciam.com/ScientificAmerican-Global"),
        ("New Scientist", "https://www.newscientist.com/feed/home/"),
        ("Quanta Magazine", "https://www.quantamagazine.org/feed/"),
        ("Ars Technica (Science)", "https://feeds.arstechnica.com/arstechnica/science"),
        ("The Conversation", "https://theconversation.com/global/articles.atom"),
        ("Phys.org", "https://phys.org/rss-feed/"),
        ("Jornal da USP", "https://jornal.usp.br/feed/"),
        ("Folha (Ciência)", "https://feeds.folha.uol.com.br/ciencia/rss091.xml"),
        ("STAT News", "https://www.statnews.com/feed/"),
    ],
}

PROMPT_SELECAO = """Você é o editor-chefe de uma newsletter diária para uma comunidade de debate competitivo. Para cada uma das 9 seções abaixo, escolha qual notícia será a matéria do dia, entre as NOTÍCIAS COLETADAS.

SEÇÕES: ia, relacoes_internacionais, portugal, politica_brasileira, meio_ambiente, economia, cultura, movimentos_sociais (grupos ativistas e movimentos organizados, nacionais e internacionais: o que estão fazendo, pautas, mobilizações, declarações e tensões; NÃO divulgação de oportunidades, editais, vagas ou eventos), ciencia.

CRITÉRIOS
- Relevância para quem debate: há um dilema, uma decisão, um conflito de interesses ou um mecanismo a explicar.
- Novidade: fato do dia ou dos últimos dois dias.
- NÃO repetir tema ou ângulo já usado na mesma seção (LOG abaixo).
- Prefira reportagem (fatos, atores, números) a nota curta, release ou opinião.
- Escolha 1 matéria principal e, se houver no material, até 2 complementares que cubram o mesmo fato em outros veículos ou com enquadramento contrastante.
- Se nada for utilizável (vazio, fraco ou repetido), deixe "principal" vazio, "overview" true e sugira em "tema_overview" um conceito ou mecanismo da área que não repita o log.

SAÍDA: apenas JSON válido, uma chave para cada uma das 9 seções, neste formato:
{
  "ia": {"principal": "ia-03", "complementares": ["ia-07"], "overview": false, "tema_overview": ""},
  "relacoes_internacionais": {...}
}
Use somente ids que existam em NOTÍCIAS COLETADAS.

LOG DOS ÚLTIMOS 7 DIAS, POR SEÇÃO:
__LOG__

NOTÍCIAS COLETADAS, POR SEÇÃO:
__NOTICIAS__

DATA DE HOJE: __DATA__
"""

PROMPT_REDACAO = """Você é o redator de uma newsletter diária para uma comunidade de debate competitivo. Produza, hoje, uma edição estruturada em exatamente 9 seções temáticas.

SEÇÕES (gere uma entrada para cada uma das 9):
1. ia
2. relacoes_internacionais
3. portugal
4. politica_brasileira
5. meio_ambiente
6. economia
7. cultura
8. movimentos_sociais (cobertura de grupos ativistas e movimentos organizados, nacionais e internacionais: o que estão fazendo, pautas que estão priorizando, mobilizações, declarações e tensões internas ou com o poder público. NÃO é espaço para divulgação de oportunidades, editais, vagas ou eventos)
9. ciencia

FONTE DE INFORMAÇÃO (regra mais importante)
Você NÃO tem acesso à internet. Para cada seção você recebe, em MATÉRIAS SELECIONADAS, o texto de 1 a 3 matérias já escolhidas pelo editor. Quando "nivel" for "integral", o texto é a matéria lida de verdade; quando for "resumo", só há manchete e resumo curto.
- Leia cada matéria inteira antes de escrever. Extraia dela nomes, cargos, números, datas, o que cada ator disse ou fez e as posições em disputa. Parafraseie sempre; nunca copie trechos longos.
- Use a matéria principal como espinha do "fato" e do "mecanismo". Use as complementares para ampliar fatos e, sobretudo, para as "visoes": quando veículos diferentes enquadram o fato de modo diferente, atribua cada enquadramento ao veículo.
- Com "nivel" = "resumo", escreva só o que consta e explique o mecanismo em termos gerais; não invente detalhes.
- Complemente com conhecimento geral de contexto e mecanismos apenas quando tiver segurança. Nunca invente nomes, números, datas, citações ou eventos.
- Se a seção vier com "overview": true (sem matérias), faça um overview didático do conceito indicado em "tema_overview", com a mesma densidade, e marque "tipo": "overview".

REGRA DE ANTIRREPETIÇÃO
O editor já evitou repetir temas do LOG DOS ÚLTIMOS 7 DIAS (abaixo). Escolha um ângulo da matéria que também não repita o que consta no log para a seção.

NÍVEL DE PROFUNDIDADE EXIGIDO (isso é o que mais importa)
Cada seção precisa ter a densidade de uma matéria bem apurada, não um resumo genérico. Isso significa:
- Use nomes próprios reais (pessoas, partidos, instituições, empresas) e números concretos (datas, percentuais, valores) presentes nas matérias.
- Em "mecanismo", cada ponto numerado deve explicar uma causa, etapa ou força em jogo específica, não uma frase vaga.
- Em "visoes", atribua cada posição a atores concretos (partidos, blocos, especialistas, instituições nomeadas), não a "alguns" ou "especialistas dizem". Quando uma posição vier de um veículo com linha editorial declarada, diga qual veículo.
- Em "impactos", cada item deve ser uma corrente causal explícita: evento A leva a consequência B, que por sua vez abre ou fecha possibilidade C. Não liste efeitos soltos e desconectados.

ESTRUTURA DE CADA SEÇÃO
- titulo: manchete curta e específica, no estilo de jornal.
- tipo: "noticia" ou "overview".
- tema: frase curta (até 10 palavras) resumindo o ângulo escolhido, para o log de antirrepetição.
- fato: um parágrafo denso (4 a 6 frases): o que aconteceu, quem fez o quê, quando, resultado imediato.
- mecanismo: lista de 3 a 5 textos, cada um uma causa, etapa ou força em jogo específica.
- visoes: lista de 2 ou mais posições, cada uma com "rotulo" (quem defende) e "pontos" (1 a 3 argumentos).
- impactos: lista de 2 a 4 correntes causais encadeadas, cada uma usável como premissa de debate.
- fontes: lista de {"nome": veículo, "url": link} das matérias que você de fato usou, copiando os links de MATÉRIAS SELECIONADAS. Lista vazia se for overview.

SAÍDA: responda SOMENTE com JSON válido, sem texto antes ou depois, neste formato:
{
  "ia": {
    "titulo": "...", "tipo": "noticia", "tema": "...", "fato": "...",
    "mecanismo": ["...", "..."],
    "visoes": [{"rotulo": "...", "pontos": ["..."]}],
    "impactos": ["...", "..."],
    "fontes": [{"nome": "...", "url": "..."}]
  },
  "relacoes_internacionais": {...},
  "portugal": {...},
  "politica_brasileira": {...},
  "meio_ambiente": {...},
  "economia": {...},
  "cultura": {...},
  "movimentos_sociais": {...},
  "ciencia": {...}
}

LOG DOS ÚLTIMOS 7 DIAS, POR SEÇÃO:
__LOG__

MATÉRIAS SELECIONADAS, POR SEÇÃO:
__MATERIAL__

DATA DE HOJE: __DATA__
"""

CAMPOS_OBRIGATORIOS = ["titulo", "tema", "fato", "mecanismo", "visoes", "impactos"]


# ---------- Planilha ----------

def conectar_planilha():
    credenciais_dict = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    escopos = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credenciais = Credentials.from_service_account_info(credenciais_dict, scopes=escopos)
    cliente = gspread.authorize(credenciais)
    return cliente.open_by_key(os.environ["SPREADSHEET_ID"])


def carregar_log(planilha):
    linhas = planilha.worksheet("log").get_all_records()
    limite = datetime.date.today() - datetime.timedelta(days=DIAS_DE_HISTORICO)
    log = {secao: [] for secao in SECOES}
    for linha in linhas:
        try:
            data_linha = datetime.date.fromisoformat(str(linha["data"]))
        except (KeyError, ValueError):
            continue
        if data_linha < limite:
            continue
        secao, tema = linha.get("secao"), linha.get("tema")
        if secao in log and tema:
            log[secao].append(tema)
    return log


def salvar_edicao(planilha, data_hoje, edicao):
    planilha.worksheet("edicoes").append_row(
        [str(data_hoje), json.dumps(edicao, ensure_ascii=False)]
    )


def atualizar_log(planilha, data_hoje, edicao):
    novas = [
        [str(data_hoje), secao, edicao[secao]["tema"]]
        for secao in SECOES
        if edicao.get(secao, {}).get("tema")
    ]
    if novas:
        planilha.worksheet("log").append_rows(novas)


# ---------- Coleta de notícias (RSS) ----------

def limpar_texto(texto, limite=350):
    texto = re.sub(r"<[^>]+>", " ", texto or "")
    texto = html.unescape(re.sub(r"\s+", " ", texto)).strip()
    return texto[:limite]


def data_da_entrada(entrada):
    parsed = entrada.get("published_parsed") or entrada.get("updated_parsed")
    if not parsed:
        return None
    return datetime.datetime.fromtimestamp(calendar.timegm(parsed), tz=datetime.timezone.utc)


def coletar_feed(nome, url, agora):
    resposta = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=TIMEOUT_FEED)
    resposta.raise_for_status()
    feed = feedparser.parse(resposta.content)
    corte = agora - datetime.timedelta(hours=JANELA_HORAS)

    itens, sem_data = [], 0
    for entrada in feed.entries:
        publicado = data_da_entrada(entrada)
        if publicado is None:
            sem_data += 1
            if sem_data > 2:  # sem data, aceita no máximo 2 por feed
                continue
        elif publicado < corte:
            continue
        titulo = limpar_texto(entrada.get("title"), 200)
        if not titulo:
            continue
        itens.append({
            "fonte": nome,
            "titulo": titulo,
            "resumo": limpar_texto(entrada.get("summary") or entrada.get("description")),
            "link": entrada.get("link", ""),
            "publicado": publicado.isoformat() if publicado else "",
        })
    itens.sort(key=lambda i: i["publicado"], reverse=True)
    return itens[:MAX_POR_FEED]


def coletar_noticias():
    agora = datetime.datetime.now(datetime.timezone.utc)
    noticias = {}
    for secao in SECOES:
        itens, vistos, falhas = [], set(), []
        for nome, url in FEEDS[secao]:
            try:
                for item in coletar_feed(nome, url, agora):
                    chave = item["titulo"].lower()
                    if chave in vistos:
                        continue
                    vistos.add(chave)
                    itens.append(item)
            except Exception as erro:
                falhas.append(nome)
                print(f"[AVISO] feed falhou ({secao}) {nome}: {url} -> {type(erro).__name__}: {erro}")
        itens.sort(key=lambda i: i["publicado"], reverse=True)
        noticias[secao] = itens[:MAX_POR_SECAO]
        for n, item in enumerate(noticias[secao], start=1):
            item["id"] = f"{secao}-{n:02d}"
        print(f"[COLETA] {secao}: {len(noticias[secao])} notícias, {len(falhas)} feeds com falha")
    return noticias


def indexar(noticias):
    return {item["id"]: item for itens in noticias.values() for item in itens}


# ---------- Gemini ----------

def interpretar_json(texto):
    texto = texto.strip()
    if texto.startswith("```"):
        texto = texto.strip("`")
        if texto.lower().startswith("json"):
            texto = texto[4:]
        texto = texto.strip()
    return json.loads(texto)


def chamar_gemini(prompt, max_tokens=30000, tentativas=3):
    client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])
    ultimo_erro = None
    for n in range(1, tentativas + 1):
        try:
            resposta = client.models.generate_content(
                model=MODELO,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    max_output_tokens=max_tokens,
                ),
            )
            return interpretar_json(resposta.text)
        except Exception as erro:
            ultimo_erro = erro
            print(f"[AVISO] tentativa {n}/{tentativas} falhou: {type(erro).__name__}: {erro}")
            if n < tentativas:
                time.sleep(20 * n)
    raise ultimo_erro


# ---------- Etapa 1: seleção das matérias ----------

def selecionar_materias(log, noticias, data_hoje):
    enxuto = {
        secao: [
            {
                "id": i["id"],
                "fonte": i["fonte"],
                "titulo": i["titulo"],
                "resumo": i["resumo"][:250],
                "publicado": i["publicado"][:10],
            }
            for i in noticias[secao]
        ]
        for secao in SECOES
    }
    prompt = (
        PROMPT_SELECAO
        .replace("__LOG__", json.dumps(log, ensure_ascii=False, indent=2))
        .replace("__NOTICIAS__", json.dumps(enxuto, ensure_ascii=False))
        .replace("__DATA__", str(data_hoje))
    )
    bruto = chamar_gemini(prompt, max_tokens=6000)
    return normalizar_selecao(bruto, noticias)


def normalizar_selecao(bruto, noticias):
    """Descarta ids inexistentes e garante uma entrada válida para cada seção."""
    validos = indexar(noticias)
    resultado = {}
    for secao in SECOES:
        item = bruto.get(secao) or {}
        principal = item.get("principal")
        if principal not in validos or not principal.startswith(secao + "-"):
            principal = None
        compl = [
            i for i in (item.get("complementares") or [])
            if i in validos and i != principal and i.startswith(secao + "-")
        ][:2]
        if principal is None and compl:
            principal, compl = compl[0], compl[1:]
        resultado[secao] = {
            "principal": principal,
            "complementares": compl if principal else [],
            "overview": principal is None,
            "tema_overview": item.get("tema_overview", "") or "",
        }
    return resultado


# ---------- Etapa 2: leitura das matérias ----------

def ler_materia(item, limite):
    """Abre o link e extrai o texto. Se falhar ou parecer cortado, devolve só o resumo."""
    texto = ""
    try:
        resposta = requests.get(
            item["link"], headers={"User-Agent": USER_AGENT}, timeout=TIMEOUT_MATERIA
        )
        resposta.raise_for_status()
        texto = trafilatura.extract(
            resposta.content, include_comments=False, include_tables=False, favor_precision=True
        ) or ""
    except Exception as erro:
        print(f"[AVISO] não consegui ler {item['fonte']}: {item['link']} -> {type(erro).__name__}: {erro}")
    if len(texto) >= MIN_CHARS_INTEGRAL:
        return texto[:limite], "integral"
    return item["resumo"], "resumo"


def reunir_material(selecao, noticias):
    validos = indexar(noticias)
    material = {}
    for secao in SECOES:
        escolha = selecao[secao]
        pares = [("principal", escolha["principal"])] + [
            ("complementar", i) for i in escolha["complementares"]
        ]
        lista = []
        for papel, ident in pares:
            if not ident:
                continue
            item = validos[ident]
            limite = MAX_CHARS_PRINCIPAL if papel == "principal" else MAX_CHARS_COMPLEMENTAR
            texto, nivel = ler_materia(item, limite)
            lista.append({
                "papel": papel,
                "fonte": item["fonte"],
                "titulo": item["titulo"],
                "link": item["link"],
                "nivel": nivel,
                "texto": texto,
            })
            time.sleep(1)
        # se a principal não pôde ser lida mas uma complementar sim, a lida vira principal
        lista.sort(key=lambda m: m["nivel"] != "integral")
        for pos, materia in enumerate(lista):
            materia["papel"] = "principal" if pos == 0 else "complementar"
        material[secao] = {
            "overview": escolha["overview"],
            "tema_overview": escolha["tema_overview"],
            "materias": lista,
        }
        lidas = sum(1 for m in lista if m["nivel"] == "integral")
        print(f"[LEITURA] {secao}: {lidas}/{len(lista)} matérias lidas na íntegra")
    return material


# ---------- Etapa 3: redação ----------

def montar_prompt_redacao(log, material, data_hoje):
    return (
        PROMPT_REDACAO
        .replace("__LOG__", json.dumps(log, ensure_ascii=False, indent=2))
        .replace("__MATERIAL__", json.dumps(material, ensure_ascii=False))
        .replace("__DATA__", str(data_hoje))
    )


def validar(edicao):
    problemas = []
    for secao in SECOES:
        bloco = edicao.get(secao)
        if not isinstance(bloco, dict):
            problemas.append(f"{secao}: ausente")
            continue
        faltando = [c for c in CAMPOS_OBRIGATORIOS if not bloco.get(c)]
        if faltando:
            problemas.append(f"{secao}: faltam {faltando}")
    if problemas:
        raise ValueError("Edição incompleta, nada foi salvo: " + "; ".join(problemas))


# ---------- Execução ----------

def main():
    planilha = conectar_planilha()
    data_hoje = datetime.date.today()

    log = carregar_log(planilha)
    noticias = coletar_noticias()
    selecao = selecionar_materias(log, noticias, data_hoje)
    material = reunir_material(selecao, noticias)
    edicao = chamar_gemini(montar_prompt_redacao(log, material, data_hoje))
    validar(edicao)

    salvar_edicao(planilha, data_hoje, edicao)
    atualizar_log(planilha, data_hoje, edicao)
    print(f"Edição de {data_hoje} gerada e salva com sucesso.")


if __name__ == "__main__":
    main()

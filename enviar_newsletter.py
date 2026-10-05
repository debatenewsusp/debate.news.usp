"""
Script 2 - Envio da newsletter (Gmail SMTP)

Fluxo, a cada execução:
1. Lê a edição de hoje na aba "edicoes".
2. Lê os assinantes ativos na aba "assinantes" e vê quem está na janela de horário.
3. Pula quem já recebeu hoje (aba "envios").
4. Monta um e-mail por pessoa, só com os temas que ela escolheu, e envia pelo Gmail.
5. Registra cada envio na aba "envios".

Variáveis de ambiente (GitHub Secrets):
- GMAIL_ADDRESS              e-mail do projeto, que envia as mensagens
- GMAIL_APP_PASSWORD         senha de app de 16 caracteres (não é a senha da conta)
- URL_APP                    URL do Apps Script publicado (termina em /exec), usada nos links de cancelar
- GOOGLE_SERVICE_ACCOUNT_JSON, SPREADSHEET_ID   (os mesmos do Script 1)
Opcionais:
- REMETENTE_NOME             nome exibido como remetente (padrão: Debate News)
- SITE_URL                   link do site de cadastro, mostrado no rodapé
- DRY_RUN=true               só simula: não envia e-mails e não grava nada
- HORA_SIMULADA=7            testa um horário (hora de São Paulo, 0 a 23) fora do horário real

Os logs do GitHub Actions podem ser públicos: este script nunca imprime e-mails de assinantes.
"""

import os
import ssl
import sys
import json
import time
import html
import smtplib
import datetime
from email.message import EmailMessage
from email.utils import formataddr
from urllib.parse import quote
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials

FUSO = ZoneInfo("America/Sao_Paulo")
HORARIOS = {"7h": 7, "9h": 9, "12h": 12, "14h": 14}
MAX_ATRASO_HORAS = 6       # se um envio atrasar, ainda recupera até 6h depois do horário
LIMITE_DIARIO = 400        # proteção da conta Gmail (o limite do Google é de cerca de 500 por dia)
PAUSA_ENTRE_ENVIOS = 2     # segundos
GRAVAR_ENVIOS_A_CADA = 20

SECOES = [
    "ia", "relacoes_internacionais", "portugal", "politica_brasileira", "meio_ambiente",
    "economia", "cultura", "movimentos_sociais", "ciencia",
]
ROTULOS = {
    "ia": "IA",
    "relacoes_internacionais": "Relações Internacionais",
    "portugal": "Portugal",
    "politica_brasileira": "Política Brasileira",
    "meio_ambiente": "Meio Ambiente",
    "economia": "Economia",
    "cultura": "Cultura",
    "movimentos_sociais": "Movimentos Sociais",
    "ciencia": "Ciência",
}
DIAS_SEMANA = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira",
               "sexta-feira", "sábado", "domingo"]
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]

CABECALHO_ENVIOS = ["data", "email", "horario"]


# ---------- Planilha ----------

def conectar_planilha():
    credenciais_dict = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    escopos = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credenciais = Credentials.from_service_account_info(credenciais_dict, scopes=escopos)
    return gspread.authorize(credenciais).open_by_key(os.environ["SPREADSHEET_ID"])


def carregar_edicao(planilha, data_hoje):
    """Devolve a edição de hoje (a mais recente do dia) ou None."""
    linhas = planilha.worksheet("edicoes").get_all_values()
    for linha in reversed(linhas):
        if linha and linha[0] == str(data_hoje):
            texto = "".join(linha[1:])   # a edição pode estar dividida em várias células
            try:
                return json.loads(texto)
            except json.JSONDecodeError:
                return None
    return None


def carregar_assinantes(planilha):
    try:
        aba = planilha.worksheet("assinantes")
    except gspread.WorksheetNotFound:
        return []
    return aba.get_all_records()


def obter_aba_envios(planilha):
    try:
        return planilha.worksheet("envios")
    except gspread.WorksheetNotFound:
        aba = planilha.add_worksheet(title="envios", rows=1000, cols=len(CABECALHO_ENVIOS))
        aba.append_row(CABECALHO_ENVIOS)
        return aba


def emails_ja_enviados_hoje(aba_envios, data_hoje):
    return {
        str(linha.get("email", "")).strip().lower()
        for linha in aba_envios.get_all_records()
        if str(linha.get("data", "")) == str(data_hoje)
    }


# ---------- Quem recebe agora ----------

def hora_atual(agora):
    simulada = os.environ.get("HORA_SIMULADA", "").strip()
    return int(simulada) if simulada else agora.hour


def escolher_destinatarios(assinantes, ja_enviados, hora):
    """Ativos, na janela do horário escolhido e que ainda não receberam hoje."""
    escolhidos = []
    for a in assinantes:
        email = str(a.get("email", "")).strip().lower()
        status = str(a.get("status", "")).strip().lower()
        horario = str(a.get("horario", "")).strip().lower()
        token = str(a.get("token", "")).strip()
        temas = [t.strip() for t in str(a.get("temas", "")).split(",") if t.strip() in SECOES]
        if not email or status != "ativo" or horario not in HORARIOS or not temas or not token:
            continue
        atraso = hora - HORARIOS[horario]
        if not (0 <= atraso <= MAX_ATRASO_HORAS):
            continue
        if email in ja_enviados:
            continue
        escolhidos.append({"email": email, "temas": temas, "horario": horario, "token": token})
    return escolhidos


# ---------- Montagem do e-mail ----------

def como_texto(valor):
    if valor is None:
        return ""
    if isinstance(valor, str):
        return valor
    if isinstance(valor, dict):
        return " ".join(como_texto(v) for v in valor.values())
    if isinstance(valor, (list, tuple)):
        return " ".join(como_texto(v) for v in valor)
    return str(valor)


def esc(valor):
    return html.escape(como_texto(valor))


def fontes_validas(bloco):
    fontes = []
    for f in bloco.get("fontes") or []:
        if isinstance(f, dict) and str(f.get("url", "")).startswith(("http://", "https://")):
            fontes.append((como_texto(f.get("nome")) or "fonte", str(f["url"])))
    return fontes


def secao_html(chave, bloco):
    estilo_rotulo = "margin:18px 0 6px;font:600 12px Arial,sans-serif;letter-spacing:.06em;color:#555"
    estilo_lista = "margin:0 0 6px;padding-left:22px"
    estilo_item = "margin:0 0 8px"

    partes = [
        f'<h2 style="margin:36px 0 4px;font:700 13px Arial,sans-serif;letter-spacing:.08em;'
        f'text-transform:uppercase;color:#8a3b12">{esc(ROTULOS[chave])}</h2>',
        f'<h3 style="margin:0 0 10px;font:700 21px/1.3 Georgia,serif;color:#111">{esc(bloco.get("titulo"))}</h3>',
        f'<p style="{estilo_rotulo}">O FATO</p>',
        f'<p style="margin:0 0 6px">{esc(bloco.get("fato"))}</p>',
    ]

    mecanismo = bloco.get("mecanismo") or []
    if mecanismo:
        itens = "".join(f'<li style="{estilo_item}">{esc(m)}</li>' for m in mecanismo)
        partes += [f'<p style="{estilo_rotulo}">MECANISMO E CONTEXTO</p>',
                   f'<ol style="{estilo_lista}">{itens}</ol>']

    visoes = bloco.get("visoes") or []
    if visoes:
        partes.append(f'<p style="{estilo_rotulo}">VISÕES E DISCUSSÕES</p>')
        for v in visoes:
            if isinstance(v, dict):
                rotulo, pontos = v.get("rotulo"), v.get("pontos") or []
            else:
                rotulo, pontos = "", [v]
            itens = "".join(f'<li style="{estilo_item}">{esc(p)}</li>' for p in pontos)
            partes += [f'<p style="margin:10px 0 4px"><em>{esc(rotulo)}</em></p>',
                       f'<ol style="{estilo_lista}">{itens}</ol>']

    impactos = bloco.get("impactos") or []
    if impactos:
        itens = "".join(f'<li style="{estilo_item}">{esc(i)}</li>' for i in impactos)
        partes += [f'<p style="{estilo_rotulo}">IMPACTOS</p>',
                   f'<ol style="{estilo_lista}">{itens}</ol>']

    fontes = fontes_validas(bloco)
    if fontes:
        links = " &middot; ".join(
            f'<a href="{html.escape(url, quote=True)}" style="color:#555">{html.escape(nome)}</a>'
            for nome, url in fontes
        )
        partes.append(f'<p style="margin:12px 0 0;font:12px Arial,sans-serif;color:#777">Fontes: {links}</p>')
    return "\n".join(partes)


def secao_texto(chave, bloco):
    linhas = [ROTULOS[chave].upper(), como_texto(bloco.get("titulo")), "", "O FATO", como_texto(bloco.get("fato"))]
    if bloco.get("mecanismo"):
        linhas += ["", "MECANISMO E CONTEXTO"]
        linhas += [f"{n}. {como_texto(m)}" for n, m in enumerate(bloco["mecanismo"], start=1)]
    if bloco.get("visoes"):
        linhas += ["", "VISÕES E DISCUSSÕES"]
        for v in bloco["visoes"]:
            if isinstance(v, dict):
                linhas.append(f"{como_texto(v.get('rotulo'))}:")
                linhas += [f"  {n}. {como_texto(p)}" for n, p in enumerate(v.get("pontos") or [], start=1)]
            else:
                linhas.append(como_texto(v))
    if bloco.get("impactos"):
        linhas += ["", "IMPACTOS"]
        linhas += [f"{n}. {como_texto(i)}" for n, i in enumerate(bloco["impactos"], start=1)]
    fontes = fontes_validas(bloco)
    if fontes:
        linhas += ["", "Fontes: " + " | ".join(f"{nome} ({url})" for nome, url in fontes)]
    return "\n".join(linhas)


def data_por_extenso(d):
    return f"{DIAS_SEMANA[d.weekday()]}, {d.day} de {MESES[d.month - 1]}"


def montar_conteudo(edicao, temas, data_hoje, url_cancelar, site_url=""):
    secoes = [c for c in SECOES if c in temas and isinstance(edicao.get(c), dict)]
    data_txt = data_por_extenso(data_hoje)

    corpo_html = "\n".join(secao_html(c, edicao[c]) for c in secoes)
    rodape_html = (
        f'<p style="margin:0 0 6px">Você recebe este e-mail porque se inscreveu no Debate News.</p>'
        f'<p style="margin:0"><a href="{html.escape(url_cancelar, quote=True)}" style="color:#777">Cancelar inscrição</a>'
        + (f' &middot; <a href="{html.escape(site_url, quote=True)}" style="color:#777">Alterar temas ou horário</a>' if site_url else "")
        + "</p>"
    )
    html_final = (
        '<div style="max-width:640px;margin:0 auto;padding:24px 16px;font:16px/1.6 Georgia,serif;color:#222">'
        '<p style="margin:0;font:700 22px Georgia,serif;color:#111">Debate News</p>'
        f'<p style="margin:2px 0 0;font:14px Arial,sans-serif;color:#777">{esc(data_txt)}</p>'
        f"{corpo_html}"
        '<hr style="margin:36px 0 16px;border:0;border-top:1px solid #ddd">'
        f'<div style="font:12px/1.5 Arial,sans-serif;color:#777">{rodape_html}</div>'
        "</div>"
    )

    texto_final = "\n\n".join(
        [f"DEBATE NEWS\n{data_txt}"]
        + [secao_texto(c, edicao[c]) for c in secoes]
        + ["---\nVocê recebe este e-mail porque se inscreveu no Debate News.\n"
           f"Cancelar inscrição: {url_cancelar}"
           + (f"\nAlterar temas ou horário: {site_url}" if site_url else "")]
    )
    return f"Debate News | {data_txt}", texto_final, html_final


def montar_mensagem(nome, remetente, destino, assunto, texto, corpo_html, url_cancelar):
    msg = EmailMessage()
    msg["From"] = formataddr((nome, remetente))
    msg["To"] = destino
    msg["Subject"] = assunto
    msg["List-Unsubscribe"] = f"<{url_cancelar}>"
    msg.set_content(texto)
    msg.add_alternative(corpo_html, subtype="html")
    return msg


# ---------- Envio ----------

def enviar(destinatarios, edicao, data_hoje, aba_envios):
    remetente = os.environ["GMAIL_ADDRESS"].strip()
    senha = os.environ["GMAIL_APP_PASSWORD"].replace(" ", "")
    nome = os.environ.get("REMETENTE_NOME", "Debate News")
    url_app = os.environ["URL_APP"].strip()
    site_url = os.environ.get("SITE_URL", "").strip()

    enviados, falhas, pendentes = 0, 0, []

    def gravar():
        if pendentes and aba_envios is not None:
            aba_envios.append_rows(pendentes, value_input_option="RAW")
            pendentes.clear()

    servidor = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context())
    try:
        try:
            servidor.login(remetente, senha)
        except smtplib.SMTPAuthenticationError:
            print("[ERRO] O Gmail recusou o login. Confira GMAIL_ADDRESS e GMAIL_APP_PASSWORD "
                  "(precisa ser uma senha de app, com verificação em duas etapas ativa).")
            sys.exit(1)

        for n, d in enumerate(destinatarios, start=1):
            url_cancelar = f"{url_app}?acao=cancelar&t={quote(d['token'])}"
            assunto, texto, corpo_html = montar_conteudo(edicao, d["temas"], data_hoje, url_cancelar, site_url)
            msg = montar_mensagem(nome, remetente, d["email"], assunto, texto, corpo_html, url_cancelar)
            try:
                servidor.send_message(msg)
                enviados += 1
                pendentes.append([str(data_hoje), d["email"], d["horario"]])
            except Exception as erro:
                falhas += 1
                print(f"[AVISO] falha no envio #{n}: {type(erro).__name__}")
            if len(pendentes) >= GRAVAR_ENVIOS_A_CADA:
                gravar()
            time.sleep(PAUSA_ENTRE_ENVIOS)
    finally:
        gravar()
        try:
            servidor.quit()
        except Exception:
            pass
    return enviados, falhas


# ---------- Execução ----------

def main():
    simular = os.environ.get("DRY_RUN", "").strip().lower() in ("1", "true", "yes")
    agora = datetime.datetime.now(FUSO)
    data_hoje = agora.date()
    hora = hora_atual(agora)

    planilha = conectar_planilha()
    edicao = carregar_edicao(planilha, data_hoje)
    if edicao is None or not all(c in edicao for c in SECOES):
        print(f"[ERRO] Não há edição completa de {data_hoje} na aba 'edicoes'. "
              "Rode o workflow 'Gerar newsletter diária' antes.")
        sys.exit(1)

    aba_envios = obter_aba_envios(planilha)
    ja_enviados = emails_ja_enviados_hoje(aba_envios, data_hoje)
    assinantes = carregar_assinantes(planilha)
    destinatarios = escolher_destinatarios(assinantes, ja_enviados, hora)

    restante = LIMITE_DIARIO - len(ja_enviados)
    if len(destinatarios) > restante:
        print(f"[AVISO] Limite diário de {LIMITE_DIARIO} e-mails: enviando só {max(restante, 0)}.")
        destinatarios = destinatarios[:max(restante, 0)]

    por_horario = {}
    for d in destinatarios:
        por_horario[d["horario"]] = por_horario.get(d["horario"], 0) + 1
    print(f"[ENVIO] {data_hoje}, hora {hora}h: {len(assinantes)} linhas em 'assinantes', "
          f"{len(ja_enviados)} já receberam hoje, {len(destinatarios)} para enviar agora {por_horario}")

    if simular:
        print("[SIMULAÇÃO] DRY_RUN ligado: nenhum e-mail foi enviado e nada foi gravado.")
        return
    if not destinatarios:
        print("[ENVIO] Ninguém para enviar neste horário.")
        return

    enviados, falhas = enviar(destinatarios, edicao, data_hoje, aba_envios)
    print(f"[ENVIO] Concluído: {enviados} enviados, {falhas} falhas.")
    if falhas and not enviados:
        sys.exit(1)


if __name__ == "__main__":
    main()

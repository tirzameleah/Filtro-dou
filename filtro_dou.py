"""
Relatório do Diário Oficial da União por cliente — sem IA, sem token.

Uso:  python filtro_dou.py                     (edição de hoje, envia e-mail)
      python filtro_dou.py 22-09-2026          (outra data)
      python filtro_dou.py --sem-email         (só gera o HTML, pra testar)
      python filtro_dou.py --cliente="Cliente A"
      python filtro_dou.py --rapido            (pula o inteiro teor da Seção 2: bem mais
                                                rápido, mas perde nomeações)
      python filtro_dou.py --completo          (inteiro teor das Seções 1 e 2: mais lento ainda)

Gera dou_AAAA-MM-DD.html (cliente > edição > seção > ministério > ato)
e envia o relatório por e-mail.
Requisito: pip install requests
"""
import sys, os, re, json, html, unicodedata, datetime, smtplib
from email.mime.text import MIMEText
import requests

# ==== E-MAIL (EDITE AQUI) ====
# A senha NÃO é a senha normal do Gmail: é uma "senha de app" de 16 letras, criada em
# myaccount.google.com/apppasswords (precisa da verificação em 2 etapas ativa).
# Nunca escreva a senha neste arquivo: defina a variável de ambiente DOU_SENHA_APP.
#   Windows (PowerShell):  $env:DOU_SENHA_APP = "suasenhadeapp"
#   Linux/macOS:           export DOU_SENHA_APP="suasenhadeapp"
EMAIL_REMETENTE = os.environ.get("DOU_EMAIL_REMETENTE", "seu_email@gmail.com")
EMAIL_DESTINO = ["destinatario@exemplo.com"]     # pode pôr mais de um
SENHA_APP = os.environ.get("DOU_SENHA_APP", "")

SECOES = ["do1", "do2", "do3"]   # do1 = atos normativos, do2 = pessoal, do3 = contratos/avisos
EXTRAS = ["do1_extra_a", "do1_extra_b", "do1_extra_c", "do1_extra_d"]  # edições extras, se houver

# ==== FILTRO POR CLIENTE (EDITE AQUI) ====
# Escreva os termos em minúsculas e SEM acento (o texto do DOU é normalizado assim antes da busca).
# A busca é por trecho: "ministerio de minas" casa com "Ministério de Minas e Energia".
#
# obrigatorias   -> se aparecer, entra sempre (nome do cliente), em qualquer órgão
# temas          -> termos próprios do cliente: valem em QUALQUER órgão, inclusive os que não
#                   estão na lista abaixo (publicação de interesse fora do órgão costumeiro)
# tipos_sempre   -> (opcional) tipos de ato que entram sempre quando vêm do
#                   Legislativo/Executivo/Presidência
# temas_orgao    -> termos genéricos: valem em órgão observado ou junto de um termo próprio
# temas_fracos   -> (opcional) termos muito genéricos: só entram acompanhados de outro tema
# orgaos_pessoal -> onde uma nomeação/exoneração importa (Seção 2)
# orgaos         -> ministérios/órgãos que costumam publicar sobre esse cliente
#
# Os valores entre < > abaixo são marcadores: troque pelos termos reais do seu cliente.

CLIENTES = {
    "Cliente A": {
        "obrigatorias": [
            "<razao social do cliente>", "<nome fantasia ou marca>", "<subsidiaria do grupo>",
        ],
        "temas": [
            "<produto ou servico principal do cliente>", "<insumo estrategico do setor>",
            "<agencia reguladora do setor>", "<programa de governo que afeta o cliente>",
            "<tributo ou encargo especifico do setor>", "<tema regulatorio do setor>",
            "<numero de processo ou acao judicial acompanhada>",
        ],
        "tipos_sempre": [
            "lei", "lei complementar", "medida provisoria", "emenda constitucional",
            "portaria conjunta", "portaria interministerial",
        ],
        "temas_orgao": [
            "audiencia publica", "consulta publica", "tomada de subsidios",
            "<tema generico que so interessa vindo de orgao do setor>",
        ],
        "orgaos_pessoal": [
            "atos do poder", "presidencia",
            "<ministerio que regula o setor>", "<agencia reguladora do setor>",
        ],
        "orgaos": [
            "atos do poder", "presidencia", "advocacia-geral da uniao",
            "<ministerio que regula o setor>", "<ministerio da area economica>",
            "<agencia reguladora do setor>", "<orgao ambiental ou de fiscalizacao>",
        ],
    },
    "Cliente B": {
        "obrigatorias": [
            "<nome do cliente>",
        ],
        "temas": [
            "<tema regulatorio do setor>", "<conselho ou colegiado que normatiza o setor>",
            "<cadastro ou sistema publico usado pelo setor>",
        ],
        "temas_orgao": [
            "<tema generico que so interessa vindo de orgao do setor>",
            "audiencia publica", "consulta publica",
        ],
        "temas_fracos": [
            "<termo muito amplo, que sozinho traria ruido>",
        ],
        "orgaos_pessoal": [
            "atos do poder", "presidencia", "<ministerio que regula o setor>",
        ],
        "orgaos": [
            "atos do poder", "presidencia", "<ministerio que regula o setor>",
            "<autarquia ou banco publico ligado ao setor>",
        ],
    },
}
ORDEM_CLIENTES = ["Cliente A", "Cliente B"]

# ==== SEÇÃO 2: ação de pessoal, cargo estratégico, afastamento e ruído ====
ACAO_PESSOAL = re.compile(
    r"(nomea|exonera|designa|dispensa|reconduz|destitu)", re.I)

CARGO_ESTRATEGICO = re.compile(
    r"(ministr[oa] de estado|gabinete d[oa] ministr|chefe de gabinete|secretári[oa]-executiv|"
    r"secretaria-executiva|secretári[oa] nacional|secretaria nacional|secretári[oa] especial|"
    r"secretaria especial|secretári[oa]-adjunt|secretaria-adjunta|secretári[oa] d[aeo]|"
    r"secretaria d[aeo]|assessor[a]? especial|diretor|diretoria|departamento|coordenador|"
    r"coordenação|superintendent|presidente|conselheir|cce 1\.|cce 2\.1[3-9]|fce 1\.1[0-9])", re.I)

AFASTAMENTO = re.compile(
    r"(afastamento do país|afastamento do pais|afastar-se do país|afastamento do titular|"
    r"viagem ao exterior|missão ao exterior|missão no exterior)", re.I)

RUIDO_PESSOAL = re.compile(
    r"(substitut.{0,4} eventual|substituição eventual|para substituir|nos afastamentos|"
    r"impedimentos legais|férias|licença|progressão|aposentador|remover|remoção|gratificação|"
    r"chefe da divisão|chefe de divisão|chefe de projeto|chefe do setor|chefe de seção|"
    r"chefe de serviço|chefe da seção|chefe do serviço|assessor técnico|assistente técnico|"
    r"comissão processante|comissão de sindicância|concurso público|professor|magistério|"
    r"analista judiciári|técnico judiciári|pró-tempore|pro tempore|reitor|campus|penalidade|"
    r"processo administrativo disciplinar|prorrogar.{0,30}prazo|exercício descentralizado|tradutor|"
    r"intérprete|fce [2-9]\.|fce 1\.0)", re.I)

RUIDO_LICITACAO = re.compile(
    r"(aviso de licitação|aviso de pregão|aviso de alienação|licitação para alienação|^aviso$|"
    r"aviso de anulação|aviso de suspensão|aviso de adiamento|pregão eletrônico|"
    r"aviso de convocação|pré-qualificação|processo seletivo|programa de estágio|concurso público)", re.I)

RUIDO_ROTINA_TITULO = re.compile(
    r"(alvará|decisão sur|decisão suf|edital de notificação|edital de intimação|edital de citação|"
    r"atos declaratórios cvm|pauta de julgamento|extrato de acordo de cooperação|"
    r"extrato de apostilamento|extrato de rescisão|extrato de doação|extrato de credenciamento|"
    r"solução de divergência|acórdão|extrato da ata)", re.I)

# Atos individuais de rotina do seu setor (autorizações, credenciamentos em massa etc.)
# que costumam casar com os temas mas não interessam. Troque pelos padrões reais.
RUIDO_ROTINA_CONTEUDO = re.compile(
    r"(<autorizacao de rotina do setor>|<credenciamento individual repetitivo>|"
    r"<regime ou cadastro que gera muitos atos individuais>)", re.I)

SECAO3_RELEVANTE = re.compile(
    r"(audiência pública|consulta pública|tomada de subsídios|manifestação de interesse|"
    r"chamamento público|comunicado relevante|extrato de compromisso|acordo de cooperação|leilão|"
    r"processo competitivo|adesão)", re.I)

SECAO3_RUIDO = re.compile(
    r"(aviso de licitação|aviso de pregão|extrato de contrato|termo aditivo|ordem de compra|"
    r"dispensa de licitação|inexigibilidade|registro de preços|resultado de julgamento|homologação|"
    r"adjudicação|apostilamento|rescisão|notificação|intimação)", re.I)



def normaliza(t):
    t = unicodedata.normalize("NFKD", t or "").encode("ascii", "ignore").decode()
    return t.lower()


def baixa_secao(data, secao):
    url = f"https://www.in.gov.br/leiturajornal?data={data}&secao={secao}"
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=60)
        r.raise_for_status()
    except Exception as e:
        print(f"  aviso: nao consegui baixar {secao} ({e})")
        return []
    m = re.search(r'id="params"[^>]*>(.*?)</script>', r.text, re.S)
    if not m:
        return []
    try:
        dados = json.loads(html.unescape(m.group(1)))
    except json.JSONDecodeError:
        return []
    return dados.get("jsonArray", []) if isinstance(dados, dict) else dados


def inteiro_teor(pub):
    """Texto completo do ato (o resumo do site corta em ~400 caracteres)."""
    url = "https://www.in.gov.br/web/dou/-/" + pub.get("urlTitle", "")
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=60)
        r.raise_for_status()
    except Exception:
        return pub.get("content", "")
    m = re.search(r'<div class="texto-dou">(.*?)<div class="dou-paragraph-container',
                  r.text, re.S) or re.search(r'<div class="texto-dou">(.*)', r.text, re.S)
    if not m:
        return pub.get("content", "")
    texto = re.sub(r"<br\s*/?>|</p>", "\n", m.group(1))
    texto = html.unescape(re.sub("<[^>]+>", " ", texto))
    return texto.split("Este conteúdo não substitui")[0][:12000]


def completa_textos(publicacoes, so_secao2=False):
    """Baixa o inteiro teor das publicacoes (o resumo do site corta cedo demais)."""
    from concurrent.futures import ThreadPoolExecutor
    if so_secao2:
        nucleo = set()
        for c in CLIENTES.values():
            nucleo |= set(c.get("orgaos_pessoal", c["orgaos"]))
        alvos = [(s, p) for s, p in publicacoes if s == "do2"
                 and any(o in normaliza(p.get("hierarchyStr", "")) for o in nucleo)]
    else:
        alvos = [(s, p) for s, p in publicacoes if not s.startswith("do3")]
    print(f"  buscando inteiro teor de {len(alvos)} atos (pode levar alguns minutos)...")

    def carrega(item):
        _, pub = item
        pub["content"] = inteiro_teor(pub)
        return True

    with ThreadPoolExecutor(16) as ex:
        list(ex.map(carrega, alvos))


def ementa(pub):
    """Texto curto do ato: o conteudo sem repetir o titulo."""
    txt = re.sub("<[^>]+>", " ", pub.get("content", "") or "")
    txt = html.unescape(re.sub(r"\s+", " ", txt)).strip()
    titulo = (pub.get("title", "") or "").strip()
    if titulo and txt.lower().startswith(titulo.lower()):
        txt = txt[len(titulo):].strip(" :-—")
    corte = txt[:420]
    ponto = corte.rfind(". ")
    if ponto > 120:
        corte = corte[:ponto + 1]
    return corte.strip()


def orgao_principal(pub):
    partes = [p.strip() for p in (pub.get("hierarchyStr", "") or "").split("/") if p.strip()]
    if not partes:
        return "Outros"
    topo = partes[0]
    # agencias e autarquias ficam mais claras com o orgao de 2o nivel
    if len(partes) > 1 and re.search(r"(agencia|instituto|banco|conselho|superintendencia|comite)",
                                     normaliza(partes[1])):
        return f"{topo} — {partes[1]}"
    return topo


def relevante(pub, cli, secao):
    """Decide se a publicacao entra para o cliente. Na duvida, inclui."""
    texto = normaliza(" ".join(str(pub.get(k, "")) for k in
                               ("title", "content", "hierarchyStr", "artType")))
    org = normaliza(pub.get("hierarchyStr", ""))
    titulo = pub.get("title", "") or ""
    corpo = titulo + " " + re.sub("<[^>]+>", " ", pub.get("content", "") or "")

    # licitacao/alienacao nao entra no recorte, nem com o nome do cliente
    if RUIDO_LICITACAO.search(titulo.strip()):
        return []

    # 0) lei, medida provisoria e portaria conjunta do alto escalao entram sempre
    if secao == "do1" and cli.get("tipos_sempre"):
        t_norm = normaliza(titulo).lstrip()
        if (any(t_norm.startswith(tp) for tp in cli["tipos_sempre"])
                and any(o in org for o in ("atos do poder", "presidencia",
                                           "atos do congresso", "ministerio"))):
            return ["ato normativo"]

    # 1) nome do cliente: entra sempre, em qualquer secao
    for termo in cli["obrigatorias"]:
        if termo in texto:
            return [termo]

    # o tema tem de estar na ementa/comeco do ato, nao numa mencao solta no fim do texto
    texto_tema = normaliza(titulo + " " + pub.get("hierarchyStr", "") + " " + corpo[:2500])
    orgao_observado = any(o in org for o in cli["orgaos"])

    # 2) Secao 2: nomeacao/exoneracao em cargo estrategico dos orgaos que interessam ao
    #    cliente entra mesmo sem palavra-chave; afastamento do pais so com termo proprio
    if secao == "do2":
        # o cargo tem de estar na parte operativa (depois do "resolve:"), nao no preambulo,
        # senao quem assina o ato (o secretario-executivo, por exemplo) ja valeria como cargo
        corte = re.search(r"(?is)resolve[m]?\s*:?(.*)", corpo)
        operativo = (corte.group(1) if corte else corpo)[:4000]
        if not any(o in org for o in cli.get("orgaos_pessoal", cli["orgaos"])):
            return []
        proprios_do2 = [t for t in cli["temas"] if t in texto_tema]
        # AGU e orgaos de apoio juridico so quando o ato toca um tema do cliente
        if "advocacia-geral" in org and not proprios_do2:
            return []
        if AFASTAMENTO.search(operativo):
            return proprios_do2
        if not ACAO_PESSOAL.search(operativo) or not CARGO_ESTRATEGICO.search(operativo):
            return []
        if RUIDO_PESSOAL.search(operativo):
            return []
        return proprios_do2 or ["cargo estrategico"]

    # 3) termos proprios do cliente valem em qualquer orgao, mesmo fora do costumeiro
    proprios = [t for t in cli["temas"] if t in texto_tema]
    # 4) termos genericos valem em orgao observado ou junto com um termo proprio
    genericos = [t for t in cli.get("temas_orgao", []) if t in texto_tema]
    if not proprios and not (orgao_observado and genericos):
        return []
    temas = proprios + (genericos if (orgao_observado or proprios) else [])
    # 5) termos muito genericos so contam acompanhados de outro tema
    temas += [t for t in cli.get("temas_fracos", []) if t in texto_tema]

    # atos individuais de rotina so entram com o nome do cliente, com o tema no proprio titulo
    # ou quando tratam de participacao social (consulta/audiencia publica)
    tema_no_titulo = (any(len(t) >= 6 and t in normaliza(titulo) for t in cli["temas"])
                      or bool(SECAO3_RELEVANTE.search(corpo[:600])))
    if not tema_no_titulo and (RUIDO_ROTINA_TITULO.search(titulo)
                               or RUIDO_ROTINA_CONTEUDO.search(corpo)):
        return []

    if secao == "do3":
        # o titulo da Secao 3 costuma ser generico ("EXTRATO DE CONTRATO"):
        # o marcador de relevancia tambem e procurado no inicio do texto
        if not SECAO3_RELEVANTE.search(titulo + " " + corpo[:400]):
            return []
    return temas


def nome_secao(secao):
    if secao.startswith("do1"):
        return "Seção Extra" if "extra" in secao else "Seção 1"
    return {"do2": "Seção 2", "do3": "Seção 3"}.get(secao, secao.upper())


def dia_util_anterior(data):
    """Data anterior em dd-mm-aaaa, pulando sabado e domingo."""
    d = datetime.datetime.strptime(data, "%d-%m-%Y").date() - datetime.timedelta(days=1)
    while d.weekday() >= 5:            # 5 = sabado, 6 = domingo
        d -= datetime.timedelta(days=1)
    return d.strftime("%d-%m-%Y")


def main():
    argv = sys.argv[1:]
    sem_email = "--sem-email" in argv
    so_cliente = next((a.split("=", 1)[1] for a in argv if a.startswith("--cliente=")), None)
    datas = [a for a in argv if re.fullmatch(r"\d{2}-\d{2}-\d{4}", a)]
    data = datas[0] if datas else datetime.date.today().strftime("%d-%m-%Y")
    # As edicoes extras saem no fim do dia anterior e so entram no recorte do dia seguinte
    anterior = dia_util_anterior(data)

    print(f"Lendo o DOU de {data} (edicoes extras de {anterior})...")
    publicacoes = []
    for secao in SECOES:
        pubs = baixa_secao(data, secao)
        if pubs:
            print(f"  {secao}: {len(pubs)} publicacoes")
        publicacoes += [(secao, p) for p in pubs]

    vistos_extra = set()
    for secao in EXTRAS:
        for p in baixa_secao(anterior, secao):
            if p.get("urlTitle") in vistos_extra:
                continue
            vistos_extra.add(p.get("urlTitle"))
            publicacoes.append((secao, p))
    if vistos_extra:
        print(f"  extras de {anterior}: {len(vistos_extra)} publicacoes")

    if "--completo" in argv:
        completa_textos(publicacoes)
    elif "--rapido" not in argv:
        # a Secao 2 precisa do texto completo: o resumo do site para no preambulo,
        # antes de dizer quem foi nomeado e para qual cargo
        completa_textos(publicacoes, so_secao2=True)

    clientes = [c for c in ORDEM_CLIENTES if c in CLIENTES]
    if so_cliente:
        clientes = [c for c in clientes if normaliza(c) == normaliza(so_cliente)]

    # cliente -> edicao -> secao -> orgao -> [atos]
    relatorio, total = {}, 0
    for nome in clientes:
        cli = CLIENTES[nome]
        vistos, por_edicao = set(), {}
        for secao, pub in publicacoes:
            chave = pub.get("urlTitle", "")
            if chave in vistos:
                continue
            base = "do1" if secao.startswith("do1") else secao
            if not relevante(pub, cli, base):
                continue
            vistos.add(chave)
            total += 1
            edicao = (f"Diário Oficial da União {pub.get('pubDate', data)} | "
                      f"Edição {pub.get('editionNumber', '')}").strip()
            por_edicao.setdefault(edicao, {}).setdefault(nome_secao(secao), {}) \
                .setdefault(orgao_principal(pub), []).append({
                    "titulo": (pub.get("title") or "(sem titulo)").strip(),
                    "ementa": ementa(pub),
                    "link": "https://www.in.gov.br/web/dou/-/" + chave,
                })
        relatorio[nome] = por_edicao

    # ---- HTML do relatorio ----
    ordem_secao = {"Seção Extra": 0, "Seção 1": 1, "Seção 2": 2, "Seção 3": 3}
    L = [f"<h1 style='font-size:22px'>Diário Oficial da União — {data}</h1>"]
    for nome in clientes:
        L.append(f"<h2 style='font-size:20px;margin-top:28px;color:#073763'>{html.escape(nome)}</h2>")
        por_edicao = relatorio.get(nome) or {}
        if not por_edicao:
            L.append("<p>Não foram encontradas publicações de interesse nesta edição.</p>")
            continue
        for edicao in sorted(por_edicao):
            L.append(f"<p style='margin:14px 0 4px'><b>{html.escape(edicao)}</b></p>")
            secoes = por_edicao[edicao]
            for sec in sorted(secoes, key=lambda s: ordem_secao.get(s, 9)):
                L.append(f"<p style='margin:10px 0 4px'><b>{html.escape(sec)}</b></p>")
                for org in sorted(secoes[sec]):
                    L.append(f"<p style='margin:8px 0 2px'>"
                             f"<b style='color:#00b050'>{html.escape(org)}</b></p>")
                    for ato in secoes[sec][org]:
                        texto = html.escape(ato["titulo"])
                        if ato["ementa"]:
                            texto += ": " + html.escape(ato["ementa"])
                        L.append(f"<p style='margin:4px 0'>"
                                 f"<a href='{ato['link']}' target='_blank'>{texto}</a></p>")

    corpo_html = ("<meta charset='utf-8'><body style='font-family:Source Sans Pro,Calibri,sans-serif;"
                  "font-size:15px;line-height:1.5;max-width:900px;margin:auto;text-align:justify'>"
                  + "".join(L) + "</body>")

    d = datetime.datetime.strptime(data, "%d-%m-%Y").date()
    saida = f"dou_{d.isoformat()}.html"
    with open(saida, "w", encoding="utf-8") as f:
        f.write(corpo_html)
    print(f"{total} recortes -> {saida}")

    if sem_email:
        return
    if not SENHA_APP:
        print("SENHA_APP vazia: defina a variavel de ambiente DOU_SENHA_APP. "
              "E-mail nao enviado.")
        return
    msg = MIMEText(corpo_html, "html", "utf-8")
    msg["Subject"] = f"DOU {data} — recortes por cliente ({total})"
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = ", ".join(EMAIL_DESTINO)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60) as s:
        s.login(EMAIL_REMETENTE, SENHA_APP.replace(" ", ""))
        s.sendmail(EMAIL_REMETENTE, EMAIL_DESTINO, msg.as_string())
    print("E-mail enviado para", ", ".join(EMAIL_DESTINO))


if __name__ == "__main__":
    main()

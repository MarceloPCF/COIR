# ruff: noqa: E501  (modelos HTML em linhas longas)
"""Gera o index.html do site (o visual do site da v2, marinho e verde, com o texto da v3; escolha
do usuário de 26/10/2026). Uso: uv run python tools/site/gerar_sites.py . tools/site/site-v2.html
"""

import sys
from pathlib import Path

REPO, V2 = Path(sys.argv[1]), Path(sys.argv[2])
GH = "https://github.com/MarceloPCF/COIR"

# ------------------------------------------------------------------ conteúdo comum (v3)
RESUMO = (
    "O imposto de renda da bolsa calculado no seu computador: importa as notas de "
    "corretagem e as planilhas da B3, calcula o preço médio, apura o IR de cada mês, prepara "
    "o DARF e as fichas da declaração anual. Gratuito, e seus dados não saem do computador."
)

PASSOS = [
    (
        "Baixe e extraia",
        "Use o botão <strong>Baixar para Windows</strong>, extraia o .zip numa pasta e "
        "dê dois cliques em <code>COIR.exe</code>. Não é preciso instalar Python, Excel nem outro programa.",
        "Se o Windows avisar que o programa não é reconhecido, clique em “Mais informações” e "
        "“Executar assim mesmo”: o aviso aparece porque o programa ainda não tem assinatura digital paga.",
    ),
    (
        "Escolha a pasta dos seus dados",
        "Na primeira vez, o COIR pergunta onde guardar os dados e o "
        "nome do titular. Use uma pasta fora da pasta do programa: assim, para atualizar, basta trocar a "
        "pasta do programa.",
        "",
    ),
    (
        "Importe as notas e as planilhas da B3",
        "Arraste os PDFs das notas de corretagem (padrão "
        "SINACOR) e as planilhas da Área do Investidor da B3. Cada nota é conferida com o resumo dela; "
        "divergência é avisada, nunca ignorada.",
        "Testado com XP, Clear, Rico, BTG, Necton e Ágora. Se a sua corretora não for lida, use o "
        '<a href="{R}#corretora">pedido de corretora</a> (sem mandar a nota nem dados pessoais).',
    ),
    (
        "Confira e pague",
        "Veja a apuração mês a mês com a memória de cálculo, o DARF com vencimento, "
        "multa e juros, e as fichas da declaração anual prontas para copiar no programa da Receita.",
        "",
    ),
]

MERCADOS = [
    ("Ações, ETF, BDR e units", "Operações comuns e day trade"),
    ("Opções", "Compra, lançamento, exercício e vencimento"),
    ("Futuros e termo", "Minicontratos, commodities e termo"),
    ("FII e Fiagro", "Ganho na venda e rendimentos"),
]

RECURSOS = [
    "Seus dados ficam com você: tudo roda no seu computador e nada é enviado a terceiros.",
    "Preço médio por titular, somando todas as corretoras, com eventos corporativos "
    "(desdobramento, bonificação, incorporação, cisão).",
    "Comum, day trade e FII apurados à parte, com isenção de R$ 20 mil, prejuízo a compensar, "
    "IRRF e DARF com mínimo de R$ 10.",
    "Proventos (dividendos, JCP e reembolsos de aluguel) e a declaração anual na ordem do programa.",
    "Todo valor mostra as operações e a regra aplicada, com a norma da Receita.",
    "Relatórios em Excel, CSV e PDF sem precisar do Excel instalado.",
]

HISTORIA = [
    ("2017", "Primeira versão, em planilhas Excel, com operações do mercado à vista."),
    ("2021", "Extração automática das notas de corretagem com Python."),
    ("2022", "Controles de futuros, commodities e opções."),
    ("2026", "Versão 3: programa próprio, sem Excel, com apuração, DARF e declaração anual."),
]

TELAS = [
    ("Painel", "A DARF da vez, as pendências e o resultado mês a mês."),
    ("Importar", "Notas em PDF e planilhas da B3, com conferência de cada nota."),
    ("Notas e operações", "A nota ao lado do original e o livro de operações com filtros."),
    ("Apuração mensal", "Resultado por categoria com a memória de cálculo de cada mês."),
    ("DARF", "Vencimento em dia útil, multa e juros por atraso e o registro do pagamento."),
    ("Proventos", "Dividendos, JCP e reembolsos, com a data do crédito do JCP."),
    ("Declaração anual", "Fichas de renda variável, bens e direitos e rendimentos para copiar."),
    ("Carteira e preço médio", "A posição em qualquer data e a evolução do preço médio."),
    ("Eventos corporativos", "Os eventos da B3 que atingem a sua posição, para confirmar."),
    ("Opções e futuros", "O ciclo de cada série e os ajustes de cada contrato."),
    ("Conciliação com a B3", "A posição calculada comparada com a da B3."),
    ("Relatórios", "Exportação em Excel, CSV e PDF."),
]

# capturas do banco de demonstração (tools/capturas.py), tema claro
IMAGENS = [
    ("tela-painel.jpg", "Painel do ano com o banco de demonstração", "Painel"),
    ("tela-apuracao.jpg", "Apuração mensal por categoria", "Apuração mensal"),
    ("tela-darf.jpg", "DARFs do ano, pagas e em aberto", "DARF"),
    ("tela-carteira.jpg", "Carteira e preço médio", "Carteira e preço médio"),
    ("tela-declaracao.jpg", "Fichas da declaração anual", "Declaração anual"),
    ("tela-notas.jpg", "Notas de corretagem importadas", "Notas"),
]

DUVIDAS = [
    ("O COIR é gratuito?", "Sim. É software livre, com licença AGPL-3.0."),
    (
        "O que preciso ter instalado?",
        "Nada: só o Windows 10 ou 11. Basta extrair o .zip e abrir o <code>COIR.exe</code>.",
    ),
    (
        "Meus dados saem do computador?",
        "Não. Notas, planilhas e o banco ficam numa pasta sua. As "
        "únicas consultas à internet (tabelas públicas da B3 e versão nova do programa) são feitas "
        "quando você pede, ou ao abrir, se você ligar essa opção.",
    ),
    (
        "Como atualizo para uma versão nova?",
        "Baixe o .zip novo e troque a pasta do programa. Os dados "
        "continuam na pasta deles; na primeira abertura o COIR faz um backup do banco e o atualiza.",
    ),
    (
        "Usei a versão 2 (planilha). Por que alguns números mudaram?",
        "A versão 3 corrige pontos da "
        "versão 2: isenção pela soma das vendas, saldo de IRRF transportado entre os meses, DARF com "
        "mínimo de R$ 10 e vencimento em dia útil, FII separado e eventos corporativos no preço médio. "
        'As <a href="{R}novidades.html">notas da versão</a> explicam cada caso.',
    ),
    (
        "O COIR substitui um contador?",
        "Não. Ele faz as contas e mostra de onde vem cada valor, mas "
        "confira sempre antes de pagar um DARF ou entregar a declaração. As "
        '<a href="{R}fontes.html">fontes fiscais</a> usadas estão listadas.',
    ),
]

CONTATO_TEXTO = (
    "Quer saber quando sair uma versão nova? Deixe o seu e-mail. É opcional: o "
    "download não depende disso, e o próprio programa também avisa quando há versão "
    "nova."
)
LGPD = (
    "Ao enviar, você aceita receber e-mails só sobre versões novas do COIR. O e-mail nunca é "
    "repassado, e você pode pedir a remoção a qualquer momento (LGPD). O formulário é de um "
    "serviço externo e só recebe o que foi digitado."
)
PEDIDO_TEXTO = (
    "A sua corretora não foi lida? Conte qual é e o que o COIR mostrou. Não mande a nota nem "
    "dados pessoais (nome, CPF, conta): se for preciso ver o layout, a resposta explica como "
    "mandar um texto da nota sem esses dados."
)

SCRIPT_FORMS = """
  // Endereço do formulário de avisos (serviço externo que aceita POST com resposta JSON,
  // como o Formspree: https://formspree.io/f/<código>). Vazio: o formulário fica "em breve".
  // O mesmo endereço recebe os avisos e os pedidos de corretora (o assunto separa os dois).
  var FORM_AVISOS = 'https://formspree.io/f/mdeeenqe';
  function ligar(idForm, idSituacao, pronto) {
    var form = document.getElementById(idForm);
    var situacao = document.getElementById(idSituacao);
    if (!form) return;
    if (!FORM_AVISOS) {
      Array.prototype.forEach.call(form.elements, function (e) { e.disabled = true; });
      situacao.textContent = 'Em breve.';
      return;
    }
    form.action = FORM_AVISOS;
    form.addEventListener('submit', function (ev) {
      if (!window.fetch) return;
      ev.preventDefault();
      situacao.textContent = 'Enviando...';
      fetch(FORM_AVISOS, { method: 'POST', body: new FormData(form),
                           headers: { 'Accept': 'application/json' } })
        .then(function (r) {
          if (!r.ok) throw new Error();
          form.reset();
          situacao.textContent = pronto;
        })
        .catch(function () { situacao.textContent = 'Não foi possível enviar agora. Tente mais tarde.'; });
    });
  }
  ligar('form-avisos', 'avisos-situacao', 'Pronto: você vai receber um aviso quando sair versão nova.');
  ligar('form-pedido', 'pedido-situacao', 'Pedido recebido. Obrigado!');
  // Botão "Baixar": aponta direto para o zip da última versão (sem passar pela página do GitHub).
  var botao = document.getElementById('baixar');
  var versao = document.getElementById('versao');
  if (window.fetch) {
    fetch('https://api.github.com/repos/MarceloPCF/COIR/releases/latest', { headers: { 'Accept': 'application/vnd.github+json' } })
      .then(function (r) { return r.ok ? r.json() : Promise.reject(); })
      .then(function (rel) {
        var zip = (rel.assets || []).filter(function (a) { return /^COIR-.*-windows\\.zip$/i.test(a.name); })[0];
        if (!zip) { return; }
        botao.href = zip.browser_download_url;
        var mb = (zip.size / 1048576).toFixed(0);
        var num = (rel.tag_name || '').replace(/^v/i, '');
        versao.textContent = 'Versão ' + num + ', arquivo .zip de ' + mb + ' MB. Gratuito, para Windows 10 ou 11.';
      })
      .catch(function () {});
  }
"""

SIMBOLO = (
    '<svg class="simbolo" viewBox="0 0 200 200" role="img" aria-label="Símbolo do COIR"><defs>'
    '<linearGradient id="g" x1="0" y1="0" x2="1" y2="0.35"><stop offset="0" stop-color="#050F2A"/>'
    '<stop offset="0.55" stop-color="#0E3A78"/><stop offset="1" stop-color="#1B5DA8"/></linearGradient>'
    '</defs><g transform="scale(1.25) translate(-20 -20)"><rect x="20" y="20" width="160" height="160" '
    'rx="36" fill="url(#g)"/><rect x="30" y="30" width="140" height="140" rx="28" fill="none" '
    'stroke="#fff" stroke-opacity="0.16" stroke-width="1.2"/><path d="M132.70 70.56 A44 44 0 1 0 '
    '103.07 143.89" fill="none" stroke="#fff" stroke-width="9"/><path d="M108.40 143.19 A44 44 0 0 0 '
    '132.70 129.44" fill="none" stroke="#8FBBEA" stroke-width="9"/></g></svg>'
)

HEAD_META = """<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>COIR — imposto de renda da bolsa calculado no seu computador</title>
<meta name="description" content="Programa gratuito para Windows que importa notas de corretagem SINACOR e planilhas da B3, calcula o preço médio, apura o IR mensal de renda variável, gera o DARF e prepara a declaração anual. Seus dados não saem do computador.">
<meta name="google-site-verification" content="vDjlNKkJJRVfJJTE-ybAxmtGzT-z72kWETuAConlxOg">
<link rel="canonical" href="https://marcelopcf.github.io/COIR/">{ROBOTS}
<link rel="icon" href="{R}favicon.svg" type="image/svg+xml">
<meta name="theme-color" content="{TEMA}">
<meta property="og:title" content="COIR — IR da bolsa no seu computador">
<meta property="og:description" content="Notas de corretagem → preço médio → IR mensal → DARF → declaração anual. Gratuito, sem enviar seus dados.">
<meta property="og:url" content="https://marcelopcf.github.io/COIR/">
<meta property="og:type" content="website">
<meta property="og:locale" content="pt_BR">
<meta property="og:image" content="https://marcelopcf.github.io/COIR/docs/social-preview.png">
<meta name="twitter:card" content="summary_large_image">
<script type="application/ld+json">
{"@context":"https://schema.org","@type":"SoftwareApplication","name":"COIR","applicationCategory":"FinanceApplication",
 "operatingSystem":"Windows 10, Windows 11","offers":{"@type":"Offer","price":"0","priceCurrency":"BRL"},
 "url":"https://marcelopcf.github.io/COIR/","license":"https://www.gnu.org/licenses/agpl-3.0.html",
 "description":"Controle de operações e imposto de renda de renda variável para pessoa física na B3."}
</script>"""


def head(raiz: str, tema: str, exemplo: bool) -> str:
    robots = '\n<meta name="robots" content="noindex">' if exemplo else ""
    return HEAD_META.replace("{ROBOTS}", robots).replace("{R}", raiz).replace("{TEMA}", tema)


# ------------------------------------------------------------------ página
def inicial(raiz: str, exemplo: bool) -> str:
    v2 = V2.read_text(encoding="utf-8")
    css = v2[v2.index("<style>") : v2.index("</style>") + len("</style>")]
    css = css.replace(
        "</style>",
        (
            ".topo{padding:14px 0 36px}\n"
            ".topo .wrap{gap:22px}\n"
            ".barra{display:flex;align-items:center;justify-content:space-between;gap:16px;"
            "flex-wrap:wrap}\n"
            ".marca{display:flex;align-items:center;gap:10px;color:#fff;text-decoration:none;"
            "font-weight:800;font-size:22px;letter-spacing:.02em}\n"
            ".simbolo{width:34px;height:34px;flex:none}\n"
            ".apresentacao{margin-top:0;max-width:none;display:grid;gap:36px;align-items:center}\n"
            "@media (min-width:980px){.apresentacao{grid-template-columns:minmax(0,1fr) "
            "minmax(0,1.05fr)}}\n"
            ".apresentacao h1 .sigla{position:absolute;width:1px;height:1px;overflow:hidden;"
            "clip:rect(0 0 0 0)}\n"  # a marca da barra já diz COIR; o h1 continua com o nome
            ".apresentacao h1 .nome{margin:0;font-size:clamp(28px,4vw,42px);color:#fff;"
            "line-height:1.15;font-weight:800}\n"
            ".vitrine{margin:0}\n"
            ".vitrine img{display:block;width:100%;height:auto;border-radius:8px;"
            "box-shadow:0 18px 40px #0007,0 0 0 1px #ffffff22}\n"
            ".privacidade{margin-top:34px;padding:16px 20px;background:#e8f5ec;border-left:5px solid "
            "var(--verde);border-radius:0 6px 6px 0;max-width:46em}\n"
            ".pequeno{font-size:15px;color:var(--suave)}\n"
            ".botao.claro{color:var(--marinho);border-color:var(--linha)}\n"
            ".avisos{display:grid;gap:8px;max-width:560px;margin-top:12px}\n"
            ".avisos .linha-form{display:flex;gap:8px;flex-wrap:wrap}\n"
            ".avisos input[type=email]{flex:1 1 240px;padding:10px 12px;border-radius:8px;border:1px solid var(--linha,#999);font:inherit;background:transparent;color:inherit}\n"
            ".avisos textarea,.avisos input[type=text]{padding:10px 12px;border-radius:8px;border:1px solid var(--linha,#999);font:inherit;background:transparent;color:inherit}\n"
            ".avisos :disabled{opacity:.6;cursor:not-allowed}\n"
            ".botao.claro:hover{border-color:var(--marinho)}\n"
            "</style>"
        ),
    )
    passos = "\n".join(
        f"      <li>\n        <h3>{t}</h3>\n        <p>{p}</p>\n"
        + (f'        <p class="nota">{n}</p>\n' if n else "")
        + "      </li>"
        for t, p, n in PASSOS
    )
    mercados = "\n".join(
        f"      <li><strong>{t}</strong><span>{s}</span></li>" for t, s in MERCADOS
    )
    recursos = "\n".join(f"        <li>{r}</li>" for r in RECURSOS)
    historia = "\n".join(
        f"      <li><strong>{a}</strong><span>{t}</span></li>" for a, t in HISTORIA
    )
    telas = "\n".join(f"      <div><dt>{t}</dt><dd>{d}</dd></div>" for t, d in TELAS)
    figuras = "\n".join(
        f'      <figure><button type="button"><span class="moldura"><img src="{raiz}docs/site/imagens/{a}" '
        f'loading="lazy" alt="{alt}"></span><figcaption>{leg}</figcaption></button></figure>'
        for a, alt, leg in IMAGENS
    )
    duvidas = "\n".join(
        f"      <h3>{q}</h3>\n      <p>{r.replace('{R}', raiz)}</p>" for q, r in DUVIDAS
    )
    script = v2[v2.index("  // Carrossel de telas.") : v2.index("})();")]
    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<!-- Exemplo 1 do site da v3: o visual do site inicial (marinho e verde), texto da v3. -->
{head(raiz, "#102a43", exemplo)}
{css}
</head>
<body>

<header class="topo">
  <div class="wrap">
    <div class="barra">
      <a class="marca" href="#" aria-label="COIR, início">{SIMBOLO}<span>COIR</span></a>
      <nav aria-label="Atalhos">
        <a href="#como-usar">Como usar</a>
        <a href="#controle">O que faz</a>
        <a href="#telas-lista">Telas</a>
        <a href="#duvidas">Dúvidas</a>
        <a href="#contato">Avisos</a>
        <a href="#corretora">Corretora</a>
      </nav>
    </div>
    <div class="apresentacao">
      <div class="texto">
        <h1><span class="sigla">COIR</span> <span class="nome">Controle de Operações e Imposto de Renda da bolsa</span></h1>
        <p class="resumo">{RESUMO}</p>
        <div class="acoes">
          <a class="botao principal" id="baixar" href="{GH}/releases/latest">Baixar para Windows</a>
          <a class="botao secundario" href="{raiz}docs/manual/index.html">Ler o manual</a>
        </div>
        <p class="versao" id="versao">Gratuito, para Windows 10 ou 11. Não precisa instalar nada.</p>
      </div>
      <figure class="vitrine"><img src="{raiz}docs/site/imagens/tela-passeio.webp" width="1280" height="800" alt="As telas do COIR em sequência: painel, notas, apuração, DARF, carteira e declaração (banco de demonstração)"></figure>
    </div>
  </div>
</header>

<main>

<section id="como-usar">
  <div class="wrap">
    <h2>Como usar</h2>
    <p class="intro">Não precisa de Excel, Python nem Java: o programa vem pronto. Em quatro passos, as notas viram apuração, DARF e declaração.</p>
    <ol class="passos">
{passos}
    </ol>
    <div class="privacidade">
      <strong>Seus dados ficam com você.</strong> O COIR roda só no seu computador. Nada da sua carteira é enviado para lugar nenhum, e nada sai do computador sem uma ação sua.
    </div>
  </div>
</section>

<section id="controle">
  <div class="wrap">
    <h2>O que o COIR controla</h2>
    <p class="intro">Notas de corretagem da B3 (Brasil, Bolsa e Balcão) nos mercados:</p>
    <ul class="mercados">
{mercados}
    </ul>
    <div class="historia coluna-leitura">
      <ul class="recursos">
{recursos}
      </ul>
    </div>
    <h2 style="margin-top:56px">História</h2>
    <ol class="linha-tempo">
{historia}
    </ol>
    <p class="coluna-leitura" style="margin-top:28px">Desde o início, o objetivo foi a simplicidade e a correção dos dados extraídos e contabilizados.</p>
  </div>
</section>

<section id="telas-lista">
  <div class="wrap">
    <h2>As telas do COIR</h2>
    <p class="intro">O que cada tela do programa faz. O <a href="{raiz}docs/manual/index.html">manual</a> tem o passo a passo de cada uma.</p>
    <dl class="planilhas">
{telas}
    </dl>
  </div>
</section>

<section id="telas">
  <div class="wrap">
    <h2>Telas do sistema</h2>
    <p class="intro">Use as setas ou deslize para ver as telas. Clique em uma imagem para ampliar.</p>
    <div class="carrossel" role="region" aria-roledescription="carrossel" aria-label="Telas do sistema">
    <button type="button" class="seta ant" aria-label="Telas anteriores" disabled><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M15 4 7 12l8 8"/></svg></button>
    <button type="button" class="seta prox" aria-label="Próximas telas"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m9 4 8 8-8 8"/></svg></button>
    <div class="galeria" id="galeria">
{figuras}
    </div>
    </div>
  </div>
</section>

<section id="duvidas">
  <div class="wrap">
    <h2>Dúvidas frequentes</h2>
    <div class="faq coluna-leitura">
{duvidas}
    </div>
  </div>
</section>

<section id="contato">
  <div class="wrap">
    <h2>Avisos de versão nova</h2>
    <p class="intro">{CONTATO_TEXTO}</p>
    <form id="form-avisos" class="avisos" method="post">
      <label for="avisos-email">Seu e-mail</label>
      <div class="linha-form">
        <input id="avisos-email" type="email" name="email" required autocomplete="email" placeholder="voce@exemplo.com">
        <button type="submit" class="botao principal">Quero ser avisado</button>
      </div>
      <input type="hidden" name="_subject" value="COIR: aviso de versão nova">
      <p id="avisos-situacao" class="pequeno" role="status"></p>
    </form>
    <p class="pequeno">{LGPD}</p>
  </div>
</section>

<section id="corretora">
  <div class="wrap">
    <h2>Pedido de corretora</h2>
    <p class="intro">{PEDIDO_TEXTO}</p>
    <form id="form-pedido" class="avisos" method="post">
      <label for="pedido-corretora">Corretora</label>
      <input id="pedido-corretora" type="text" name="corretora" required placeholder="Ex.: Nova Corretora">
      <label for="pedido-segmento">Segmento e ano das notas</label>
      <input id="pedido-segmento" type="text" name="segmento" placeholder="Ex.: Bovespa, 2024 e 2025">
      <label for="pedido-mensagem">O que o COIR mostrou (a mensagem da importação)</label>
      <textarea id="pedido-mensagem" name="mensagem" rows="4"></textarea>
      <label for="pedido-email">Seu e-mail, se quiser resposta (opcional)</label>
      <div class="linha-form">
        <input id="pedido-email" type="email" name="email" autocomplete="email" placeholder="voce@exemplo.com">
        <button type="submit" class="botao principal">Enviar pedido</button>
      </div>
      <input type="hidden" name="_subject" value="COIR: pedido de corretora">
      <p id="pedido-situacao" class="pequeno" role="status"></p>
    </form>
  </div>
</section>

</main>

<footer>
  <div class="wrap">
    <span>COIR é gratuito e software livre (licença AGPL-3.0). <a href="{GH}">Código-fonte</a> · <a href="{raiz}legacy/index.html">Versão 2 (planilha)</a></span>
    <span>Não substitui a orientação de um contador.</span>
  </div>
</footer>

<dialog id="ampliada" aria-label="Imagem ampliada">
  <form method="dialog"><button type="submit">Fechar</button></form>
  <div class="moldura"><img id="ampliada-img" src="" alt=""></div>
  <p id="ampliada-legenda"></p>
</dialog>

<script>
(function () {{
{SCRIPT_FORMS}
{script}}})();
</script>
</body>
</html>
"""


def gravar(caminho: Path, texto: str) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_text(texto, encoding="utf-8", newline="\n")
    print(caminho.relative_to(REPO).as_posix(), len(texto.splitlines()), "linhas")


gravar(REPO / "index.html", inicial("./", exemplo=False))

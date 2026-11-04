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
        "Testado com XP, Clear, Rico, BTG, Necton e Ágora, inclusive as notas em PDF com senha que "
        "chegam por e-mail. Se a sua corretora não for lida, a tela "
        '<a href="#corretora">Suporte</a> manda a nota sem os seus dados pessoais.',
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
    "Um evento corporativo que falta tira só aquele ativo do cálculo, com DARF provisória nos meses "
    "afetados; o restante da apuração continua.",
    "Um sino no topo da janela junta as pendências de todos os anos: DARF em atraso, nota "
    "rejeitada, evento para conferir e avisos da apuração.",
    "Desdobramentos, grupamentos e bonificações com custo publicado entram sozinhos; o que "
    "depende de você (subscrição, oferta, incorporação) fica numa lista para conferir.",
    "Reprocessa notas já importadas quando a corretora corrige uma nota, com backup e a versão "
    "anterior guardada; corretora ainda sem perfil é lida pelos layouts padrão, com aviso.",
    "Todo valor mostra as operações e a regra aplicada, com a norma da Receita, e o botão "
    "<strong>Como faço?</strong> abre o manual ao lado de cada tela.",
    "Relatórios em Excel, CSV e PDF sem precisar do Excel instalado.",
]

HISTORIA = [
    ("2017", "Primeira versão, em planilhas Excel, com operações do mercado à vista."),
    ("2021", "Extração automática das notas de corretagem com Python."),
    ("2022", "Controles de futuros, commodities e opções."),
    ("2026", "Versão 3: programa próprio, sem Excel, com apuração, DARF e declaração anual."),
]

TELAS = [
    ("Painel", "A DARF da vez, as pendências de todos os anos e o resultado mês a mês."),
    (
        "Importar",
        "Notas em PDF (também com senha) e planilhas da B3, com conferência de cada nota.",
    ),
    ("Notas e operações", "A nota ao lado do original e o livro de operações com filtros."),
    ("Apuração mensal", "Resultado por categoria e por ativo, com a memória de cálculo do mês."),
    ("DARF", "Vencimento em dia útil, multa e juros por atraso e o registro do pagamento."),
    ("Proventos", "Dividendos, JCP e reembolsos, com a data do crédito do JCP."),
    ("Declaração anual", "Fichas de renda variável, bens e direitos e rendimentos para copiar."),
    ("Carteira e preço médio", "A posição em qualquer data e a evolução do preço médio."),
    (
        "Eventos corporativos",
        "Os eventos da B3 que atingem a sua posição, em abas: a conferir, "
        "lançados, subscrições, transferências, conversões e ofertas.",
    ),
    ("Opções e futuros", "O ciclo de cada série e os ajustes de cada contrato."),
    ("Conciliação com a B3", "A posição calculada comparada com a da B3."),
    ("Relatórios", "Exportação em Excel, CSV e PDF."),
    (
        "Suporte",
        "Manda uma nota que o COIR não lê, com os seus dados trocados e tarjados, ou o registro "
        "de um erro, sem dados pessoais.",
    ),
]

# capturas do banco de demonstração (tools/capturas.py), tema claro
IMAGENS = [
    ("tela-painel.jpg", "Painel do ano com o banco de demonstração", "Painel"),
    ("tela-apuracao.jpg", "Apuração mensal por categoria", "Apuração mensal"),
    ("tela-darf.jpg", "DARFs do ano, pagas e em aberto", "DARF"),
    ("tela-carteira.jpg", "Carteira e preço médio", "Carteira e preço médio"),
    ("tela-declaracao.jpg", "Fichas da declaração anual", "Declaração anual"),
    ("tela-notas.jpg", "Notas de corretagem importadas", "Notas"),
    ("tela-suporte.jpg", "Dados para ocultar antes de enviar uma nota", "Suporte"),
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
        "As notas chegam por e-mail com senha. O COIR lê?",
        "Lê. O COIR tenta sozinho o começo do CPF dos titulares cadastrados, que é a senha "
        "comum dessas notas. Se for outra, digite-a no campo <strong>Senha dos PDFs</strong> da tela "
        "Importar: ela vale para o lote e não é gravada.",
    ),
    (
        "A minha corretora não foi lida. Como mando a nota sem expor os meus dados?",
        "Pela tela <strong>Suporte</strong> do COIR. Você informa o que aparece nas suas notas "
        "(nome, CPF, conta, endereço, telefone) e o COIR grava uma cópia com esses dados trocados "
        "por fictícios, e as imagens das páginas com tarja preta por cima. A cópia só sai se nada "
        "pessoal sobrar. Você confere as imagens, e o botão abre o seu e-mail com a mensagem "
        "pronta; você anexa o .zip e envia. As operações (ativos, datas e valores) continuam "
        "visíveis, porque é delas que a leitura depende.",
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
EMAIL_SUPORTE = "suporte.coir@gmail.com"
PERFIL = "@coirapp"  # o mesmo nome no Instagram e no YouTube
INSTAGRAM = "https://www.instagram.com/coirapp/"
YOUTUBE = "https://www.youtube.com/@coirapp"
SUPORTE_TEXTO = (
    "Uma nota não foi lida, ou foi lida errado? Use a tela <strong>Suporte</strong> do COIR: ela "
    "troca os seus dados por fictícios, cobre-os com tarja nas imagens e prepara o e-mail para "
    f'<a href="mailto:{EMAIL_SUPORTE}">{EMAIL_SUPORTE}</a>. Para um erro numa tela, ela também '
    "envia o registro de atividades, sem CPF, nome, conta nem valores "
    '(<a href="docs/manual/telas/suporte.html">veja como no manual</a>).'
)
PEDIDO_TEXTO = (
    "Prefere só avisar? Conte qual é a corretora e o que o COIR mostrou. Aqui não mande a nota nem "
    "dados pessoais (nome, CPF, conta)."
)

SCRIPT_FORMS = """
  // Endereço do formulário de avisos (serviço externo que aceita POST com resposta JSON,
  // como o Formspree: https://formspree.io/f/<código>). Vazio: o formulário fica "em breve".
  // O mesmo endereço recebe os avisos e os pedidos de corretora (o assunto separa os dois).
  var FORM_AVISOS = 'https://formspree.io/f/mdeeenqe';
  var PRAZO_MS = 20000;  // sem resposta nesse tempo, o envio é dado como perdido
  var FALHOU = 'Não foi possível enviar agora. Tente mais tarde.';
  // A resposta de erro do Formspree (JSON): { errors: [{ code, field, message }] }; 429 = muitos
  // envios seguidos (o limite é de 20 por minuto).
  function mensagemDeErro(status, corpo) {
    if (status === 429) return 'Muitos envios seguidos. Espere um minuto e tente de novo.';
    var erros = (corpo && corpo.errors) || [];
    for (var i = 0; i < erros.length; i++) {
      if (erros[i].code === 'TYPE_EMAIL') return 'Confira o e-mail: ele parece incompleto ou com erro de digitação.';
      if (/^REQUIRED_FIELD/.test(erros[i].code || '')) return 'Preencha os campos obrigatórios.';
    }
    return FALHOU;
  }
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
    var botao = form.querySelector('button[type=submit]');
    var enviando = false;
    var obrigatorios = Array.prototype.filter.call(form.elements, function (e) { return e.required; });
    obrigatorios.forEach(function (e) {
      e.addEventListener('input', function () { e.setCustomValidity(''); });
    });
    form.addEventListener('submit', function (ev) {
      if (!window.fetch) return;
      ev.preventDefault();
      if (enviando) return;  // clique duplo: um envio só
      // campo obrigatório só com espaços passa no "required" do navegador
      var vazio = obrigatorios.filter(function (e) { return !String(e.value).trim(); })[0];
      if (vazio) {
        vazio.setCustomValidity('Preencha este campo.');
        vazio.reportValidity();
        return;
      }
      enviando = true;
      if (botao) botao.disabled = true;
      situacao.textContent = 'Enviando...';
      var controle = window.AbortController ? new AbortController() : null;
      var relogio = controle ? setTimeout(function () { controle.abort(); }, PRAZO_MS) : null;
      fetch(FORM_AVISOS, { method: 'POST', body: new FormData(form),
                           headers: { 'Accept': 'application/json' },
                           signal: controle ? controle.signal : undefined })
        .then(function (r) {
          if (r.ok) {
            form.reset();
            situacao.textContent = pronto;
            return;
          }
          return r.json().catch(function () { return null; }).then(function (corpo) {
            situacao.textContent = mensagemDeErro(r.status, corpo);
          });
        })
        .catch(function (erro) {
          situacao.textContent = erro && erro.name === 'AbortError'
            ? 'O envio demorou demais. Confira a conexão e tente de novo.' : FALHOU;
        })
        .then(function () {
          if (relogio) clearTimeout(relogio);
          enviando = false;
          if (botao) botao.disabled = false;
        });
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
<link rel="canonical" href="https://coir.com.br/">{ROBOTS}
<link rel="icon" href="{R}favicon.svg" type="image/svg+xml">
<meta name="theme-color" content="{TEMA}">
<meta property="og:title" content="COIR — IR da bolsa no seu computador">
<meta property="og:description" content="Notas de corretagem → preço médio → IR mensal → DARF → declaração anual. Gratuito, sem enviar seus dados.">
<meta property="og:url" content="https://coir.com.br/">
<meta property="og:type" content="website">
<meta property="og:locale" content="pt_BR">
<meta property="og:image" content="https://coir.com.br/docs/social-preview-v3.png">
<meta property="og:image:width" content="1280">
<meta property="og:image:height" content="640">
<meta property="og:image:alt" content="COIR versão 3: o Painel do programa com a DARF da vez e o resultado mês a mês">
<meta name="twitter:card" content="summary_large_image">
<script type="application/ld+json">
{"@context":"https://schema.org","@type":"SoftwareApplication","name":"COIR","applicationCategory":"FinanceApplication",
 "operatingSystem":"Windows 10, Windows 11","offers":{"@type":"Offer","price":"0","priceCurrency":"BRL"},
 "url":"https://coir.com.br/","license":"https://www.gnu.org/licenses/agpl-3.0.html",
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
            ".topo{padding:10px 0 26px}\n"
            ".topo .wrap{gap:0}\n"
            ".barra{padding-bottom:14px;margin-bottom:26px;border-bottom:1px solid #ffffff1f}\n"
            ".barra{display:flex;align-items:center;justify-content:space-between;gap:16px;"
            "flex-wrap:wrap}\n"
            ".marca{display:flex;align-items:center;gap:10px;color:#fff;text-decoration:none;"
            "font-weight:800;font-size:22px;letter-spacing:.02em}\n"
            ".simbolo{width:34px;height:34px;flex:none}\n"
            ".apresentacao{margin-top:0;max-width:52em}\n"
            ".apresentacao h1 .sigla{position:absolute;width:1px;height:1px;overflow:hidden;"
            "clip:rect(0 0 0 0)}\n"  # a marca da barra já diz COIR; o h1 continua com o nome
            ".apresentacao h1 .nome{margin:0;font-size:clamp(24px,3.2vw,34px);color:#fff;"
            "line-height:1.15;font-weight:800}\n"
            ".apresentacao .resumo{font-size:17px;margin-top:10px}\n"
            ".acoes{margin-top:16px}\n"
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
        <a href="#corretora">Suporte</a>
      </nav>
    </div>
    <div class="apresentacao">
      <div class="texto">
        <h1><span class="sigla">COIR</span> <span class="nome">Controle de Operações e Imposto de Renda da bolsa</span></h1>
        <p class="resumo">{RESUMO}</p>
        <div class="acoes">
          <a class="botao principal" id="baixar" href="{GH}/releases/latest/download/COIR-windows.zip">Baixar para Windows</a>
          <a class="botao secundario" href="{raiz}docs/manual/index.html">Ler o manual</a>
        </div>
        <p class="versao" id="versao">Gratuito, para Windows 10 ou 11. Não precisa instalar nada.</p>
      </div>
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
        <input id="avisos-email" type="email" name="email" required autocomplete="email" pattern="[^@ ]+@[^@ ]+[.][^@ ]+" title="E-mail completo, como voce@exemplo.com" maxlength="254" placeholder="voce@exemplo.com">
        <button type="submit" class="botao principal">Quero ser avisado</button>
      </div>
      <input type="hidden" name="_subject" value="COIR: aviso de versão nova">
      <input type="text" name="_gotcha" tabindex="-1" autocomplete="off" aria-hidden="true" style="position:absolute;left:-9999px;width:1px;height:1px;opacity:0">
      <p id="avisos-situacao" class="pequeno" role="status"></p>
    </form>
    <p class="pequeno">{LGPD}</p>
  </div>
</section>

<section id="corretora">
  <div class="wrap">
    <h2>Suporte e pedido de corretora</h2>
    <p class="intro">{SUPORTE_TEXTO}</p>
    <p>{PEDIDO_TEXTO}</p>
    <form id="form-pedido" class="avisos" method="post">
      <label for="pedido-corretora">Corretora</label>
      <input id="pedido-corretora" type="text" name="corretora" required maxlength="80" placeholder="Ex.: Nova Corretora">
      <label for="pedido-segmento">Segmento e ano das notas</label>
      <input id="pedido-segmento" type="text" name="segmento" maxlength="120" placeholder="Ex.: Bovespa, 2024 e 2025">
      <label for="pedido-mensagem">O que o COIR mostrou (a mensagem da importação)</label>
      <textarea id="pedido-mensagem" name="mensagem" rows="4" maxlength="2000"></textarea>
      <label for="pedido-email">Seu e-mail, se quiser resposta (opcional)</label>
      <div class="linha-form">
        <input id="pedido-email" type="email" name="email" autocomplete="email" pattern="[^@ ]+@[^@ ]+[.][^@ ]+" title="E-mail completo, como voce@exemplo.com" maxlength="254" placeholder="voce@exemplo.com">
        <button type="submit" class="botao principal">Enviar pedido</button>
      </div>
      <input type="hidden" name="_subject" value="COIR: pedido de corretora">
      <input type="text" name="_gotcha" tabindex="-1" autocomplete="off" aria-hidden="true" style="position:absolute;left:-9999px;width:1px;height:1px;opacity:0">
      <p id="pedido-situacao" class="pequeno" role="status"></p>
    </form>
  </div>
</section>

</main>

<footer>
  <div class="wrap">
    <span>COIR é gratuito e software livre (licença AGPL-3.0). <a href="{GH}">Código-fonte</a> · <a href="{raiz}legacy/index.html">Versão 2 (planilha)</a> · <a href="{raiz}docs/manual/se-nao-abrir.html">Se não abrir</a></span>
    <span>Suporte: <a href="mailto:{EMAIL_SUPORTE}">{EMAIL_SUPORTE}</a> · <a href="{INSTAGRAM}">Instagram</a> e <a href="{YOUTUBE}">YouTube</a>: {PERFIL} · Não substitui a orientação de um contador.</span>
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

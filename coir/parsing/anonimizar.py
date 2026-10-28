"""Cópia anonimizada do TEXTO de uma nota (o que o COIR lê), para enviar a quem desenvolve.

Nada muda de lugar: cada dado pessoal é trocado no próprio lugar por outro do mesmo formato e
comprimento (dígito por dígito, letra por letra, maiúsculas e pontuação mantidas). É o texto, e
não o PDF, que se anonimiza: mexer no PDF muda a ordem do texto extraído (na 3.0.1 a conta da BTG
foi parar na última linha das cópias, e o leitor ajustado a elas errou nas notas reais).

O que é trocado:
- CPF (com ou sem máscara) por um CPF fictício com dígitos verificadores válidos;
- conta do cliente, em todas as ocorrências (com ou sem zeros à esquerda), pela conta fictícia do
  mapa: a mesma conta real vira sempre a mesma fictícia;
- nome do cliente, onde aparecer (também com as palavras coladas do layout antigo da XP);
- o bloco de endereço que segue o nome (rua, número, bairro, CEP, cidade, telefone, e-mail);
- código do assessor, CEP solto sob "Cliente" (BTG BM&F) e banco, agência e conta corrente.

Dois portões decidem se a cópia vale (``conferir``): o de fidelidade extrai o original e a cópia e
exige o mesmo resultado, fora a conta (trocada pela do mapa); o de privacidade procura na cópia
cada termo pessoal achado no original e os termos da lista local. Qualquer achado recusa a cópia.
"""

import hashlib
import hmac
import re
import secrets
import unicodedata
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field

from pydantic import BaseModel, ConfigDict, Field

from coir.domain.cpf import cpf_valido, so_digitos
from coir.erros import AnonimizacaoError, ExtracaoError
from coir.parsing.cliente import RE_CODIGO, RE_CPF
from coir.parsing.extrair import ResultadoExtracao, extrair_textos
from coir.parsing.numeros import compactar
from coir.parsing.paginas import classificar
from coir.parsing.texto_pdf import SEPARADOR_PAGINAS

SEM_CORRETORA = "?"  # página que nenhum perfil reconhece (corretora nova)

RE_LINHA = re.compile(r"[^\r\n]+")
RE_PALAVRA = re.compile(r"[^\W_]+")  # letras (com acento) e dígitos
RE_LETRAS = re.compile(r"[^\W\d_]+")
RE_CODIGO_CONTA = re.compile(r"^\s*\d{1,4}-\d\s+(\d+)\s*$")  # "<código>-<dígito> <conta>"
RE_SO_NUMERO = re.compile(r"^\s*\d[\d.\-/ ]*\s*$")
# rótulos que ficam no bloco do cliente (o resto do bloco é dado pessoal)
ROTULOS_DO_BLOCO = frozenset({"codigo", "cliente", "assessor", "tel", "fone", "fax", "e", "mail"})
# linha que encerra o bloco de endereço (rótulo de outro quadro do cabeçalho)
FIM_DO_BLOCO = (
    "participantedestino", "corretora", "cliente", "c/vmercadoria", "agentedecompensa",
    "negociosrealizados", "assessor", "banco", "c.n.p.j", "numerodacorretora", "c.p.f",
    "custodiante", "qnegociacao", "ouvidoria", "internet", "sac:",
)  # fmt: skip
NAO_E_NOME = (*FIM_DO_BLOCO, "valor", "agencia", "contacorrente")
RE_SO_LETRAS = re.compile(r"[^\W\d_]+(?:[ '?][^\W\d_?]*)*")
LINHAS_DO_BLOCO = 5  # no máximo, depois da linha do nome
VALOR_NA_LINHA_SEGUINTE = (
    "cliente", "assessor", "participantedestinodorepasse", "codigoclienteassessor",
)  # fmt: skip
VALOR_NA_LINHA_ANTERIOR = ("banco", "agencia", "contacorrente")
VOGAIS = "AEIOU"
CONSOANTES = "BCDFGHJKLMNPQRSTVXZ"
MINIMO_TERMO = 4  # termos menores (DA, 60, PE) não identificam ninguém e colidem com o texto


class MapaAnonimizacao(BaseModel):
    """Segredo e contas/CPFs fictícios já atribuídos. Fica fora do repositório (pasta de dados):
    com ele, a mesma conta real vira sempre a mesma fictícia. Nunca aparece em repr/log."""

    model_config = ConfigDict(extra="forbid")

    segredo: str = Field(repr=False)  # hex; sem ele não se reverte a troca por força bruta
    contas: dict[str, str] = Field(default_factory=dict, repr=False)  # "CORRETORA|real" -> fict.
    cpfs: dict[str, str] = Field(default_factory=dict, repr=False)  # real -> fictício (dígitos)

    @classmethod
    def novo(cls) -> "MapaAnonimizacao":
        """Mapa vazio com segredo aleatório (só na primeira anonimização da pasta de dados)."""
        return cls(segredo=secrets.token_hex(32))

    def __repr__(self) -> str:
        return f"MapaAnonimizacao({len(self.contas)} conta(s), {len(self.cpfs)} CPF(s))"

    __str__ = __repr__

    def _fluxo(self, chave: str) -> Iterable[int]:
        """Bytes pseudoaleatórios determinísticos (HMAC-SHA256 do segredo) para a chave."""
        contador = 0
        while True:
            bloco = hmac.new(
                bytes.fromhex(self.segredo), f"{chave}#{contador}".encode(), hashlib.sha256
            ).digest()
            yield from bloco
            contador += 1

    def digitos(self, chave: str, n: int) -> str:
        fluxo = iter(self._fluxo(chave))
        return "".join(str(next(fluxo) % 10) for _ in range(n))

    def conta(self, corretora: str, real: str) -> str:
        """Conta fictícia (sem zeros à esquerda, com no máximo os dígitos da real)."""
        real = real.lstrip("0") or "0"
        chave = f"{corretora}|{real}"
        if chave not in self.contas:
            usadas = set(self.contas.values()) | {k.split("|", 1)[1] for k in self.contas}
            tentativa = 0
            while True:
                d = self.digitos(f"conta:{chave}:{tentativa}", len(real))
                nova = str(1 + int(d[0]) % 9) + d[1:]
                if nova != real and nova not in usadas:
                    break
                tentativa += 1
            self.contas[chave] = nova
        return self.contas[chave]

    def cpf(self, real: str) -> str:
        """CPF fictício (11 dígitos, verificadores válidos) para o CPF real (só dígitos)."""
        if real not in self.cpfs:
            usados = set(self.cpfs.values())
            tentativa = 0
            while True:
                base = self.digitos(f"cpf:{real}:{tentativa}", 9)
                nova = base + _dv_cpf(base)
                if cpf_valido(nova) and nova != real and nova not in usados:
                    break
                tentativa += 1
            self.cpfs[real] = nova
        return self.cpfs[real]

    def trocar(self, tipo: str, texto: str) -> str:
        """Mesmo formato: dígito por dígito, vogal por vogal, consoante por consoante (caixa e
        pontuação mantidas; letra acentuada vira letra simples). Sempre diferente do original."""
        tentativa = 0
        while True:
            fluxo = iter(self._fluxo(f"{tipo}:{compactar(texto)}:{tentativa}"))
            saida = []
            for c in texto:
                if c.isdigit():
                    saida.append(str(next(fluxo) % 10))
                elif c.isalpha():
                    base = _sem_acento(c).upper()
                    grupo = VOGAIS if base in VOGAIS else CONSOANTES
                    novo = grupo[next(fluxo) % len(grupo)]
                    saida.append(novo if c.isupper() else novo.lower())
                else:
                    saida.append(c)
            nova = "".join(saida)
            if compactar(nova) != compactar(texto) or not any(c.isalnum() for c in texto):
                return nova
            tentativa += 1


def _dv_cpf(base: str) -> str:
    nums = [int(c) for c in base]
    for pos in (9, 10):
        soma = sum(n * peso for n, peso in zip(nums[:pos], range(pos + 1, 1, -1), strict=True))
        nums.append(soma * 10 % 11 % 10)
    return f"{nums[9]}{nums[10]}"


def _sem_acento(texto: str) -> str:
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode() or texto


# ------------------------------------------------------------------------------- anonimização
@dataclass
class _Pagina:
    """Texto de uma página com as trocas já feitas e o que foi trocado (para o portão)."""

    original: str
    corretora: str
    chars: list[str] = field(default_factory=list)
    trocado: bytearray = field(default_factory=bytearray)
    termos: set[str] = field(default_factory=set)
    com_bloco: bool = False  # achou o nome e trocou o bloco do cliente

    def __post_init__(self) -> None:
        self.chars = list(self.original)
        self.trocado = bytearray(len(self.original))

    def livre(self, inicio: int, fim: int) -> bool:
        return not any(self.trocado[inicio:fim])

    def trocar(self, inicio: int, fim: int, novo: str) -> None:
        if len(novo) != fim - inicio:
            raise AnonimizacaoError("a troca mudaria o comprimento do texto")  # nunca deveria
        if not self.livre(inicio, fim):
            return
        self.termos.add(self.original[inicio:fim])
        self.chars[inicio:fim] = list(novo)
        self.trocado[inicio:fim] = b"\x01" * (fim - inicio)

    @property
    def texto(self) -> str:
        return "".join(self.chars)

    def publico(self) -> str:
        """Linhas do original sem nenhuma troca (o que é da corretora, não do cliente). Página em
        que o bloco do cliente não foi achado não conta: o endereço dela pode ter ficado."""
        if not self.com_bloco:
            return ""
        linhas = _linhas(self.original)
        return "\n".join(m.group(0) for m in linhas if self.livre(m.start(), m.end()))


@dataclass(frozen=True)
class Anonimizada:
    textos: list[str]
    termos: frozenset[str]  # tudo o que foi trocado no original (para o portão de privacidade)
    publico: str  # o original sem os trechos trocados, de todas as páginas


def anonimizar(textos: Sequence[str], mapa: MapaAnonimizacao) -> Anonimizada:
    """Troca os dados pessoais de cada página; o mapa ganha as contas e CPFs novos."""
    contas = _contas(textos)
    cpfs = {so_digitos(m.group(0)) for t in textos for m in RE_CPF.finditer(t)}
    achados = {n for t in textos if (n := _nome(_linhas(t))) is not None}
    nomes: list[str] = sorted(achados, key=len, reverse=True)
    padroes = [_padrao_do_nome(n) for n in nomes]
    paginas = [_Pagina(t, _corretora(t)) for t in textos]
    for p in paginas:
        _trocar_cpfs(p, cpfs, mapa)
        _trocar_contas(p, contas, mapa)
        for nome, padrao in zip(nomes, padroes, strict=True):
            _trocar_nome(p, nome, padrao, mapa)
        p.com_bloco = _trocar_bloco_do_cliente(p, padroes, mapa)
        _trocar_valores_rotulados(p, mapa)
    termos = frozenset(t for p in paginas for t in p.termos) | frozenset(nomes)
    return Anonimizada([p.texto for p in paginas], termos, "\n".join(p.publico() for p in paginas))


def _linhas(texto: str) -> list[re.Match[str]]:
    return [m for m in RE_LINHA.finditer(texto) if m.group(0).strip()]


def _corretora(texto: str) -> str:
    try:
        return classificar(texto).perfil.corretora
    except ExtracaoError:
        return SEM_CORRETORA


def _contas(textos: Sequence[str]) -> set[tuple[str, str]]:
    """(corretora, conta) de cada nota lida; nas páginas que o leitor recusa, a conta impressa
    no formato SINACOR ("<código>-<dígito> <conta>") ou sob "Codigo do Cliente"."""
    res = extrair_textos("", textos)
    contas = {(n.corretora, n.conta) for n in res.notas}
    for texto in textos:
        linhas = _linhas(texto)
        corretora = _corretora(texto)
        for i, m in enumerate(linhas):
            achado = RE_CODIGO_CONTA.match(m.group(0))
            if achado is None and i and compactar(linhas[i - 1].group(0)) == "codigodocliente":
                achado = re.match(r"^\s*(\d+)\s*$", m.group(0))
            if achado:
                contas.add((corretora, achado.group(1)))
    return {(c, conta.lstrip("0")) for c, conta in contas if len(conta.lstrip("0")) >= MINIMO_TERMO}


def _trocar_cpfs(p: _Pagina, cpfs: set[str], mapa: MapaAnonimizacao) -> None:
    for m in RE_CPF.finditer(p.original):
        p.trocar(m.start(), m.end(), _formatado(m.group(0), mapa.cpf(so_digitos(m.group(0)))))
    for real in cpfs:  # o mesmo CPF sem máscara
        for m in re.finditer(rf"(?<!\d){real}(?!\d)", p.original):
            p.trocar(m.start(), m.end(), mapa.cpf(real))


def _formatado(modelo: str, digitos: str) -> str:
    fila = iter(digitos)
    return "".join(next(fila) if c.isdigit() else c for c in modelo)


def _trocar_contas(p: _Pagina, contas: set[tuple[str, str]], mapa: MapaAnonimizacao) -> None:
    for corretora, real in contas:
        if corretora not in (p.corretora, SEM_CORRETORA) and p.corretora != SEM_CORRETORA:
            continue
        ficticia = mapa.conta(corretora, real)
        for m in re.finditer(rf"(?<![\d.,/])(0*){real}(?!\d)(?![.,]\d)", p.original):
            largura = m.end() - m.start()
            if len(ficticia) > largura:
                raise AnonimizacaoError("conta fictícia com mais dígitos que a impressa")
            p.trocar(m.start(), m.end(), ficticia.zfill(largura))


def _nome(linhas: list[re.Match[str]]) -> str | None:
    i = _indice_do_nome(linhas)
    return None if i is None else _nome_da_linha(linhas[i].group(0))


def _nome_da_linha(linha: str) -> str | None:
    """A linha sem o código do cliente à esquerda e sem o CPF; None se não parecer um nome."""
    m = RE_CPF.search(linha)
    nome = " ".join(RE_CODIGO.sub("", linha[: m.start()] if m else linha).split())
    if not RE_SO_LETRAS.fullmatch(nome) or compactar(nome).startswith(NAO_E_NOME):
        return None
    return nome if re.search(r"[^\W\d_]{2}", nome) else None


def _indice_do_nome(linhas: list[re.Match[str]]) -> int | None:
    """Linha do nome: a do CPF (Ágora, BTG); a primeira depois do rodapé de atendimento
    ("... Ouvidoria ..."), pulando o rótulo "Cliente" (XP, Rico); ou, se essa for rótulo, a
    última linha só de letras antes do CPF que não seja rótulo (XP BM&F junto da Bovespa)."""
    i_cpf = next((i for i, m in enumerate(linhas) if RE_CPF.search(m.group(0))), None)
    if i_cpf is None:
        return None
    if _nome_da_linha(linhas[i_cpf].group(0)) is not None:
        return i_cpf
    rodape = [j for j in range(i_cpf) if "ouvidoria" in linhas[j].group(0).lower()]
    if rodape:
        j = rodape[-1] + 1
        if j < i_cpf and compactar(linhas[j].group(0)) == "cliente":
            j += 1
        if j < i_cpf and _nome_da_linha(linhas[j].group(0)) is not None:
            return j
    return next(
        (j for j in range(i_cpf - 1, -1, -1) if _nome_da_linha(linhas[j].group(0)) is not None),
        None,
    )


def _padrao_do_nome(nome: str) -> re.Pattern[str]:
    """Palavras do nome com ou sem espaço entre elas, com ou sem acento ("?" no lugar de uma
    letra acentuada, como sai em notas antigas da Rico), sem diferenciar caixa."""

    def letra(c: str) -> str:
        base = _sem_acento(c).casefold()
        return f"(?:{re.escape(c)}|[{base}{base.upper()}]|\\?|[^\\x00-\\x7f])"

    palavras = ["".join(letra(c) for c in p) for p in RE_LETRAS.findall(nome)]
    return re.compile(r"(?<![^\W\d_])" + r"\s*".join(palavras) + r"(?![^\W\d_])", re.IGNORECASE)


def _trocar_nome(p: _Pagina, nome: str, padrao: re.Pattern[str], mapa: MapaAnonimizacao) -> None:
    ficticio = "".join(mapa.trocar("nome", w) for w in RE_LETRAS.findall(nome))
    for m in padrao.finditer(p.original):
        letras = iter(ficticio)
        novo = "".join(c if c.isspace() else next(letras) for c in m.group(0))
        original = m.group(0)
        # caixa como no original (o layout antigo da XP mistura)
        novo = "".join(
            n.lower() if o.islower() else n.upper() for n, o in zip(novo, original, strict=True)
        )
        p.trocar(m.start(), m.end(), novo)


def _trocar_bloco_do_cliente(
    p: _Pagina, padroes: Sequence[re.Pattern[str]], mapa: MapaAnonimizacao
) -> bool:
    """A linha do nome e as de endereço que a seguem, até o próximo rótulo do cabeçalho. A linha
    do nome é a primeira com um dos nomes achados no arquivo (o layout pode variar entre as
    páginas); sem nenhum, a que a página indicar."""
    linhas = _linhas(p.original)
    i = next((j for j, m in enumerate(linhas)
              if any(x.search(m.group(0)) for x in padroes)), None)  # fmt: skip
    if i is None:
        i = _indice_do_nome(linhas)
    if i is None:
        return False
    bloco = [linhas[i]]
    for m in linhas[i + 1 : i + 1 + LINHAS_DO_BLOCO]:
        c = compactar(m.group(0))
        if (
            c.startswith(FIM_DO_BLOCO)
            or RE_CODIGO_CONTA.match(m.group(0))
            or RE_CPF.fullmatch(m.group(0).strip())
        ):
            break
        bloco.append(m)
    for m in bloco:
        _trocar_palavras(p, m.start(), m.end(), mapa)
    return True


def _trocar_palavras(p: _Pagina, inicio: int, fim: int, mapa: MapaAnonimizacao) -> None:
    for w in RE_PALAVRA.finditer(p.original, inicio, fim):
        if compactar(w.group(0)) in ROTULOS_DO_BLOCO or not p.livre(w.start(), w.end()):
            continue
        p.trocar(w.start(), w.end(), mapa.trocar("bloco", w.group(0)))


def _trocar_valores_rotulados(p: _Pagina, mapa: MapaAnonimizacao) -> None:
    """Números junto a rótulos do cabeçalho: assessor e CEP/código (valor na linha seguinte),
    banco, agência e conta corrente (valor antes do rótulo, na Rico; na linha seguinte, na
    Ágora)."""
    linhas = _linhas(p.original)
    for i, m in enumerate(linhas):
        c = compactar(m.group(0))
        alvo = None
        if c in VALOR_NA_LINHA_SEGUINTE and i + 1 < len(linhas):
            alvo = linhas[i + 1]
        elif c.startswith(VALOR_NA_LINHA_ANTERIOR) and i:
            alvo = linhas[i - 1]
        if (
            alvo is not None
            and RE_SO_NUMERO.match(alvo.group(0))
            and re.search(r"[1-9]", alvo.group(0))
        ):
            _trocar_palavras(p, alvo.start(), alvo.end(), mapa)
        if c.startswith("bancoagenciacontacorrente") and i + 1 < len(linhas):
            seguinte = linhas[i + 1]
            for w in re.finditer(r"(?<!\S)\d+(?!\S)", seguinte.group(0)):
                ini, fim = seguinte.start() + w.start(), seguinte.start() + w.end()
                if p.livre(ini, fim):
                    p.trocar(ini, fim, mapa.trocar("banco", w.group(0)))


# ------------------------------------------------------------------------------------ portões
def divergencias_de_fidelidade(
    original: ResultadoExtracao, copia: ResultadoExtracao, mapa: MapaAnonimizacao
) -> list[str]:
    """O que a cópia extrai diferente do original. Só a conta pode mudar, e só pela do mapa."""
    erros: list[str] = []
    if len(original.notas) != len(copia.notas):
        erros.append(f"{len(original.notas)} nota(s) no original e {len(copia.notas)} na cópia")
    for a, b in zip(original.notas, copia.notas, strict=False):
        onde = f"nota {a.numero} de {a.data_pregao:%d/%m/%Y}"
        esperada = mapa.contas.get(f"{a.corretora}|{a.conta.lstrip('0') or '0'}", a.conta)
        if b.conta != esperada:
            erros.append(f"{onde}: campo conta (a fictícia não é a do mapa)")
        da, db = a.model_dump(exclude={"conta"}), b.model_dump(exclude={"conta"})
        erros += [f"{onde}: campo {campo}" for campo in da if da[campo] != db.get(campo)]
        ca = original.clientes.get(a.chave)
        cb = copia.clientes.get(b.chave)
        if (ca is None) != (cb is None) or (ca and cb and mapa.cpfs.get(ca.cpf) != cb.cpf):
            erros.append(f"{onde}: campo cliente (CPF)")
    rej_a = [(r.paginas, r.motivo, r.nota, r.data) for r in original.rejeitadas]
    rej_b = [(r.paginas, r.motivo, r.nota, r.data) for r in copia.rejeitadas]
    for ra, rb in zip(rej_a, rej_b, strict=False):
        if ra != rb:
            erros.append(f"página(s) {list(ra[0])}: recusa diferente na cópia")
    if len(rej_a) != len(rej_b):
        erros.append(f"{len(rej_a)} recusa(s) no original e {len(rej_b)} na cópia")
    if original.avisos != copia.avisos:
        erros.append("avisos da extração diferentes na cópia")
    return erros


def divergencias_de_forma(original: Sequence[str], copia: Sequence[str]) -> list[str]:
    """Nenhuma linha muda de lugar: mesmo comprimento e, posição a posição, o mesmo caractere ou
    outro da mesma classe (dígito por dígito, letra por letra)."""
    if len(original) != len(copia):
        return [f"{len(original)} página(s) no original e {len(copia)} na cópia"]
    erros = []
    for n, (a, b) in enumerate(zip(original, copia, strict=True), start=1):
        if len(a) != len(b) or any(
            x != y and not ((x.isdigit() and y.isdigit()) or (x.isalpha() and y.isalpha()))
            for x, y in zip(a, b, strict=True)
        ):
            erros.append(f"página {n}: texto fora do lugar")
    return erros


def vazamentos(
    copia: Sequence[str], termos: Iterable[str], publico: str, sensiveis: Sequence[str] = ()
) -> list[str]:
    """Termos pessoais do original que ainda aparecem na cópia (descrição, nunca o valor).

    Cada termo trocado (nome, palavra ou número do endereço, CPF, conta, assessor...) com 4 ou
    mais letras/dígitos é procurado na cópia: palavras sem espaço nem acento (o layout antigo da
    XP cola as palavras); números com ou sem separadores. Termo que também está na parte pública
    do original (dados da corretora) não identifica o cliente e fica de fora. ``sensiveis`` é a
    lista local de termos (como em tools/verificar_dados_pessoais.py), procurada sem caixa."""
    achados: list[str] = []
    publico_compacto = compactar(publico)
    partes = sorted({p for t in termos for p in _partes(t)})
    for n, texto in enumerate(copia, start=1):
        compacto, caixa = compactar(texto), texto.casefold()
        for parte in partes:
            if parte.isdigit():
                padrao = r"(?<!\d)" + r"[.\- ]?".join(parte) + r"(?!\d)"
                if re.search(padrao, texto) and not re.search(padrao, publico):
                    achados.append(f"página {n}: número pessoal de {len(parte)} dígitos na cópia")
            elif compactar(parte) in compacto and compactar(parte) not in publico_compacto:
                achados.append(f"página {n}: palavra pessoal de {len(parte)} letras na cópia")
        if any(t in caixa for t in sensiveis):
            achados.append(f"página {n}: termo da lista local de termos sensíveis na cópia")
    return list(dict.fromkeys(achados))


def _partes(termo: str) -> list[str]:
    """Pedaços que identificam: número sem máscara (CPF, conta sem zeros à esquerda) e palavras."""
    partes: list[str] = []
    digitos = so_digitos(termo)
    if len(digitos) >= MINIMO_TERMO and re.fullmatch(r"[\d.\-/ ]+", termo.strip()):
        partes.append(digitos.lstrip("0") or digitos)
    else:
        partes += [w for w in RE_PALAVRA.findall(termo) if len(w) >= MINIMO_TERMO]
        partes = [w.lstrip("0") if w.isdigit() else w for w in partes]
    nome_compacto = "".join(RE_LETRAS.findall(termo))
    if " " in termo.strip() and len(nome_compacto) >= MINIMO_TERMO and not digitos:
        partes.append(nome_compacto)
    return [p for p in partes if len(p) >= MINIMO_TERMO]


@dataclass(frozen=True)
class Conferida:
    textos: list[str]
    recusas: list[str]  # vazio: a cópia pode ser enviada

    @property
    def aceita(self) -> bool:
        return not self.recusas


def conferir(
    arquivo: str, textos: Sequence[str], mapa: MapaAnonimizacao, sensiveis: Sequence[str] = ()
) -> Conferida:
    """Anonimiza e passa pelos três portões (forma, fidelidade e privacidade)."""
    anon = anonimizar(textos, mapa)
    recusas = divergencias_de_forma(textos, anon.textos)
    if SEPARADOR_PAGINAS in "".join(textos):
        recusas.append("o texto tem o separador de páginas (\\f) dentro de uma página")
    recusas += divergencias_de_fidelidade(
        extrair_textos(arquivo, textos), extrair_textos(arquivo, anon.textos), mapa
    )
    recusas += vazamentos(anon.textos, anon.termos, anon.publico, sensiveis)
    partes = [p for t in anon.termos for p in _partes(t)]
    if any(compactar(p) in compactar(arquivo) for p in partes if not p.isdigit()) or any(
        p in so_digitos(arquivo) for p in partes if p.isdigit() and len(p) >= 6
    ):
        recusas.append("o nome do arquivo tem dado pessoal")
    return Conferida(anon.textos, recusas)

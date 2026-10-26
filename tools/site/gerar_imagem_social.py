"""Imagem de prévia do link do site (og:image, 1280x640): texto da v3 e a tela do Painel do banco
de demonstração. O nome do arquivo leva a versão porque os aplicativos de mensagem guardam a prévia
pelo endereço da imagem. Uso: uv run --with pillow python tools/site/gerar_imagem_social.py
"""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

RAIZ = Path(__file__).resolve().parents[2]
SAIDA = RAIZ / "docs" / "social-preview-v3.png"
TELA = RAIZ / "docs" / "site" / "imagens" / "tela-painel.jpg"
FONTES = Path("C:/Windows/Fonts")

MARINHO, VERDE, AZUL_CLARO, CLARO = "#102a43", "#2ea043", "#a8d4ff", "#d9e2ec"
LARGURA, ALTURA, MARGEM = 1280, 640, 64


def fonte(nome: str, tamanho: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(FONTES / nome), tamanho)


def quebrar(desenho: ImageDraw.ImageDraw, texto: str, f: ImageFont.FreeTypeFont, largura: int):
    linhas, atual = [], ""
    for palavra in texto.split():
        teste = f"{atual} {palavra}".strip()
        if desenho.textlength(teste, font=f) <= largura:
            atual = teste
        else:
            linhas.append(atual)
            atual = palavra
    return [*linhas, atual]


def main() -> None:
    img = Image.new("RGB", (LARGURA, ALTURA), MARINHO)
    d = ImageDraw.Draw(img)
    d.rectangle((0, 0, 15, ALTURA), fill=VERDE)

    largura_texto = 520
    y = 56
    d.text((MARGEM, y), "COIR", font=fonte("segoeuib.ttf", 112), fill="white")
    y += 150
    for linha in ("Controle de Operações", "e Imposto de Renda"):
        d.text((MARGEM, y), linha, font=fonte("segoeuib.ttf", 36), fill=AZUL_CLARO)
        y += 46
    y += 22
    corpo = fonte("segoeui.ttf", 26)
    texto = (
        "Importa as notas de corretagem e as planilhas da B3 e calcula o preço médio, "
        "o IR de cada mês, o DARF e a declaração anual."
    )
    for linha in quebrar(d, texto, corpo, largura_texto):
        d.text((MARGEM, y), linha, font=corpo, fill=CLARO)
        y += 36
    rodape = fonte("segoeuib.ttf", 22)
    d.text((MARGEM, ALTURA - 112), "Versão 3 · Windows · Gratuito · AGPL-3.0", font=rodape,
           fill=VERDE)  # fmt: skip
    d.text((MARGEM, ALTURA - 76), "Seus dados não saem do computador", font=rodape, fill=CLARO)

    tela = Image.open(TELA).convert("RGB")
    larg = LARGURA - 640 - MARGEM + 24
    tela = tela.resize((larg, round(tela.height * larg / tela.width)), Image.Resampling.LANCZOS)
    x, y_tela = 640, (ALTURA - tela.height) // 2
    d.rounded_rectangle((x - 6, y_tela - 6, x + tela.width + 5, y_tela + tela.height + 5), 10,
                        fill="#3d5a7a")  # fmt: skip
    img.paste(tela, (x, y_tela))
    img.save(SAIDA, optimize=True)
    print(SAIDA)


if __name__ == "__main__":
    main()

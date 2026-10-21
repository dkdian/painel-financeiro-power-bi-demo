"""Monta um vídeo de demonstração no padrão dos vídeos do portfólio (1280x720, H.264, sem som, tela escura).

Cada cena: círculo verde com o número da etapa + título no alto, etiqueta "dados fictícios" no canto, captura no meio, 1 a 3 tópicos embaixo.
Marcação nos tópicos: **negrito** (branco) e ~~destaque~~ (amarelo).
Uso: python montar_video.py <roteiro.py> <pasta de saída>   (o roteiro define SLUG, SLIDES; ver roteiro_painel.py)
"""
import importlib.util
import re
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

W, H = 1280, 720
BG, PAINEL, BORDA = (9, 20, 33), (14, 29, 45), (30, 49, 68)
VERDE, TXT, MUDO = (41, 174, 139), (207, 224, 242), (124, 162, 206)
AMARELO, VERM, VERDE_T, BRANCO, ETIQ = (255, 211, 110), (255, 143, 143), (127, 227, 168), (255, 255, 255), (128, 143, 168)
FD = "C:/Windows/Fonts/"


def F(tam, estilo="reg"):
    nome = {"reg": "segoeui.ttf", "bold": "segoeuib.ttf", "semi": "seguisb.ttf", "luz": "segoeuil.ttf"}[estilo]
    try:
        return ImageFont.truetype(FD + nome, tam)
    except OSError:
        return ImageFont.truetype(FD + "segoeuib.ttf" if estilo != "reg" else FD + "segoeui.ttf", tam)


def largura(d, txt, fonte):
    return d.textlength(txt, font=fonte)


def novo():
    img = Image.new("RGB", (W, H), BG)
    return img, ImageDraw.Draw(img)


def cabecalho(d, num, titulo):
    if num is not None:
        d.ellipse([40, 26, 76, 62], fill=VERDE)
        f = F(18 if len(str(num)) < 2 else 16, "bold")
        d.text((58, 44), str(num), font=f, fill=BRANCO, anchor="mm")
    d.text((90, 44), titulo, font=F(29, "semi"), fill=BRANCO, anchor="lm")
    d.text((1240, 43), "dados fictícios", font=F(14), fill=ETIQ, anchor="rm")


def partes(txt):
    """'**a** b ~~c~~' -> [(texto, cor, estilo)]"""
    out = []
    for p in re.split(r"(\*\*.+?\*\*|~~.+?~~)", txt):
        if not p:
            continue
        if p.startswith("**"):
            out.append((p[2:-2], BRANCO, "semi"))
        elif p.startswith("~~"):
            out.append((p[2:-2], AMARELO, "reg"))
        else:
            out.append((p, TXT, "reg"))
    return out


def topicos(d, itens, y0=583, tam=21, passo=35, x0=64, maxw=1180):
    y = y0
    for it in itens:
        d.ellipse([44, y + 3, 54, y + 13], fill=VERDE)
        x = x0
        # quebra simples por palavras mantendo cor/estilo
        for texto, cor, est in partes(it):
            f = F(tam, est)
            if texto[:1].isspace() and x > x0:
                x += largura(d, ' ', f)
            for palavra in re.findall(r"\S+\s*", texto):
                w = largura(d, palavra, f)
                if x + w > x0 + maxw:
                    y += passo - 6
                    x = x0
                d.text((x, y), palavra, font=f, fill=cor)
                x += w
        y += passo


def slide_titulo(linhas, sub, texto):
    img, d = novo()
    y = 255 - (len(linhas) - 1) * 28
    for ln in linhas:
        d.text((W // 2, y), ln, font=F(52, "semi"), fill=BRANCO, anchor="mm")
        y += 58
    y += 18
    d.text((W // 2, y), sub, font=F(26, "semi"), fill=AMARELO, anchor="mm")
    y += 58
    for ln in texto:
        d.text((W // 2, y), ln, font=F(21), fill=MUDO, anchor="mm")
        y += 37
    return img


def slide_cena(num, titulo, captura, itens):
    img, d = novo()
    cabecalho(d, num, titulo)
    x0, y0, x1, y1 = 40, 80, 1240, 565
    d.rounded_rectangle([x0, y0, x1, y1], radius=8, fill=PAINEL, outline=BORDA)
    cap = captura.convert("RGB")
    esc = min((x1 - x0 - 4) / cap.width, (y1 - y0 - 4) / cap.height)
    cap = cap.resize((max(1, int(cap.width * esc)), max(1, int(cap.height * esc))), Image.LANCZOS)
    img.paste(cap, (x0 + (x1 - x0 - cap.width) // 2, y0 + (y1 - y0 - cap.height) // 2))
    topicos(d, itens)
    return img


def slide_antes_depois(num, linhas, itens, destaque=None):
    """linhas: [(indicador, antes, depois)] como na tabela da página; destaque = índice da linha em negrito."""
    img, d = novo()
    cabecalho(d, num, "Antes e depois")
    d.rounded_rectangle([40, 78, 1240, 568], radius=12, fill=PAINEL, outline=BORDA)
    xs = (89, 400, 745)
    wcol = (290, 320, 470)
    for x, t in zip(xs, ("INDICADOR", "ANTES", "DEPOIS")):
        d.text((x, 104), t, font=F(14, "semi"), fill=MUDO, anchor="lm")
    d.line([(75, 126), (1205, 126)], fill=(33, 48, 64))
    y = 126
    altura = (568 - 20 - y) // len(linhas)
    for i, (ind, ant, dep) in enumerate(linhas):
        forte = i == destaque
        for x, w, txt, cor in zip(xs, wcol, (ind, ant, dep), (BRANCO, VERM, VERDE_T)):
            f = F(21, "bold" if forte else "reg")
            palavras, ls, atual = txt.split(), [], ""
            for p in palavras:
                t = (atual + " " + p).strip()
                if largura(d, t, f) > w and atual:
                    ls.append(atual)
                    atual = p
                else:
                    atual = t
            ls.append(atual)
            ty = y + altura // 2 - (len(ls) - 1) * 13
            for ln in ls:
                d.text((x, ty), ln, font=f, fill=cor, anchor="lm")
                ty += 27
        y += altura
        if i < len(linhas) - 1:
            d.line([(75, y), (1205, y)], fill=(33, 48, 64))
    topicos(d, itens, y0=588)
    return img


def construir(roteiro, saida):
    spec = importlib.util.spec_from_file_location("roteiro", roteiro)
    r = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(r)
    saida = Path(saida)
    quadros = saida / ("_quadros_" + r.SLUG)
    quadros.mkdir(parents=True, exist_ok=True)
    for f in quadros.glob("*.png"):
        f.unlink()
    lista = []
    for i, s in enumerate(r.SLIDES):
        tipo = s["tipo"]
        if tipo == "titulo":
            img = slide_titulo(s["linhas"], s["sub"], s["texto"])
        elif tipo == "cena":
            img = slide_cena(s["num"], s["titulo"], Image.open(s["captura"]), s["itens"])
        else:
            img = slide_antes_depois(s["num"], s["linhas"], s["itens"], s.get("destaque"))
        p = quadros / ("%02d.png" % i)
        img.save(p)
        lista.append((p, s["dur"]))
    capa = saida / (r.SLUG + "-demo-capa.jpg")
    Image.open(lista[0][0]).save(capa, quality=90)
    # concat com fade de entrada e saída (0,5 s) em cada cena
    filtros, entradas = [], []
    for k, (p, dur) in enumerate(lista):
        entradas += ["-loop", "1", "-framerate", "15", "-t", str(dur), "-i", str(p)]
        filtros.append("[%d:v]fade=t=in:st=0:d=0.5,fade=t=out:st=%.2f:d=0.5,format=yuv420p[v%d]" % (k, dur - 0.5, k))
    cadeia = "".join("[v%d]" % k for k in range(len(lista))) + "concat=n=%d:v=1:a=0[v]" % len(lista)
    mp4 = saida / (r.SLUG + "-demo.mp4")
    cmd = ["ffmpeg", "-v", "error", "-y", *entradas, "-filter_complex", ";".join(filtros) + ";" + cadeia, "-map", "[v]", "-an",
           "-c:v", "libx264", "-preset", getattr(r, "PRESET", "slow"), "-crf", str(getattr(r, "CRF", 24)), "-pix_fmt", "yuv420p", "-r", "15",
           "-map_metadata", "-1", "-movflags", "+faststart", str(mp4)]
    subprocess.run(cmd, check=True)
    print("vídeo:", mp4, "%.1f MB" % (mp4.stat().st_size / 1e6), "| capa:", capa, "| duração total: %d s" % sum(d for _, d in lista))


if __name__ == "__main__":
    construir(sys.argv[1], sys.argv[2])

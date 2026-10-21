"""Cobre áreas de uma captura (conta logada, nome de usuário) com a cor vizinha. Uso: python mascara.py entrada.png saida.png x0,y0,x1,y1 [recorte x0,y0,x1,y1]"""
import sys
from PIL import Image, ImageDraw
img = Image.open(sys.argv[1]).convert("RGB")
x0, y0, x1, y1 = [int(v) for v in sys.argv[3].split(",")]
# preenche cada linha com a cor do pixel imediatamente à esquerda da área (mantém o degradê da barra de título)
for y in range(y0, y1 + 1):
    cor = img.getpixel((max(x0 - 4, 0), y))
    ImageDraw.Draw(img).line([(x0, y), (x1, y)], fill=cor)
if len(sys.argv) > 4:
    img = img.crop(tuple(int(v) for v in sys.argv[4].split(",")))
img.save(sys.argv[2]); print("ok", img.size)

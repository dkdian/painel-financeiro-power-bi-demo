"""Quadros desenhados (sem captura de tela) para as cenas em que não há uma janela para mostrar. Tudo fictício."""
import sys
from pathlib import Path

from PIL import Image, ImageDraw

sys.path.insert(0, str(Path(__file__).parent))
from montar_video import AMARELO, BG, BORDA, BRANCO, MUDO, PAINEL, TXT, VERDE, VERDE_T, F  # noqa: E402

SAIDA = Path(r"C:\Demo\Painel_BI\_capturas")


def cartao(d, box, titulo):
    d.rounded_rectangle(box, radius=10, fill=(17, 35, 54), outline=BORDA)
    d.text((box[0] + 24, box[1] + 26), titulo, font=F(21, "semi"), fill=BRANCO, anchor="lm")


def seta(d, x0, x1, y, texto):
    d.line([(x0, y), (x1 - 12, y)], fill=VERDE, width=4)
    d.polygon([(x1, y), (x1 - 16, y - 10), (x1 - 16, y + 10)], fill=VERDE)
    d.text(((x0 + x1) // 2, y - 22), texto, font=F(15), fill=MUDO, anchor="mm")


def pasta_para_power_bi(destino):
    img = Image.new("RGB", (1200, 485), PAINEL)
    d = ImageDraw.Draw(img)
    cartao(d, (30, 40, 470, 445), "00_Painel BI  (a pasta do projeto)")
    arquivos = [("Controle Financeiro 2025.xlsx", "618 lançamentos", TXT), ("Controle Financeiro 2026.xlsm", "938 lançamentos", TXT),
                ("Controle Financeiro 2027.xlsx", "em janeiro, basta criar o arquivo do ano", AMARELO), ("AnáliseFinanceira.pbix", "o painel", MUDO)]
    y = 110
    for nome, nota, cor in arquivos:
        d.rounded_rectangle((54, y, 446, y + 62), radius=8, fill=(14, 29, 45), outline=BORDA)
        d.rectangle((70, y + 16, 92, y + 46), fill=(33, 115, 70) if "xls" in nome else (214, 168, 41))
        d.text((108, y + 22), nome, font=F(18, "semi"), fill=cor, anchor="lm")
        d.text((108, y + 44), nota, font=F(15), fill=MUDO, anchor="lm")
        y += 78
    seta(d, 490, 640, 242, "lê todos os arquivos")
    cartao(d, (655, 40, 1170, 445), "Power BI: consulta ControleFinanceiro")
    linhas = [("Folder.Contents", "lista os arquivos 'Controle Financeiro AAAA'"), ("Excel.Workbook", "abre a tabela ControleFinanceiro de cada ano"),
              ("Table.Combine", "empilha os anos numa tabela só")]
    y = 118
    for a, b in linhas:
        d.text((684, y), a, font=F(19, "semi"), fill=VERDE_T, anchor="lm")
        d.text((684, y + 26), b, font=F(16), fill=TXT, anchor="lm")
        y += 66
    d.rounded_rectangle((684, 330, 1140, 420), radius=10, fill=(14, 29, 45), outline=VERDE)
    d.text((912, 358), "1.556 lançamentos", font=F(26, "bold"), fill=BRANCO, anchor="mm")
    d.text((912, 392), "2025 + 2026, mesmas colunas e categorias", font=F(15), fill=MUDO, anchor="mm")
    img.save(destino)
    return destino


if __name__ == "__main__":
    SAIDA.mkdir(parents=True, exist_ok=True)
    print(pasta_para_power_bi(SAIDA / "painel" / "p04_pasta_power_bi.png"))

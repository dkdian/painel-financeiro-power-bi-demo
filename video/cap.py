"""Captura a área de uma janela (por trecho do título) ou da tela inteira e grava em PNG. Só lê a tela; não mexe em foco, mouse nem teclado.

  python cap.py listar
  python cap.py janela "<trecho do título>" <saida.png> [--topo N] [--esq N] [--dir N] [--base N]   (N = pixels cortados de cada borda)
  python cap.py tela <saida.png> [--monitor K]
Os cortes servem para tirar da imagem a conta logada (canto superior direito da barra de título do Office e do Power BI) e o que mais identifique.
"""
import ctypes
import sys

import win32gui
from PIL import ImageGrab

ctypes.windll.user32.SetProcessDPIAware()


def janelas():
    out = []

    def cb(h, _):
        if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h):
            out.append((h, win32gui.GetWindowText(h), win32gui.GetWindowRect(h)))
    win32gui.EnumWindows(cb, None)
    return out


def arg(nome, padrao=0):
    return int(sys.argv[sys.argv.index(nome) + 1]) if nome in sys.argv else padrao


def main():
    cmd = sys.argv[1]
    if cmd == "listar":
        for h, t, r in janelas():
            print(h, r, t.encode("ascii", "replace").decode())
    elif cmd == "janela":
        trecho, saida = sys.argv[2].lower(), sys.argv[3]
        achadas = [j for j in janelas() if trecho in j[1].lower()]
        if not achadas:
            sys.exit("janela não encontrada: " + trecho)
        h, t, (x0, y0, x1, y1) = max(achadas, key=lambda j: (j[2][2] - j[2][0]) * (j[2][3] - j[2][1]))
        # a janela maximizada tem uma borda invisível de 8 px: ajuste
        img = ImageGrab.grab(bbox=(x0, y0, x1, y1), all_screens=True)
        img = img.crop((arg("--esq"), arg("--topo"), img.width - arg("--dir"), img.height - arg("--base")))
        img.save(saida)
        print("ok", saida, img.size, "| janela:", t.encode("ascii", "replace").decode())
    elif cmd == "regiao":
        x0, y0, x1, y1 = [int(v) for v in sys.argv[2].split(",")]
        img = ImageGrab.grab(bbox=(x0, y0, x1, y1), all_screens=True)
        img.save(sys.argv[3])
        print("ok", sys.argv[3], img.size)
    elif cmd == "tela":
        img = ImageGrab.grab(all_screens=True)
        img.save(sys.argv[2])
        print("ok", sys.argv[2], img.size)


if __name__ == "__main__":
    main()

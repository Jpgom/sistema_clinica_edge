from __future__ import annotations

import copy
import json
import os
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageDraw, ImageFont, ImageTk
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "posicionamento.json"
PREVIEW_BG = BASE_DIR / "modelo_preview.png"

try:
    import win32print
    import win32ui
    import win32con
    WINDOWS_PRINT_AVAILABLE = os.name == "nt"
except ImportError:
    WINDOWS_PRINT_AVAILABLE = False


CAMPOS_CALIBRAVEIS = [
    ("cliente", "Nome do cliente"),
    ("data", "Data"),
    ("quantidade", "Quantidade"),
    ("unidade", "Unidade"),
    ("descricao", "Descrição"),
    ("preco_unitario", "Preço unitário"),
    ("total_item", "Total do item"),
    ("total_geral", "Total geral"),
]

PADRAO_CAMPO_OFFSETS = {
    "cliente": {"x": 0.0, "y": -2.0},
    "data": {"x": 0.0, "y": 0.0},
    "quantidade": {"x": -4.0, "y": 1.0},
    "unidade": {"x": -3.3, "y": 1.0},
    "descricao": {"x": 0.0, "y": 0.8},
    "preco_unitario": {"x": -2.0, "y": 1.0},
    "total_item": {"x": -2.0, "y": 1.0},
    "total_geral": {"x": -2.0, "y": 2.0},
}


@dataclass
class Item:
    quantidade: Decimal
    unidade: str
    descricao: str
    preco_unitario: Decimal

    @property
    def total(self) -> Decimal:
        return (self.quantidade * self.preco_unitario).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def decimal_br(texto: str, default: Decimal = Decimal("0")) -> Decimal:
    texto = (texto or "").strip().replace("R$", "").replace(" ", "")
    if not texto:
        return default
    # Aceita 1.234,56 e 1234.56
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return Decimal(texto)
    except InvalidOperation:
        return default


def float_br(texto: str, default: float = 0.0) -> float:
    texto = str(texto or "").strip().replace(",", ".")
    if not texto:
        return default
    return float(texto)


def moeda_br(valor: Decimal) -> str:
    valor = valor.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    s = f"{valor:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def qtd_br(valor: Decimal) -> str:
    if valor == valor.to_integral():
        return str(int(valor))
    s = format(valor.normalize(), "f")
    return s.replace(".", ",")


def carregar_config() -> dict:
    with CONFIG_PATH.open("r", encoding="utf-8") as f:
        return json.load(f)


def salvar_config(config: dict) -> None:
    with CONFIG_PATH.open("w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def normalizar_config(config: dict) -> dict:
    """Atualiza configurações antigas sem apagar calibrações já feitas."""
    versao = int(config.get("config_version", 1))

    # Ajuste solicitado no teste físico: cliente mais à direita e um pouco acima.
    if versao < 2:
        config.setdefault("cliente", {})
        config["cliente"]["x_mm"] = 24.0
        config["cliente"]["y_mm"] = 39.2

    config.setdefault("offset_x_mm", 0.0)
    config.setdefault("offset_y_mm", 0.0)
    config.setdefault("fonte", "Arial")
    config.setdefault("fonte_tamanho_pt", 8.0)

    # Correção exclusiva da impressão direta. O primeiro teste físico recebido
    # mostrou um deslocamento forte para a esquerda e pequeno para baixo.
    # Estes valores não alteram a prévia nem o PDF gerado.
    if versao < 3:
        config.setdefault("impressao_direta_offset_x_mm", 35.0)
        config.setdefault("impressao_direta_offset_y_mm", -2.0)
    else:
        config.setdefault("impressao_direta_offset_x_mm", 0.0)
        config.setdefault("impressao_direta_offset_y_mm", 0.0)

    offsets = config.setdefault("campo_offsets_mm", {})
    # Calibração padrão definida após o teste físico do usuário.
    # Ao atualizar uma configuração v3 para v4, estes valores passam a ser
    # a nova base. Depois disso, novas calibrações manuais são preservadas.
    if versao < 4:
        for chave, valores in PADRAO_CAMPO_OFFSETS.items():
            offsets[chave] = {"x": valores["x"], "y": valores["y"]}

    for chave, _ in CAMPOS_CALIBRAVEIS:
        campo = offsets.setdefault(chave, {})
        padrao = PADRAO_CAMPO_OFFSETS.get(chave, {"x": 0.0, "y": 0.0})
        campo.setdefault("x", padrao["x"])
        campo.setdefault("y", padrao["y"])

    linhas = config.setdefault("linha_offsets_mm", [])
    while len(linhas) < 10:
        linhas.append({"x": 0.0, "y": 0.0})
    if len(linhas) > 10:
        del linhas[10:]
    for linha in linhas:
        linha.setdefault("x", 0.0)
        linha.setdefault("y", 0.0)

    # Ajuste por célula: cada campo de cada produto pode ser movido sozinho.
    celulas = config.setdefault("celula_offsets_mm", [])
    chaves_celula = ["quantidade", "unidade", "descricao", "preco_unitario", "total_item"]
    while len(celulas) < 10:
        celulas.append({})
    if len(celulas) > 10:
        del celulas[10:]
    for linha in celulas:
        for chave in chaves_celula:
            campo = linha.setdefault(chave, {})
            campo.setdefault("x", 0.0)
            campo.setdefault("y", 0.0)

    config["config_version"] = 4
    return config


class NotaFiscalApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Preenchedor de Nota de Balcão - v4")
        self.geometry("1180x760")
        self.minsize(1040, 680)
        self.configuracao = normalizar_config(carregar_config())
        salvar_config(self.configuracao)
        self.preview_photo = None
        self._montar_tela()
        self._atualizar_totais()
        self.after(200, self._atualizar_previa)

    def _montar_tela(self):
        topo = ttk.Frame(self, padding=10)
        topo.pack(fill="x")

        ttk.Label(topo, text="Cliente / O Sr.:").grid(row=0, column=0, sticky="w")
        self.var_cliente = tk.StringVar()
        ttk.Entry(topo, textvariable=self.var_cliente, width=56).grid(row=0, column=1, padx=(6, 18), sticky="ew")

        ttk.Label(topo, text="Data:").grid(row=0, column=2, sticky="w")
        self.var_data = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        ttk.Entry(topo, textvariable=self.var_data, width=14).grid(row=0, column=3, padx=(6, 0))
        topo.columnconfigure(1, weight=1)

        corpo = ttk.Frame(self, padding=(10, 0, 10, 8))
        corpo.pack(fill="both", expand=True)

        esquerda = ttk.Frame(corpo)
        esquerda.pack(side="left", fill="both", expand=True)

        cabecalhos = ["Quant.", "Unid.", "Discriminação das mercadorias", "P. Unitário", "Total"]
        larguras = [9, 8, 46, 14, 14]
        for col, (texto, largura) in enumerate(zip(cabecalhos, larguras)):
            ttk.Label(esquerda, text=texto, anchor="center").grid(row=0, column=col, padx=2, pady=(0, 4), sticky="ew")
            esquerda.columnconfigure(col, weight=1 if col == 2 else 0)

        self.linhas = []
        for i in range(10):
            vq = tk.StringVar()
            vu = tk.StringVar(value="UN")
            vd = tk.StringVar()
            vp = tk.StringVar()
            vt = tk.StringVar(value="0,00")

            vars_linha = [vq, vu, vd, vp, vt]
            self.linhas.append(vars_linha)

            ttk.Entry(esquerda, textvariable=vq, width=larguras[0], justify="right").grid(row=i + 1, column=0, padx=2, pady=2)
            ttk.Entry(esquerda, textvariable=vu, width=larguras[1], justify="center").grid(row=i + 1, column=1, padx=2, pady=2)
            ttk.Entry(esquerda, textvariable=vd, width=larguras[2]).grid(row=i + 1, column=2, padx=2, pady=2, sticky="ew")
            ttk.Entry(esquerda, textvariable=vp, width=larguras[3], justify="right").grid(row=i + 1, column=3, padx=2, pady=2)
            ent_total = ttk.Entry(esquerda, textvariable=vt, width=larguras[4], justify="right", state="readonly")
            ent_total.grid(row=i + 1, column=4, padx=2, pady=2)

            vq.trace_add("write", lambda *_: self._atualizar_totais())
            vp.trace_add("write", lambda *_: self._atualizar_totais())

        rodape = ttk.Frame(esquerda, padding=(0, 12, 0, 0))
        rodape.grid(row=12, column=0, columnspan=5, sticky="ew")
        ttk.Label(rodape, text="TOTAL R$", font=("TkDefaultFont", 11, "bold")).pack(side="left")
        self.var_total_geral = tk.StringVar(value="0,00")
        ttk.Label(rodape, textvariable=self.var_total_geral, font=("TkDefaultFont", 14, "bold")).pack(side="left", padx=10)

        botoes = ttk.Frame(esquerda, padding=(0, 14, 0, 0))
        botoes.grid(row=13, column=0, columnspan=5, sticky="w")
        ttk.Button(botoes, text="Atualizar prévia", command=self._atualizar_previa).pack(side="left", padx=(0, 8))
        ttk.Button(botoes, text="Gerar PDF para impressão", command=self._gerar_pdf_dialog).pack(side="left", padx=8)
        ttk.Button(botoes, text="Imprimir direto", command=self._imprimir_direto).pack(side="left", padx=8)
        ttk.Button(botoes, text="Calibração avançada", command=self._abrir_calibracao).pack(side="left", padx=8)
        ttk.Button(botoes, text="Limpar", command=self._limpar).pack(side="left", padx=8)

        dica = (
            "O layout da nota NÃO é impresso. A imagem à direita é apenas uma prévia. "
            "Na calibração avançada você pode mover tudo, cada campo/coluna e também cada linha de produto separadamente."
        )
        ttk.Label(esquerda, text=dica, wraplength=690, foreground="#555").grid(row=14, column=0, columnspan=5, sticky="w", pady=(16, 0))

        direita = ttk.LabelFrame(corpo, text="Prévia de alinhamento", padding=8)
        direita.pack(side="right", fill="both", padx=(12, 0))
        self.preview_label = ttk.Label(direita)
        self.preview_label.pack(fill="both", expand=True)

    def _coletar_itens(self) -> list[Item]:
        itens: list[Item] = []
        for vq, vu, vd, vp, _ in self.linhas:
            q = decimal_br(vq.get())
            p = decimal_br(vp.get())
            desc = vd.get().strip()
            un = vu.get().strip()
            if q != 0 or p != 0 or desc or (un and un.upper() != "UN"):
                itens.append(Item(q, un, desc, p))
            else:
                itens.append(Item(Decimal("0"), un, "", Decimal("0")))
        return itens

    def _atualizar_totais(self):
        total_geral = Decimal("0")
        for vars_linha in self.linhas:
            vq, _, _, vp, vt = vars_linha
            total = (decimal_br(vq.get()) * decimal_br(vp.get())).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            vt.set(moeda_br(total))
            total_geral += total
        self.var_total_geral.set(moeda_br(total_geral))

    def _offset_campo(self, chave: str) -> tuple[float, float]:
        obj = self.configuracao.get("campo_offsets_mm", {}).get(chave, {})
        return float(obj.get("x", 0.0)), float(obj.get("y", 0.0))

    def _offset_linha(self, indice: int) -> tuple[float, float]:
        linhas = self.configuracao.get("linha_offsets_mm", [])
        if 0 <= indice < len(linhas):
            return float(linhas[indice].get("x", 0.0)), float(linhas[indice].get("y", 0.0))
        return 0.0, 0.0

    def _offset_celula(self, indice: int, chave: str) -> tuple[float, float]:
        linhas = self.configuracao.get("celula_offsets_mm", [])
        if 0 <= indice < len(linhas):
            obj = linhas[indice].get(chave, {})
            return float(obj.get("x", 0.0)), float(obj.get("y", 0.0))
        return 0.0, 0.0

    def _xy(self, chave: str, base_x: float, base_y: float, linha: int | None = None) -> tuple[float, float]:
        """Combina posição base + geral + campo + linha + célula individual."""
        x = float(base_x) + float(self.configuracao.get("offset_x_mm", 0.0))
        y = float(base_y) + float(self.configuracao.get("offset_y_mm", 0.0))
        cx, cy = self._offset_campo(chave)
        x += cx
        y += cy
        if linha is not None:
            lx, ly = self._offset_linha(linha)
            x += lx
            y += ly
            ex, ey = self._offset_celula(linha, chave)
            x += ex
            y += ey
        return x, y

    def _campos_impressao(self):
        cfg = self.configuracao
        campos = []

        cliente = self.var_cliente.get().strip()
        if cliente:
            x, y = self._xy("cliente", cfg["cliente"]["x_mm"], cfg["cliente"]["y_mm"])
            campos.append(("left", x, y, cliente))

        data_txt = self.var_data.get().strip()
        if data_txt:
            x, y = self._xy("data", cfg["data"]["x_centro_mm"], cfg["data"]["y_mm"])
            campos.append(("center", x, y, data_txt))

        itens = self._coletar_itens()
        ys = cfg["linhas_y_mm"]
        for i, item in enumerate(itens[:len(ys)]):
            if item.quantidade == 0 and item.preco_unitario == 0 and not item.descricao:
                continue

            base_y = float(ys[i])
            if item.quantidade != 0:
                x, y = self._xy("quantidade", cfg["quantidade"]["x_direita_mm"], base_y, i)
                campos.append(("right", x, y, qtd_br(item.quantidade)))
            if item.unidade:
                x, y = self._xy("unidade", cfg["unidade"]["x_centro_mm"], base_y, i)
                campos.append(("center", x, y, item.unidade))
            if item.descricao:
                x, y = self._xy("descricao", cfg["descricao"]["x_mm"], base_y, i)
                campos.append(("left", x, y, item.descricao[:48]))
            if item.preco_unitario != 0:
                x, y = self._xy("preco_unitario", cfg["preco_unitario"]["x_direita_mm"], base_y, i)
                campos.append(("right", x, y, moeda_br(item.preco_unitario)))
                x, y = self._xy("total_item", cfg["total_item"]["x_direita_mm"], base_y, i)
                campos.append(("right", x, y, moeda_br(item.total)))

        total = sum((i.total for i in itens), Decimal("0"))
        x, y = self._xy("total_geral", cfg["total_geral"]["x_direita_mm"], cfg["total_geral"]["y_mm"])
        campos.append(("right", x, y, moeda_br(total)))
        return campos

    def _atualizar_previa(self):
        try:
            img = Image.open(PREVIEW_BG).convert("RGB")
            draw = ImageDraw.Draw(img)
            sx = img.width / float(self.configuracao["papel_largura_mm"])
            sy = img.height / float(self.configuracao["papel_altura_mm"])
            tamanho_px = max(10, int(float(self.configuracao.get("fonte_tamanho_pt", 8.0)) * 1.9))
            try:
                font = ImageFont.truetype("arial.ttf", tamanho_px)
            except Exception:
                try:
                    font = ImageFont.truetype("DejaVuSans.ttf", tamanho_px)
                except Exception:
                    font = ImageFont.load_default()

            for alinhamento, x_mm, y_mm, texto in self._campos_impressao():
                x = x_mm * sx
                y = y_mm * sy
                box = draw.textbbox((0, 0), texto, font=font)
                w = box[2] - box[0]
                h = box[3] - box[1]
                if alinhamento == "center":
                    x -= w / 2
                elif alinhamento == "right":
                    x -= w
                # Aproxima a coordenada configurada à linha-base, como ocorre no PDF/GDI.
                draw.text((x, y - h), texto, fill="black", font=font)

            max_w = 470
            ratio = min(1.0, max_w / img.width)
            resized = img.resize((int(img.width * ratio), int(img.height * ratio)))
            self.preview_photo = ImageTk.PhotoImage(resized)
            self.preview_label.configure(image=self.preview_photo)
        except Exception as exc:
            messagebox.showerror("Prévia", f"Não foi possível gerar a prévia:\n{exc}")

    def _gerar_pdf(self, destino: str):
        cfg = self.configuracao
        largura = float(cfg["papel_largura_mm"]) * mm
        altura = float(cfg["papel_altura_mm"]) * mm
        c = canvas.Canvas(destino, pagesize=(largura, altura))
        tamanho = float(cfg.get("fonte_tamanho_pt", 8.0))
        c.setFont("Helvetica", tamanho)

        for alinhamento, x_mm, y_mm, texto in self._campos_impressao():
            x = x_mm * mm
            # Config usa coordenadas a partir do topo; PDF usa a partir de baixo.
            y = altura - (y_mm * mm)
            if alinhamento == "right":
                c.drawRightString(x, y, texto)
            elif alinhamento == "center":
                c.drawCentredString(x, y, texto)
            else:
                c.drawString(x, y, texto)
        c.showPage()
        c.save()

    def _gerar_pdf_dialog(self):
        self._atualizar_totais()
        destino = filedialog.asksaveasfilename(
            title="Salvar PDF de impressão",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile="nota_preenchida.pdf",
        )
        if not destino:
            return
        try:
            self._gerar_pdf(destino)
            messagebox.showinfo(
                "PDF gerado",
                "PDF gerado com somente os dados.\n\nNa impressão, use escala 100% / Tamanho real e papel 144 x 105,4 mm.",
            )
        except Exception as exc:
            messagebox.showerror("Erro", f"Falha ao gerar PDF:\n{exc}")

    def _imprimir_direto(self):
        if not WINDOWS_PRINT_AVAILABLE:
            messagebox.showwarning(
                "Impressão direta",
                "A impressão direta está disponível no Windows com o pacote pywin32.\nUse 'Gerar PDF para impressão' neste computador.",
            )
            return

        try:
            printer_name = win32print.GetDefaultPrinter()
            dc = win32ui.CreateDC()
            dc.CreatePrinterDC(printer_name)

            dpi_x = dc.GetDeviceCaps(win32con.LOGPIXELSX)
            dpi_y = dc.GetDeviceCaps(win32con.LOGPIXELSY)
            off_x = dc.GetDeviceCaps(win32con.PHYSICALOFFSETX)
            off_y = dc.GetDeviceCaps(win32con.PHYSICALOFFSETY)

            tamanho_pt = float(self.configuracao.get("fonte_tamanho_pt", 8.0))
            font = win32ui.CreateFont({
                "name": self.configuracao.get("fonte", "Arial"),
                "height": -int(tamanho_pt * dpi_y / 72.0),
                "weight": 400,
            })
            dc.SelectObject(font)
            dc.SetBkMode(win32con.TRANSPARENT)

            # Ajuste separado da impressora: serve para corrigir diferença entre
            # a coordenada lógica do driver e a posição real no formulário.
            driver_dx = float(self.configuracao.get("impressao_direta_offset_x_mm", 0.0))
            driver_dy = float(self.configuracao.get("impressao_direta_offset_y_mm", 0.0))

            def pos_px(x_mm, y_mm):
                x = int((x_mm + driver_dx) * dpi_x / 25.4) - off_x
                y = int((y_mm + driver_dy) * dpi_y / 25.4) - off_y
                return x, y

            dc.StartDoc("Nota de Balcão - preenchimento")
            dc.StartPage()
            for alinhamento, x_mm, y_mm, texto in self._campos_impressao():
                x, y = pos_px(x_mm, y_mm)
                w, h = dc.GetTextExtent(texto)
                if alinhamento == "right":
                    x -= w
                elif alinhamento == "center":
                    x -= w // 2
                # y configurado como linha-base aproximada; GDI usa topo do texto.
                y -= int(h * 0.78)
                dc.TextOut(x, y, texto)
            dc.EndPage()
            dc.EndDoc()
            dc.DeleteDC()
            messagebox.showinfo(
                "Enviado à impressora",
                f"Nota enviada para:\n{printer_name}\n\nSe houver deslocamento, ajuste em 'Calibração avançada'.",
            )
        except Exception as exc:
            messagebox.showerror(
                "Erro de impressão",
                "Não foi possível imprimir diretamente. Verifique se o tamanho de papel personalizado está configurado na impressora.\n\n"
                f"Detalhes: {exc}",
            )

    def _abrir_calibracao(self):
        janela = tk.Toplevel(self)
        janela.title("Calibração avançada de alinhamento")
        janela.geometry("760x610")
        janela.minsize(700, 560)
        janela.transient(self)

        original = copy.deepcopy(self.configuracao)

        externo = ttk.Frame(janela, padding=12)
        externo.pack(fill="both", expand=True)

        ttk.Label(
            externo,
            text=(
                "Ajustes em milímetros. X positivo move para a direita; X negativo para a esquerda. "
                "Y positivo move para baixo; Y negativo para cima."
            ),
            wraplength=720,
        ).pack(fill="x", pady=(0, 8))

        notebook = ttk.Notebook(externo)
        notebook.pack(fill="both", expand=True)

        # ----- Aba Geral -----
        aba_geral = ttk.Frame(notebook, padding=14)
        notebook.add(aba_geral, text="Geral")

        vx = tk.StringVar(value=str(self.configuracao.get("offset_x_mm", 0.0)).replace(".", ","))
        vy = tk.StringVar(value=str(self.configuracao.get("offset_y_mm", 0.0)).replace(".", ","))
        vf = tk.StringVar(value=str(self.configuracao.get("fonte_tamanho_pt", 8.0)).replace(".", ","))
        vdx = tk.StringVar(value=str(self.configuracao.get("impressao_direta_offset_x_mm", 0.0)).replace(".", ","))
        vdy = tk.StringVar(value=str(self.configuracao.get("impressao_direta_offset_y_mm", 0.0)).replace(".", ","))

        ttk.Label(aba_geral, text="Mover toda a impressão em X (mm):").grid(row=0, column=0, sticky="w", pady=7)
        ttk.Entry(aba_geral, textvariable=vx, width=12, justify="right").grid(row=0, column=1, padx=8)
        ttk.Label(aba_geral, text="Mover toda a impressão em Y (mm):").grid(row=1, column=0, sticky="w", pady=7)
        ttk.Entry(aba_geral, textvariable=vy, width=12, justify="right").grid(row=1, column=1, padx=8)
        ttk.Label(aba_geral, text="Tamanho da fonte (pt):").grid(row=2, column=0, sticky="w", pady=7)
        ttk.Entry(aba_geral, textvariable=vf, width=12, justify="right").grid(row=2, column=1, padx=8)

        ttk.Separator(aba_geral, orient="horizontal").grid(row=3, column=0, columnspan=2, sticky="ew", pady=14)
        ttk.Label(aba_geral, text="Correção da impressora direta", font=("TkDefaultFont", 10, "bold")).grid(row=4, column=0, columnspan=2, sticky="w")
        ttk.Label(aba_geral, text="Correção X somente na impressão (mm):").grid(row=5, column=0, sticky="w", pady=7)
        ttk.Entry(aba_geral, textvariable=vdx, width=12, justify="right").grid(row=5, column=1, padx=8)
        ttk.Label(aba_geral, text="Correção Y somente na impressão (mm):").grid(row=6, column=0, sticky="w", pady=7)
        ttk.Entry(aba_geral, textvariable=vdy, width=12, justify="right").grid(row=6, column=1, padx=8)
        ttk.Label(
            aba_geral,
            text=(
                "Essa correção existe porque o driver da impressora pode deslocar o papel real. "
                "Ela NÃO muda a prévia nem o PDF. No teste físico recebido, deixei inicialmente X = +35 mm e Y = -2 mm."
            ),
            wraplength=620,
            foreground="#555",
        ).grid(row=7, column=0, columnspan=2, sticky="w", pady=(10, 0))

        ttk.Label(
            aba_geral,
            text=(
                "O ajuste geral acima muda prévia, PDF e impressão. Use a correção da impressora quando a prévia estiver certa "
                "mas o papel físico sair deslocado."
            ),
            wraplength=620,
            foreground="#555",
        ).grid(row=8, column=0, columnspan=2, sticky="w", pady=(10, 0))

        # ----- Aba Campos -----
        aba_campos = ttk.Frame(notebook, padding=12)
        notebook.add(aba_campos, text="Cada campo")

        ttk.Label(aba_campos, text="Campo", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=0, sticky="w", padx=4, pady=(0, 6))
        ttk.Label(aba_campos, text="X (mm)", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=1, padx=4, pady=(0, 6))
        ttk.Label(aba_campos, text="Y (mm)", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=2, padx=4, pady=(0, 6))

        vars_campos: dict[str, tuple[tk.StringVar, tk.StringVar]] = {}
        for idx, (chave, rotulo) in enumerate(CAMPOS_CALIBRAVEIS, start=1):
            atual = self.configuracao.get("campo_offsets_mm", {}).get(chave, {"x": 0.0, "y": 0.0})
            vcx = tk.StringVar(value=str(atual.get("x", 0.0)).replace(".", ","))
            vcy = tk.StringVar(value=str(atual.get("y", 0.0)).replace(".", ","))
            vars_campos[chave] = (vcx, vcy)
            ttk.Label(aba_campos, text=rotulo).grid(row=idx, column=0, sticky="w", padx=4, pady=4)
            ttk.Entry(aba_campos, textvariable=vcx, width=12, justify="right").grid(row=idx, column=1, padx=4, pady=4)
            ttk.Entry(aba_campos, textvariable=vcy, width=12, justify="right").grid(row=idx, column=2, padx=4, pady=4)

        ttk.Label(
            aba_campos,
            text=(
                "Exemplo: se apenas o NOME estiver 2 mm à esquerda, coloque X = +2,0 somente em 'Nome do cliente'. "
                "Esse ajuste não altera os demais campos."
            ),
            wraplength=650,
            foreground="#555",
        ).grid(row=len(CAMPOS_CALIBRAVEIS) + 1, column=0, columnspan=3, sticky="w", pady=(12, 0))

        # ----- Aba Linhas -----
        aba_linhas = ttk.Frame(notebook, padding=12)
        notebook.add(aba_linhas, text="Cada linha de produto")

        ttk.Label(aba_linhas, text="Linha", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=0, sticky="w", padx=4, pady=(0, 6))
        ttk.Label(aba_linhas, text="X (mm)", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=1, padx=4, pady=(0, 6))
        ttk.Label(aba_linhas, text="Y (mm)", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=2, padx=4, pady=(0, 6))

        vars_linhas: list[tuple[tk.StringVar, tk.StringVar]] = []
        for i in range(10):
            atual = self.configuracao.get("linha_offsets_mm", [])[i]
            vlx = tk.StringVar(value=str(atual.get("x", 0.0)).replace(".", ","))
            vly = tk.StringVar(value=str(atual.get("y", 0.0)).replace(".", ","))
            vars_linhas.append((vlx, vly))
            ttk.Label(aba_linhas, text=f"Produto / linha {i + 1}").grid(row=i + 1, column=0, sticky="w", padx=4, pady=3)
            ttk.Entry(aba_linhas, textvariable=vlx, width=12, justify="right").grid(row=i + 1, column=1, padx=4, pady=3)
            ttk.Entry(aba_linhas, textvariable=vly, width=12, justify="right").grid(row=i + 1, column=2, padx=4, pady=3)

        ttk.Label(
            aba_linhas,
            text=(
                "O ajuste de uma linha move juntos Quantidade, Unidade, Descrição, Preço Unitário e Total daquela linha. "
                "Assim você pode corrigir uma linha sem mexer nas outras."
            ),
            wraplength=650,
            foreground="#555",
        ).grid(row=11, column=0, columnspan=3, sticky="w", pady=(10, 0))

        # ----- Aba Cada item/célula -----
        aba_celulas = ttk.Frame(notebook, padding=12)
        notebook.add(aba_celulas, text="Cada item individual")

        ttk.Label(aba_celulas, text="Produto / linha:").grid(row=0, column=0, sticky="w", padx=4, pady=(0, 10))
        var_linha_celula = tk.StringVar(value="1")
        seletor_linha = ttk.Combobox(
            aba_celulas, textvariable=var_linha_celula, values=[str(i) for i in range(1, 11)], width=8, state="readonly"
        )
        seletor_linha.grid(row=0, column=1, sticky="w", padx=4, pady=(0, 10))

        ttk.Label(aba_celulas, text="Campo", font=("TkDefaultFont", 9, "bold")).grid(row=1, column=0, sticky="w", padx=4)
        ttk.Label(aba_celulas, text="X (mm)", font=("TkDefaultFont", 9, "bold")).grid(row=1, column=1, padx=4)
        ttk.Label(aba_celulas, text="Y (mm)", font=("TkDefaultFont", 9, "bold")).grid(row=1, column=2, padx=4)

        campos_celula = [
            ("quantidade", "Quantidade"),
            ("unidade", "Unidade"),
            ("descricao", "Descrição"),
            ("preco_unitario", "Preço unitário"),
            ("total_item", "Total do item"),
        ]
        vars_celulas: list[dict[str, tuple[tk.StringVar, tk.StringVar]]] = []
        cfg_celulas = self.configuracao.get("celula_offsets_mm", [])
        for i in range(10):
            mapa = {}
            for chave, _ in campos_celula:
                atual = cfg_celulas[i].get(chave, {"x": 0.0, "y": 0.0})
                mapa[chave] = (
                    tk.StringVar(value=str(atual.get("x", 0.0)).replace(".", ",")),
                    tk.StringVar(value=str(atual.get("y", 0.0)).replace(".", ",")),
                )
            vars_celulas.append(mapa)

        entradas_celula = {}
        for idx, (chave, rotulo) in enumerate(campos_celula, start=2):
            ttk.Label(aba_celulas, text=rotulo).grid(row=idx, column=0, sticky="w", padx=4, pady=5)
            ex = ttk.Entry(aba_celulas, width=12, justify="right")
            ey = ttk.Entry(aba_celulas, width=12, justify="right")
            ex.grid(row=idx, column=1, padx=4, pady=5)
            ey.grid(row=idx, column=2, padx=4, pady=5)
            entradas_celula[chave] = (ex, ey)

        def mostrar_linha_celula(*_):
            i = max(0, min(9, int(var_linha_celula.get() or "1") - 1))
            for chave, _ in campos_celula:
                ex, ey = entradas_celula[chave]
                vx_c, vy_c = vars_celulas[i][chave]
                ex.configure(textvariable=vx_c)
                ey.configure(textvariable=vy_c)

        seletor_linha.bind("<<ComboboxSelected>>", mostrar_linha_celula)
        mostrar_linha_celula()

        ttk.Label(
            aba_celulas,
            text=(
                "Aqui cada célula é independente. Exemplo: você pode mover somente o PREÇO UNITÁRIO da linha 4, "
                "sem alterar quantidade, unidade, descrição, total ou qualquer outra linha."
            ),
            wraplength=650,
            foreground="#555",
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(14, 0))

        def ler_tela_para_config():
            self.configuracao["offset_x_mm"] = float_br(vx.get())
            self.configuracao["offset_y_mm"] = float_br(vy.get())
            self.configuracao["impressao_direta_offset_x_mm"] = float_br(vdx.get())
            self.configuracao["impressao_direta_offset_y_mm"] = float_br(vdy.get())
            fonte = float_br(vf.get(), 8.0)
            if fonte <= 0:
                raise ValueError("O tamanho da fonte deve ser maior que zero.")
            self.configuracao["fonte_tamanho_pt"] = fonte

            self.configuracao.setdefault("campo_offsets_mm", {})
            for chave, (vcx, vcy) in vars_campos.items():
                self.configuracao["campo_offsets_mm"][chave] = {
                    "x": float_br(vcx.get()),
                    "y": float_br(vcy.get()),
                }

            self.configuracao["linha_offsets_mm"] = []
            for vlx, vly in vars_linhas:
                self.configuracao["linha_offsets_mm"].append({
                    "x": float_br(vlx.get()),
                    "y": float_br(vly.get()),
                })

            self.configuracao["celula_offsets_mm"] = []
            for mapa in vars_celulas:
                linha_cfg = {}
                for chave, (vcx, vcy) in mapa.items():
                    linha_cfg[chave] = {
                        "x": float_br(vcx.get()),
                        "y": float_br(vcy.get()),
                    }
                self.configuracao["celula_offsets_mm"].append(linha_cfg)

        def aplicar_previa():
            try:
                ler_tela_para_config()
                self._atualizar_previa()
            except Exception as exc:
                messagebox.showerror("Calibração", f"Verifique os valores informados.\n\n{exc}", parent=janela)

        def salvar():
            try:
                ler_tela_para_config()
                salvar_config(self.configuracao)
                self._atualizar_previa()
                janela.destroy()
            except Exception as exc:
                messagebox.showerror("Calibração", f"Verifique os valores informados.\n\n{exc}", parent=janela)

        def zerar_individuais():
            for vcx, vcy in vars_campos.values():
                vcx.set("0")
                vcy.set("0")
            for vlx, vly in vars_linhas:
                vlx.set("0")
                vly.set("0")
            for mapa in vars_celulas:
                for vcx, vcy in mapa.values():
                    vcx.set("0")
                    vcy.set("0")

        def cancelar():
            self.configuracao = original
            self._atualizar_previa()
            janela.destroy()

        barra = ttk.Frame(externo, padding=(0, 10, 0, 0))
        barra.pack(fill="x")
        ttk.Button(barra, text="Aplicar na prévia", command=aplicar_previa).pack(side="left")
        ttk.Button(barra, text="Zerar ajustes individuais", command=zerar_individuais).pack(side="left", padx=8)
        ttk.Button(barra, text="Cancelar", command=cancelar).pack(side="right")
        ttk.Button(barra, text="Salvar calibração", command=salvar).pack(side="right", padx=8)

        janela.protocol("WM_DELETE_WINDOW", cancelar)

    def _limpar(self):
        self.var_cliente.set("")
        self.var_data.set(date.today().strftime("%d/%m/%Y"))
        for vq, vu, vd, vp, _ in self.linhas:
            vq.set("")
            vu.set("UN")
            vd.set("")
            vp.set("")
        self._atualizar_totais()
        self._atualizar_previa()


if __name__ == "__main__":
    app = NotaFiscalApp()
    app.mainloop()

import os
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser, font, ttk

import fitz  # PyMuPDF
from PIL import Image, ImageTk, ImageOps
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer
from xml.sax.saxutils import escape


class DaviaPDFMaker:
    def __init__(self, root):
        self.root = root
        self.root.title("davIA PDF Maker")
        self.root.geometry("1000x900")
        self.root.minsize(850, 600)
        self.root.configure(bg="#142D50")
        self.criar_interface()

    def criar_interface(self):
        # A imagem ocupa toda a janela; os controles ficam sobre ela.
        self.caminho_banner = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "banner_davia.png"
        )
        self.imagem_fundo_original = None
        self.imagem_fundo_tk = None

        try:
            self.imagem_fundo_original = Image.open(
                self.caminho_banner
            ).convert("RGB")
        except (FileNotFoundError, OSError):
            self.imagem_fundo_original = None

        self.canvas = tk.Canvas(
            self.root,
            highlightthickness=0,
            bd=0,
            bg="#142D50"
        )
        self.canvas.pack(fill="both", expand=True)

        # Título e subtítulo desenhados diretamente sobre o fundo.
        self.titulo = self.canvas.create_text(
            500, 55,
            text="davIA PDF Maker",
            font=("Arial", 30, "bold"),
            fill="white",
            anchor="center"
        )
        self.subtitulo = self.canvas.create_text(
            500, 105,
            text="Gerenciador Inteligente de Documentos PDF",
            font=("Arial", 12, "bold"),
            fill="#D9ECFF",
            anchor="center"
        )

        botoes = [
            ("Converter imagens em PDF", self.imagens_para_pdf),
            ("Juntar arquivos PDF", self.juntar_pdfs),
            ("Dividir arquivo PDF", self.dividir_pdf),
            ("Extrair imagens do PDF", self.extrair_imagens),
            ("Abrir arquivo PDF", self.abrir_pdf),
            ("Converter TXT/DOCX para PDF", self.converter_texto_para_pdf),
            ("Digitar texto e criar PDF", self.criar_pdf_texto_digitado),
        ]

        self.janelas_botoes = []
        for texto, comando in botoes:
            botao = tk.Button(
                self.root,
                text=texto,
                command=comando,
                font=("Arial", 12, "bold"),
                bg="#173B68",
                fg="white",
                activebackground="#238BE6",
                activeforeground="white",
                relief="flat",
                width=34,
                height=2,
                cursor="hand2",
                bd=0
            )
            janela_botao = self.canvas.create_window(
                500, 200,
                window=botao,
                anchor="center"
            )
            self.janelas_botoes.append(janela_botao)

        self.rodape = self.canvas.create_text(
            500, 870,
            text="DAVIA | Tecnologia e produtividade",
            font=("Arial", 9, "bold"),
            fill="#D9ECFF",
            anchor="center"
        )

        self.canvas.bind("<Configure>", self.redimensionar_interface)
        self.root.after(100, self.redimensionar_interface)

    def redimensionar_interface(self, evento=None):
        largura = max(self.canvas.winfo_width(), 1)
        altura = max(self.canvas.winfo_height(), 1)

        # Redimensiona a imagem para preencher toda a janela sem deformar.
        if self.imagem_fundo_original is not None:
            fundo = ImageOps.fit(
                self.imagem_fundo_original,
                (largura, altura),
                method=Image.Resampling.LANCZOS
            )
            self.imagem_fundo_tk = ImageTk.PhotoImage(fundo)
            if hasattr(self, "item_fundo"):
                self.canvas.itemconfigure(
                    self.item_fundo,
                    image=self.imagem_fundo_tk
                )
            else:
                self.item_fundo = self.canvas.create_image(
                    0, 0,
                    image=self.imagem_fundo_tk,
                    anchor="nw"
                )
                self.canvas.tag_lower(self.item_fundo)

        self.canvas.coords(self.titulo, largura / 2, altura * 0.09)
        self.canvas.coords(self.subtitulo, largura / 2, altura * 0.145)

        # Distribui os sete botões verticalmente na região central.
        inicio = altura * 0.27
        espaco = min(altura * 0.095, 78)
        for indice, item in enumerate(self.janelas_botoes):
            self.canvas.coords(
                item,
                largura / 2,
                inicio + indice * espaco
            )

        self.canvas.coords(
            self.rodape,
            largura / 2,
            altura - 24
        )

    def converter_texto_para_pdf(self):
        arquivo = filedialog.askopenfilename(
            parent=self.root,
            title="Selecionar arquivo de texto",
            filetypes=[
                ("Arquivos de texto", "*.txt"),
                ("Documentos Word", "*.docx"),
                ("Todos os arquivos", "*.*")
            ]
        )

        if not arquivo:
            return

        try:
            extensao = os.path.splitext(arquivo)[1].lower()

            if extensao == ".txt":
                with open(arquivo, "r", encoding="utf-8-sig") as f:
                    texto = f.read()
            elif extensao == ".docx":
                documento_word = Document(arquivo)
                texto = "\n".join(
                    paragrafo.text
                    for paragrafo in documento_word.paragraphs
                )
            else:
                messagebox.showerror(
                    "Formato não suportado",
                    "Selecione um arquivo TXT ou DOCX.",
                    parent=self.root
                )
                return

            destino = filedialog.asksaveasfilename(
                parent=self.root,
                title="Salvar PDF",
                defaultextension=".pdf",
                filetypes=[("Arquivo PDF", "*.pdf")],
                initialfile=os.path.splitext(
                    os.path.basename(arquivo)
                )[0] + ".pdf"
            )

            if not destino:
                return

            criar_pdf_texto(texto, destino)
            messagebox.showinfo(
                "DAVIA PDF Maker",
                "Arquivo convertido para PDF com sucesso!",
                parent=self.root
            )

        except Exception as erro:
            messagebox.showerror(
                "Erro na conversão",
                f"Não foi possível converter o arquivo:\n{erro}",
                parent=self.root
            )

    def criar_pdf_texto_digitado(self):
        """Abre um editor de texto com ferramentas básicas de formatação."""
        janela = tk.Toplevel(self.root)
        janela.title("davIA PDF Maker - Editor de texto")
        janela.geometry("1000x720")
        janela.minsize(720, 520)
        janela.configure(bg="#EAF1F8")

        barra = tk.Frame(janela, bg="#DCE7F2", padx=8, pady=8)
        barra.pack(fill="x")

        campo = tk.Text(
            janela,
            wrap="word",
            font=("Arial", 12),
            undo=True,
            padx=12,
            pady=12,
            spacing1=2,
            spacing3=4,
            selectbackground="#B8D7F5"
        )
        rolagem = ttk.Scrollbar(janela, orient="vertical", command=campo.yview)
        campo.configure(yscrollcommand=rolagem.set)
        rolagem.pack(side="right", fill="y")
        campo.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)

        # Formatos de fonte compatíveis com a interface e com a exportação PDF.
        familias = sorted(set(font.families(janela)))
        preferidas = ["Arial", "Times New Roman", "Courier New", "Verdana", "Calibri"]
        familia_inicial = "Arial" if "Arial" in familias else (familias[0] if familias else "Arial")

        fonte_var = tk.StringVar(value=familia_inicial)
        tamanho_var = tk.StringVar(value="12")
        cor_atual = {"valor": "#000000"}
        ativos = {"bold": False, "italic": False, "underline": False}
        alinhamento_atual = {"valor": "left"}

        def tag_fonte(familia):
            return "family_" + re.sub(r"[^A-Za-z0-9]", "_", familia)

        def aplicar_tag(tag, configuracao, inicio=None, fim=None):
            campo.tag_configure(tag, **configuracao)
            if inicio is None or fim is None:
                try:
                    inicio, fim = campo.index("sel.first"), campo.index("sel.last")
                except tk.TclError:
                    # Sem seleção: a formatação passa a valer para o próximo texto digitado.
                    ativos[tag] = not ativos.get(tag, False)
                    atualizar_status()
                    return
            campo.tag_add(tag, inicio, fim)
            atualizar_status()

        def atualizar_status():
            nomes = []
            if ativos["bold"]:
                nomes.append("Negrito")
            if ativos["italic"]:
                nomes.append("Itálico")
            if ativos["underline"]:
                nomes.append("Sublinhado")
            status_var.set(" | ".join(nomes) if nomes else "Pronto")

        def alternar_formato(chave, tag, configuracao):
            try:
                inicio, fim = campo.index("sel.first"), campo.index("sel.last")
                ja_aplicado = tag in campo.tag_names(inicio)
                campo.tag_configure(tag, **configuracao)
                if ja_aplicado:
                    campo.tag_remove(tag, inicio, fim)
                else:
                    campo.tag_add(tag, inicio, fim)
            except tk.TclError:
                ativos[chave] = not ativos[chave]
            atualizar_status()
            campo.focus_set()

        def escolher_fonte(_evento=None):
            familia = fonte_var.get()
            try:
                inicio, fim = campo.index("sel.first"), campo.index("sel.last")
                tag = tag_fonte(familia)
                campo.tag_configure(tag, font=(familia, int(tamanho_var.get())))
                campo.tag_add(tag, inicio, fim)
            except (tk.TclError, ValueError):
                pass
            campo.focus_set()

        def escolher_tamanho(_evento=None):
            try:
                tamanho = max(6, min(96, int(tamanho_var.get())))
                tamanho_var.set(str(tamanho))
                inicio, fim = campo.index("sel.first"), campo.index("sel.last")
                tag = f"size_{tamanho}"
                campo.tag_configure(tag, font=(fonte_var.get(), tamanho))
                campo.tag_add(tag, inicio, fim)
            except (tk.TclError, ValueError):
                pass
            campo.focus_set()

        def escolher_cor():
            resultado = colorchooser.askcolor(
                color=cor_atual["valor"],
                parent=janela,
                title="Escolher cor do texto"
            )
            if resultado and resultado[1]:
                cor_atual["valor"] = resultado[1]
                try:
                    inicio, fim = campo.index("sel.first"), campo.index("sel.last")
                    tag = "color_" + resultado[1].replace("#", "")
                    campo.tag_configure(tag, foreground=resultado[1])
                    campo.tag_add(tag, inicio, fim)
                except tk.TclError:
                    pass
            campo.focus_set()

        def alinhar(tipo):
            try:
                inicio, fim = campo.index("sel.first"), campo.index("sel.last")
            except tk.TclError:
                inicio = campo.index("insert linestart")
                fim = campo.index("insert lineend")
            inicio = campo.index(f"{inicio} linestart")
            fim = campo.index(f"{fim} lineend")
            for nome in ("align_left", "align_center", "align_right", "align_justify"):
                campo.tag_remove(nome, inicio, fim)
            tag = "align_" + tipo
            campo.tag_configure(
                tag,
                justify={
                    "left": "left",
                    "center": "center",
                    "right": "right",
                    "justify": "left"
                }[tipo]
            )
            campo.tag_add(tag, inicio, fim)
            alinhamento_atual["valor"] = tipo
            campo.focus_set()

        def inserir_titulo():
            try:
                inicio, fim = campo.index("sel.first"), campo.index("sel.last")
            except tk.TclError:
                inicio, fim = campo.index("insert linestart"), campo.index("insert lineend")
            campo.tag_configure("heading", font=("Arial", 20, "bold"), foreground="#173B68", spacing1=10, spacing3=8)
            campo.tag_add("heading", inicio, fim)
            campo.focus_set()

        def inserir_subtitulo():
            try:
                inicio, fim = campo.index("sel.first"), campo.index("sel.last")
            except tk.TclError:
                inicio, fim = campo.index("insert linestart"), campo.index("insert lineend")
            campo.tag_configure("subheading", font=("Arial", 15, "bold"), foreground="#285F91", spacing1=8, spacing3=5)
            campo.tag_add("subheading", inicio, fim)
            campo.focus_set()

        def aplicar_estilos_ao_texto_digitado(_evento=None):
            # Aplica os botões de negrito/itálico/sublinhado ao caractere recém-digitado.
            try:
                fim = campo.index("insert")
                inicio = campo.index(f"{fim} - 1 chars")
                if inicio == fim:
                    return
                if ativos["bold"]:
                    campo.tag_add("bold", inicio, fim)
                if ativos["italic"]:
                    campo.tag_add("italic", inicio, fim)
                if ativos["underline"]:
                    campo.tag_add("underline", inicio, fim)
                familia = fonte_var.get()
                tamanho = int(tamanho_var.get())
                campo.tag_configure(tag_fonte(familia), font=(familia, tamanho))
                campo.tag_add(tag_fonte(familia), inicio, fim)
                campo.tag_configure("current_color", foreground=cor_atual["valor"])
                campo.tag_add("current_color", inicio, fim)
            except (tk.TclError, ValueError):
                pass

        # Barra superior: família e tamanho da fonte.
        tk.Label(barra, text="Fonte:", bg="#DCE7F2", fg="#142D50").grid(row=0, column=0, padx=(0, 4))
        combo_fonte = ttk.Combobox(
            barra, textvariable=fonte_var, values=familias,
            width=20, state="readonly"
        )
        combo_fonte.grid(row=0, column=1, padx=4, pady=3)
        combo_fonte.bind("<<ComboboxSelected>>", escolher_fonte)

        tk.Label(barra, text="Tamanho:", bg="#DCE7F2", fg="#142D50").grid(row=0, column=2, padx=(8, 4))
        combo_tamanho = ttk.Combobox(
            barra, textvariable=tamanho_var,
            values=[str(n) for n in (8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36, 48, 72)],
            width=5
        )
        combo_tamanho.grid(row=0, column=3, padx=4, pady=3)
        combo_tamanho.bind("<<ComboboxSelected>>", escolher_tamanho)
        combo_tamanho.bind("<Return>", escolher_tamanho)

        def botao(texto, comando, coluna, largura=5):
            tk.Button(
                barra, text=texto, command=comando, width=largura,
                bg="#FFFFFF", fg="#142D50", activebackground="#B8D7F5",
                relief="groove", bd=1, cursor="hand2",
                font=("Arial", 10, "bold"), padx=3, pady=3
            ).grid(row=0, column=coluna, padx=2, pady=3)

        botao("B", lambda: alternar_formato("bold", "bold", {"font": ("Arial", 12, "bold")}), 4)
        botao("I", lambda: alternar_formato("italic", "italic", {"font": ("Arial", 12, "italic")}), 5)
        botao("U", lambda: alternar_formato("underline", "underline", {"underline": True}), 6)
        botao("Cor", escolher_cor, 7, 6)

        barra2 = tk.Frame(janela, bg="#DCE7F2", padx=8, pady=5)
        barra2.pack(fill="x", before=campo)

        tk.Label(barra2, text="Alinhamento:", bg="#DCE7F2", fg="#142D50").pack(side="left", padx=(0, 5))
        tk.Button(barra2, text="☰ Esquerda", command=lambda: alinhar("left"), relief="groove").pack(side="left", padx=2)
        tk.Button(barra2, text="≡ Centro", command=lambda: alinhar("center"), relief="groove").pack(side="left", padx=2)
        tk.Button(barra2, text="☷ Direita", command=lambda: alinhar("right"), relief="groove").pack(side="left", padx=2)
        tk.Button(barra2, text="Justificar", command=lambda: alinhar("justify"), relief="groove").pack(side="left", padx=2)
        tk.Button(barra2, text="Título", command=inserir_titulo, relief="groove").pack(side="left", padx=(12, 2))
        tk.Button(barra2, text="Subtítulo", command=inserir_subtitulo, relief="groove").pack(side="left", padx=2)

        rodape = tk.Frame(janela, bg="#EAF1F8", padx=10, pady=8)
        rodape.pack(fill="x")
        status_var = tk.StringVar(value="Pronto")
        tk.Label(rodape, textvariable=status_var, bg="#EAF1F8", fg="#52677D").pack(side="left")

        def exportar_pdf():
            texto = campo.get("1.0", "end-1c")
            if not texto.strip():
                messagebox.showwarning(
                    "Atenção", "Digite algum texto antes de converter.", parent=janela
                )
                return

            destino = filedialog.asksaveasfilename(
                parent=janela,
                title="Salvar documento PDF",
                defaultextension=".pdf",
                filetypes=[("Arquivo PDF", "*.pdf")],
                initialfile="documento.pdf"
            )
            if not destino:
                return

            try:
                criar_pdf_formatado(campo, destino)
                messagebox.showinfo(
                    "davIA PDF Maker", "PDF formatado criado com sucesso!", parent=janela
                )
            except Exception as erro:
                messagebox.showerror(
                    "Erro ao criar PDF",
                    f"Não foi possível gerar o documento:\n{erro}",
                    parent=janela
                )

        tk.Button(
            rodape, text="CANCELAR", command=janela.destroy,
            bg="#718096", fg="white", relief="flat",
            font=("Arial", 10, "bold"), padx=14, pady=8
        ).pack(side="right", padx=5)
        tk.Button(
            rodape, text="GERAR PDF", command=exportar_pdf,
            bg="#1565C0", fg="white", activebackground="#0D47A1",
            relief="flat", cursor="hand2",
            font=("Arial", 10, "bold"), padx=20, pady=8
        ).pack(side="right", padx=5)

        campo.bind("<KeyRelease>", aplicar_estilos_ao_texto_digitado)
        campo.focus_set()

    def imagens_para_pdf(self):
        arquivos = filedialog.askopenfilenames(
            parent=self.root,
            title="Selecionar imagens",
            filetypes=[
                ("Imagens", "*.png *.jpg *.jpeg *.bmp *.tiff *.tif")
            ]
        )
        if not arquivos:
            return

        destino = filedialog.asksaveasfilename(
            parent=self.root,
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")]
        )
        if not destino:
            return

        imagens = []
        try:
            imagens = [
                Image.open(arquivo).convert("RGB")
                for arquivo in arquivos
            ]
            imagens[0].save(
                destino,
                save_all=True,
                append_images=imagens[1:]
            )
            messagebox.showinfo(
                "Sucesso",
                "PDF criado com sucesso!",
                parent=self.root
            )
        except Exception as erro:
            messagebox.showerror(
                "Erro",
                f"Não foi possível converter as imagens:\n{erro}",
                parent=self.root
            )
        finally:
            for imagem in imagens:
                imagem.close()

    def juntar_pdfs(self):
        arquivos = filedialog.askopenfilenames(
            parent=self.root,
            title="Selecionar PDFs",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if not arquivos:
            return

        destino = filedialog.asksaveasfilename(
            parent=self.root,
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")]
        )
        if not destino:
            return

        resultado = fitz.open()
        try:
            for arquivo in arquivos:
                with fitz.open(arquivo) as documento:
                    resultado.insert_pdf(documento)

            resultado.save(destino)
            messagebox.showinfo(
                "Sucesso",
                "PDFs unidos com sucesso!",
                parent=self.root
            )
        except Exception as erro:
            messagebox.showerror(
                "Erro",
                f"Não foi possível juntar os PDFs:\n{erro}",
                parent=self.root
            )
        finally:
            resultado.close()

    def dividir_pdf(self):
        arquivo = filedialog.askopenfilename(
            parent=self.root,
            title="Selecionar PDF",
            filetypes=[("PDF", "*.pdf")]
        )
        if not arquivo:
            return

        pasta = filedialog.askdirectory(
            parent=self.root,
            title="Selecionar pasta de destino"
        )
        if not pasta:
            return

        documento = None
        try:
            documento = fitz.open(arquivo)
            for numero in range(len(documento)):
                novo = fitz.open()
                try:
                    novo.insert_pdf(
                        documento,
                        from_page=numero,
                        to_page=numero
                    )
                    destino = os.path.join(
                        pasta,
                        f"pagina_{numero + 1}.pdf"
                    )
                    novo.save(destino)
                finally:
                    novo.close()

            messagebox.showinfo(
                "Sucesso",
                "PDF dividido em páginas!",
                parent=self.root
            )
        except Exception as erro:
            messagebox.showerror(
                "Erro",
                f"Não foi possível dividir o PDF:\n{erro}",
                parent=self.root
            )
        finally:
            if documento is not None:
                documento.close()

    def extrair_imagens(self):
        arquivo = filedialog.askopenfilename(
            parent=self.root,
            title="Selecionar PDF",
            filetypes=[("PDF", "*.pdf")]
        )
        if not arquivo:
            return

        pasta = filedialog.askdirectory(
            parent=self.root,
            title="Selecionar pasta de destino"
        )
        if not pasta:
            return

        documento = None
        try:
            documento = fitz.open(arquivo)
            quantidade = 0

            for pagina_num, pagina in enumerate(documento):
                imagens = pagina.get_images(full=True)

                for imagem_num, info in enumerate(imagens):
                    dados = documento.extract_image(info[0])
                    nome = (
                        f"pagina_{pagina_num + 1}_"
                        f"imagem_{imagem_num + 1}."
                        f"{dados['ext']}"
                    )
                    destino = os.path.join(pasta, nome)

                    with open(destino, "wb") as arquivo_img:
                        arquivo_img.write(dados["image"])

                    quantidade += 1

            messagebox.showinfo(
                "Resultado",
                f"{quantidade} imagem(ns) extraída(s).",
                parent=self.root
            )
        except Exception as erro:
            messagebox.showerror(
                "Erro",
                f"Não foi possível extrair as imagens:\n{erro}",
                parent=self.root
            )
        finally:
            if documento is not None:
                documento.close()

    def abrir_pdf(self):
        arquivo = filedialog.askopenfilename(
            parent=self.root,
            title="Abrir PDF",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if arquivo:
            try:
                os.startfile(arquivo)
            except Exception as erro:
                messagebox.showerror(
                    "Erro",
                    f"Não foi possível abrir o PDF:\n{erro}",
                    parent=self.root
                )



def criar_pdf_formatado(campo, destino):
    """Exporta o conteúdo do Tk Text preservando estilos e alinhamentos básicos."""
    estilos = getSampleStyleSheet()
    documento = SimpleDocTemplate(
        destino,
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm
    )

    alinhamentos = {
        "align_left": TA_LEFT,
        "align_center": TA_CENTER,
        "align_right": TA_RIGHT,
        "align_justify": TA_JUSTIFY
    }

    def estilo_da_posicao(posicao):
        tags = set(campo.tag_names(posicao))
        if "heading" in tags:
            return {"family": "Helvetica", "size": 20, "color": "#173B68", "bold": True, "italic": False, "underline": False}
        if "subheading" in tags:
            return {"family": "Helvetica", "size": 15, "color": "#285F91", "bold": True, "italic": False, "underline": False}

        resultado = {
            "family": "Helvetica", "size": 11, "color": "#000000",
            "bold": False, "italic": False, "underline": False
        }
        if "bold" in tags:
            resultado["bold"] = True
        if "italic" in tags:
            resultado["italic"] = True
        if "underline" in tags:
            resultado["underline"] = True

        for tag in tags:
            if tag.startswith("color_"):
                resultado["color"] = "#" + tag[6:]
            elif tag.startswith("size_"):
                try:
                    resultado["size"] = int(tag.split("_", 1)[1])
                except ValueError:
                    pass
            elif tag.startswith("family_"):
                # O PDF usa as famílias-base do ReportLab para manter compatibilidade.
                nome = tag[len("family_"):].lower()
                if "times" in nome:
                    resultado["family"] = "Times-Roman"
                elif "courier" in nome:
                    resultado["family"] = "Courier"
                else:
                    resultado["family"] = "Helvetica"
        if "current_color" in tags:
            info = campo.tag_cget("current_color", "foreground")
            if info:
                resultado["color"] = info
        return resultado

    def nome_fonte(estilo):
        familia = estilo["family"]
        if familia == "Times-Roman":
            if estilo["bold"] and estilo["italic"]:
                return "Times-BoldItalic"
            if estilo["bold"]:
                return "Times-Bold"
            if estilo["italic"]:
                return "Times-Italic"
            return "Times-Roman"
        if familia == "Courier":
            if estilo["bold"] and estilo["italic"]:
                return "Courier-BoldOblique"
            if estilo["bold"]:
                return "Courier-Bold"
            if estilo["italic"]:
                return "Courier-Oblique"
            return "Courier"
        if estilo["bold"] and estilo["italic"]:
            return "Helvetica-BoldOblique"
        if estilo["bold"]:
            return "Helvetica-Bold"
        if estilo["italic"]:
            return "Helvetica-Oblique"
        return "Helvetica"

    elementos = []
    ultima_linha = int(float(campo.index("end-1c")))
    for numero in range(1, ultima_linha + 1):
        inicio = f"{numero}.0"
        fim = f"{numero}.end"
        texto_linha = campo.get(inicio, fim)
        if not texto_linha:
            elementos.append(Spacer(1, 8))
            continue

        segmentos = []
        posicao = 0
        while posicao < len(texto_linha):
            indice = f"{numero}.{posicao}"
            estilo = estilo_da_posicao(indice)
            fim_segmento = posicao + 1
            while fim_segmento < len(texto_linha):
                proximo = estilo_da_posicao(f"{numero}.{fim_segmento}")
                if proximo != estilo:
                    break
                fim_segmento += 1

            trecho = escape(texto_linha[posicao:fim_segmento])
            trecho = trecho.replace(" ", "&nbsp;") if trecho.startswith(" ") or trecho.endswith(" ") else trecho
            atributos = (
                f'<font name="{nome_fonte(estilo)}" size="{estilo["size"]}" '
                f'color="{estilo["color"]}">{trecho}</font>'
            )
            if estilo["underline"]:
                atributos = f"<u>{atributos}</u>"
            segmentos.append(atributos)
            posicao = fim_segmento

        tags_linha = set(campo.tag_names(inicio))
        alinhamento = next(
            (alinhamentos[tag] for tag in ("align_left", "align_center", "align_right", "align_justify") if tag in tags_linha),
            TA_LEFT
        )
        estilo_paragrafo = ParagraphStyle(
            f"linha_{numero}",
            parent=estilos["Normal"],
            fontName="Helvetica",
            fontSize=11,
            leading=15,
            alignment=alinhamento,
            spaceAfter=5
        )
        elementos.append(Paragraph("".join(segmentos) or " ", estilo_paragrafo))

    documento.build(elementos)


def criar_pdf_texto(texto, destino):
    estilos = getSampleStyleSheet()

    documento = SimpleDocTemplate(
        destino,
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm
    )

    elementos = []
    for linha in texto.splitlines():
        linha_formatada = escape(linha)
        elementos.append(
            Paragraph(
                linha_formatada if linha_formatada else " ",
                estilos["Normal"]
            )
        )
        elementos.append(Spacer(1, 5))

    documento.build(elementos)


if __name__ == "__main__":
    janela = tk.Tk()
    aplicativo = DaviaPDFMaker(janela)
    janela.mainloop()

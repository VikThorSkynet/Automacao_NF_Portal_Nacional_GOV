import argparse
import logging
import os
from datetime import datetime
from pathlib import Path
from queue import Empty, Queue
from threading import Event, Thread
from tkinter import BooleanVar, StringVar, Tk, filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from playwright.sync_api import sync_playwright

from lote import Lote, PORTAL, PararLote, mensagem_segura
from planilha import ler_intervalo


class Janela:
    def __init__(self, raiz, *, somente_teste=False, arquivo=None):
        self.raiz = raiz
        self.arquivo = arquivo
        self.somente_teste = somente_teste
        self.log = None
        self.fila = Queue()
        self.retomar = Event()
        self.cancelar = Event()
        self.thread = None
        self.fechando = False
        self.cnpj = BooleanVar(value=False)
        self.planilha = StringVar(value=arquivo.name if arquivo else "Escolha a planilha para carregar G10 e G13.")
        self.atual = StringVar(value="Aguardando.")
        self.totais = StringVar(value="Emitidas: 0   Ignoradas: 0   Erros: 0")

        raiz.title("Teste de NFS-e até revisão" if somente_teste else "Emissão de NFS-e")
        raiz.minsize(650, 390)
        quadro = ttk.Frame(raiz, padding=16)
        quadro.pack(fill="both", expand=True)
        self.botao_excel = ttk.Button(quadro, text="Escolher Excel", command=self.escolher)
        self.botao_excel.pack(anchor="w")
        ttk.Label(quadro, textvariable=self.planilha, wraplength=615).pack(anchor="w", pady=6)
        self.botao_cnpj = ttk.Checkbutton(quadro, text="Aceitar CNPJ na coluna C", variable=self.cnpj)
        self.botao_cnpj.pack(anchor="w")

        acoes = ttk.Frame(quadro)
        acoes.pack(fill="x", pady=10)
        self.botao_teste = ttk.Button(acoes, text="Testar até revisão", command=lambda: self.iniciar(True))
        self.botao_emitir = ttk.Button(acoes, text="Emitir lote", command=lambda: self.iniciar(False))
        if somente_teste:
            self.botao_emitir.configure(state="disabled")
        self.botao_retomar = ttk.Button(acoes, text="Portal pronto", command=self.liberar, state="disabled")
        self.botao_parar = ttk.Button(acoes, text="Encerrar", command=self.parar, state="disabled")
        self.botao_log = ttk.Button(acoes, text="Abrir log completo", command=self.abrir_log, state="disabled")
        for botao in (self.botao_teste, self.botao_emitir, self.botao_retomar, self.botao_parar, self.botao_log):
            botao.pack(side="left", padx=(0, 6))

        ttk.Label(quadro, textvariable=self.totais).pack(anchor="w")
        ttk.Label(quadro, textvariable=self.atual, wraplength=615).pack(anchor="w", pady=8)
        self.eventos = ScrolledText(quadro, height=10, state="disabled", wrap="word")
        self.eventos.pack(fill="both", expand=True)
        raiz.protocol("WM_DELETE_WINDOW", self.fechar)
        raiz.after(100, self.atualizar)

    def escolher(self):
        caminho = filedialog.askopenfilename(parent=self.raiz, title="Escolha a planilha da NFS-e",
                                            filetypes=[("Planilhas Excel", "*.xlsx *.xlsm")])
        if caminho:
            self.arquivo = Path(caminho)
            self.planilha.set(self.arquivo.name)

    def iniciar(self, teste):
        if self.thread is not None or (self.somente_teste and not teste):
            return
        if self.arquivo is None:
            self.escolher()
        if self.arquivo is None:
            return
        try:
            intervalo = ler_intervalo(self.arquivo)
            pasta = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "AutomacaoNFSe" / "logs"
            pasta.mkdir(parents=True, exist_ok=True)
            self.log = pasta / f"nfse-{datetime.now():%Y%m%d-%H%M%S-%f}.log"
            handler = logging.FileHandler(self.log, encoding="utf-8")
        except (OSError, ValueError) as erro:
            messagebox.showerror("Não foi possível iniciar", mensagem_segura(erro), parent=self.raiz)
            return
        handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
        logger = logging.getLogger(f"nfse.{self.log.stem}")
        logger.setLevel(logging.INFO)
        logger.addHandler(handler)
        logger.propagate = False
        self.planilha.set(f"{self.arquivo.name} · linhas {intervalo.start} a {intervalo.stop - 1}")
        self.totais.set("Emitidas: 0   Ignoradas: 0   Erros: 0")
        self.retomar.clear()
        self.cancelar.clear()
        for botao in (self.botao_excel, self.botao_cnpj, self.botao_teste, self.botao_emitir):
            botao.configure(state="disabled")
        self.botao_parar.configure(state="normal")
        self.botao_log.configure(state="normal")
        self.atual.set("Abrindo o Chrome…")
        self.thread = Thread(target=self.trabalhar, args=(self.arquivo, teste, self.cnpj.get(), logger, handler))
        self.thread.start()

    def aguardar(self):
        self.retomar.wait()
        self.retomar.clear()
        if self.cancelar.is_set():
            raise PararLote("Lote encerrado pelo operador.")

    def trabalhar(self, arquivo, teste, cnpj, logger, handler):
        try:
            with sync_playwright() as playwright:
                navegador = playwright.chromium.launch(channel="chrome", headless=False)
                pagina = navegador.new_page()
                pagina.set_default_timeout(7000)
                pagina.goto(PORTAL, wait_until="domcontentloaded", timeout=30000)
                lote = Lote(pagina, arquivo, self.fila.put, self.aguardar, self.cancelar, logger)
                lote.evento("Login", "Entre com o certificado, abra Emissão Completa e clique em Portal pronto.")
                self.aguardar()
                lote.executar(teste=teste, aceitar_cnpj=cnpj)
                # Mantém a última tela aberta para inspeção até o operador encerrar.
                self.cancelar.wait()
        except PararLote as erro:
            logger.info("%s", erro)
            self.fila.put({"etapa": "Encerrado", "mensagem": str(erro)})
        except Exception as erro:
            mensagem = mensagem_segura(erro)
            logger.error("Lote interrompido | %s | %s", type(erro).__name__, mensagem)
            self.fila.put({"etapa": "Interrompido", "mensagem": mensagem})
        finally:
            logger.removeHandler(handler)
            handler.close()
            self.fila.put(None)

    def atualizar(self):
        while True:
            try:
                evento = self.fila.get_nowait()
            except Empty:
                break
            if evento is None:
                self.thread = None
                if self.fechando:
                    self.raiz.destroy()
                    return
                for botao in (self.botao_excel, self.botao_cnpj, self.botao_teste, self.botao_emitir):
                    botao.configure(state="normal")
                if self.somente_teste:
                    self.botao_emitir.configure(state="disabled")
                self.botao_retomar.configure(state="disabled")
                self.botao_parar.configure(state="disabled")
                continue
            etapa, mensagem = evento["etapa"], evento["mensagem"]
            linha = evento.get("linha")
            nome = evento.get("nome")
            self.atual.set(" · ".join(filter(None, [f"Linha {linha}" if linha else "", nome, etapa, mensagem])))
            if "emitidas" in evento:
                self.totais.set(f"Emitidas: {evento['emitidas']}   Ignoradas: {evento['ignoradas']}   Erros: {evento['erros']}")
            if etapa in ("Login", "Pausa") and not self.cancelar.is_set():
                self.botao_retomar.configure(state="normal", text="Portal pronto" if etapa == "Login" else "Retomar")
            self.eventos.configure(state="normal")
            prefixo = f"Linha {linha} · " if linha else ""
            self.eventos.insert("end", f"{datetime.now():%H:%M:%S} · {prefixo}{etapa} · {mensagem}\n")
            self.eventos.see("end")
            self.eventos.configure(state="disabled")
        self.raiz.after(100, self.atualizar)

    def abrir_log(self):
        if self.log is not None:
            os.startfile(self.log)

    def liberar(self):
        self.botao_retomar.configure(state="disabled")
        self.retomar.set()

    def parar(self):
        self.botao_parar.configure(state="disabled")
        self.cancelar.set()
        self.liberar()

    def fechar(self):
        if self.thread is None:
            self.raiz.destroy()
        else:
            self.fechando = True
            self.parar()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--somente-teste", action="store_true")
    parser.add_argument("--planilha", type=Path)
    args = parser.parse_args()
    raiz = Tk()
    Janela(raiz, somente_teste=args.somente_teste, arquivo=args.planilha)
    raiz.mainloop()

import logging
import re
from time import monotonic
from xml.etree import ElementTree

from playwright.sync_api import Error, TimeoutError, expect

import pessoas
import servico
import valores
from planilha import ler_intervalo, validar_linha
from registro import CHAVE_NFSE, Registro


PORTAL = "https://www.nfse.gov.br/EmissorNacional/"
CAMPO_PESSOAS = "#DataCompetencia"
NOVA_NFSE = "Nova NFS-e"
ESPERA_TELA_MS = 3000
INTERVALO_TENTATIVA_S = 3
FALHA_REDE = re.compile(r"ERR_(?:INTERNET_DISCONNECTED|NAME_NOT_RESOLVED|CONNECTION_\w+|NETWORK_\w+|TIMED_OUT)")
SESSAO_EXPIRADA = re.compile(r"sessão.{0,45}(?:expir|encerr)|tempo.{0,35}(?:esgot|sessão)", re.I)
TOMADOR_REJEITADO = re.compile(
    r"(?:CPF(?:/CNPJ)?|CNPJ|tomador)[^.!?\n]{0,40}(?:inválido|não encontrad[oa])", re.I
)


class PararLote(Exception):
    pass


class ErroDaLinha(Exception):
    pass


def mensagem_segura(texto, nome=""):
    texto = str(texto).splitlines()[0] if str(texto) else ""
    if nome:
        texto = re.sub(re.escape(nome), "[nome]", texto, flags=re.I)
    texto = re.sub(r"https?://\S+", "[endereço]", texto)
    return re.sub(r"(?<!\d)(?:\d[./\- ]?){11,}\d?(?!\d)", "[identificador]", texto)[:240]


def causas(erro):
    while erro is not None:
        yield erro
        erro = erro.__cause__


def ler_chave_xml(caminho):
    raiz = ElementTree.parse(caminho).getroot()
    chaves = {
        chave for no in raiz.iter()
        if no.tag.rsplit("}", 1)[-1] == "infNFSe"
        and (chave := no.get("Id", "").removeprefix("NFS"))
        and CHAVE_NFSE.fullmatch(chave)
    }
    if len(chaves) != 1:
        raise ValueError("O XML não contém uma chave de acesso inequívoca da NFS-e.")
    return chaves.pop()


def capturar_chave(pagina):
    # Limita a busca ao campo de chave, sem extrair números de outras áreas da página.
    valores_chave = pagina.evaluate(r"""() => {
      const rotulo = /^chave de acesso(?:\s+(?:da|do)\s+(?:nfs-?e|nota))?\s*:?$/i;
      const resultados = [];
      for (const el of document.querySelectorAll('label,dt,th,strong,b,span,div,input')) {
        if (!el.getClientRects().length) continue;
        const campo = el instanceof HTMLInputElement &&
          /chave\s*de\s*acesso|chaveacesso/i.test(
            [el.id, el.name, el.getAttribute('aria-label')].join(' ')
          );
        if (campo) resultados.push(el.value);
        if (el.children.length || !rotulo.test(el.textContent.trim())) continue;
        const vinculado = el.control || el.nextElementSibling || el.parentElement?.nextElementSibling;
        if (vinculado?.getClientRects().length) resultados.push(vinculado.value || vinculado.innerText);
        if (el.parentElement) resultados.push(el.parentElement.innerText.replace(el.innerText, ''));
      }
      return resultados;
    }""")
    chaves = {re.sub(r"\s+", "", texto) for texto in valores_chave if texto}
    chaves = {chave for chave in chaves if CHAVE_NFSE.fullmatch(chave)}
    if len(chaves) == 1:
        return chaves.pop()
    if len(chaves) > 1:
        raise RuntimeError("A tela mostrou mais de uma chave de acesso.")

    nome = re.compile(r"^\s*(?:baixar|download(?: do)?)\s+(?:o\s+)?xml\b", re.I)
    download = pagina.get_by_role("link", name=nome).or_(pagina.get_by_role("button", name=nome))
    with pagina.expect_download(timeout=15000) as evento:
        download.filter(visible=True).click()
    arquivo = evento.value
    try:
        return ler_chave_xml(arquivo.path())
    finally:
        arquivo.delete()


class Lote:
    def __init__(self, pagina, excel, publicar, esperar_retomada, cancelado, logger: logging.Logger):
        self.pagina = pagina
        self.registro = Registro(excel)
        self.publicar = publicar
        self.esperar_retomada = esperar_retomada
        self.cancelado = cancelado
        self.logger = logger
        self.linha = None
        self.nome = ""
        self.contagem = {"emitidas": 0, "ignoradas": 0, "erros": 0}
        self.destino = pagina.get_by_text(NOVA_NFSE, exact=True).filter(visible=True).first
        self.emitir = pagina.get_by_role("button", name=re.compile(r"^Emitir NFS-e$", re.I))

    def evento(self, etapa, mensagem, nivel=logging.INFO):
        self.logger.log(nivel, "linha=%s | %s | %s", self.linha or "-", etapa,
                        mensagem_segura(mensagem, self.nome))
        self.publicar({"etapa": etapa, "mensagem": mensagem_segura(mensagem, self.nome),
                       "linha": self.linha, "nome": self.nome, **self.contagem})

    def conferir_parada(self):
        if self.cancelado.is_set() or self.pagina.is_closed():
            raise PararLote("Lote interrompido.")

    def pausa(self, motivo):
        self.evento("Pausa", motivo, logging.WARNING)
        self.esperar_retomada()
        self.conferir_parada()

    def aviso(self):
        return " ".join(self.pagina.locator(
            ".jconfirm.jconfirm-open .jconfirm-content:visible, [data-valmsg-for]:visible"
        ).all_text_contents())

    def pausa_portal(self):
        self.conferir_parada()
        if "/login" in self.pagina.url.lower() or SESSAO_EXPIRADA.search(self.aviso()):
            return "Sessão expirada. Entre novamente e volte à etapa do lote."
        if not self.pagina.evaluate("navigator.onLine"):
            return "Sem conexão. Restabeleça a internet e clique em Retomar."
        return None

    def passo(self, etapa, acao, *, pronto=None, tela=None):
        # Apenas preenchimentos e navegação de rascunho passam por esta rotina.
        while True:
            for tentativa in range(1, 4):
                inicio = monotonic()
                try:
                    motivo = self.pausa_portal()
                    if motivo:
                        break
                    if pronto is not None and pronto.is_visible():
                        self.evento(etapa, "A tela de destino já está aberta.")
                        return
                    if tela is not None and not tela.is_visible():
                        self.evento(etapa, f"Aguardando a tela esperada na mesma aba ({tentativa}/3).")
                        alvo = tela.filter(visible=True)
                        if pronto is not None:
                            alvo = alvo.or_(pronto.filter(visible=True))
                        try:
                            alvo.first.wait_for(state="visible", timeout=ESPERA_TELA_MS)
                        except TimeoutError:
                            motivo = "A tela esperada não apareceu. Volte a ela na mesma aba antes de retomar."
                            continue
                        self.conferir_parada()
                        if pronto is not None and pronto.is_visible():
                            self.evento(etapa, "Avanço confirmado após a espera.")
                            return
                        if not tela.is_visible():
                            motivo = "A tela esperada ainda não está disponível na mesma aba."
                            continue
                    self.evento(etapa, f"Tentativa {tentativa}/3.")
                    acao()
                    if pronto is not None:
                        pronto.wait_for(state="visible", timeout=ESPERA_TELA_MS)
                    self.evento(etapa, f"Concluído em {monotonic() - inicio:.1f} s.")
                    return
                except (Error, RuntimeError, AssertionError) as erro:
                    self.conferir_parada()
                    self.evento(etapa, f"Falha na tentativa {tentativa} após "
                                f"{monotonic() - inicio:.1f} s. {mensagem_segura(erro, self.nome)}",
                                logging.WARNING)
                    if any(FALHA_REDE.search(str(causa)) for causa in causas(erro)):
                        motivo = "Falha de conexão. Restabeleça a internet e clique em Retomar."
                        break
                    if pronto is not None and pronto.is_visible():
                        self.evento(etapa, "Avanço confirmado após a espera.")
                        return
                    aviso = self.aviso()
                    interrupcao = self.pausa_portal()
                    if not interrupcao and etapa == "Pessoas" and TOMADOR_REJEITADO.search(aviso):
                        raise ErroDaLinha("O portal rejeitou o documento do tomador.") from erro
                    motivo = interrupcao or mensagem_segura(erro, self.nome)
                    if interrupcao or aviso or not any(isinstance(causa, TimeoutError) for causa in causas(erro)):
                        break
                    if tentativa < 3:
                        self.evento(etapa, f"Nova tentativa automática em {INTERVALO_TENTATIVA_S:g} s.")
                        if self.cancelado.wait(INTERVALO_TENTATIVA_S):
                            self.conferir_parada()
            self.pausa(f"{etapa}: {motivo}")

    def reconciliar(self, *, teste):
        dados = self.registro.ler_pendente()
        if dados is None:
            return
        if teste:
            raise RuntimeError("Há uma emissão pendente nesta planilha. Confira o resultado antes do teste.")
        while True:
            try:
                self.registro.reconciliar(dados)
                self.evento("Excel", f"Pendência da linha {dados['linha']} reconciliada.")
                return
            except (OSError, RuntimeError) as erro:
                self.pausa(str(erro))

    def enviar_nota(self, dados):
        self.conferir_parada()
        expect(self.emitir).to_be_enabled()
        self.registro.conferir_linha(dados)
        self.registro.salvar_pendente(dados)
        self.evento("Emissão", "Tentativa registrada. Enviando um único clique em Emitir NFS-e.")
        try:
            self.emitir.click(timeout=15000)
            expect(self.emitir).not_to_be_visible(timeout=20000)
            expect(self.destino).to_be_visible(timeout=20000)
            chave = capturar_chave(self.pagina)
            dados["chave"] = chave
            self.registro.salvar_pendente(dados)
            self.evento("Emissão", "Portal confirmou a nota e a chave foi capturada.")
        except (Error, AssertionError, RuntimeError, ValueError, OSError, ElementTree.ParseError) as erro:
            self.evento("Emissão", mensagem_segura(erro, self.nome), logging.ERROR)
        self.reconciliar(teste=False)
        self.contagem["emitidas"] += 1
        self.evento("Excel", f"Chave de acesso conferida em E{self.linha}.")

    def abrir_proxima(self):
        pessoas_aberta = self.pagina.locator(CAMPO_PESSOAS)
        if pessoas_aberta.is_visible():
            return

        def abrir():
            if self.destino.is_visible():
                self.destino.click()
            if not pessoas_aberta.is_visible():
                completa = self.pagina.get_by_text("Emissão Completa", exact=True).filter(visible=True)
                if completa.count() == 1:
                    completa.click()
            expect(pessoas_aberta).to_be_visible()

        self.passo("Nova NFS-e", abrir, pronto=pessoas_aberta)

    def executar(self, *, teste=False, aceitar_cnpj=False):
        excel = self.registro.excel
        with self.registro.exclusivo():
            self.reconciliar(teste=teste)
            intervalo = ler_intervalo(excel)
            descricao = servico.ler_descricao_servico(excel)
            self.evento("Lote", f"Intervalo de G10/G13: linhas {intervalo.start} a {intervalo.stop - 1}.")
            for self.linha in intervalo:
                self.conferir_parada()
                self.nome = ""

                def ler_nome(aba):
                    validar_linha(aba, self.linha, aceitar_cnpj=aceitar_cnpj)
                    return str(aba[f"B{self.linha}"].value or "")

                try:
                    self.nome, dados = self.registro.ler_linha(self.linha, ler_nome)
                except ValueError as erro:
                    self.contagem["ignoradas"] += 1
                    self.evento("Excel", str(erro))
                    continue
                data, documento = pessoas.ler_dados_pessoas(excel, self.linha, aceitar_cnpj=aceitar_cnpj)
                valor, aliquota = valores.ler_dados_valores(excel, self.linha, aceitar_cnpj=aceitar_cnpj)
                self.evento("Excel", "Linha validada.")
                self.abrir_proxima()
                tela_pessoas = self.pagina.locator(CAMPO_PESSOAS)
                tela_servico = self.pagina.locator(servico.CAMPO_DESCRICAO)
                tela_valores = self.pagina.locator(valores.CAMPO_VALOR)
                try:
                    self.passo("Pessoas", lambda: pessoas.preencher_pessoas(
                        self.pagina, data, documento, informar=lambda acao: self.evento("Pessoas", acao)
                    ), tela=tela_pessoas)
                except ErroDaLinha as erro:
                    while True:
                        try:
                            self.registro.gravar(dados)
                            break
                        except (OSError, RuntimeError) as falha:
                            self.pausa(f"Não foi possível gravar F{self.linha}: {falha}")
                    self.contagem["erros"] += 1
                    self.evento("Excel", f"F{self.linha} marcado com Erro. {erro}", logging.ERROR)
                    if self.pagina.locator(".jconfirm.jconfirm-open").first.is_visible():
                        self.pausa("Feche o aviso do tomador para seguir para a próxima linha.")
                    continue
                self.passo("Avançar de Pessoas", lambda: pessoas.avancar_pessoas(self.pagina),
                           tela=tela_pessoas, pronto=tela_servico)
                self.passo("Serviço", lambda: servico.preencher_servico(
                    self.pagina, descricao, informar=lambda acao: self.evento("Serviço", acao)
                ),
                           tela=tela_servico)
                self.passo("Avançar de Serviço", lambda: servico.avancar_servico(self.pagina),
                           tela=tela_servico, pronto=tela_valores)
                self.passo("Valores", lambda: valores.preencher_valores(
                    self.pagina, valor, aliquota, informar=lambda acao: self.evento("Valores", acao)
                ),
                           tela=tela_valores)
                self.passo("Avançar de Valores", lambda: valores.avancar_valores(self.pagina),
                           tela=tela_valores, pronto=self.emitir)
                if teste:
                    self.evento("Revisão", "Teste concluído. A tela de revisão está aberta.")
                    return
                self.enviar_nota(dados)
            self.evento("Fim", "Lote concluído.")

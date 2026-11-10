import logging
import sys
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Event
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from playwright.sync_api import Error, TimeoutError, sync_playwright

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "Novo programa de NF"))
import lote


CHAVE = "0123456789" * 5
PESSOAS = '<input id="DataCompetencia">'
SERVICO = '<textarea id="ServicoPrestado_Descricao"></textarea>'
VALORES = '<input id="Valores_ValorServico">'
RESULTADO = (
    f'<label for="chave">Chave de acesso da NFS-e</label><input id="chave" value="{CHAVE}">'
    '<button onclick="document.body.innerHTML=\'<input id=DataCompetencia>\'">Nova NFS-e</button>'
)


class TesteLote(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.playwright = sync_playwright().start()
        cls.chrome = cls.playwright.chromium.launch(channel="chrome", headless=True)

    @classmethod
    def tearDownClass(cls):
        cls.chrome.close()
        cls.playwright.stop()

    def setUp(self):
        self.pasta = TemporaryDirectory()
        self.addCleanup(self.pasta.cleanup)
        self.excel = Path(self.pasta.name) / "teste.xlsx"
        workbook = Workbook()
        aba = workbook.active
        aba["G2"] = "Descrição de teste"
        aba["G4"] = 4.16
        aba["G7"] = date(2026, 11, 10)
        aba["G10"] = 2
        aba["G13"] = 4
        for linha in range(2, 6):
            aba[f"B{linha}"] = f"Tomador {linha}"
            aba[f"C{linha}"] = f"0123456789{linha}"
            aba[f"D{linha}"] = 100 + linha
        aba["E3"] = "Nota existente"
        workbook.save(self.excel)
        workbook.close()
        self.pagina = self.chrome.new_page()
        self.addCleanup(self.pagina.close)
        self.pagina.set_content(PESSOAS)
        self.eventos = []
        self.logger = logging.getLogger("teste-lote")
        self.logger.addHandler(logging.NullHandler())

    def criar_lote(self, retomar):
        return lote.Lote(self.pagina, self.excel, self.eventos.append, retomar, Event(), self.logger)

    def criar_tela_emissao(self, pagina):
        pagina.set_content('<button id="enviar">Emitir NFS-e</button>')
        pagina.evaluate(
            """html => document.querySelector('#enviar').onclick = () => {
              window.envios = (window.envios || 0) + 1;
              document.body.innerHTML = html;
            }""", RESULTADO
        )

    def parar(self):
        raise lote.PararLote("Parada solicitada no teste.")

    def test_lote_grava_e_f_e_passa_para_outra_nota_na_mesma_aba(self):
        def preencher_pessoas(pagina, data, documento, **kwargs):
            if documento.endswith("4"):
                pagina.set_content(PESSOAS + '<span data-valmsg-for="Tomador">CPF inválido</span>')
                raise TimeoutError("Pesquisa recusada.")

        operacao = self.criar_lote(self.parar)
        with (
            patch.object(lote.pessoas, "preencher_pessoas", side_effect=preencher_pessoas),
            patch.object(lote.pessoas, "avancar_pessoas", side_effect=lambda p: p.set_content(SERVICO)),
            patch.object(lote.servico, "preencher_servico"),
            patch.object(lote.servico, "avancar_servico", side_effect=lambda p: p.set_content(VALORES)),
            patch.object(lote.valores, "preencher_valores"),
            patch.object(lote.valores, "avancar_valores", side_effect=self.criar_tela_emissao),
        ):
            operacao.executar()

        workbook = load_workbook(self.excel)
        try:
            aba = workbook.active
            for celula in ("E2", "E5"):
                self.assertEqual(aba[celula].value, CHAVE)
                self.assertTrue(aba[celula].quotePrefix)
                self.assertEqual(aba[celula].data_type, "s")
            self.assertEqual(aba["E3"].value, "Nota existente")
            self.assertEqual(aba["F4"].value, "Erro")
        finally:
            workbook.close()
        self.assertEqual(operacao.contagem, {"emitidas": 2, "ignoradas": 1, "erros": 1})
        self.assertEqual(len(self.pagina.context.pages), 1)
        self.assertEqual(self.pagina.evaluate("window.envios"), 2)

    def test_resposta_incerta_bloqueia_novo_envio_ate_conferencia(self):
        self.criar_tela_emissao(self.pagina)
        operacao = self.criar_lote(self.parar)
        operacao.linha = 2
        _, dados = operacao.registro.ler_linha(2, lambda aba: None)
        with patch.object(lote, "capturar_chave", side_effect=ValueError("Resposta incerta.")):
            with self.assertRaises(lote.PararLote):
                operacao.enviar_nota(dados)
        self.assertIsNotNone(operacao.registro.ler_pendente())
        with self.assertRaises(lote.PararLote):
            self.criar_lote(self.parar).executar()
        self.assertEqual(self.pagina.evaluate("window.envios"), 1)
        workbook = load_workbook(self.excel)
        try:
            self.assertIsNone(workbook.active["E2"].value)
            self.assertIsNone(workbook.active["F2"].value)
        finally:
            workbook.close()

    def test_timeout_repete_acao_segura_e_erro_de_rede_pausa(self):
        operacao = self.criar_lote(self.parar)
        tentativas = iter([TimeoutError("Espera excedida."), None])

        def acao():
            erro = next(tentativas)
            if erro:
                raise erro

        with patch.object(lote, "INTERVALO_TENTATIVA_S", 0):
            operacao.passo("Serviço", acao)
        self.assertTrue(any("Tentativa 2" in e["mensagem"] for e in self.eventos))
        with self.assertRaises(lote.PararLote):
            operacao.passo("Serviço", lambda: self.falhar_rede())
        self.assertIn("Falha de conexão", self.eventos[-1]["mensagem"])

    def test_avanco_de_pessoas_aguarda_o_formulario_de_servico_visivel(self):
        self.pagina.set_content(
            PESSOAS + """
            <button id="btnAvancar" onclick="
              document.body.innerHTML = '<select id=ServicoPrestado_CodigoTributacaoNacional></select>'
                + '<textarea id=ServicoPrestado_Descricao hidden></textarea>';
              setTimeout(() => document.querySelector('textarea').hidden = false, 300);
            ">Avançar</button>
            """
        )
        lote.pessoas.avancar_pessoas(self.pagina)
        self.assertTrue(self.pagina.locator(lote.servico.CAMPO_DESCRICAO).is_visible())

    def test_servico_reverifica_tela_sem_pausar_quando_formulario_demora(self):
        self.pagina.set_content(
            '<textarea id="ServicoPrestado_Descricao" hidden></textarea>'
            '<script>setTimeout(() => document.querySelector("textarea").hidden = false, 350)</script>'
        )
        operacao = self.criar_lote(self.parar)
        tela = self.pagina.locator(lote.servico.CAMPO_DESCRICAO)
        with patch.object(lote, "ESPERA_TELA_MS", 200):
            operacao.passo("Serviço", lambda: tela.fill("Descrição"), tela=tela)

        self.assertEqual(tela.input_value(), "Descrição")
        self.assertTrue(any("Aguardando a tela" in e["mensagem"] for e in self.eventos))
        self.assertFalse(any(e["etapa"] == "Pausa" for e in self.eventos))

    def test_avanco_nao_repete_clique_quando_destino_aparece_apos_timeout(self):
        self.pagina.set_content(PESSOAS + '<textarea id="ServicoPrestado_Descricao" hidden></textarea>')
        operacao = self.criar_lote(self.parar)
        tela = self.pagina.locator(lote.CAMPO_PESSOAS)
        destino = self.pagina.locator(lote.servico.CAMPO_DESCRICAO)
        envios = []

        def avancar():
            envios.append(True)
            self.pagina.evaluate("() => {setTimeout(() => document.querySelector('textarea').hidden = false, 100);}")
            raise TimeoutError("Resposta do avanço ainda não chegou.")

        with patch.object(lote, "INTERVALO_TENTATIVA_S", 0.25):
            operacao.passo("Avançar de Pessoas", avancar, tela=tela, pronto=destino)

        self.assertEqual(len(envios), 1)
        self.assertTrue(destino.is_visible())

    def preparar_lista_dependente(self, seletor, origem, codigo, *, prefixo_visual=""):
        self.pagina.evaluate(
            """({seletor, origem, codigo, prefixoVisual}) => {
              const campo = document.querySelector(seletor);
              const menu = campo.nextElementSibling.querySelector('.chosen-single');
              campo.disabled = true;
              window.jQuery = (el) => ({trigger: (evento) => {
                if (evento === 'chosen:updated') menu.textContent = prefixoVisual + el.selectedOptions[0].text;
              }});
              window.selecoesDependentes = 0;
              campo.addEventListener('change', () => window.selecoesDependentes++);
              document.querySelector(origem).addEventListener('change', () => {
                setTimeout(() => {
                  campo.add(new Option(codigo + ' - opção correta', codigo));
                  campo.disabled = false;
                }, 250);
              }, {once: true});
            }""",
            {"seletor": seletor, "origem": origem, "codigo": codigo, "prefixoVisual": prefixo_visual},
        )

    def conferir_lista_dependente(self, seletor, codigo, *, prefixo_visual=""):
        self.assertEqual(self.pagina.locator(seletor).input_value(), codigo)
        menu = self.pagina.locator(f"{seletor} + .chosen-container .chosen-single")
        self.assertEqual(menu.inner_text(), f"{prefixo_visual}{codigo} - opção correta")
        self.assertEqual(self.pagina.evaluate("window.selecoesDependentes"), 1)

    def test_servico_aguarda_indicador_carregado_apos_nbs(self):
        self.pagina.set_content(
            """
            <select id="LocalPrestacao_CodigoMunicipioPrestacao"><option>Lagoa Santa/MG</option></select>
            <select id="ServicoPrestado_CodigoTributacaoNacional"><option>08.02.01 - Instrução</option></select>
            <label><input type="radio" name="ServicoPrestado.HaExportacaoImunidadeNaoIncidencia">Não</label>
            <select id="ServicoPrestado_CodigoNBS">
              <option>Selecione</option><option>122051300 - Educação</option>
            </select>
            <select id="ServicoPrestado_CodigoIndOp" hidden><option>Selecione</option></select>
            <div class="chosen-container"><button class="chosen-single">Selecione</button></div>
            <textarea id="ServicoPrestado_Descricao"></textarea>
            """
        )
        seletor = "#ServicoPrestado_CodigoIndOp"
        self.preparar_lista_dependente(seletor, "#ServicoPrestado_CodigoNBS", "100301")
        lote.servico.preencher_servico(self.pagina, "Descrição de teste")
        self.conferir_lista_dependente(seletor, "100301")

    def test_valores_aguarda_classificacao_carregada_apos_cst(self):
        self.pagina.set_content(
            """
            <input id="Valores_ValorServico">
            <label><input type="radio" name="ISSQN.HaSuspensao">Não</label>
            <label><input type="radio" name="ISSQN.HaRetencao">Não</label>
            <label><input type="radio" name="ISSQN.HaBeneficioMunicipal">Não</label>
            <select id="TributacaoFederal_PISCofins_SituacaoTributaria"><option>00 - Nenhum</option></select>
            <select id="TributacaoFederal_PISCofins_TipoRetencao"><option>0 - Sem retenção</option></select>
            <select id="ValorTributos_CodigoSituacaoTributaria">
              <option>Selecione</option><option>000 - Tributação integral</option>
            </select>
            <select id="ValorTributos_CodigoClassificacaoTributaria" hidden><option>Selecione</option></select>
            <div class="chosen-container">
              <button class="chosen-single"
                onclick="this.parentElement.classList.add('chosen-with-drop'); this.nextElementSibling.hidden=false"
              >Selecione</button>
              <ul class="chosen-results" hidden>
                <li class="active-result" onclick="
                  const campo = document.querySelector('#ValorTributos_CodigoClassificacaoTributaria');
                  campo.value = '000001';
                  campo.dispatchEvent(new Event('change'));
                  this.closest('.chosen-container').querySelector('button').textContent = this.textContent;
                ">000001 - opção correta</li>
              </ul>
            </div>
            <label><input type="radio" name="ValorTributos.TipoValorTributos">Informar alíquota do Simples Nacional</label>
            <input id="ValorTributos_AliquotaSN">
            """
        )
        seletor = "#ValorTributos_CodigoClassificacaoTributaria"
        self.preparar_lista_dependente(seletor, "#ValorTributos_CodigoSituacaoTributaria", "000001")
        lote.valores.preencher_valores(self.pagina, Decimal("1.00"), Decimal("4.16"))
        self.conferir_lista_dependente(seletor, "000001")

    @staticmethod
    def falhar_rede():
        raise Error("net::ERR_INTERNET_DISCONNECTED")

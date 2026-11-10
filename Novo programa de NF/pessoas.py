import re
from collections.abc import Callable
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import Page, expect

from planilha import validar_linha


CEP_EMITENTE = "33230190"
REGIME_SIMPLES = re.compile(
    r"^\s*Regime de apuração dos tributos federais e municipa(?:l|is) pelo Simples Nacional\s*$",
    re.IGNORECASE,
)


def ler_dados_pessoas(caminho_excel: Path, linha: int, *, aceitar_cnpj: bool = False) -> tuple[str, str]:
    planilha = load_workbook(caminho_excel, read_only=True, data_only=True)
    try:
        aba = planilha.active
        data = aba["G7"].value
        documento, _ = validar_linha(aba, linha, aceitar_cnpj=aceitar_cnpj)
    finally:
        planilha.close()
    if not isinstance(data, date):
        raise ValueError("G7 da aba ativa deve conter uma data do Excel.")

    return data.strftime("%d/%m/%Y"), documento


def preencher_data_competencia(pagina: Page, data: str) -> None:
    campo = pagina.locator("input#DataCompetencia:visible")
    try:
        if campo.get_attribute("type") == "date":
            iso = datetime.strptime(data, "%d/%m/%Y").date().isoformat()
            campo.fill(iso)
            campo.blur()
            expect(campo).to_have_value(iso)
            return

        campo.fill("")
        campo.press_sequentially(data.replace("/", ""))
        campo.blur()
        if campo.input_value() != data:
            campo.fill(data)
            campo.blur()
        expect(campo).to_have_value(data)
    except AssertionError as erro:
        raise RuntimeError(f"O portal não manteve a data de competência {data}.") from erro


def marcar_opcao(pagina: Page, grupo: str, texto: str) -> None:
    radios = pagina.locator(f'input[type="radio"][name="{grupo}"]')
    radios.first.wait_for(state="attached")
    correspondencias = radios.evaluate_all(
        """(radios, texto) => {
            const rotulos = [...document.querySelectorAll('label')];
            return radios.flatMap((radio, indice) =>
                [...radio.labels]
                    .filter(label => label.textContent.trim() === texto)
                    .map(label => [indice, rotulos.indexOf(label)])
            );
        }""",
        texto,
    )
    if len(correspondencias) != 1:
        raise RuntimeError(f"Não foi possível identificar a opção {texto} de {grupo}.")
    indice_radio, indice_rotulo = correspondencias[0]
    opcao = radios.nth(indice_radio)
    if not opcao.is_checked():
        label = pagina.locator("label").nth(indice_rotulo)
        if label.is_visible():
            label.click()
        else:
            opcao.check()
    expect(opcao).to_be_checked()


def selecionar_regime_simples(pagina: Page) -> None:
    campo = pagina.locator("#SimplesNacional_RegimeApuracaoTributosSN")
    opcao = campo.locator("option").filter(has_text=REGIME_SIMPLES)
    try:
        expect(opcao).to_have_count(1)
        # Alguns menus exibem um controle visual sobre o select nativo.
        campo.select_option(label=opcao.text_content().strip(), force=not campo.is_visible())
        expect(opcao.and_(pagina.locator("option:checked"))).to_have_count(1)
    except AssertionError as erro:
        raise RuntimeError("O portal não confirmou o regime federal e municipal do Simples Nacional.") from erro


def preencher_cep_emitente(pagina: Page) -> None:
    """Abre os detalhes do emitente e preenche o CEP na etapa Pessoas."""
    campo_cep = pagina.locator("#Prestador_EnderecoNacional_CEP")
    if not campo_cep.is_visible():
        detalhes = pagina.locator("#btnMaisInfoEmitente")
        try:
            expect(detalhes).to_be_enabled(timeout=5000)
        except AssertionError as erro:
            raise RuntimeError(
                "Exibir detalhes do emitente está desabilitado. "
                "Confira IBS/CBS e a data de competência no portal."
            ) from erro
        detalhes.click()

    campo_cep.fill(CEP_EMITENTE)
    campo_cep.blur()
    cep_digitado = campo_cep.input_value()
    if "".join(filter(str.isdigit, cep_digitado)) != CEP_EMITENTE:
        raise RuntimeError("O portal não manteve o CEP 33230-190 após o preenchimento.")


def pesquisar_tomador(pagina: Page, documento: str) -> None:
    marcar_opcao(pagina, "Tomador.LocalDomicilio", "Brasil")
    campo = pagina.locator("#Tomador_Inscricao")
    campo.fill(documento)
    campo.blur()
    if re.sub(r"\D", "", campo.input_value()) != documento:
        campo.fill("")
        campo.press_sequentially(documento)
        campo.blur()
    if re.sub(r"\D", "", campo.input_value()) != documento:
        raise RuntimeError("O portal não manteve o documento completo do tomador.")
    pagina.locator("#btn_Tomador_Inscricao_pesquisar").click()
    expect(pagina.locator("#Tomador_Nome")).to_have_value(re.compile(r"\S"), timeout=10000)


def preencher_pessoas(
    pagina: Page,
    data: str,
    documento: str,
    informar: Callable[[str], None] | None = None,
    *,
    retomar_apos_data: bool = False,
) -> None:
    if retomar_apos_data:
        data_na_tela = pagina.locator("input#DataCompetencia:visible").input_value()
        if not data_na_tela.strip():
            raise RuntimeError("Preencha a competência no portal antes de retomar.")
        if informar:
            informar(f"Competência preservada no portal: {data_na_tela}")
    passos = (
        ("IBS/CBS: Sim", marcar_opcao, (pagina, "PreencherInfoIBSCBS", "Sim")),
        ("Data de competência", preencher_data_competencia, (pagina, data)),
        ("Regime do Simples Nacional", selecionar_regime_simples, (pagina,)),
        ("Detalhes e CEP do emitente", preencher_cep_emitente, (pagina,)),
        ("Compras Governamentais: Não", marcar_opcao, (pagina, "EhCompraGovernamental", "Não")),
        ("Pesquisa do CPF/CNPJ do tomador", pesquisar_tomador, (pagina, documento)),
        ("Destinatário do Serviço: Sim", marcar_opcao, (pagina, "DestinatarioEhOAdquirente", "Sim")),
    )
    for descricao, acao, argumentos in passos[2:] if retomar_apos_data else passos:
        if informar:
            informar(descricao)
        acao(*argumentos)
    if not retomar_apos_data:
        campo_data = pagina.locator("input#DataCompetencia:visible")
        esperado = datetime.strptime(data, "%d/%m/%Y").date().isoformat() if (
            campo_data.get_attribute("type") == "date"
        ) else data
        if campo_data.input_value() != esperado:
            raise RuntimeError(f"O portal alterou a competência de G7 ({data}). Confira antes de avançar.")


def avancar_pessoas(pagina: Page) -> None:
    expect(pagina.locator("#DataCompetencia")).to_be_visible()
    pagina.locator("#btnAvancar").click()
    pagina.locator("#ServicoPrestado_Descricao").wait_for(state="visible", timeout=15000)

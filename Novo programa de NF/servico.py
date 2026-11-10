import re
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import Error, Page, expect

from pessoas import marcar_opcao


MUNICIPIO = "Lagoa Santa/MG"
CODIGO_TRIBUTACAO = "08.02.01"
CODIGO_NBS = "122051300"
CODIGO_INDICADOR = "100301"
OPCAO_TRIBUTACAO = re.compile(r"^\s*08\.02\.01\s*-")
CAMPO_MUNICIPIO = "#LocalPrestacao_CodigoMunicipioPrestacao"
CAMPO_TRIBUTACAO = "#ServicoPrestado_CodigoTributacaoNacional"
CAMPO_DESCRICAO = "#ServicoPrestado_Descricao"
TEMPO_LISTA_MS = 15000


def ler_descricao_servico(caminho_excel: Path) -> str:
    planilha = load_workbook(caminho_excel, read_only=True, data_only=True)
    try:
        descricao = planilha.active["G2"].value
    finally:
        planilha.close()
    if not isinstance(descricao, str) or not descricao.strip():
        raise ValueError("G2 da aba ativa deve conter a descrição do serviço.")
    return descricao


def selecionar_resultado(
    pagina: Page, seletor: str, busca: str, texto: re.Pattern[str], *, nativo: bool = False
) -> None:
    campo = pagina.locator(seletor)
    campo.wait_for(state="attached")
    selecionado = campo.locator("option:checked").filter(has_text=texto)
    if selecionado.count() == 1:
        return

    try:
        if nativo:
            selecionar_opcao_nativa(pagina, seletor, texto)
        else:
            escolher_opcao(pagina, seletor, busca, texto)
        expect(selecionado).to_have_count(1, timeout=TEMPO_LISTA_MS)
    except (Error, AssertionError) as erro:
        raise RuntimeError(
            f"O portal não confirmou a opção {busca} em {seletor}. "
            f"{str(erro).splitlines()[0]}"
        ) from erro


def escolher_opcao(pagina: Page, seletor: str, busca: str, texto: re.Pattern[str]) -> None:
    campo = pagina.locator(seletor)
    expect(campo).to_be_enabled(timeout=TEMPO_LISTA_MS)
    menu = pagina.locator(f"{seletor} + .select2 [role=combobox]")
    chosen = pagina.locator(f"{seletor} + .chosen-container")
    opcao = campo.locator("option").filter(has_text=texto)
    if menu.is_visible():
        if menu.get_attribute("aria-expanded") != "true":
            menu.click(timeout=TEMPO_LISTA_MS)
        pesquisa = pagina.locator(".select2-container--open input.select2-search__field")
        pesquisa.fill(busca, timeout=TEMPO_LISTA_MS)
        resultado = pagina.locator(
            ".select2-container--open .select2-results__option"
            ":not([aria-disabled='true']):not(.loading-results)"
        ).filter(
            has_text=texto, visible=True,
        )
        resultado.click(timeout=TEMPO_LISTA_MS)
    elif chosen.is_visible():
        expect(opcao).to_have_count(1, timeout=TEMPO_LISTA_MS)
        if "chosen-with-drop" not in (chosen.get_attribute("class") or ""):
            chosen.locator(".chosen-single").click(timeout=TEMPO_LISTA_MS)
        pesquisa = chosen.locator(".chosen-search input")
        if pesquisa.is_visible() and pesquisa.is_editable():
            pesquisa.fill(busca)
        resultado = chosen.locator(".chosen-results li.active-result").filter(
            has_text=texto, visible=True
        )
        resultado.click(timeout=TEMPO_LISTA_MS)
    else:
        expect(opcao).to_have_count(1, timeout=TEMPO_LISTA_MS)
        campo.select_option(label=opcao.text_content().strip(), timeout=TEMPO_LISTA_MS)


def selecionar_opcao_nativa(pagina: Page, seletor: str, texto: re.Pattern[str]) -> None:
    # O menu Chosen pode fechar enquanto o portal atualiza uma lista dependente.
    campo = pagina.locator(seletor)
    opcao = campo.locator("option:enabled").filter(has_text=texto)
    expect(campo).to_be_enabled(timeout=TEMPO_LISTA_MS)
    expect(opcao).to_have_count(1, timeout=TEMPO_LISTA_MS)
    campo.select_option(label=opcao.text_content().strip(), force=True, timeout=TEMPO_LISTA_MS)

    menu = pagina.locator(f"{seletor} + .chosen-container .chosen-single")
    if menu.is_visible():
        campo.evaluate("el => window.jQuery(el).trigger('chosen:updated')")
        # A confirmação é feita na opção do formulário, que o portal envia.
        # O texto do menu pode ter uma apresentação diferente da opção.


def preencher_servico(pagina: Page, descricao: str, informar=None) -> None:
    if informar:
        informar("Selecionando município de prestação.")
    selecionar_resultado(
        pagina, CAMPO_MUNICIPIO, MUNICIPIO, re.compile(r"^\s*Lagoa Santa\s*/\s*MG\s*$")
    )
    if informar:
        informar("Selecionando código de tributação e incidência de ISSQN.")
    selecionar_resultado(pagina, CAMPO_TRIBUTACAO, CODIGO_TRIBUTACAO, OPCAO_TRIBUTACAO)
    marcar_opcao(pagina, "ServicoPrestado.HaExportacaoImunidadeNaoIncidencia", "Não")
    if informar:
        informar(f"Selecionando NBS {CODIGO_NBS}.")
    selecionar_resultado(
        pagina, "#ServicoPrestado_CodigoNBS", CODIGO_NBS,
        re.compile(rf"^\s*{CODIGO_NBS}\s*-"), nativo=True,
    )
    if informar:
        informar(f"Selecionando indicador da operação {CODIGO_INDICADOR}.")
    selecionar_resultado(
        pagina, "#ServicoPrestado_CodigoIndOp", CODIGO_INDICADOR,
        re.compile(rf"^\s*{CODIGO_INDICADOR}\s*-"), nativo=True,
    )
    if informar:
        informar("Preenchendo a descrição de G2.")
    campo_descricao = pagina.locator(CAMPO_DESCRICAO)
    campo_descricao.fill(descricao)
    expect(campo_descricao).to_have_value(descricao.replace("\r\n", "\n").replace("\r", "\n"))


def avancar_servico(pagina: Page) -> None:
    campo_descricao = pagina.locator(CAMPO_DESCRICAO)
    expect(campo_descricao).to_be_visible()
    try:
        pagina.get_by_role("button", name="Avançar", exact=True).click(timeout=TEMPO_LISTA_MS)
        expect(pagina.locator("#Valores_ValorServico")).to_be_visible(timeout=TEMPO_LISTA_MS)
    except (Error, AssertionError) as erro:
        raise RuntimeError("Serviço: o portal não confirmou o avanço para Valores.") from erro

import re
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import Locator, Page, expect

from pessoas import marcar_opcao
from planilha import validar_linha
from servico import selecionar_resultado


CAMPO_VALOR = "#Valores_ValorServico"
CAMPO_ALIQUOTA_SN = "#ValorTributos_AliquotaSN"


def ler_dados_valores(caminho_excel: Path, linha: int, *, aceitar_cnpj: bool = False) -> tuple[Decimal, Decimal]:
    planilha = load_workbook(caminho_excel, read_only=True, data_only=True)
    try:
        aba = planilha.active
        _, valor = validar_linha(aba, linha, aceitar_cnpj=aceitar_cnpj)
        celula = aba["G4"]
        bruto = celula.value
        if type(bruto) in (int, float):
            aliquota = Decimal(str(bruto))
            if "%" in celula.number_format:
                aliquota *= 100
        elif isinstance(bruto, str) and re.fullmatch(r"\d+(?:[,.]\d+)?\s*%?", bruto.strip()):
            aliquota = Decimal(bruto.strip().rstrip("%").strip().replace(",", "."))
        else:
            raise ValueError("G4 da aba ativa deve conter a alíquota do Simples Nacional.")
        if not aliquota.is_finite() or not 0 <= aliquota <= 100:
            raise ValueError("G4 da aba ativa deve conter uma alíquota entre 0 e 100%.")
        return valor, aliquota
    except InvalidOperation as erro:
        raise ValueError("G4 da aba ativa contém uma alíquota inválida.") from erro
    finally:
        planilha.close()


def preencher_decimal(campo: Locator, valor: Decimal, casas: int = 2) -> None:
    formatado = f"{valor:,.{casas}f}".translate(str.maketrans({",": ".", ".": ","}))
    sem_milhar = formatado.replace(".", "")
    esperado = re.compile(rf"^\s*(?:{re.escape(formatado)}|{re.escape(sem_milhar)})\s*$")

    campo.fill(formatado)
    campo.blur()
    if not esperado.fullmatch(campo.input_value()) and valor >= 0:
        campo.fill("")
        campo.press_sequentially(sem_milhar.replace(",", ""))
        campo.blur()
    expect(campo).to_have_value(esperado)


def preencher_valores(pagina: Page, valor: Decimal, aliquota: Decimal) -> None:
    preencher_decimal(pagina.locator(CAMPO_VALOR), valor)
    marcar_opcao(pagina, "ISSQN.HaSuspensao", "Não")
    marcar_opcao(pagina, "ISSQN.HaRetencao", "Não")
    marcar_opcao(pagina, "ISSQN.HaBeneficioMunicipal", "Não")

    for seletor, codigo in (
        ("#TributacaoFederal_PISCofins_SituacaoTributaria", "00"),
        ("#TributacaoFederal_PISCofins_TipoRetencao", "0"),
        ("#ValorTributos_CodigoSituacaoTributaria", "000"),
        ("#ValorTributos_CodigoClassificacaoTributaria", "000001"),
    ):
        selecionar_resultado(pagina, seletor, codigo, re.compile(rf"^\s*{codigo}\s*-"))

    marcar_opcao(pagina, "ValorTributos.TipoValorTributos", "Informar alíquota do Simples Nacional")
    casas_aliquota = max(2, -aliquota.normalize().as_tuple().exponent)
    preencher_decimal(pagina.locator(CAMPO_ALIQUOTA_SN), aliquota, casas_aliquota)


def avancar_valores(pagina: Page) -> None:
    expect(pagina.locator(CAMPO_VALOR)).to_be_visible()
    pagina.get_by_role("button", name="Avançar", exact=True).click()
    botao_emissao = pagina.get_by_role("button", name=re.compile(r"^Emitir NFS-e$", re.IGNORECASE))
    try:
        expect(botao_emissao).to_be_visible()
    except AssertionError as erro:
        campos = pagina.locator("[data-valmsg-for]:visible").evaluate_all(
            """els => [...new Set(els.filter(el => el.textContent.trim())
                .map(el => el.getAttribute('data-valmsg-for')))]"""
        )
        detalhe = f" Confira os campos: {', '.join(campos)}." if campos else ""
        raise RuntimeError(f"O portal não abriu a tela Emitir NFS-e.{detalhe}") from erro

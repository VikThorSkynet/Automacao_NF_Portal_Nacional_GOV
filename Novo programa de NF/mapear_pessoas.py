import json
import logging
from importlib import reload
from pathlib import Path
from threading import Event
from tkinter import Tk, filedialog

from playwright.sync_api import Error, sync_playwright

import pessoas
import servico
import valores
from lote import Lote, PararLote
from planilha import ler_intervalo
from registro import Registro


PORTAL_NFSE = "https://www.nfse.gov.br/EmissorNacional/"


def escolher_excel() -> Path | None:
    janela = Tk()
    janela.withdraw()
    janela.attributes("-topmost", True)
    try:
        caminho = filedialog.askopenfilename(
            parent=janela,
            title="Escolha a planilha Excel para emissão de NFS-e",
            filetypes=[("Planilhas Excel", "*.xlsx *.xlsm")],
        )
    finally:
        janela.destroy()
    return Path(caminho) if caminho else None


def inspecionar_campos(pagina):
    """Lista os controles visíveis, sem ler os valores preenchidos."""
    campos = pagina.locator("input, select, textarea, button, [role=combobox]")
    return campos.evaluate_all(
        """elementos => elementos
            .filter(el => el.type !== 'hidden' && (
                el.getClientRects().length || el.tagName === 'SELECT' ||
                Array.from(el.labels || []).some(label => label.getClientRects().length)
            ))
            .map(el => ({
                tag: el.tagName.toLowerCase(),
                tipo: el.getAttribute('type'),
                id: el.id || null,
                nome: el.getAttribute('name'),
                rotulo: Array.from(el.labels || [], label => label.innerText.trim()),
                opcoes: el.tagName === 'SELECT'
                    ? Array.from(el.options, opcao => opcao.text.trim())
                    : undefined
            }))"""
    )


def listar_campos(pagina):
    for tentativa in range(3):
        try:
            pagina.wait_for_load_state("domcontentloaded")
            return inspecionar_campos(pagina)
        except Error as erro:
            if "Execution context was destroyed" not in str(erro) or tentativa == 2:
                raise


def testar_ate_emissao(
    pagina, caminho_excel: Path, linha: int, apos_data: bool, *, aceitar_cnpj: bool = False
) -> None:
    etapa_pessoas = reload(pessoas)
    etapa_servico = reload(servico)
    etapa_valores = reload(valores)
    data, documento = etapa_pessoas.ler_dados_pessoas(caminho_excel, linha, aceitar_cnpj=aceitar_cnpj)
    descricao = etapa_servico.ler_descricao_servico(caminho_excel)
    valor, aliquota = etapa_valores.ler_dados_valores(caminho_excel, linha, aceitar_cnpj=aceitar_cnpj)
    print(f"Pessoas — linha {linha}.", flush=True)
    etapa_pessoas.preencher_pessoas(
        pagina, data, documento,
        informar=lambda etapa: print(etapa, flush=True),
        retomar_apos_data=apos_data,
    )
    etapa_pessoas.avancar_pessoas(pagina)
    print("Serviço — município, códigos e descrição de G2.", flush=True)
    etapa_servico.preencher_servico(pagina, descricao)
    etapa_servico.avancar_servico(pagina)
    print("Valores — tributos e alíquota de G4.", flush=True)
    etapa_valores.preencher_valores(pagina, valor, aliquota)
    etapa_valores.avancar_valores(pagina)
    print(f"Teste concluído na tela Emitir NFS-e. Alíquota: {aliquota}%.", flush=True)


def main():
    caminho_excel = escolher_excel()
    if caminho_excel is None:
        return
    print(f"Planilha selecionada: {caminho_excel.name}", flush=True)
    intervalo = ler_intervalo(caminho_excel)
    registro = Registro(caminho_excel)
    print(f"Linhas configuradas no Excel: {intervalo.start} a {intervalo.stop - 1}.", flush=True)
    with sync_playwright() as playwright:
        navegador = playwright.chromium.launch(channel="chrome", headless=False, slow_mo=650)
        sessao = navegador.new_context()
        sessao.set_default_timeout(5000)
        sessao.new_page().goto(PORTAL_NFSE, wait_until="domcontentloaded", timeout=30000)
        print("Entre com o certificado e abra Emissão Completa > Pessoas.", flush=True)
        linha_em_teste = None
        emissao_preparada = None
        emissao_tentada = False

        while sessao.pages:
            try:
                comando = input(
                    "'teste N' roda até a tela Emitir NFS-e; 'teste-apos-data N' usa a data manual; "
                    "'teste-cnpj N' prepara CNPJ; 'teste-cnpj-apos-data N' preserva a data; 'emitir N' envia uma nota; "
                    "Enter lista campos; 'pessoas N' testa; 'apos-data N' retoma; "
                    "'avancar' passa de etapa; 'servico' preenche; "
                    "'valores N' preenche a linha N; 'sair' fecha: "
                )
            except EOFError:
                break
            comando = comando.strip().lower()
            if comando == "sair" or not sessao.pages:
                break
            try:
                pagina = sessao.pages[-1]
                if comando.startswith(("teste ", "teste-apos-data ", "teste-cnpj ", "teste-cnpj-apos-data ")):
                    emissao_preparada = None
                    linha = int(comando.split(maxsplit=1)[1])
                    if linha not in intervalo:
                        raise ValueError("A linha está fora do intervalo configurado em G10 e G13.")
                    aceitar_cnpj = comando.startswith(("teste-cnpj ", "teste-cnpj-apos-data "))
                    if aceitar_cnpj and len(intervalo) != 1:
                        raise ValueError("O teste com CNPJ requer uma única linha em G13.")
                    _, dados = registro.ler_linha(linha, lambda aba: None)
                    testar_ate_emissao(
                        pagina, caminho_excel, linha,
                        apos_data=comando.startswith(("teste-apos-data ", "teste-cnpj-apos-data ")),
                        aceitar_cnpj=aceitar_cnpj,
                    )
                    linha_em_teste = linha
                    emissao_preparada = (pagina, linha, dados)
                elif comando.startswith("emitir "):
                    linha = int(comando.split(maxsplit=1)[1])
                    if emissao_tentada:
                        raise ValueError("Já houve uma tentativa. Confira o resultado no portal antes de repetir.")
                    if intervalo != range(linha, linha + 1) or ler_intervalo(caminho_excel) != intervalo:
                        raise ValueError("A emissão deste teste exige G10 na linha escolhida e G13 igual a 1.")
                    if emissao_preparada is None or emissao_preparada[:2] != (pagina, linha):
                        raise ValueError("Prepare a linha nesta mesma aba antes de emitir.")

                    def parar_para_conferencia():
                        raise PararLote("Confira a pendência no portal ou abra a interface para retomar.")

                    operacao = Lote(
                        pagina, caminho_excel,
                        lambda evento: print(evento["mensagem"], flush=True),
                        parar_para_conferencia, Event(), logging.getLogger("nfse.depuracao"),
                    )
                    operacao.linha = linha
                    with registro.exclusivo():
                        if registro.ler_pendente() is not None:
                            raise ValueError("Já existe uma emissão pendente. Confira antes de emitir.")
                        emissao_tentada = True
                        operacao.enviar_nota(emissao_preparada[2])
                elif comando.startswith(("pessoas ", "apos-data ")):
                    emissao_preparada = None
                    linha = int(comando.split(maxsplit=1)[1])
                    reload(pessoas)
                    data, cpf = pessoas.ler_dados_pessoas(caminho_excel, linha)
                    print(f"Linha {linha} do Excel. Competência: {data}.", flush=True)
                    pessoas.preencher_pessoas(
                        pagina, data, cpf,
                        informar=lambda etapa: print(etapa, flush=True),
                        retomar_apos_data=comando.startswith("apos-data "),
                    )
                    linha_em_teste = linha
                    print("Preenchimento concluído. Confira a tela antes de Avançar.", flush=True)
                elif comando == "avancar":
                    emissao_preparada = None
                    if pagina.locator(valores.CAMPO_VALOR).is_visible():
                        reload(valores).avancar_valores(pagina)
                        print("Página Conferência aberta.", flush=True)
                    elif pagina.locator(servico.CAMPO_DESCRICAO).is_visible():
                        reload(servico).avancar_servico(pagina)
                        print("Página Serviço avançada.", flush=True)
                    else:
                        reload(pessoas).avancar_pessoas(pagina)
                        print("Página Serviço aberta.", flush=True)
                elif comando == "servico":
                    emissao_preparada = None
                    etapa_servico = reload(servico)
                    descricao = etapa_servico.ler_descricao_servico(caminho_excel)
                    etapa_servico.preencher_servico(pagina, descricao)
                    print("Serviço preenchido. Confira antes de Avançar.", flush=True)
                elif comando.startswith("valores "):
                    emissao_preparada = None
                    linha = int(comando.split(maxsplit=1)[1])
                    if linha_em_teste is not None and linha != linha_em_teste:
                        raise ValueError(f"Use a linha {linha_em_teste}, já preenchida em Pessoas.")
                    etapa_valores = reload(valores)
                    valor, aliquota = etapa_valores.ler_dados_valores(caminho_excel, linha)
                    etapa_valores.preencher_valores(pagina, valor, aliquota)
                    print(f"Valores da linha {linha} preenchidos. Confira antes de Avançar.", flush=True)
                elif not comando:
                    print(json.dumps(listar_campos(pagina), ensure_ascii=False, indent=2), flush=True)
            except (Error, RuntimeError, AssertionError, ValueError, OSError, PararLote) as erro:
                print(f"Preenchimento interrompido: {str(erro).splitlines()[0]}", flush=True)
                if not pagina.is_closed():
                    aviso = pagina.locator(".jconfirm.jconfirm-open .jconfirm-content").first
                    if aviso.is_visible():
                        print(f"Aviso do portal: {aviso.inner_text().strip()}", flush=True)


if __name__ == "__main__":
    main()

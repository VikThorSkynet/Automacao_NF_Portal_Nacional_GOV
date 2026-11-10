import re
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook


def ler_intervalo(caminho_excel: Path) -> range:
    planilha = load_workbook(caminho_excel, read_only=True, data_only=True)
    try:
        inicio = planilha.active["G10"].value
        quantidade = planilha.active["G13"].value
    finally:
        planilha.close()
    if type(inicio) is not int or not 2 <= inicio <= 1_048_576:
        raise ValueError("G10 deve conter uma linha inicial inteira a partir de 2.")
    if type(quantidade) is not int or not 1 <= quantidade <= 1_048_577 - inicio:
        raise ValueError("G13 deve conter uma quantidade inteira positiva dentro dos limites do Excel.")
    return range(inicio, inicio + quantidade)


def validar_linha(aba, linha: int, *, aceitar_cnpj: bool = False) -> tuple[str, Decimal]:
    """Valida os dados antes de usar uma linha da planilha para uma nota."""
    if linha < 2:
        raise ValueError("Os dados da planilha começam na linha 2.")

    numero_nf = aba[f"E{linha}"].value
    erro = aba[f"F{linha}"].value
    if numero_nf is not None and str(numero_nf).strip():
        raise ValueError(f"Linha {linha} ignorada: coluna E já preenchida.")
    if isinstance(erro, str) and erro.strip().upper() == "ERRO":
        raise ValueError(f"Linha {linha} ignorada: coluna F marcada com Erro.")

    valor_cpf = aba[f"C{linha}"].value
    if isinstance(valor_cpf, int) and valor_cpf > 0:
        documento = str(valor_cpf).zfill(11)
    else:
        pontuacao = r"[./\s-]" if aceitar_cnpj else r"[.\s-]"
        documento = re.sub(pontuacao, "", str(valor_cpf or ""))
    formato = r"(?:[0-9]{11}|[0-9]{14})" if aceitar_cnpj else r"[0-9]{11}"
    if not re.fullmatch(formato, documento):
        tipo = "CPF ou CNPJ" if aceitar_cnpj else "CPF de 11 dígitos"
        raise ValueError(f"Linha {linha} ignorada: C{linha} sem {tipo}.")

    celula = aba[f"D{linha}"]
    preenchimento = getattr(celula, "fill", None)
    if preenchimento and preenchimento.patternType == "solid":
        cor = preenchimento.start_color
        if cor.type == "rgb" and cor.rgb[-6:].upper() == "FFFF00":
            raise ValueError(f"Linha {linha} ignorada: D{linha} está em amarelo.")

    if type(celula.value) not in (int, float):
        raise ValueError(f"Linha {linha} ignorada: D{linha} precisa ser um número do Excel.")
    valor = Decimal(str(celula.value))
    if not valor.is_finite():
        raise ValueError(f"Linha {linha} ignorada: D{linha} contém um valor inválido.")
    return documento, valor

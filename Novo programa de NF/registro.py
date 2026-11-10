import json
import msvcrt
import os
import re
from contextlib import contextmanager
from hashlib import sha256
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import load_workbook


CHAVE_NFSE = re.compile(r"[0-9]{50}")


@contextmanager
def arquivo_temporario(destino: Path):
    with NamedTemporaryFile(dir=destino.parent, suffix=destino.suffix, delete=False) as arquivo:
        temporario = Path(arquivo.name)
    try:
        yield temporario
    finally:
        temporario.unlink(missing_ok=True)


class Registro:
    def __init__(self, excel: Path):
        self.excel = excel.resolve()
        self.pendente = self.excel.with_suffix(self.excel.suffix + ".nfse-pendente.json")
        self.trava = self.excel.with_suffix(self.excel.suffix + ".nfse.lock")

    @contextmanager
    def exclusivo(self):
        with self.trava.open("a+b") as arquivo:
            if arquivo.tell() == 0:
                arquivo.write(b"\0")
                arquivo.flush()
            arquivo.seek(0)
            try:
                msvcrt.locking(arquivo.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError as erro:
                raise RuntimeError("Já existe outro lote usando esta planilha.") from erro
            try:
                yield
            finally:
                arquivo.seek(0)
                msvcrt.locking(arquivo.fileno(), msvcrt.LK_UNLCK, 1)

    def ler_pendente(self):
        return json.loads(self.pendente.read_text(encoding="utf-8")) if self.pendente.exists() else None

    def salvar_pendente(self, dados):
        with arquivo_temporario(self.pendente) as temporario:
            temporario.write_text(json.dumps(dados), encoding="utf-8")
            os.replace(temporario, self.pendente)

    def limpar_pendente(self):
        self.pendente.unlink(missing_ok=True)

    def ler_linha(self, linha, preparar):
        planilha = load_workbook(self.excel, read_only=True, data_only=True)
        try:
            aba = planilha.active
            return preparar(aba), self.identidade(aba, linha)
        finally:
            planilha.close()

    @staticmethod
    def identidade(aba, linha):
        campos = [aba[f"{coluna}{linha}"].value for coluna in "BCD"]
        assinatura = sha256(repr(campos).encode("utf-8")).hexdigest()
        return {"aba": aba.title, "linha": linha, "assinatura": assinatura}

    def conferir_linha(self, dados, *, aceitar_gravada=False):
        linha = dados["linha"]
        _, atual = self.ler_linha(
            linha, lambda aba: self.conferir_celulas(aba, dados, aceitar_gravada)
        )
        if any(atual[campo] != dados[campo] for campo in atual):
            raise RuntimeError(f"Os dados da linha {linha} mudaram no Excel.")

    def conferir_celulas(self, aba, dados, aceitar_gravada):
        linha = dados["linha"]
        atual = aba[f"E{linha}"].value
        if atual not in (None, ""):
            chave = dados.get("chave")
            texto = str(atual).lstrip("'")
            if aceitar_gravada and CHAVE_NFSE.fullmatch(texto) and chave in (None, texto):
                return texto
            raise RuntimeError(f"E{linha} já está preenchida. Confira a nota no Excel.")
        if str(aba[f"F{linha}"].value).strip().upper() == "ERRO":
            raise RuntimeError(f"F{linha} já está marcada com Erro.")
        return None

    def reconciliar(self, dados):
        linha = dados["linha"]
        atual, identidade = self.ler_linha(linha, lambda aba: self.conferir_celulas(aba, dados, True))
        if any(identidade[campo] != dados[campo] for campo in identidade):
            raise RuntimeError(f"Os dados da linha {linha} mudaram no Excel.")
        if atual is None:
            if not dados.get("chave"):
                raise RuntimeError(f"Emissão pendente na linha {linha}. Confira no portal e grave a chave em E{linha}.")
            self.gravar(dados, dados["chave"])
        self.limpar_pendente()

    def gravar(self, dados, chave=None):
        linha = dados["linha"]
        if chave is not None and not CHAVE_NFSE.fullmatch(chave):
            raise ValueError("A chave de acesso deve ter 50 dígitos.")
        original = self.excel.stat()
        self.conferir_linha(dados)
        planilha = load_workbook(self.excel, keep_vba=self.excel.suffix.lower() == ".xlsm")
        coluna = "E" if chave is not None else "F"
        esperado = chave if chave is not None else "Erro"
        try:
            aba = planilha[dados["aba"]]
            celula = aba[f"{coluna}{linha}"]
            if celula.value not in (None, "", esperado):
                raise RuntimeError(f"{celula.coordinate} mudou no Excel.")
            celula.value = esperado
            if chave is not None:
                celula.number_format = "@"
                celula.quotePrefix = True
            with arquivo_temporario(self.excel) as temporario:
                planilha.save(temporario)
                salvo = load_workbook(temporario, read_only=True)
                try:
                    confirmado = salvo[dados["aba"]][f"{coluna}{linha}"]
                    if confirmado.value != esperado or confirmado.data_type != "s":
                        raise RuntimeError(f"Não foi possível verificar {coluna}{linha}.")
                finally:
                    salvo.close()
                if self.excel.stat() != original:
                    raise RuntimeError("A planilha mudou durante a gravação. Feche o Excel e confira.")
                os.replace(temporario, self.excel)
        finally:
            planilha.close()

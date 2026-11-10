# Automacao_NF_Portal_Nacional_GOV

Para abrir a janela de acompanhamento no Windows, com Google Chrome instalado:

```powershell
python -m pip install openpyxl playwright
python "Novo programa de NF/interface.py"
```

Escolha a planilha e confira o intervalo de G10/G13. **Testar até revisão** prepara a
primeira linha elegível. **Emitir lote** envia as notas do intervalo. Ao abrir o Chrome,
entre com o certificado, abra **Emissão Completa** e clique em **Portal pronto**.

A chave de acesso confirmada é gravada como texto em E, e uma rejeição definitiva do
documento do tomador é marcada em F. Feche a planilha no Excel durante a gravação.
Em caso de pausa, use **Retomar** após resolver o motivo exibido. Se a tela tiver perdido
o rascunho, encerre e reabra o lote. O programa pula linhas com E preenchida.

Se uma emissão ficar incerta, confira a nota no portal e salve a chave correta em E
antes de retomar. O arquivo `.nfse-pendente.json` junto da planilha guarda o estado
da tentativa. A janela mostra o progresso e abre o log de cada execução.

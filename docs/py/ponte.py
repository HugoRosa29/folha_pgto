"""
Ponte entre o navegador e o núcleo da conferência.

Carregado pelo Pyodide (ver docs/app.js). O JavaScript entrega
os arquivos como bytes e recebe o resultado como texto JSON,
evitando conversões automáticas entre Python e JavaScript.

O resultado completo fica guardado aqui para que a exportação
em Excel não precise reprocessar os arquivos.
"""

from decimal import Decimal
import json

from exportar import montar_planilha
from folha import comparar


# Último resultado processado, usado pela exportação.
ULTIMO = {"resultado": None}


class Arquivo:
    """
    Embala o conteúdo enviado pelo navegador no formato que as
    funções de leitura esperam (nome + .read()).
    """

    def __init__(self, nome, conteudo):
        self.filename = nome or ""
        self._conteudo = bytes(conteudo)

    def read(self):
        return self._conteudo


def _json_seguro(objeto):
    if isinstance(objeto, Decimal):
        return float(objeto)

    return str(objeto)


def para_tela(resultado):
    """
    A tela recebe tudo menos a lista de lançamentos fora do
    portal, que é longa e só interessa no Excel.
    """

    return {
        chave: valor
        for chave, valor in resultado.items()
        if chave != "lancamentos_fora"
    }


def processar(arquivos):
    """
    Recebe {"txts": [{nome, dados}], "folha": {nome, dados}} e
    devolve o resultado em JSON.

    Em caso de erro de leitura, devolve {"erro": "..."} em vez de
    levantar exceção, para a tela mostrar uma mensagem clara.
    """

    txts = [
        Arquivo(item.get("nome"), item.get("dados"))
        for item in (arquivos.get("txts") or [])
    ]

    folha = arquivos.get("folha")
    folha = Arquivo(folha.get("nome"), folha.get("dados")) if folha else None

    if not txts or folha is None:
        return json.dumps(
            {"erro": "Envie os TXT do portal e a planilha da Folha."},
            ensure_ascii=False,
        )

    try:
        resultado = comparar(txts, folha)
    except Exception as erro:
        return json.dumps({
            "erro": (
                "Não foi possível processar os arquivos: "
                f"{type(erro).__name__}: {erro}"
            )
        }, ensure_ascii=False)

    ULTIMO["resultado"] = resultado

    return json.dumps(para_tela(resultado), default=_json_seguro, ensure_ascii=False)


def exportar():
    """
    Devolve os bytes do .xlsx da última conferência processada.
    """

    resultado = ULTIMO["resultado"]

    if resultado is None:
        raise RuntimeError("Nenhuma conferência foi processada ainda.")

    return montar_planilha(resultado)

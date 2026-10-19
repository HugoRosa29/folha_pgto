"""
Exportação da conferência da Folha para Excel.

Usa xlsxwriter direto (sem pandas), para que o mesmo código
gere o arquivo na linha de comando e no navegador, via Pyodide.
"""

from decimal import Decimal
from io import BytesIO

from folha import SITUACAO_TEXTO, SITUACOES, formatar_moeda


AZUL = "#1f3864"


def _valor(celula):
    """
    Decimal vira float para o xlsxwriter escrever como número.
    """

    if isinstance(celula, Decimal):
        return float(celula)

    return celula


def _escrever_aba(livro, titulo, colunas, linhas, colunas_moeda=(), vazio=None):
    """
    Cria uma aba com cabeçalho fixo, filtro, largura automática
    e formato de moeda nas colunas indicadas.
    """

    pagina = livro.add_worksheet(titulo[:31])

    cabecalho = livro.add_format({
        "bold": True,
        "bg_color": AZUL,
        "font_color": "white",
        "border": 1,
    })
    moeda = livro.add_format({"num_format": "R$ #,##0.00"})

    for indice, coluna in enumerate(colunas):
        pagina.write(0, indice, coluna, cabecalho)

    larguras = [len(coluna) for coluna in colunas]

    if not linhas and vazio:
        pagina.write(1, 0, vazio)

    for numero, linha in enumerate(linhas, start=1):
        for indice, celula in enumerate(linha):
            valor = _valor(celula)
            formato = moeda if indice in colunas_moeda and isinstance(valor, float) else None
            pagina.write(numero, indice, valor, formato)
            larguras[indice] = max(larguras[indice], len(str(valor if valor is not None else "")))

    for indice, largura in enumerate(larguras):
        pagina.set_column(indice, indice, min(max(largura + 2, 10), 60))

    pagina.freeze_panes(1, 0)

    if linhas:
        pagina.autofilter(0, 0, len(linhas), len(colunas) - 1)

    return pagina


def _resumo(livro, resultado):

    pagina = livro.add_worksheet("Resumo")

    titulo = livro.add_format({"bold": True, "font_size": 13, "font_color": AZUL})
    negrito = livro.add_format({"bold": True})
    cabecalho = livro.add_format({"bold": True, "bg_color": AZUL, "font_color": "white", "border": 1})
    moeda = livro.add_format({"num_format": "R$ #,##0.00"})
    alerta = livro.add_format({"bold": True, "font_color": "#9C0006", "bg_color": "#FFC7CE"})
    ok = livro.add_format({"bold": True, "font_color": "#006100", "bg_color": "#C6EFCE"})

    linha = 0

    pagina.write(linha, 0, "Conferência Portal da Transparência x Folha", titulo)
    linha += 2

    folha = resultado["folha"]
    pagina.write(linha, 0, "Folha", negrito)
    pagina.write(linha, 1, f"{folha['arquivo']} ({resultado['tipo_folha']}, {folha['quantidade']} lançamentos)")
    linha += 1

    for txt in resultado["txts"]:
        pagina.write(linha, 0, f"TXT {txt['empresa']}", negrito)
        pagina.write(linha, 1, f"{txt['arquivo']} ({txt['quantidade']} lançamentos)")
        linha += 1

    linha += 1

    if resultado["confere"]:
        pagina.write(linha, 0, "Os TXT conferem com a Folha.", ok)
    else:
        pagina.write(linha, 0, "Há inconsistências entre os TXT e a Folha.", alerta)

    linha += 2

    for aviso in resultado["avisos"]:
        pagina.write(linha, 0, "Aviso: " + aviso, alerta)
        linha += 1

    if resultado["avisos"]:
        linha += 1

    # Totais por empresa.
    colunas = [
        "Empresa", "Linhas TXT", "Linhas Folha", "Linhas Folha (rubricas publicadas)",
        "Colaboradores TXT", "Colaboradores Folha",
        "Proventos TXT", "Proventos Folha (publicadas)", "Diferença proventos",
        "Descontos TXT", "Descontos Folha (publicadas)", "Diferença descontos",
        "Proventos Folha (todas)", "Descontos Folha (todas)",
        "Lançamentos conferidos", "Divergências",
    ]

    for indice, coluna in enumerate(colunas):
        pagina.write(linha, indice, coluna, cabecalho)

    linha += 1

    for total in resultado["totais"]:
        valores = [
            total["empresa"], total["linhas_txt"], total["linhas_folha"],
            total["linhas_folha_publicaveis"], total["colaboradores_txt"],
            total["colaboradores_folha"],
            total["proventos_txt"], total["proventos_folha"],
            total["proventos_txt"] - total["proventos_folha"],
            total["descontos_txt"], total["descontos_folha"],
            total["descontos_txt"] - total["descontos_folha"],
            total["proventos_folha_total"], total["descontos_folha_total"],
            total["conferidos"], total["divergencias"],
        ]

        for indice, valor in enumerate(valores):
            valor = _valor(valor)
            pagina.write(linha, indice, valor, moeda if isinstance(valor, float) else None)

        linha += 1

    linha += 1

    # Contagem das inconsistências.
    pagina.write(linha, 0, "Inconsistências", negrito)
    linha += 1

    contagens = [
        (SITUACAO_TEXTO[s], resultado["por_situacao"][s]) for s in SITUACOES
    ] + [
        ("Colaboradores só no TXT", sum(1 for c in resultado["colaboradores"] if c["lado"] == "Só no TXT")),
        ("Colaboradores só na Folha com rubricas publicadas",
         sum(1 for c in resultado["colaboradores"] if c["lado"] == "Só na Folha" and c["relevante"])),
        ("Nomes diferentes (mesma matrícula)", len(resultado["nomes_diferentes"])),
        ("Problemas internos do TXT", len(resultado["internos"]["problemas"])),
        ("Repetições nome + rubrica sem explicação",
         sum(1 for r in resultado["internos"]["repeticoes"] if not r["explicada"])),
    ]

    for rotulo, quantidade in contagens:
        pagina.write(linha, 0, rotulo)
        pagina.write(linha, 1, quantidade, alerta if quantidade else None)
        linha += 1

    linha += 1
    pagina.write(linha, 0, "Informativo", negrito)
    linha += 1

    fora = resultado["rubricas_fora"]
    informativo = [
        ("Lançamentos conferidos (empresa + matrícula + rubrica)", resultado["quantidade_conferidos"]),
        ("Valor conferido", formatar_moeda(resultado["valor_conferido"])),
        ("Rubricas da Folha que nenhum TXT publica", len(fora)),
        ("   das quais proventos", sum(1 for r in fora if r["provento"])),
        ("Colaboradores só na Folha, só com rubricas não publicadas",
         sum(1 for c in resultado["colaboradores"] if not c["relevante"])),
        ("Repetições nome + rubrica explicadas", sum(1 for r in resultado["internos"]["repeticoes"] if r["explicada"])),
    ]

    for rotulo, quantidade in informativo:
        pagina.write(linha, 0, rotulo)
        pagina.write(linha, 1, quantidade)
        linha += 1

    pagina.set_column(0, 0, 58)
    pagina.set_column(1, 1, 60)
    pagina.set_column(2, len(colunas), 18)


def montar_planilha(resultado):
    """
    Monta o .xlsx da conferência em memória e devolve os bytes.
    """

    import xlsxwriter

    buffer = BytesIO()
    livro = xlsxwriter.Workbook(buffer, {"in_memory": True})

    _resumo(livro, resultado)

    _escrever_aba(
        livro,
        "Divergencias",
        [
            "Situação", "Empresa", "Matrícula", "Nome", "Rubrica", "Descrição",
            "Tipo TXT", "Tipo Folha", "Qtde TXT", "Qtde Folha",
            "Valor TXT", "Valor Folha", "Diferença (TXT - Folha)",
            "Natureza (portal)", "Observação",
        ],
        [
            [
                SITUACAO_TEXTO[d["situacao"]], d["empresa"], d["matricula"], d["nome"],
                d["codigo"], d["rubrica"], d["tipo_txt"], d["tipo_folha"],
                d["quantidade_txt"], d["quantidade_folha"],
                d["valor_txt"], d["valor_folha"], d["diferenca"],
                d["natureza"], d["observacao"],
            ]
            for d in resultado["divergencias"]
        ],
        colunas_moeda=(10, 11, 12),
        vazio="Nenhuma divergência linha a linha.",
    )

    _escrever_aba(
        livro,
        "Colaboradores",
        [
            "Situação", "Relevante", "Empresa", "Matrícula", "Nome", "Grupo",
            "Cargo / tipo", "Lançamentos", "Proventos", "Descontos",
            "Rubricas", "Observação",
        ],
        [
            [
                c["lado"], "Sim" if c["relevante"] else "Não (informativo)",
                c["empresa"], c["matricula"], c["nome"], c["grupo"], c["cargo"],
                c["lancamentos"], c["proventos"], c["descontos"],
                c["rubricas"], c["observacao"],
            ]
            for c in sorted(resultado["colaboradores"], key=lambda c: (not c["relevante"], c["lado"], c["empresa"], c["nome"]))
        ],
        colunas_moeda=(8, 9),
        vazio="Todos os colaboradores estão nos dois lados.",
    )

    _escrever_aba(
        livro,
        "Rubricas fora do portal",
        [
            "Rubrica", "Descrição", "Tipo", "Lançamentos", "Colaboradores",
            "Valor", "Empresas (lançamentos)", "Grupos",
        ],
        [
            [
                r["codigo"], r["rubrica"], r["tipo"], r["lancamentos"],
                r["colaboradores"], r["valor"], r["empresas"], r["grupos"],
            ]
            for r in resultado["rubricas_fora"]
        ],
        colunas_moeda=(5,),
        vazio="Todas as rubricas da Folha aparecem em algum TXT.",
    )

    _escrever_aba(
        livro,
        "Lancamentos fora do portal",
        ["Empresa", "Matrícula", "Nome", "Grupo", "Rubrica", "Descrição", "Tipo", "Valor"],
        [
            [l["empresa"], l["matricula"], l["nome"], l["grupo"], l["codigo"], l["rubrica"], l["tipo"], l["valor"]]
            for l in resultado["lancamentos_fora"]
        ],
        colunas_moeda=(7,),
        vazio="Nenhum lançamento fora do portal.",
    )

    _escrever_aba(
        livro,
        "Nomes diferentes",
        ["Empresa", "Matrícula", "Nome no TXT", "Nome na Folha"],
        [
            [n["empresa"], n["matricula"], n["nome_txt"], n["nome_folha"]]
            for n in resultado["nomes_diferentes"]
        ],
        vazio="Os nomes conferem em todas as matrículas.",
    )

    _escrever_aba(
        livro,
        "Repeticoes nome+rubrica",
        ["Empresa", "Nome", "Matrículas", "Rubrica", "Descrição", "Vezes", "Valores", "Explicação", "Explicada"],
        [
            [
                r["empresa"], r["nome"], r["matriculas"], r["codigo"], r["rubrica"],
                r["quantidade"], r["valores"], r["explicacao"],
                "Sim" if r["explicada"] else "Não",
            ]
            for r in resultado["internos"]["repeticoes"]
        ],
        vazio="Nenhum nome com a mesma rubrica repetida.",
    )

    _escrever_aba(
        livro,
        "Problemas no TXT",
        ["Arquivo", "Linha", "Problema", "Conteúdo"],
        [
            [p["arquivo"], p["linha"], p["problema"], p["conteudo"]]
            for p in resultado["internos"]["problemas"]
        ],
        vazio="Nenhum problema de layout ou cadastro nos TXT.",
    )

    livro.close()

    return buffer.getvalue()

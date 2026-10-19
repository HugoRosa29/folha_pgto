"""
Conferência Portal da Transparência x Folha pela linha de comando.

Mesma conferência da página web, para quem prefere rodar direto
no terminal.

Uso:

    python main.py --pasta ARQUIVOS_REF

    python main.py --pasta ARQUIVOS_REF --excel conferencia.xlsx

    python main.py --folha "Folha 3 empresas - SET 26.xlsx" \\
                   --txt REMUNERACAO_TERRACAP_20260901.txt \\
                   --txt REMUNERACAO_ETR_20260901.txt \\
                   --txt REMUNERACAO_BIOTIC_20260901.txt

Com --pasta, os arquivos são reconhecidos assim:

    TXT    todo .txt cujo nome começa com REMUNERACAO
    Folha  a planilha (.xlsx) que tem as colunas do SELECT do
           BENNER (EMPRESA, MATRICULA, CODIGO, VALOR, TIPOVERBA...)
"""

from pathlib import Path
import argparse
import sys

from folha import SITUACAO_TEXTO, SITUACOES, comparar, formatar_moeda, ler_folha


LARGURA = 78


# ============================================================
# DESCOBERTA DE ARQUIVOS
# ============================================================

def descobrir(pasta):
    """
    Encontra os TXT do portal e a planilha da Folha numa pasta.
    """

    pasta = Path(pasta)

    txts = sorted(
        caminho for caminho in pasta.iterdir()
        if caminho.is_file()
        and caminho.suffix.lower() == ".txt"
        and caminho.name.upper().startswith("REMUNERACAO")
    )

    # A pasta pode ter outras planilhas (a conferência manual,
    # por exemplo): a Folha é a que tem as colunas do SELECT.
    folhas = []

    for caminho in sorted(pasta.iterdir()):
        if not caminho.is_file() or caminho.name.startswith("~$"):
            continue

        if caminho.suffix.lower() not in (".xlsx", ".xlsm"):
            continue

        try:
            ler_folha(caminho)
        except Exception:
            continue

        folhas.append(caminho)

    if len(folhas) > 1:
        raise ValueError(
            "Mais de uma planilha de Folha na pasta: "
            + ", ".join(f.name for f in folhas)
            + ". Indique a certa com --folha."
        )

    return txts, (folhas[0] if folhas else None)


# ============================================================
# RELATÓRIO
# ============================================================

def titulo(texto):
    print()
    print("=" * LARGURA)
    print(texto)
    print("=" * LARGURA)


def secao(texto):
    print()
    print(texto)
    print("-" * LARGURA)


def linha(rotulo, valor, largura=50):
    print(f"  {rotulo:<{largura}} {valor:>24}")


def relatar(resultado):

    titulo("CONFERÊNCIA PORTAL DA TRANSPARÊNCIA x FOLHA")

    folha = resultado["folha"]
    print(f"  Folha: {folha['arquivo']} ({resultado['tipo_folha']}, "
          f"{folha['quantidade']} lançamentos)")

    for txt in resultado["txts"]:
        print(f"  TXT {txt['empresa']:<9} {txt['arquivo']} ({txt['quantidade']} lançamentos)")

    for aviso in resultado["avisos"]:
        print(f"\n  AVISO: {aviso}")

    # ---------- totais ----------

    for total in resultado["totais"]:

        secao(f"{total['empresa']}" + ("" if total["comparada"] else " (não comparada)"))

        linha("Lançamentos no TXT", str(total["linhas_txt"]))
        linha("Lançamentos na Folha (rubricas publicadas)",
              f"{total['linhas_folha_publicaveis']} de {total['linhas_folha']}")
        linha("Colaboradores TXT / Folha",
              f"{total['colaboradores_txt']} / {total['colaboradores_folha']}")
        linha("Proventos TXT", formatar_moeda(total["proventos_txt"]))
        linha("Proventos Folha (rubricas publicadas)", formatar_moeda(total["proventos_folha"]))
        linha("   diferença", formatar_moeda(total["proventos_txt"] - total["proventos_folha"]))
        linha("Descontos TXT", formatar_moeda(total["descontos_txt"]))
        linha("Descontos Folha (rubricas publicadas)", formatar_moeda(total["descontos_folha"]))
        linha("   diferença", formatar_moeda(total["descontos_txt"] - total["descontos_folha"]))
        linha("Lançamentos conferidos", str(total["conferidos"]))
        linha("Divergências linha a linha", str(total["divergencias"]))

    # ---------- inconsistências ----------

    titulo("INCONSISTÊNCIAS")

    colaboradores = [c for c in resultado["colaboradores"] if c["relevante"]]

    if colaboradores:
        secao(f"Colaboradores de um lado só ({len(colaboradores)})")

        for c in colaboradores:
            print(f"  {c['lado']:<12} {c['empresa']:<9} {c['matricula']:>9}  {c['nome'][:38]:<38}"
                  f" {formatar_moeda(c['proventos']):>16}")
            codigos = ", ".join(r.split(" ", 1)[0] for r in c["rubricas"].split(", "))
            print(f"  {'':<12} {c['lancamentos']} lançamento(s), rubricas {codigos}")
            print(f"  {'':<12} {c['observacao']}")

    soltas = [d for d in resultado["divergencias"] if not d["colaborador_ausente"]]

    for situacao in SITUACOES:
        itens = [d for d in soltas if d["situacao"] == situacao]

        if not itens:
            continue

        secao(f"{SITUACAO_TEXTO[situacao]} ({len(itens)})")

        for d in itens:
            print(f"  {d['empresa']:<9} {d['matricula']:>9}  {d['nome'][:28]:<28} "
                  f"{str(d['codigo']):>6} {d['rubrica'][:22]:<22}")
            print(f"  {'':<20} TXT {formatar_moeda(d['valor_txt'] or 0):>15} ({d['quantidade_txt']}x {d['tipo_txt'] or '—'})"
                  f"   Folha {formatar_moeda(d['valor_folha'] or 0):>15} ({d['quantidade_folha']}x {d['tipo_folha'] or '—'})")

    if resultado["nomes_diferentes"]:
        secao(f"Nomes diferentes na mesma matrícula ({len(resultado['nomes_diferentes'])})")

        for n in resultado["nomes_diferentes"]:
            print(f"  {n['empresa']:<9} {n['matricula']:>9}  TXT: {n['nome_txt']}")
            print(f"  {'':<20} Folha: {n['nome_folha']}")

    internos = resultado["internos"]

    if internos["problemas"]:
        secao(f"Problemas no próprio TXT ({len(internos['problemas'])})")

        for p in internos["problemas"]:
            onde = f"{p['arquivo']}:{p['linha']}" if p["linha"] else p["arquivo"]
            print(f"  {onde}  {p['problema']}")

    repeticoes = internos["repeticoes"]

    if repeticoes:
        secao(f"Nome + verba repetidos no TXT, checagem da Conf ({len(repeticoes)})")

        for r in repeticoes:
            print(f"  {r['empresa']:<9} {r['chave_conf'][:44]:<44} {r['quantidade']}x  {r['valores']}")
            print(f"  {'':<9} {r['explicacao']}")

    if resultado["confere"]:
        print("\n  Nenhuma inconsistência: os TXT conferem com a Folha.")

    # ---------- informativo ----------

    titulo("INFORMATIVO")

    fora = resultado["rubricas_fora"]
    proventos = [r for r in fora if r["provento"]]

    secao(f"Rubricas da Folha que nenhum TXT publica ({len(fora)}; {len(proventos)} de provento)")
    print("  Não são comparadas linha a linha. Proventos fora do portal merecem")
    print("  uma segunda olhada; a lista completa está no Excel.\n")

    for r in proventos:
        print(f"  {str(r['codigo']):>6} {r['rubrica'][:40]:<40} {r['lancamentos']:>4}x"
              f" {formatar_moeda(r['valor']):>16}  {r['grupos'][:20]}")

    descontos = [r for r in fora if not r["provento"]]

    if descontos:
        print(f"\n  + {len(descontos)} rubricas de desconto "
              f"({formatar_moeda(sum(r['valor'] for r in descontos))}): "
              "consignações, planos, empréstimos...")

    informativos = [c for c in resultado["colaboradores"] if not c["relevante"]]

    if informativos:
        grupos = {}
        for c in informativos:
            grupos[c["grupo"]] = grupos.get(c["grupo"], 0) + 1

        secao(f"Colaboradores só na Folha, só com rubricas não publicadas ({len(informativos)})")
        for grupo, quantidade in sorted(grupos.items(), key=lambda g: -g[1]):
            print(f"  {quantidade:>4}  {grupo}")

    print()


# ============================================================
# ENTRADA
# ============================================================

def main(argumentos=None):

    parser = argparse.ArgumentParser(
        description="Confere os TXT do Portal da Transparência com a Folha do BENNER.",
    )
    parser.add_argument("--pasta", help="pasta com os TXT e a planilha da Folha")
    parser.add_argument("--txt", action="append", default=[], help="TXT do portal (repita para cada empresa)")
    parser.add_argument("--folha", help="planilha da Folha (.xlsx)")
    parser.add_argument("--excel", help="grava o resultado completo neste .xlsx")

    args = parser.parse_args(argumentos)

    txts = [Path(t) for t in args.txt]
    folha = Path(args.folha) if args.folha else None

    if args.pasta:
        achados, achada = descobrir(args.pasta)
        txts = txts or achados
        folha = folha or achada

    if not txts:
        parser.error("nenhum TXT do portal: use --txt ou --pasta")

    if folha is None:
        parser.error("nenhuma planilha da Folha: use --folha ou --pasta")

    resultado = comparar(txts, folha)

    relatar(resultado)

    if args.excel:
        from exportar import montar_planilha

        Path(args.excel).write_bytes(montar_planilha(resultado))
        print(f"Excel gravado em {args.excel}")

    return 0 if resultado["confere"] else 1


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    sys.exit(main())

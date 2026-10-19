# Conferência da Folha — Portal da Transparência

Confere os arquivos de remuneração enviados ao **Portal da
Transparência** (`REMUNERACAO_<EMPRESA>_<AAAAMMDD>.txt`) contra a
**Folha** extraída do BENNER para as três empresas: TERRACAP, ETR
e BIOTIC.

A conferência manual (*Conf Portal Transparência*) só olhava o TXT
da Terracap: montava a chave nome + verba e contava as repetições
dentro do próprio arquivo. Não comparava nada com a Folha. Este
programa casa cada lançamento do TXT com a Folha, centavo por
centavo, e aponta nome por nome o que não confere.

Assim como a conciliação financeira, a aplicação roda **inteiramente
dentro do navegador**. Não há servidor, upload nem armazenamento: os
arquivos de folha nunca saem da máquina de quem usa.

## O que é conferido

### Lançamento por lançamento

A chave é **empresa + matrícula + rubrica**. O nome não entra na
chave: duas pessoas com o mesmo nome, ou a mesma pessoa com dois
vínculos, não se confundem.

| Situação | O que significa |
| --- | --- |
| Valor diferente | o mesmo lançamento, por outro valor |
| Provento x desconto | `VL_RECEBIMENTO` de um lado, `TIPOVERBA = 2` do outro (ou o inverso) |
| Na Folha, ausente do TXT | a rubrica é publicada, mas esse lançamento não foi |
| No TXT, ausente da Folha | foi publicado e não está na Folha |
| Quantidade de lançamentos diferente | mesmo total, número de linhas diferente |

### Colaboradores de um lado só

- **Só no TXT**: estão no portal e não têm nenhum lançamento na Folha.
  Em geral receberam por outra folha (rescisão, férias, complementar),
  porque o SELECT da Folha traz só `TIPOFOLHA = 11` (Folha Normal).
- **Só na Folha**: têm rubricas publicadas e não aparecem no TXT.

### Rubricas fora do portal

O portal não publica todas as rubricas: planos de saúde, empréstimos,
consignações e bolsa de estágio ficam de fora. Uma rubrica que não
aparece em **nenhum** TXT não é cobrada linha a linha. Ela vai para um
resumo à parte, com os **proventos** em destaque, porque um provento
fora do portal merece uma segunda olhada.

Os colaboradores que só recebem essas rubricas (os estagiários, por
exemplo) aparecem como informativos, não como inconsistência.

### O próprio TXT

- layout de 30 campos e valor em recebimento **ou** desconto
- `CPF_FORMATADO` coerente com o `NR_CPF`
- mesma matrícula com o mesmo CPF, nome, grupo e cargo em todas as linhas
- a mesma rubrica sempre na mesma natureza
- **nome + rubrica repetidos**, a checagem da conferência manual, agora
  com a explicação: dois vínculos (matrículas diferentes) ou a Folha
  também traz os mesmos lançamentos repetidos

## Como usar

### No navegador (GitHub Pages)

Abra a página publicada, arraste os três TXT para a primeira caixa e a
planilha da Folha para a segunda, e clique em **Conferir**. Cada TXT
se identifica pelo CNPJ, então o nome do arquivo não importa.

**Baixar Excel** gera a planilha com uma aba por bloco: *Resumo*,
*Divergencias*, *Colaboradores*, *Rubricas fora do portal*,
*Lancamentos fora do portal*, *Nomes diferentes*,
*Repeticoes nome+rubrica* e *Problemas no TXT*.

### Na linha de comando

```powershell
pip install -r requirements.txt

python main.py --pasta ARQUIVOS_REF
python main.py --pasta ARQUIVOS_REF --excel conferencia.xlsx

python main.py --folha "Folha 3 empresas - SET 26.xlsx" `
               --txt REMUNERACAO_TERRACAP_20260901.txt `
               --txt REMUNERACAO_ETR_20260901.txt `
               --txt REMUNERACAO_BIOTIC_20260901.txt
```

Com `--pasta`, os TXT são os `.txt` que começam com `REMUNERACAO`, e a
Folha é a planilha que tem as colunas do SELECT (`EMPRESA`,
`MATRICULA`, `CODIGO`, `VALOR`, `TIPOVERBA`...). Outras planilhas da
pasta, como a conferência manual, são ignoradas.

O programa sai com código 0 quando tudo confere e 1 quando há
inconsistência.

## Validação

```powershell
python -m unittest -v test_folha

$env:FOLHA_REF = "C:\Users\hugos\Downloads\FOLHA_PGTO\ARQUIVOS_REF"
python -m unittest -v test_folha
```

Os testes sintéticos montam TXT e planilha em memória e rodam no CI.
Os testes com os arquivos reais de setembro/2026 só rodam com
`FOLHA_REF` definido, para que nenhum dado de folha vá para o
repositório.

## Publicar no GitHub Pages

A pasta `docs/` já é o site pronto.

1. `git push` para o branch `main`.
2. Em **Settings → Pages**, escolha **GitHub Actions** como origem
   (o workflow `.github/workflows/pages.yml` testa e publica `docs/`).

> Depois de alterar qualquer módulo do núcleo, rode `python build.py`
> para atualizar `docs/py/`. O workflow confere isso antes de publicar.

## Como está organizado

```
folha.py       núcleo: leitura do TXT e da Folha, conferência
exportar.py    geração do Excel (xlsxwriter)
ponte.py       ligação entre o navegador e o núcleo
main.py        interface de linha de comando
build.py       copia o núcleo para docs/py/
test_folha.py  testes
docs/          o site publicado no GitHub Pages
```

O núcleo depende apenas de `openpyxl` e `XlsxWriter`, bibliotecas
Python puras. É por isso que o Pyodide consegue carregá-lo no
navegador. Os wheels ficam em `docs/vendor/`, servidos da própria
página.

## Aviso

Os TXT e a Folha trazem nome, CPF e remuneração de cada colaborador.
O `.gitignore` impede que sejam versionados por engano. Confira antes
de cada commit.

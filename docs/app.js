/*
 * Interface da conferência Portal da Transparência x Folha.
 *
 * Carrega o Pyodide, instala as bibliotecas de planilha
 * (Python puro, servidas da própria pasta vendor/) e roda o
 * mesmo núcleo em Python usado pela linha de comando.
 *
 * Nenhum arquivo é enviado para a rede: o conteúdo é lido como
 * bytes e passado direto para o Python dentro do navegador.
 */

"use strict";

const MODULOS = ["folha.py", "exportar.py", "ponte.py"];

const WHEELS = [
  "vendor/et_xmlfile-2.0.0-py3-none-any.whl",
  "vendor/openpyxl-3.1.5-py2.py3-none-any.whl",
  "vendor/xlsxwriter-3.2.9-py3-none-any.whl",
];

const SITUACAO = {
  valor_diferente: { texto: "Valor diferente", classe: "encerrado" },
  tipo_diferente: { texto: "Provento x desconto", classe: "encerrado" },
  ausente_txt: { texto: "Na Folha, ausente do TXT", classe: "encerrado" },
  ausente_folha: { texto: "No TXT, ausente da Folha", classe: "encerrado" },
  quantidade_diferente: { texto: "Qtde de lançamentos", classe: "estudo" },
};

const arquivos = { txts: [], folha: null };

let pyodide = null;
let ponte = null;
let resultadoAtual = null;

const el = (selector) => document.querySelector(selector);
const todos = (selector) => Array.from(document.querySelectorAll(selector));

const moeda = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
const inteiro = new Intl.NumberFormat("pt-BR");
const formatarMoeda = (valor) => moeda.format(Number(valor) || 0);

/* ============================================================
 * Carregamento do motor Python
 * ============================================================ */

function estado(texto, classe) {
  el("#estado-texto").textContent = texto;
  el("#estado-runtime").className = `carimbo carimbo-${classe}`;
}

async function iniciarPython() {
  try {
    estado("Carregando o motor de cálculo", "estudo");

    pyodide = await loadPyodide({
      stdout: () => {},
      stderr: (linha) => console.warn("[python]", linha),
    });

    estado("Instalando bibliotecas de planilha", "estudo");

    await pyodide.loadPackage("micropip");

    const micropip = pyodide.pyimport("micropip");
    const base = new URL(".", window.location.href).href;
    const lista = pyodide.toPy(WHEELS.map((caminho) => base + caminho));

    try {
      await micropip.install(lista, false, false);
    } finally {
      lista.destroy();
    }

    estado("Carregando a conferência", "estudo");

    for (const modulo of MODULOS) {
      const resposta = await fetch(`py/${modulo}`, { cache: "no-cache" });

      if (!resposta.ok) {
        throw new Error(`Não foi possível carregar py/${modulo}`);
      }

      pyodide.FS.writeFile(`/home/pyodide/${modulo}`, await resposta.text());
    }

    ponte = pyodide.pyimport("ponte");

    estado("Pronto · nada sai deste computador", "andamento");
    atualizarBotao();
  } catch (erro) {
    console.error(erro);
    estado("Falha ao carregar", "encerrado");
    mostrarMensagem(
      "Não foi possível iniciar o motor de cálculo. Verifique a conexão e recarregue a página.",
      true,
    );
  }
}

/* ============================================================
 * Seleção de arquivos
 * ============================================================ */

function registrarCaixa(caixa) {
  const campo = caixa.dataset.campo;
  const entrada = caixa.querySelector("input[type=file]");

  entrada.addEventListener("change", () => {
    if (entrada.files.length) {
      definirArquivos(campo, Array.from(entrada.files));
    }
  });

  ["dragenter", "dragover"].forEach((evento) =>
    caixa.addEventListener(evento, (e) => {
      e.preventDefault();
      caixa.classList.add("arrastando");
    }),
  );

  ["dragleave", "drop"].forEach((evento) =>
    caixa.addEventListener(evento, (e) => {
      e.preventDefault();
      caixa.classList.remove("arrastando");
    }),
  );

  caixa.addEventListener("drop", (e) => {
    const lista = Array.from(e.dataTransfer?.files || []);

    if (lista.length) {
      definirArquivos(campo, lista);
    }
  });
}

function definirArquivos(campo, lista) {
  const caixa = el(`.caixa-arquivo[data-campo="${campo}"]`);

  if (campo === "txts") {
    // Arrastar mais TXT soma aos que já estão; o mesmo nome
    // substitui o anterior.
    const porNome = new Map(arquivos.txts.map((a) => [a.name, a]));
    lista.forEach((arquivo) => porNome.set(arquivo.name, arquivo));
    arquivos.txts = Array.from(porNome.values());

    caixa.querySelector(".nome").textContent = arquivos.txts
      .map((a) => `${a.name} · ${(a.size / 1024).toFixed(0)} KB`)
      .join("\n");
  } else {
    arquivos.folha = lista[0];
    caixa.querySelector(".nome").textContent =
      `${lista[0].name} · ${(lista[0].size / 1024).toFixed(0)} KB`;
  }

  caixa.classList.add("preenchida");
  mostrarMensagem("");
  atualizarBotao();
}

function atualizarBotao() {
  el("#botao-conferir").disabled = !(arquivos.txts.length && arquivos.folha && ponte);
}

function limpar() {
  arquivos.txts = [];
  arquivos.folha = null;

  todos(".caixa-arquivo").forEach((caixa) => {
    caixa.classList.remove("preenchida");
    caixa.querySelector(".nome").textContent =
      caixa.dataset.campo === "txts" ? "Clique ou arraste os arquivos" : "Clique ou arraste o arquivo";
    caixa.querySelector("input[type=file]").value = "";
  });

  resultadoAtual = null;
  el("#resultado").hidden = true;
  mostrarMensagem("");
  atualizarBotao();
}

function mostrarMensagem(texto, erro = false) {
  const alvo = el("#mensagem");
  alvo.textContent = texto;
  alvo.className = erro ? "mensagem erro" : "mensagem";
}

/* ============================================================
 * Execução
 * ============================================================ */

async function bytesDe(arquivo) {
  return { nome: arquivo.name, dados: new Uint8Array(await arquivo.arrayBuffer()) };
}

async function conferir() {
  if (!ponte) {
    return;
  }

  const botao = el("#botao-conferir");
  botao.disabled = true;
  mostrarMensagem("Processando os arquivos…");

  try {
    const entrada = {
      txts: await Promise.all(arquivos.txts.map(bytesDe)),
      folha: await bytesDe(arquivos.folha),
    };

    const mapa = pyodide.toPy(entrada);
    let bruto;

    try {
      bruto = ponte.processar(mapa);
    } finally {
      mapa.destroy();
    }

    const resultado = JSON.parse(bruto);

    if (resultado.erro) {
      mostrarMensagem(resultado.erro, true);
      return;
    }

    resultadoAtual = resultado;
    desenhar(resultado);
    mostrarMensagem("Conferência concluída.");

    el("#resultado").scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (erro) {
    console.error(erro);
    mostrarMensagem(`Erro ao processar: ${erro.message || erro}`, true);
  } finally {
    botao.disabled = false;
    atualizarBotao();
  }
}

async function exportar() {
  if (!resultadoAtual || !ponte) {
    return;
  }

  const botao = el("#botao-exportar");
  botao.disabled = true;

  try {
    const bytes = ponte.exportar();
    let dados = bytes;

    if (typeof bytes?.toJs === "function") {
      dados = bytes.toJs({ create_proxies: false });
      bytes.destroy();
    }

    const blob = new Blob([dados], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const agora = new Date().toISOString().slice(0, 16).replace("T", "_").replace(":", "");
    const link = document.createElement("a");

    link.href = URL.createObjectURL(blob);
    link.download = `conferencia_folha_${agora}.xlsx`;
    link.click();

    URL.revokeObjectURL(link.href);
  } catch (erro) {
    console.error(erro);
    mostrarMensagem(`Erro ao gerar o Excel: ${erro.message || erro}`, true);
  } finally {
    botao.disabled = false;
  }
}

/* ============================================================
 * Desenho do resultado
 * ============================================================ */

function desenhar(resultado) {
  el("#resultado").hidden = false;

  el("#avisos").innerHTML = (resultado.avisos || [])
    .map((aviso) => `<div class="aviso">${escapar(aviso)}</div>`)
    .join("");

  desenharIndicadores(resultado);
  desenharPrestacaoDeContas(resultado);
  desenharColaboradores(resultado);
  desenharDivergencias(resultado);
  desenharRubricas(resultado);
  desenharRepeticoes(resultado);
  desenharCadastro(resultado);
  desenharInformativo(resultado);
}

function indicador({ titulo, valor, detalhe = "", classe = "" }) {
  return `
    <div class="indicador ${classe}">
      <div class="titulo">${escapar(titulo)}</div>
      <div class="valor">${escapar(valor)}</div>
      ${detalhe ? `<div class="detalhe">${escapar(detalhe)}</div>` : ""}
    </div>`;
}

function desenharIndicadores(resultado) {
  const cartoes = [];

  for (const total of resultado.totais) {
    const diferenca = Number(total.proventos_txt) - Number(total.proventos_folha);
    const fecha = Math.abs(diferenca) < 0.005 && total.divergencias === 0;

    cartoes.push(
      indicador({
        titulo: `${total.empresa} · proventos no portal`,
        valor: formatarMoeda(total.proventos_txt),
        detalhe: !total.comparada
          ? "Empresa não comparada"
          : fecha
            ? `Fecha com a Folha · ${inteiro.format(total.conferidos)} lançamentos conferidos`
            : `Diferença de ${formatarMoeda(diferenca)} para a Folha · ${inteiro.format(total.divergencias)} lançamento(s) divergente(s)`,
        classe: !total.comparada ? "" : fecha ? "ok" : "alerta",
      }),
    );
  }

  const fora = resultado.rubricas_fora || [];
  const proventosFora = fora.filter((r) => r.provento);

  cartoes.push(
    indicador({
      titulo: "Lançamentos conferidos",
      valor: inteiro.format(resultado.quantidade_conferidos),
      detalhe: `${formatarMoeda(resultado.valor_conferido)} iguais nos dois lados`,
      classe: "ok",
    }),
    indicador({
      titulo: "Rubricas fora do portal",
      valor: inteiro.format(fora.length),
      detalhe: `${inteiro.format(proventosFora.length)} de provento, ${formatarMoeda(
        proventosFora.reduce((s, r) => s + Number(r.valor), 0),
      )}`,
      classe: proventosFora.length ? "alerta" : "",
    }),
  );

  el("#indicadores").innerHTML = cartoes.join("");
}

function linhaConta(rotulo, valor, classe = "") {
  return `
    <div class="linha-conta ${classe}">
      <span>${escapar(rotulo)}</span>
      <span>${valor}</span>
    </div>`;
}

function desenharPrestacaoDeContas(resultado) {
  const colaboradores = resultado.colaboradores.filter((c) => c.relevante);
  const soNoTxt = colaboradores.filter((c) => c.lado === "Só no TXT");
  const soNaFolha = colaboradores.filter((c) => c.lado === "Só na Folha");
  const soltas = resultado.divergencias.filter((d) => !d.colaborador_ausente);
  const repeticoes = resultado.internos.repeticoes;

  const contas = [
    linhaConta(
      `Colaboradores no TXT sem lançamento na Folha (${inteiro.format(soNoTxt.length)})`,
      formatarMoeda(soNoTxt.reduce((s, c) => s + Number(c.proventos), 0)),
    ),
    linhaConta(
      `Colaboradores na Folha fora do TXT (${inteiro.format(soNaFolha.length)})`,
      formatarMoeda(soNaFolha.reduce((s, c) => s + Number(c.proventos), 0)),
    ),
  ];

  for (const [situacao, { texto }] of Object.entries(SITUACAO)) {
    const itens = soltas.filter((d) => d.situacao === situacao);
    contas.push(linhaConta(`${texto} (${inteiro.format(itens.length)})`,
      formatarMoeda(itens.reduce((s, d) => s + Math.abs(Number(d.diferenca)), 0))));
  }

  contas.push(
    linhaConta("Nomes diferentes na mesma matrícula", inteiro.format(resultado.nomes_diferentes.length)),
    linhaConta("Problemas de layout ou cadastro no TXT", inteiro.format(resultado.internos.problemas.length)),
    linhaConta("Nome + verba repetidos no TXT (checagem da Conf)", inteiro.format(repeticoes.length), "total"),
  );

  const selo = resultado.confere
    ? '<div class="selo ok">Os TXT conferem com a Folha em todas as rubricas publicadas.</div>'
    : `<div class="selo">Entre parênteses, a quantidade; à direita, os proventos dos colaboradores
         ou a soma das diferenças. Os lançamentos de quem está de um lado só entram
         apenas nas duas primeiras linhas. O detalhe está nas abas abaixo e no Excel.</div>`;

  el("#divergencia").innerHTML = `<div class="conta">${contas.join("")}${selo}</div>`;
}

function tabela(colunas, linhas, vazio) {
  if (!linhas.length) {
    return `<p class="vazio">${escapar(vazio)}</p>`;
  }

  const cabecalho = colunas
    .map((c) => `<th class="${c.numero ? "numero" : ""}">${escapar(c.titulo)}</th>`)
    .join("");

  const corpo = linhas
    .map((linha) => `<tr>${colunas
      .map((c) => `<td class="${c.numero ? "numero" : ""}">${
        c.html ? linha[c.chave] : escapar(linha[c.chave])}</td>`)
      .join("")}</tr>`)
    .join("");

  return `
    <div class="rolagem">
      <table>
        <thead><tr>${cabecalho}</tr></thead>
        <tbody>${corpo}</tbody>
      </table>
    </div>`;
}

function carimbo(texto, classe) {
  return `<span class="carimbo carimbo-${classe}">${escapar(texto)}</span>`;
}

function desenharColaboradores(resultado) {
  const linhas = resultado.colaboradores
    .filter((c) => c.relevante)
    .map((c) => ({
      ...c,
      lado: carimbo(c.lado, "encerrado"),
      proventos: formatarMoeda(c.proventos),
      descontos: formatarMoeda(c.descontos),
      // Só os códigos: a descrição de cada rubrica está no Excel.
      rubricas: c.rubricas.split(", ").map((r) => r.split(" ")[0]).join(", "),
    }));

  el('[data-painel="colaboradores"]').innerHTML =
    `<p class="nota-tabela">Colaboradores com lançamentos de um lado só. A Folha enviada é
      <strong>${escapar(resultado.tipo_folha)}</strong>: quem está no TXT sem nenhum lançamento
      nela em geral recebeu por outra folha (rescisão, férias, complementar). Confira no BENNER.
      As descrições das rubricas estão no Excel.</p>` +
    tabela(
      [
        { titulo: "Situação", chave: "lado", html: true },
        { titulo: "Empresa", chave: "empresa" },
        { titulo: "Matrícula", chave: "matricula" },
        { titulo: "Nome", chave: "nome" },
        { titulo: "Grupo", chave: "grupo" },
        { titulo: "Lançamentos", chave: "lancamentos", numero: true },
        { titulo: "Proventos", chave: "proventos", numero: true },
        { titulo: "Descontos", chave: "descontos", numero: true },
        { titulo: "Rubricas", chave: "rubricas" },
      ],
      linhas,
      "Todos os colaboradores com rubricas publicadas estão nos dois lados.",
    );
}

function desenharDivergencias(resultado) {
  const linhas = resultado.divergencias
    .filter((d) => !d.colaborador_ausente)
    .map((d) => ({
      ...d,
      situacao: carimbo(SITUACAO[d.situacao].texto, SITUACAO[d.situacao].classe),
      valor_txt: d.valor_txt == null ? "—" : `${formatarMoeda(d.valor_txt)} (${d.quantidade_txt}x ${d.tipo_txt})`,
      valor_folha: d.valor_folha == null ? "—" : `${formatarMoeda(d.valor_folha)} (${d.quantidade_folha}x ${d.tipo_folha})`,
      diferenca: formatarMoeda(d.diferenca),
    }));

  const deColaboradores = resultado.divergencias.length - linhas.length;

  el('[data-painel="divergencias"]').innerHTML =
    tabela(
      [
        { titulo: "Situação", chave: "situacao", html: true },
        { titulo: "Empresa", chave: "empresa" },
        { titulo: "Matrícula", chave: "matricula" },
        { titulo: "Nome", chave: "nome" },
        { titulo: "Rubrica", chave: "codigo" },
        { titulo: "Descrição", chave: "rubrica" },
        { titulo: "TXT", chave: "valor_txt", numero: true },
        { titulo: "Folha", chave: "valor_folha", numero: true },
        { titulo: "Diferença (TXT − Folha)", chave: "diferenca", numero: true },
      ],
      linhas,
      `Nenhuma divergência entre colaboradores que estão nos dois lados: valor e tipo
       (provento ou desconto) conferem em todos os ${inteiro.format(resultado.quantidade_conferidos)} lançamentos casados.`,
    ) +
    (deColaboradores
      ? `<p class="nota-tabela">Outros ${inteiro.format(deColaboradores)} lançamento(s) pertencem a
         colaboradores que estão de um lado só e ficam na aba anterior. No Excel, todos estão na aba
         <em>Divergencias</em>.</p>`
      : "");
}

function desenharRubricas(resultado) {
  const linhas = (resultado.rubricas_fora || []).map((r) => ({
    ...r,
    tipo: r.provento ? carimbo(r.tipo, "estudo") : escapar(r.tipo),
    valor: formatarMoeda(r.valor),
  }));

  el('[data-painel="rubricas"]').innerHTML =
    `<p class="nota-tabela">Rubricas da Folha que não aparecem em nenhum TXT. Não são cobradas
      linha a linha, porque o portal não as publica para ninguém. Os descontos são, em geral,
      consignações e planos; os <strong>proventos</strong> merecem uma segunda olhada.
      Os lançamentos, um a um, estão no Excel.</p>` +
    tabela(
      [
        { titulo: "Rubrica", chave: "codigo" },
        { titulo: "Descrição", chave: "rubrica" },
        { titulo: "Tipo", chave: "tipo", html: true },
        { titulo: "Lançamentos", chave: "lancamentos", numero: true },
        { titulo: "Colaboradores", chave: "colaboradores", numero: true },
        { titulo: "Valor", chave: "valor", numero: true },
        { titulo: "Empresas", chave: "empresas" },
        { titulo: "Grupos", chave: "grupos" },
      ],
      linhas,
      "Todas as rubricas da Folha aparecem em algum TXT.",
    );
}

function desenharRepeticoes(resultado) {
  const linhas = resultado.internos.repeticoes.map((r) => ({
    ...r,
    situacao: carimbo(`Repetido ${r.quantidade}x`, "encerrado"),
    causa: r.explicada
      ? escapar(r.explicacao)
      : `<span class="pendente">${escapar(r.explicacao)}</span>`,
  }));

  el('[data-painel="repeticoes"]').innerHTML =
    `<p class="nota-tabela">A mesma checagem da conferência manual (Conf Portal Transparência):
      a chave <em>nome e verba para conferência</em> (nome + código da rubrica) aparece mais de
      uma vez no TXT. Cada repetição é apontada para conferência; a causa provável é a
      análise automática contra a Folha e não dispensa a verificação.</p>` +
    tabela(
      [
        { titulo: "Situação", chave: "situacao", html: true },
        { titulo: "Empresa", chave: "empresa" },
        { titulo: "Nome e verba para conferência", chave: "chave_conf" },
        { titulo: "Matrículas", chave: "matriculas" },
        { titulo: "Descrição", chave: "rubrica" },
        { titulo: "Valores", chave: "valores", numero: true },
        { titulo: "Causa provável", chave: "causa", html: true },
      ],
      linhas,
      "Nenhum nome com a mesma verba repetida no TXT.",
    );
}

function desenharCadastro(resultado) {
  el('[data-painel="cadastro"]').innerHTML =
    "<h3>Nomes diferentes na mesma matrícula</h3>" +
    tabela(
      [
        { titulo: "Empresa", chave: "empresa" },
        { titulo: "Matrícula", chave: "matricula" },
        { titulo: "Nome no TXT", chave: "nome_txt" },
        { titulo: "Nome na Folha", chave: "nome_folha" },
      ],
      resultado.nomes_diferentes,
      "Os nomes conferem em todas as matrículas (sem considerar acentos e maiúsculas).",
    ) +
    "<h3>Layout e cadastro do TXT</h3>" +
    tabela(
      [
        { titulo: "Arquivo", chave: "arquivo" },
        { titulo: "Linha", chave: "linha", numero: true },
        { titulo: "Problema", chave: "problema" },
        { titulo: "Conteúdo", chave: "conteudo" },
      ],
      resultado.internos.problemas.map((p) => ({ ...p, linha: p.linha ?? "—" })),
      "Layout de 30 campos, CPF mascarado e cadastro por matrícula sem problemas.",
    );
}

function desenharInformativo(resultado) {
  const linhas = resultado.colaboradores
    .filter((c) => !c.relevante)
    .map((c) => ({
      ...c,
      proventos: formatarMoeda(c.proventos),
      rubricas: c.rubricas.split(", ").map((r) => r.split(" ")[0]).join(", "),
    }));

  el('[data-painel="informativo"]').innerHTML =
    `<p class="nota-tabela">Colaboradores que estão só na Folha e recebem apenas rubricas que o
      portal não publica (bolsa de estágio, por exemplo). Ficam fora do TXT por consequência
      e não contam como inconsistência.</p>` +
    tabela(
      [
        { titulo: "Empresa", chave: "empresa" },
        { titulo: "Matrícula", chave: "matricula" },
        { titulo: "Nome", chave: "nome" },
        { titulo: "Grupo", chave: "grupo" },
        { titulo: "Tipo", chave: "cargo" },
        { titulo: "Proventos", chave: "proventos", numero: true },
        { titulo: "Rubricas", chave: "rubricas" },
      ],
      linhas,
      "Nenhum.",
    );
}

function escapar(valor) {
  return String(valor ?? "").replace(
    /[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c],
  );
}

/* ============================================================
 * Abas
 * ============================================================ */

function registrarAbas() {
  todos(".aba").forEach((aba) => {
    aba.addEventListener("click", () => {
      todos(".aba").forEach((outra) => outra.classList.toggle("ativa", outra === aba));
      todos(".painel").forEach((painel) => {
        painel.hidden = painel.dataset.painel !== aba.dataset.aba;
      });
    });
  });
}

/* ============================================================
 * Tema claro / escuro
 * ============================================================ */

const CHAVE_TEMA = "conferencia-folha-tema";

function lerTemaSalvo() {
  try {
    const salvo = localStorage.getItem(CHAVE_TEMA);
    return salvo === "dark" || salvo === "light" ? salvo : null;
  } catch {
    return null;
  }
}

function temaEmUso() {
  const explicito = document.documentElement.dataset.theme;

  if (explicito === "dark" || explicito === "light") {
    return explicito;
  }

  return window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
}

function aplicarTema(tema) {
  document.documentElement.dataset.theme = tema;
  el("#tema-texto").textContent = tema === "dark" ? "Tema claro" : "Tema escuro";
  el("#botao-tema").setAttribute(
    "aria-label",
    tema === "dark" ? "Mudar para o tema claro" : "Mudar para o tema escuro",
  );

  try {
    localStorage.setItem(CHAVE_TEMA, tema);
  } catch {
    /* sem persistência, o tema vale só para esta visita */
  }
}

function registrarTema() {
  const salvo = lerTemaSalvo();

  if (salvo) {
    document.documentElement.dataset.theme = salvo;
  }

  el("#tema-texto").textContent = temaEmUso() === "dark" ? "Tema claro" : "Tema escuro";

  el("#botao-tema").addEventListener("click", () => {
    aplicarTema(temaEmUso() === "dark" ? "light" : "dark");
  });

  window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => {
    if (!lerTemaSalvo()) {
      el("#tema-texto").textContent = temaEmUso() === "dark" ? "Tema claro" : "Tema escuro";
    }
  });
}

/* ============================================================
 * Revelar ao rolar
 * ============================================================ */

function registrarRevelacao() {
  const alvos = todos(".revelar");

  if (!("IntersectionObserver" in window)) {
    alvos.forEach((alvo) => alvo.classList.add("visivel"));
    return;
  }

  const observador = new IntersectionObserver(
    (entradas) => {
      entradas.forEach((entrada) => {
        if (entrada.isIntersecting) {
          entrada.target.classList.add("visivel");
          observador.unobserve(entrada.target);
        }
      });
    },
    { threshold: 0.08, rootMargin: "0px 0px -40px 0px" },
  );

  alvos.forEach((alvo) => observador.observe(alvo));
}

/* ============================================================
 * Início
 * ============================================================ */

todos(".caixa-arquivo").forEach(registrarCaixa);
registrarAbas();
registrarRevelacao();
registrarTema();

el("#botao-conferir").addEventListener("click", conferir);
el("#botao-limpar").addEventListener("click", limpar);
el("#botao-exportar").addEventListener("click", exportar);

iniciarPython();

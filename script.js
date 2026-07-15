// ==========================================================
// Dashboard Operacional de Atendimento
// Script.js
// ==========================================================

console.clear();
console.log("Dashboard iniciado!");

// Guarda os dados originais da planilha (sem filtro nenhum aplicado)
let dadosOriginais = [];

// ==========================================================
// LEITURA AUTOMÁTICA DO ARQUIVO "report.xlsx"
// ==========================================================
// O arquivo precisa estar na MESMA PASTA do index.html/report.xlsx,
// e a página precisa ser acessada via http(s) (servidor, SharePoint etc.)
// -- abrir o index.html direto do computador (file://) NÃO funciona,
// por bloqueio de segurança do navegador.

const NOME_ARQUIVO_RELATORIO = "report.xlsx";

const botaoAtualizar = document.getElementById("btnAtualizar");

botaoAtualizar.addEventListener("click", carregarRelatorio);

async function carregarRelatorio(){

    botaoAtualizar.disabled = true;
    botaoAtualizar.classList.add("carregando");

    try{

        // "cache: no-store" + parâmetro no fim da URL garantem que
        // o navegador sempre busque a versão mais recente do arquivo
        const resposta = await fetch(
            `${NOME_ARQUIVO_RELATORIO}?_=${Date.now()}`,
            { cache: "no-store" }
        );

        if(!resposta.ok){
            throw new Error("HTTP " + resposta.status);
        }

        // Guarda o cabeçalho "Last-Modified" como alternativa,
        // caso o próprio arquivo .xlsx não tenha a data de modificação
        // salva nos seus metadados internos
        const lastModifiedHeader = resposta.headers.get("Last-Modified");

        const buffer = await resposta.arrayBuffer();

        // cellDates:false -> NÃO deixamos o SheetJS decidir sozinho quais
        // colunas são "data". Isso evita que uma coluna numérica (como o
        // número do chamado) seja convertida errado para data só porque a
        // célula ficou com alguma formatação de data no Excel.
        // Convertemos manualmente, mais abaixo, apenas as colunas que
        // realmente são datas (pelo nome da coluna).
        const workbook = XLSX.read(buffer,{
            type:"array",
            cellDates:false
        });

        // Primeira aba
        const aba = workbook.Sheets[workbook.SheetNames[0]];

        // Lemos a planilha de duas formas:
        // - jsonRaw: valores "crus" (números continuam números) -> usado
        //   para calcular corretamente as colunas que são datas de verdade.
        // - jsonFormatado: o texto exatamente como o Excel exibe na célula
        //   -> usado para todo o resto (ex: coluna "Chamado"), porque assim
        //   evitamos o problema de números "escondidos" atrás de uma
        //   formatação (ex: "13-1" que o Excel converteu sozinho para data
        //   e guardou como número de série por baixo dos panos).
        const jsonRaw = XLSX.utils.sheet_to_json(aba,{ raw:true });
        const jsonFormatado = XLSX.utils.sheet_to_json(aba,{ raw:false });

        const json = jsonRaw.map((linhaRaw, indice)=>{

            const linhaFormatada = jsonFormatado[indice] || {};
            const linhaFinal = {};

            Object.keys(linhaRaw).forEach(coluna=>{

                const valorRaw = linhaRaw[coluna];

                if(ehColunaDeNumeroChamado(coluna) && typeof valorRaw === "number"){

                    // Coluna "Chamado" quebrada pelo Excel (virou data) ->
                    // reconstrói o texto original (ex: 1-Jan -> "1-1")
                    linhaFinal[coluna] = reconstruirNumeroChamado(valorRaw);

                } else if(ehColunaDeData(coluna) && typeof valorRaw === "number"){

                    // Coluna de data de verdade -> converte o número de
                    // série do Excel em um objeto Date
                    linhaFinal[coluna] = excelSerialParaData(valorRaw);

                } else if(linhaFormatada[coluna] !== undefined){

                    // Qualquer outra coluna -> usa o texto formatado,
                    // exatamente como aparece no Excel
                    linhaFinal[coluna] = linhaFormatada[coluna];

                } else {

                    linhaFinal[coluna] = valorRaw;

                }

            });

            return linhaFinal;

        });

        // Converte para Date apenas as colunas que realmente são datas
        corrigirColunasDeData(json);

        // Renomeia a localidade "Padrão" para "Teste Cervello" em todos
        // os registros -> isso reflete automaticamente nos filtros,
        // nos gráficos, na tabela e no Excel exportado
        json.forEach(item=>{

            const localidade = item.Localidade_do_Solicitante;

            if(localidade && String(localidade).trim().toLowerCase() === "padrão"){
                item.Localidade_do_Solicitante = "Teste Cervello";
            }

        });

        // Classifica cada chamado em um Setor, com base no Analista_Atual
        json.forEach(item=>{
            item.Setor = obterSetorPorAnalista(item.Analista_Atual);
        });

        console.log(json);

        dadosOriginais = json;

        // Data do RELATÓRIO (não a data de acesso ao site):
        // 1) tenta a data de modificação salva dentro do próprio .xlsx
        // 2) se não existir, usa o cabeçalho Last-Modified do arquivo
        let dataDoRelatorio = null;

        if(workbook.Props && workbook.Props.ModifiedDate){
            dataDoRelatorio = new Date(workbook.Props.ModifiedDate);
        } else if(lastModifiedHeader){
            dataDoRelatorio = new Date(lastModifiedHeader);
        }

        atualizarData(dataDoRelatorio);

        // Popula as opções dos filtros com base nos dados importados
        popularFiltros(dadosOriginais);

        // Aplica os filtros (que, sem seleção, mostram tudo)
        aplicarFiltros();

    } catch(erro){

        console.error("Erro ao carregar " + NOME_ARQUIVO_RELATORIO + ":", erro);

        alert(
            "Não foi possível carregar o arquivo \"" + NOME_ARQUIVO_RELATORIO + "\".\n\n" +
            "Verifique se:\n" +
            "1) o arquivo \"" + NOME_ARQUIVO_RELATORIO + "\" está na mesma pasta do dashboard;\n" +
            "2) a página está sendo acessada por um endereço http(s) (ex: SharePoint), " +
            "e não aberta direto do computador."
        );

    } finally {

        botaoAtualizar.disabled = false;
        botaoAtualizar.classList.remove("carregando");

    }

}

// ==========================================================
// CLASSIFICAÇÃO: ANALISTA -> SETOR
// ==========================================================
// Mapa de qual setor cada analista pertence. Se um analista não
// estiver aqui, cai em "Não Classificado" (não quebra o dashboard,
// só sinaliza que falta cadastrar esse nome na lista abaixo).

const MAPA_SETOR_POR_ANALISTA = {
    "Agnes Cristina Costa dos Santos": "Comercial",
    "Lidia Escocio Pereira Zenzeluk": "Comercial",
    "Keilla Suzane Cotrin": "Comercial",

    "Franckllin Pereira": "Gestão de Ativos",
    "Marcos barbosa dos santos": "Gestão de Ativos",
    "Guilherme Lima Giroldo": "Gestão de Ativos",
    "Elder Placeres de Luis Junior": "Gestão de Ativos",

    "Suelen Marin Machado": "Centro de Comando",
    "Suelen Anton De Campos De Farias": "Centro de Comando",
    "Priscila Alves da Silva": "Centro de Comando",
    "Anne Karoline Aparecida Kinapp": "Centro de Comando"
};

// Remove acentos, espaços extras e diferenças de maiúsculas/minúsculas,
// para que a comparação funcione mesmo com pequenas variações de digitação
function normalizarNomeAnalista(nome){

    return String(nome || "")
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ");

}

const MAPA_SETOR_NORMALIZADO = {};

Object.entries(MAPA_SETOR_POR_ANALISTA).forEach(([nome, setor])=>{
    MAPA_SETOR_NORMALIZADO[normalizarNomeAnalista(nome)] = setor;
});

function obterSetorPorAnalista(nomeAnalista){

    const chave = normalizarNomeAnalista(nomeAnalista);

    return MAPA_SETOR_NORMALIZADO[chave] || "Não Classificado";

}

// Carrega os dados automaticamente assim que a página abre
carregarRelatorio();

// ==========================================================
// DATA DA ÚLTIMA ATUALIZAÇÃO
// ==========================================================

function atualizarData(dataDoRelatorio){

    const dataValida =
        (dataDoRelatorio instanceof Date) && !isNaN(dataDoRelatorio)
            ? dataDoRelatorio
            : null;

    if(!dataValida){

        document.getElementById("ultimaAtualizacao").innerText =
            "Data do relatório não disponível";

        return;

    }

    const texto =
        dataValida.toLocaleDateString("pt-BR") +
        " às " +
        dataValida.toLocaleTimeString("pt-BR",{
            hour:"2-digit",
            minute:"2-digit"
        });

    document.getElementById("ultimaAtualizacao").innerText =
        "Relatório de " + texto;

}

// ==========================================================
// CORREÇÃO DE COLUNAS DE DATA
// ==========================================================
// Como lemos a planilha com cellDates:false, nenhuma coluna vira Date
// automaticamente. Aqui convertemos manualmente só as colunas cujo
// NOME indica que são datas (contém "data" no nome), preservando como
// número qualquer outra coluna (ex: número do chamado, IDs, etc).
//
// Se alguma coluna de data da sua planilha não tiver "data" no nome,
// adicione o nome exato dela em COLUNAS_DE_DATA_EXTRA abaixo.

const COLUNAS_DE_DATA_EXTRA = [
    // "Nome_Exato_Da_Coluna",
];

// ==========================================================
// CORREÇÃO ESPECÍFICA: COLUNA "Chamado" (ex: 1-1, 3-1, 13-1...)
// ==========================================================
// O Excel, sozinho, interpreta textos como "1-1", "3-1", "5-1" (mês
// válido de 1 a 12) como DATAS (1-Jan, 1-Mar, 1-May...) e guarda por
// baixo um número de série. Já "13-1", "15-1" etc. o Excel não
// consegue interpretar como data (não existe mês 13), então ficam
// como texto mesmo, corretos.
//
// Aqui, para essas colunas, quando o valor cru vier como número
// (ou seja, foi convertido em data pelo Excel), reconstruímos o texto
// original juntando MÊS-DIA da data quebrada.

const COLUNAS_NUMERO_CHAMADO = ["Chamado"];

function reconstruirNumeroChamado(valorRaw){

    const data = excelSerialParaData(valorRaw);

    const mes = data.getUTCMonth() + 1;
    const dia = data.getUTCDate();

    return `${mes}-${dia}`;

}

function ehColunaDeNumeroChamado(nomeColuna){
    return COLUNAS_NUMERO_CHAMADO.includes(nomeColuna);
}

function ehColunaDeData(nomeColuna){

    if(COLUNAS_DE_DATA_EXTRA.includes(nomeColuna)) return true;

    return /data/i.test(nomeColuna);

}

// Converte um número de série do Excel (dias desde 30/12/1899) em Date
function excelSerialParaData(serial){

    const dataBase = new Date(Date.UTC(1899, 11, 30));

    dataBase.setUTCDate(dataBase.getUTCDate() + Math.floor(serial));

    // Parte fracionária do serial = horário do dia
    const fracaoDoDia = serial - Math.floor(serial);

    if(fracaoDoDia > 0){
        dataBase.setUTCMilliseconds(
            dataBase.getUTCMilliseconds() + Math.round(fracaoDoDia * 86400000)
        );
    }

    return dataBase;

}

// ==========================================================
// CARD "CHAMADOS (7 DIAS)" + GRÁFICO 7 DIAS POR CAUSA
// ==========================================================
// Precisamos saber qual coluna da planilha guarda a DATA DE ABERTURA
// do chamado, para calcular quantos chamados foram abertos nos
// últimos 7 dias.
//
// AJUSTE AQUI o nome exato da coluna de data de abertura da sua
// planilha, caso não seja "Data_Abertura":

const COLUNA_DATA_ABERTURA = "Data_de_Abertura";

// Se a coluna configurada acima não existir na planilha carregada,
// tentamos detectar automaticamente a primeira coluna que contenha
// valores do tipo Date (normalmente qualquer coluna com "data" no
// nome, já convertida em corrigirColunasDeData).
function obterColunaDataAbertura(dados){

    if(dados.length === 0) return null;

    if(dados[0][COLUNA_DATA_ABERTURA] instanceof Date){
        return COLUNA_DATA_ABERTURA;
    }

    const colunas = Object.keys(dados[0]);

    const candidata = colunas.find(coluna=> dados[0][coluna] instanceof Date);

    return candidata || null;

}

// Retorna apenas os itens cuja data de abertura está dentro dos
// últimos 7 dias (contando a partir de agora)
function filtrarUltimos7Dias(dados){

    const coluna = obterColunaDataAbertura(dados);

    if(!coluna) return [];

    const agora = new Date();

    const limite = new Date();
    limite.setDate(limite.getDate() - 7);

    return dados.filter(item=>{

        const valor = item[coluna];

        if(!(valor instanceof Date) || isNaN(valor)) return false;

        return valor >= limite && valor <= agora;

    });

}

function corrigirColunasDeData(json){

    json.forEach(linha=>{

        Object.keys(linha).forEach(coluna=>{

            const valor = linha[coluna];

            if(typeof valor !== "number") return;

            if(!ehColunaDeData(coluna)) return;

            linha[coluna] = excelSerialParaData(valor);

        });

    });

}

// ==========================================================
// KPIs
// ==========================================================

function atualizarKPIs(dados){

    let total = dados.length;

    let aguardando = 0;
    let atendimento = 0;
    let finalizados = 0;
    let cancelados = 0;

    dados.forEach(item=>{

        const estado = String(item.Estado || "")
            .trim()
            .toUpperCase();

        switch(estado){

            case "AGUARDANDO ATENDIMENTO":
                aguardando++;
                break;

            case "EM ATENDIMENTO":
                atendimento++;
                break;

            case "CANCELADO":
                cancelados++;
                break;

            case "FECHADO":
            case "RESOLVIDO":
                finalizados++;
                break;

        }

    });

    document.getElementById("totalChamados").innerText = total;
    document.getElementById("aguardando").innerText = aguardando;
    document.getElementById("atendimento").innerText = atendimento;
    document.getElementById("finalizados").innerText = finalizados;
    document.getElementById("cancelados").innerText = cancelados;

}

// ==========================================================
// FILTROS ESTILO EXCEL (MULTI-SELEÇÃO)
// ==========================================================

class MultiSelectFiltro{

    constructor(containerId){

        this.container = document.getElementById(containerId);
        this.toggle = this.container.querySelector(".ms-toggle");
        this.panel = this.container.querySelector(".ms-panel");
        this.optionsEl = this.container.querySelector(".ms-options");
        this.searchEl = this.container.querySelector(".ms-search input");

        this.todosValores = [];
        this.selecionados = new Set();

        this._bind();

    }

    // Define a lista de valores possíveis (chamado ao carregar a planilha)
    setValores(valores){

        this.todosValores = valores;
        this.selecionados = new Set(valores); // por padrão, tudo selecionado

        this._render();
        this._atualizarRotulo();

    }

    // Um item passa no filtro se: não há valores carregados ainda,
    // ou o valor dele está entre os selecionados
    filtra(valor){

        if(this.todosValores.length === 0) return true;

        return this.selecionados.has(valor);

    }

    _render(textoBusca = ""){

        this.optionsEl.innerHTML = "";

        const filtrados = this.todosValores.filter(v=>
            v.toLowerCase().includes(textoBusca.toLowerCase())
        );

        if(filtrados.length === 0){

            this.optionsEl.innerHTML = '<div class="ms-vazio">Nenhum resultado</div>';
            return;

        }

        filtrados.forEach(valor=>{

            const label = document.createElement("label");
            label.className = "ms-option";

            const input = document.createElement("input");
            input.type = "checkbox";
            input.value = valor;
            input.checked = this.selecionados.has(valor);

            input.addEventListener("change", ()=>{

                if(input.checked){
                    this.selecionados.add(valor);
                } else {
                    this.selecionados.delete(valor);
                }

                this._atualizarRotulo();
                aplicarFiltros();

            });

            const span = document.createElement("span");
            span.textContent = valor;

            label.appendChild(input);
            label.appendChild(span);

            this.optionsEl.appendChild(label);

        });

    }

    _atualizarRotulo(){

        const textoEl = this.toggle.firstChild;

        if(this.todosValores.length === 0 || this.selecionados.size === this.todosValores.length){
            textoEl.textContent = "Todos ";
        } else if(this.selecionados.size === 0){
            textoEl.textContent = "Nenhum ";
        } else {
            textoEl.textContent = `${this.selecionados.size} selecionados `;
        }

    }

    marcarTodos(){

        this.selecionados = new Set(this.todosValores);
        this._render(this.searchEl.value);
        this._atualizarRotulo();
        aplicarFiltros();

    }

    desmarcarTodos(){

        this.selecionados.clear();
        this._render(this.searchEl.value);
        this._atualizarRotulo();
        aplicarFiltros();

    }

    _bind(){

        this.toggle.addEventListener("click", (e)=>{

            e.stopPropagation();

            const jaAberto = this.container.classList.contains("aberto");

            document.querySelectorAll(".multiselect.aberto").forEach(el=>{
                el.classList.remove("aberto");
            });

            if(!jaAberto){
                this.container.classList.add("aberto");
            }

        });

        this.container.querySelector(".ms-all").addEventListener("click", ()=> this.marcarTodos());
        this.container.querySelector(".ms-none").addEventListener("click", ()=> this.desmarcarTodos());

        this.searchEl.addEventListener("input", (e)=> this._render(e.target.value));

        this.container.addEventListener("click", (e)=> e.stopPropagation());

    }

}

// Fecha qualquer dropdown aberto ao clicar fora
document.addEventListener("click", ()=>{
    document.querySelectorAll(".multiselect.aberto").forEach(el=>{
        el.classList.remove("aberto");
    });
});

const msMes = new MultiSelectFiltro("msMes");
const msSetor = new MultiSelectFiltro("msSetor");
const msLocalidade = new MultiSelectFiltro("msLocalidade");
const msCausa = new MultiSelectFiltro("msCausa");
const msStatus = new MultiSelectFiltro("msStatus");

// Status tem opções fixas, não depende da planilha carregada
msStatus.setValores([
    "AGUARDANDO ATENDIMENTO",
    "EM ATENDIMENTO",
    "FECHADO",
    "RESOLVIDO",
    "CANCELADO"
]);

const botaoLimparFiltros = document.getElementById("limparFiltros");

// Ordem cronológica correta para os meses (evita ordenação alfabética errada)
const ORDEM_MESES = ["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"];

function ordenarMeses(meses){

    return meses.sort((a, b)=>{

        const idxA = ORDEM_MESES.indexOf(String(a).trim().toUpperCase());
        const idxB = ORDEM_MESES.indexOf(String(b).trim().toUpperCase());

        if(idxA === -1 && idxB === -1) return String(a).localeCompare(String(b));
        if(idxA === -1) return 1;
        if(idxB === -1) return -1;

        return idxA - idxB;

    });

}

function popularFiltros(dados){

    const meses = ordenarMeses([...new Set(
        dados.map(item=>item.Data_de_Abertura_Mes_texto).filter(Boolean)
    )]);

    const setores = [...new Set(
        dados.map(item=>item.Setor).filter(Boolean)
    )].sort();

    const localidades = [...new Set(
        dados.map(item=>item.Localidade_do_Solicitante).filter(Boolean)
    )].sort();

    const causas = [...new Set(
        dados.map(item=>item.Causa).filter(Boolean)
    )].sort();

    msMes.setValores(meses);
    msSetor.setValores(setores);
    msLocalidade.setValores(localidades);
    msCausa.setValores(causas);

    // Status tem opções fixas definidas acima, não precisa repopular

}

function aplicarFiltros(){

    const dadosFiltrados = dadosOriginais.filter(item=>{

        if(!msMes.filtra(item.Data_de_Abertura_Mes_texto)) return false;

        if(!msSetor.filtra(item.Setor)) return false;

        if(!msLocalidade.filtra(item.Localidade_do_Solicitante)) return false;

        if(!msCausa.filtra(item.Causa)) return false;

        const estado = String(item.Estado || "").trim().toUpperCase();

        if(!msStatus.filtra(estado)) return false;

        return true;

    });

    // Chamados dos últimos 7 dias, já respeitando os filtros ativos
    const dadosUltimos7Dias = filtrarUltimos7Dias(dadosFiltrados);

    // Atualiza KPIs, gráficos e a tabela com base apenas nos dados filtrados
    atualizarKPIs(dadosFiltrados);

    document.getElementById("kpiChamados7Dias").innerText = dadosUltimos7Dias.length;

    criarGrafico7Dias(dadosUltimos7Dias);
    criarGraficoLocalidades(dadosFiltrados);
    criarGraficoCausas(dadosFiltrados);
    criarGraficoStatus(dadosFiltrados);

    atualizarTabela(dadosFiltrados);

}

// Botão "Limpar Filtros" -> volta tudo para "Todos selecionados"
botaoLimparFiltros.addEventListener("click", ()=>{

    msMes.marcarTodos();
    msSetor.marcarTodos();
    msLocalidade.marcarTodos();
    msCausa.marcarTodos();
    msStatus.marcarTodos();

    aplicarFiltros();

});

// ==========================================================
// TABELA DE DADOS FILTRADOS + EXPORTAÇÃO PARA EXCEL
// ==========================================================

let dadosFiltradosAtual = [];

// Estado atual de ordenação da tabela (qual coluna e direção)
const ordenacaoAtual = { coluna: null, direcao: null }; // direcao: "asc" | "desc"

// Compara dois valores de célula de forma inteligente: números como número,
// datas como data, e o resto como texto (ordem alfabética em pt-BR)
function compararValoresTabela(a, b){

    const aVazio = (a === undefined || a === null || a === "");
    const bVazio = (b === undefined || b === null || b === "");

    if(aVazio && bVazio) return 0;
    if(aVazio) return -1;
    if(bVazio) return 1;

    if(a instanceof Date && b instanceof Date){
        return a.getTime() - b.getTime();
    }

    const ehNumerico = (valor)=>{
        return typeof valor === "number" ||
            /^-?\d+([.,]\d+)?$/.test(String(valor).trim());
    };

    if(ehNumerico(a) && ehNumerico(b)){

        const numA = typeof a === "number" ? a : parseFloat(String(a).replace(",", "."));
        const numB = typeof b === "number" ? b : parseFloat(String(b).replace(",", "."));

        return numA - numB;

    }

    return String(a).localeCompare(String(b), "pt-BR", { sensitivity: "base" });

}

// Retorna uma cópia de "dados" ordenada conforme ordenacaoAtual
function ordenarDadosTabela(dados){

    if(!ordenacaoAtual.coluna) return dados;

    const copia = [...dados].sort((itemA, itemB)=>{

        const resultado = compararValoresTabela(
            itemA[ordenacaoAtual.coluna],
            itemB[ordenacaoAtual.coluna]
        );

        return ordenacaoAtual.direcao === "asc" ? resultado : -resultado;

    });

    return copia;

}

// Clique no cabeçalho: 1º clique = maior->menor, 2º clique (mesma coluna)
// = menor->maior, 3º clique = remove a ordenação
function alternarOrdenacaoTabela(coluna){

    if(ordenacaoAtual.coluna === coluna){

        if(ordenacaoAtual.direcao === "desc"){
            ordenacaoAtual.direcao = "asc";
        } else {
            ordenacaoAtual.coluna = null;
            ordenacaoAtual.direcao = null;
        }

    } else {

        ordenacaoAtual.coluna = coluna;
        ordenacaoAtual.direcao = "desc";

    }

    atualizarTabela(dadosFiltradosAtual);

}

// Formata o valor de uma célula para exibição na tabela.
// Datas (vindas do Excel com cellDates:true) viram texto no padrão pt-BR.
function formatarValorCelula(valor){

    if(valor instanceof Date){

        const dataFormatada = valor.toLocaleDateString("pt-BR");

        const temHorario = valor.getHours() !== 0 || valor.getMinutes() !== 0;

        if(temHorario){

            const horaFormatada = valor.toLocaleTimeString("pt-BR",{
                hour:"2-digit",
                minute:"2-digit"
            });

            return dataFormatada + " " + horaFormatada;

        }

        return dataFormatada;

    }

    return valor ?? "";

}

function atualizarTabela(dados){

    dadosFiltradosAtual = dados;

    // Aplica a ordenação (se houver alguma coluna selecionada) antes de exibir
    const dadosOrdenados = ordenarDadosTabela(dados);

    document.getElementById("contadorLinhas").innerText = dadosOrdenados.length;

    const head = document.getElementById("tabelaHead");
    const body = document.getElementById("tabelaBody");

    head.innerHTML = "";
    body.innerHTML = "";

    if(dadosOrdenados.length === 0){

        body.innerHTML = '<tr><td class="tabela-vazia" colspan="100">Nenhum dado encontrado</td></tr>';
        return;

    }

    // Monta a lista de colunas com base nas chaves presentes nos dados,
    // preservando a ordem em que aparecem na planilha
    const colunas = [];

    dadosOrdenados.forEach(item=>{

        Object.keys(item).forEach(chave=>{

            if(!colunas.includes(chave)) colunas.push(chave);

        });

    });

    // Coloca "Chamado" e "Numero de Serie form" sempre no início da
    // tabela, nessa ordem, mantendo as demais colunas como estavam
    const ORDEM_PRIORITARIA_COLUNAS = ["Chamado", "Numero de Serie form", "Analista_Atual", "Setor"];

    colunas.sort((a, b)=>{

        const idxA = ORDEM_PRIORITARIA_COLUNAS.indexOf(a);
        const idxB = ORDEM_PRIORITARIA_COLUNAS.indexOf(b);

        if(idxA === -1 && idxB === -1) return 0;
        if(idxA === -1) return 1;
        if(idxB === -1) return -1;

        return idxA - idxB;

    });

    const RENOMEAR_CABECALHO = {
        "Numero de Serie form": "Nº de Série"
    };

    colunas.forEach(coluna=>{

        const th = document.createElement("th");
        th.classList.add("th-ordenavel");

        const rotulo = document.createElement("span");
        rotulo.className = "th-rotulo";
        rotulo.textContent = RENOMEAR_CABECALHO[coluna] || coluna.replace(/_/g, " ");

        const seta = document.createElement("span");
        seta.className = "th-seta";

        if(ordenacaoAtual.coluna === coluna){
            th.classList.add("th-ordenado");
            seta.textContent = ordenacaoAtual.direcao === "desc" ? "▼" : "▲";
        }

        th.appendChild(rotulo);
        th.appendChild(seta);

        th.addEventListener("click", ()=> alternarOrdenacaoTabela(coluna));

        head.appendChild(th);

    });

    const fragmento = document.createDocumentFragment();

    dadosOrdenados.forEach(item=>{

        const tr = document.createElement("tr");

        colunas.forEach(coluna=>{

            const td = document.createElement("td");
            td.textContent = formatarValorCelula(item[coluna]);
            tr.appendChild(td);

        });

        fragmento.appendChild(tr);

    });

    body.appendChild(fragmento);

}

document.getElementById("baixarExcel").addEventListener("click", ()=>{

    if(dadosFiltradosAtual.length === 0){

        alert("Não há dados filtrados para exportar.");
        return;

    }

    const dadosParaExportar = ordenarDadosTabela(dadosFiltradosAtual);

    const planilha = XLSX.utils.json_to_sheet(dadosParaExportar);
    const livro = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(livro, planilha, "Dados Filtrados");

    const dataArquivo = new Date().toISOString().slice(0, 10);

    XLSX.writeFile(livro, `chamados_filtrados_${dataArquivo}.xlsx`);

});

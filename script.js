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
// FILTRO POR MÊS
// ==========================================================
// Usa a mesma coluna de data de abertura já detectada acima
// (COLUNA_DATA_ABERTURA / obterColunaDataAbertura) para agrupar
// os chamados por "Mês/Ano" (ex: "Junho/2026").

const NOMES_MESES = [
    "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
    "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"
];

function obterRotuloMes(data){
    return `${NOMES_MESES[data.getMonth()]}/${data.getFullYear()}`;
}

// Retorna o rótulo de mês ("Junho/2026") de um item específico,
// ou null se o item não tiver uma data de abertura válida
function obterRotuloMesDoItem(item, coluna){

    if(!coluna) return null;

    const valor = item[coluna];

    if(!(valor instanceof Date) || isNaN(valor)) return null;

    return obterRotuloMes(valor);

}

// Lista, em ordem cronológica, todos os meses presentes na planilha
function obterMesesDisponiveis(dados){

    const coluna = obterColunaDataAbertura(dados);

    if(!coluna) return [];

    const mapa = new Map(); // chave "AAAA-MM" -> { ano, mes, rotulo }

    dados.forEach(item=>{

        const valor = item[coluna];

        if(!(valor instanceof Date) || isNaN(valor)) return;

        const chave = `${valor.getFullYear()}-${String(valor.getMonth()).padStart(2,"0")}`;

        if(!mapa.has(chave)){

            mapa.set(chave, {
                ano: valor.getFullYear(),
                mes: valor.getMonth(),
                rotulo: obterRotuloMes(valor)
            });

        }

    });

    return [...mapa.values()]
        .sort((a,b)=> (a.ano - b.ano) || (a.mes - b.mes))
        .map(item=>item.rotulo);

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
const msAnalista = new MultiSelectFiltro("msAnalista");
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

function popularFiltros(dados){

    const meses = obterMesesDisponiveis(dados);

    msMes.setValores(meses);

    const analistas = [...new Set(
        dados.map(item=>item.Analista_Atual).filter(Boolean)
    )].sort();

    const localidades = [...new Set(
        dados.map(item=>item.Localidade_do_Solicitante).filter(Boolean)
    )].sort();

    const causas = [...new Set(
        dados.map(item=>item.Causa).filter(Boolean)
    )].sort();

    msAnalista.setValores(analistas);
    msLocalidade.setValores(localidades);
    msCausa.setValores(causas);

    // Status tem opções fixas definidas acima, não precisa repopular

}

function aplicarFiltros(){

    const colunaDataMes = obterColunaDataAbertura(dadosOriginais);

    const dadosFiltrados = dadosOriginais.filter(item=>{

        const rotuloMes = obterRotuloMesDoItem(item, colunaDataMes);

        if(!msMes.filtra(rotuloMes)) return false;

        if(!msAnalista.filtra(item.Analista_Atual)) return false;

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
    msAnalista.marcarTodos();
    msLocalidade.marcarTodos();
    msCausa.marcarTodos();
    msStatus.marcarTodos();

    aplicarFiltros();

});

// ==========================================================
// TABELA DE DADOS FILTRADOS + EXPORTAÇÃO PARA EXCEL
// ==========================================================

let dadosFiltradosAtual = [];

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

    document.getElementById("contadorLinhas").innerText = dados.length;

    const head = document.getElementById("tabelaHead");
    const body = document.getElementById("tabelaBody");

    head.innerHTML = "";
    body.innerHTML = "";

    if(dados.length === 0){

        body.innerHTML = '<tr><td class="tabela-vazia" colspan="100">Nenhum dado encontrado</td></tr>';
        return;

    }

    // Monta a lista de colunas com base nas chaves presentes nos dados,
    // preservando a ordem em que aparecem na planilha
    const colunas = [];

    dados.forEach(item=>{

        Object.keys(item).forEach(chave=>{

            if(!colunas.includes(chave)) colunas.push(chave);

        });

    });

    colunas.forEach(coluna=>{

        const th = document.createElement("th");
        th.textContent = coluna.replace(/_/g, " ");
        head.appendChild(th);

    });

    const fragmento = document.createDocumentFragment();

    dados.forEach(item=>{

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

    const planilha = XLSX.utils.json_to_sheet(dadosFiltradosAtual);
    const livro = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(livro, planilha, "Dados Filtrados");

    const dataArquivo = new Date().toISOString().slice(0, 10);

    XLSX.writeFile(livro, `chamados_filtrados_${dataArquivo}.xlsx`);

});

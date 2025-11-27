// =========================================================================
// 1. CONFIGURAÇÃO & ESTRUTURA
// =========================================================================

const RUAS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15];

const CONFIG = {
    arquivos: {
        enderecos: 'Endereços.xlsx',
        convocacoes: 'Convocações Matriz - Mapeamento.xlsx',
        abastecimento: 'Abastecimento-mapeamento.xlsx'
    },
    colsEnderecos: {
        rua: 'Rua', predio: 'Prédio', nivel: 'Nível', apto: 'Apartamento',
        tipoEndereco: 'Tipo de endereço' 
    },
    colsConvocacoes: {
        rua: 'Endereço de WMS - Parte 1', predio: 'Endereço de WMS - Parte 2',
        nivel: 'Endereço de WMS - Parte 3', apto: 'Endereço de WMS - Parte 4',
        tipoMov: 'Tipo de movimento', 
        data: 'Data/hora inserção' 
    },
    colsAbastecimento: {
        rua: 'Rua', predio: 'Prédio', nivel: 'Nível', apto: 'Apartamento',
        tipoMov: 'Tipo da Transferência', 
        data: 'Data/hora cadastro'
    }
};

// ===================================================================
// 2. LÓGICA DE CORES E HELPERS
// ===================================================================
const corMinimo = { r: 255, g: 255, b: 255 };
const corBaixo = { r: 255, g: 255, b: 0 };
const corMedio = { r: 255, g: 165, b: 0 };
const corAlto = { r: 255, g: 69, b: 0 };
const corMaximo = { r: 255, g: 0, b: 0 };

function interpolarCor(cor1, cor2, fator) {
    const r = Math.round(cor1.r + (cor2.r - cor1.r) * fator);
    const g = Math.round(cor1.g + (cor2.g - cor1.g) * fator);
    const b = Math.round(cor1.b + (cor2.b - cor1.b) * fator);
    return `rgb(${r}, ${g}, ${b})`;
}

function calcularCor(valor, min, max) {
    if (!valor || valor === 0) return 'rgba(200, 200, 200, 0.3)';
    const range = max - min;
    const normalized = range === 0 ? 0.5 : (valor - min) / range;

    if (normalized < 0.25) return interpolarCor(corMinimo, corBaixo, normalized / 0.25);
    else if (normalized < 0.5) return interpolarCor(corBaixo, corMedio, (normalized - 0.25) / 0.25);
    else if (normalized < 0.75) return interpolarCor(corMedio, corAlto, (normalized - 0.5) / 0.25);
    else return interpolarCor(corAlto, corMaximo, (normalized - 0.75) / 0.25);
}

const formatarNumero = (n) => (n === undefined || n === null) ? '-' : n.toString().replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, ".");
const normalizar = (v) => (v === undefined || v === null) ? "" : String(v).replace('.0', '').trim();

const processarData = (v) => {
    if (!v) return null;
    if (typeof v === 'number') {
        return new Date(Math.round((v - 25569) * 86400 * 1000));
    }
    if (typeof v === 'string') {
        if (v.includes('/')) {
            const partes = v.split(' ');
            const dataPartes = partes[0].split('/');
            if (dataPartes.length === 3) {
                return new Date(dataPartes[2], dataPartes[1] - 1, dataPartes[0]);
            }
        }
    }
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
};

// ===================================================================
// 3. ESTRUTURA FÍSICA E VALIDAÇÃO VISUAL
// ===================================================================
let mapaVisualValido = new Set(); 

function getEstruturaPredios() {
    const estrutura = {};
    mapaVisualValido = new Set();

    for (let rua of RUAS) {
        let prediosImpar = [];
        let prediosPar = [];
        
        if (rua === 1) {
            for (let i = 41; i <= 69; i += 2) prediosImpar.push(i);
            prediosPar = Array.from({ length: 5 }, (_, i) => 48 + i * 2);
        } else if (rua === 2) {
            for (let i = 3; i <= 46; i += 2) prediosImpar.push(i);
            for (let i = 2; i <= 46; i += 2) prediosPar.push(i);
        } else {
            for (let i = 1; i <= 46; i += 2) prediosImpar.push(i);
            for (let i = 2; i <= 46; i += 2) prediosPar.push(i);
        }
        
        estrutura[rua] = { Impar: prediosImpar, Par: prediosPar };

        [...prediosImpar, ...prediosPar].forEach(predio => {
            mapaVisualValido.add(`${rua}-${predio}`);
        });
    }
    return estrutura;
}

getEstruturaPredios(); 

let dadosConsolidados = [];
let filtrosAtivos = { tipoMovimento: '', picking: '', curva: '', dataInicial: null, dataFinal: null };
let predioSelecionado = null; 

// ===================================================================
// 4. CARREGAMENTO E PROCESSAMENTO
// ===================================================================
async function carregarExcel(url) {
    try {
        const res = await fetch(url + `?t=${new Date().getTime()}`);
        if (!res.ok) throw new Error(`Erro ${res.status}`);
        const ab = await res.arrayBuffer();
        return XLSX.read(ab, { type: 'array' });
    } catch (e) { console.error(`Falha ao carregar ${url}`, e); return null; }
}

function buscarDadosValidos(wb, nomeArquivo, colChave) {
    if (!wb) return [];
    for (const nomeAba of wb.SheetNames) {
        const dados = XLSX.utils.sheet_to_json(wb.Sheets[nomeAba]);
        if (dados.length > 0 && dados[0][colChave] !== undefined) {
            console.log(`[${nomeArquivo}] Aba carregada: ${nomeAba}. Linhas: ${dados.length}`);
            return dados;
        }
    }
    return [];
}

async function iniciarSistema() {
    const loading = document.getElementById('loading');
    const erroDiv = document.getElementById('errorMessage');
    loading.classList.add('active');
    
    getEstruturaPredios();

    const [wbEnd, wbConv, wbAbast] = await Promise.all([
        carregarExcel(CONFIG.arquivos.enderecos),
        carregarExcel(CONFIG.arquivos.convocacoes),
        carregarExcel(CONFIG.arquivos.abastecimento)
    ]);

    if (!wbEnd) {
        erroDiv.textContent = "Erro Crítico: Arquivo 'Endereços.xlsx' não encontrado.";
        erroDiv.classList.add('active');
        loading.classList.remove('active');
        return;
    }

    const rawEnd = buscarDadosValidos(wbEnd, "Endereços", CONFIG.colsEnderecos.rua);
    const mapaRef = {}; 

    rawEnd.forEach(row => {
        const r = normalizar(row[CONFIG.colsEnderecos.rua]);
        const p = normalizar(row[CONFIG.colsEnderecos.predio]);
        const n = normalizar(row[CONFIG.colsEnderecos.nivel]);
        const a = normalizar(row[CONFIG.colsEnderecos.apto]);
        
        if (r && p) {
            const key = `${r}-${p}-${n}-${a}`;
            const tipo = row[CONFIG.colsEnderecos.tipoEndereco] || "";
            mapaRef[key] = {
                picking: tipo.toLowerCase().includes('picking') ? "Sim" : "Não"
            };
        }
    });

    dadosConsolidados = [];

    const processarTabela = (rows, cols) => {
        rows.forEach(row => {
            const r = normalizar(row[cols.rua]);
            const p = normalizar(row[cols.predio]);
            const n = normalizar(row[cols.nivel]);
            const a = normalizar(row[cols.apto]);

            if (r && p) {
                const key = `${r}-${p}-${n}-${a}`;
                const info = mapaRef[key] || { picking: 'Indefinido' };
                const dataProc = processarData(row[cols.data]);

                dadosConsolidados.push({
                    rua: parseInt(r) || 0,
                    predio: parseInt(p) || 0,
                    nivel: n,
                    apto: a,
                    tipoMovimento: row[cols.tipoMov] || "Outros",
                    data: dataProc,
                    picking: info.picking
                });
            }
        });
    };

    if(wbConv) processarTabela(buscarDadosValidos(wbConv, "Convocações", CONFIG.colsConvocacoes.rua), CONFIG.colsConvocacoes);
    if(wbAbast) processarTabela(buscarDadosValidos(wbAbast, "Abastecimento", CONFIG.colsAbastecimento.rua), CONFIG.colsAbastecimento);
    
    atualizarFiltrosUI();
    loading.classList.remove('active');
    renderizarMapa();
}

// ===================================================================
// 5. FILTRAGEM E CÁLCULOS ABC (ENDEREÇO E PRÉDIO)
// ===================================================================

function getDadosBasicos() {
    return dadosConsolidados.filter(item => {
        const keyVisual = `${item.rua}-${item.predio}`;
        if (!mapaVisualValido.has(keyVisual)) return false;

        const matchTipo = !filtrosAtivos.tipoMovimento || item.tipoMovimento === filtrosAtivos.tipoMovimento;
        const matchPicking = !filtrosAtivos.picking || item.picking === filtrosAtivos.picking;
        
        let matchData = true;
        if ((filtrosAtivos.dataInicial || filtrosAtivos.dataFinal) && !item.data) return false;

        if (item.data) {
            if (filtrosAtivos.dataInicial) {
                const dtIni = new Date(filtrosAtivos.dataInicial);
                dtIni.setHours(0,0,0,0);
                if (item.data < dtIni) matchData = false;
            }
            if (filtrosAtivos.dataFinal && matchData) {
                const dtFim = new Date(filtrosAtivos.dataFinal);
                dtFim.setHours(23, 59, 59, 999);
                if (item.data > dtFim) matchData = false;
            }
        }
        
        return matchTipo && matchPicking && matchData;
    });
}

// --- CLASSIFICAÇÃO 1: POR ENDEREÇO (Para Tooltips/Detalhes) ---
function calcularClassificacaoABC_Endereco(dados) {
    const contagem = {};
    let totalMovimentos = 0;

    dados.forEach(item => {
        const key = `${item.rua}-${item.predio}-${item.nivel}-${item.apto}`;
        contagem[key] = (contagem[key] || 0) + 1;
        totalMovimentos++;
    });

    return classificarPorPareto(contagem, totalMovimentos);
}

// --- CLASSIFICAÇÃO 2: POR PRÉDIO (Para Filtro "Curva") ---
function calcularClassificacaoABC_Predio(dados) {
    const contagem = {};
    let totalMovimentos = 0;

    dados.forEach(item => {
        const key = `${item.rua}-${item.predio}`; // Chave é apenas o prédio
        contagem[key] = (contagem[key] || 0) + 1;
        totalMovimentos++;
    });

    return classificarPorPareto(contagem, totalMovimentos);
}

// Lógica de Pareto Genérica (Reutilizável)
function classificarPorPareto(contagemObj, totalMovimentos) {
    // Agrupa por quantidade (para tratar empates)
    const gruposPorQtd = {};
    Object.entries(contagemObj).forEach(([key, qtd]) => {
        if (!gruposPorQtd[qtd]) gruposPorQtd[qtd] = [];
        gruposPorQtd[qtd].push(key);
    });

    // Ordena quantidades (Maior -> Menor)
    const quantidadesOrdenadas = Object.keys(gruposPorQtd).map(Number).sort((a, b) => b - a);

    const mapaClassificacao = {}; 
    let acumulado = 0;

    quantidadesOrdenadas.forEach(qtd => {
        const listaItens = gruposPorQtd[qtd];
        const volumeDoGrupo = qtd * listaItens.length;
        
        // Verifica INÍCIO do grupo no acumulado para ser justo com os maiores
        const percInicio = totalMovimentos === 0 ? 0 : (acumulado / totalMovimentos);
        
        let classeGrupo = 'C';
        if (percInicio < 0.80) {
            classeGrupo = 'A';
        } else if (percInicio < 0.95) {
            classeGrupo = 'B';
        } else {
            classeGrupo = 'C';
        }

        listaItens.forEach(key => {
            mapaClassificacao[key] = classeGrupo;
        });

        acumulado += volumeDoGrupo;
    });

    return mapaClassificacao;
}

// ===================================================================
// 6. RENDERIZAÇÃO E ATUALIZAÇÃO
// ===================================================================

function atualizarTituloPeriodo() {
    const span = document.getElementById('subtituloData');
    if(!span) return;

    if (filtrosAtivos.dataInicial || filtrosAtivos.dataFinal) {
        const f = (d) => d ? d.toLocaleDateString('pt-BR') : '';
        if (filtrosAtivos.dataInicial && filtrosAtivos.dataFinal) {
            span.textContent = `Período Filtrado: ${f(filtrosAtivos.dataInicial)} até ${f(filtrosAtivos.dataFinal)}`;
        } else if (filtrosAtivos.dataInicial) {
            span.textContent = `Período Filtrado: A partir de ${f(filtrosAtivos.dataInicial)}`;
        } else {
            span.textContent = `Período Filtrado: Até ${f(filtrosAtivos.dataFinal)}`;
        }
        return;
    }

    if (dadosConsolidados.length === 0) {
        span.textContent = "Aguardando dados...";
        return;
    }

    const mesesMap = new Map();
    dadosConsolidados.forEach(d => {
        if (d.data) {
            const ano = d.data.getFullYear();
            const mes = d.data.getMonth();
            const key = `${ano}-${String(mes + 1).padStart(2, '0')}`;
            if (!mesesMap.has(key)) {
                const nomeMes = d.data.toLocaleString('pt-BR', { month: 'long', year: 'numeric' });
                mesesMap.set(key, nomeMes.charAt(0).toUpperCase() + nomeMes.slice(1));
            }
        }
    });

    const chavesOrdenadas = Array.from(mesesMap.keys()).sort();
    const nomesMeses = chavesOrdenadas.map(key => mesesMap.get(key));

    if (nomesMeses.length === 0) span.textContent = "Meses Analisados: Nenhuma data válida";
    else if (nomesMeses.length <= 3) span.textContent = `Meses Analisados: ${nomesMeses.join(', ')}`;
    else span.textContent = `Meses Analisados: ${nomesMeses[0]} a ${nomesMeses[nomesMeses.length - 1]}`;
}

function renderizarMapa() {
    const container = document.getElementById('mapaContainer');
    container.innerHTML = '';
    fecharDetalhes(); 
    atualizarTituloPeriodo();

    const dadosBase = getDadosBasicos();

    // Calcula ABC por Endereço (para Detalhes) e por Prédio (para Filtro)
    const mapaABC_Enderecos = calcularClassificacaoABC_Endereco(dadosBase);
    const mapaABC_Predios = calcularClassificacaoABC_Predio(dadosBase);

    // Filtra dados visualmente baseado na Curva do PRÉDIO
    const dadosFinais = dadosBase.filter(item => {
        if (!filtrosAtivos.curva) return true; 
        
        const keyPredio = `${item.rua}-${item.predio}`;
        const classePredio = mapaABC_Predios[keyPredio] || 'C';
        
        return classePredio === filtrosAtivos.curva;
    });

    const contagensPorPredio = {};
    let totalMovimentos = 0;
    let posicoesComDados = 0;

    for (const item of dadosFinais) {
        const key = `${item.rua}-${item.predio}`;
        if (!contagensPorPredio[key]) {
            contagensPorPredio[key] = 0;
            posicoesComDados++;
        }
        contagensPorPredio[key]++;
        totalMovimentos++;
    }

    let minVal = Infinity, maxVal = 0;
    let minLocal = '-', maxLocal = '-';
    
    const entradas = Object.entries(contagensPorPredio);
    if (entradas.length > 0) {
        for (const [loc, qtd] of entradas) {
            if (qtd < minVal) { minVal = qtd; minLocal = loc; }
            if (qtd > maxVal) { maxVal = qtd; maxLocal = loc; }
        }
    } else { minVal = 0; }

    actualizarPainelEstatisticas(minVal, maxVal, minLocal, maxLocal, totalMovimentos, posicoesComDados);

    const estrutura = getEstruturaPredios();

    for (let rua of RUAS) {
        const ruaDiv = document.createElement('div');
        ruaDiv.className = 'rua';
        
        const titulo = document.createElement('div');
        titulo.className = 'rua-titulo';
        titulo.textContent = `Rua ${rua}`;
        ruaDiv.appendChild(titulo);
        
        const content = document.createElement('div');
        content.className = 'rua-content';

        const criarLado = (lista, nomeLado) => {
            if (lista.length > 0) {
                const ladoDiv = document.createElement('div');
                ladoDiv.className = 'lado';
                ladoDiv.innerHTML = `<div class="lado-label">${nomeLado}</div>`;
                
                const prediosDiv = document.createElement('div');
                prediosDiv.className = 'predios';

                lista.forEach(predio => {
                    const key = `${rua}-${predio}`;
                    const qtd = contagensPorPredio[key] || 0;
                    // Passa mapaABC_Enderecos para o Tooltip mostrar detalhes internos
                    const el = criarElementoPredio(predio, qtd, minVal, maxVal, rua, nomeLado, dadosFinais, mapaABC_Enderecos);
                    prediosDiv.appendChild(el);
                });

                ladoDiv.appendChild(prediosDiv);
                content.appendChild(ladoDiv);
            }
        };

        criarLado(estrutura[rua].Impar, 'Impar');
        criarLado(estrutura[rua].Par, 'Par');

        ruaDiv.appendChild(content);
        container.appendChild(ruaDiv);
    }
}

function criarElementoPredio(predio, contagem, min, max, rua, lado, dadosFiltrados, mapaABC_Enderecos) {
    const div = document.createElement('div');
    div.className = 'predio';
    
    const isPAR = (rua === 4 && predio >= 25 && predio <= 44);
    if (isPAR) div.classList.add('par');

    if (contagem === 0) {
        div.classList.add('sem-dados');
        div.textContent = '-';
    } else {
        div.style.backgroundColor = calcularCor(contagem, min, max);
        div.textContent = contagem > 9999 ? (contagem/1000).toFixed(0) + 'k' : formatarNumero(contagem);

        const tooltip = document.createElement('div');
        tooltip.className = 'predio-tooltip';
        
        // Conta endereços A, B, C dentro deste prédio (baseado nos dados filtrados)
        let cA = 0, cB = 0, cC = 0;
        const unicos = new Set();
        
        dadosFiltrados.filter(d => d.rua === rua && d.predio === predio).forEach(d => {
            const keyFull = `${d.rua}-${d.predio}-${d.nivel}-${d.apto}`;
            if(!unicos.has(keyFull)) {
                unicos.add(keyFull);
                const classe = mapaABC_Enderecos[keyFull] || 'C';
                if(classe === 'A') cA++; else if(classe === 'B') cB++; else cC++;
            }
        });

        tooltip.innerHTML = `
            <strong>Prédio ${predio}</strong>${isPAR ? ' (P.A.R)' : ''}<br>
            Clique para mais detalhes.<br>
            <hr style="margin:4px 0; opacity:0.5; border-color: #aaa">
            Posições A: ${cA}<br>
            Posições B: ${cB}<br>
            Posições C: ${cC}
        `;
        div.appendChild(tooltip);

        div.addEventListener('click', (e) => {
            e.stopPropagation(); 
            if (predioSelecionado) predioSelecionado.classList.remove('selecionado');
            predioSelecionado = div;
            predioSelecionado.classList.add('selecionado');
            mostrarDetalhes(rua, predio, lado, dadosFiltrados, mapaABC_Enderecos);
        });
    }
    return div;
}

// ===================================================================
// 7. PAINÉIS E DETALHES
// ===================================================================
function actualizarPainelEstatisticas(min, max, minLoc, maxLoc, total, posicoes) {
    const el = (id) => document.getElementById(id);
    if (posicoes > 0) {
        el('minValue').textContent = formatarNumero(min);
        el('minLocal').textContent = `Rua: ${minLoc.split('-')[0]} - Prédio: ${minLoc.split('-')[1]}`;
        el('maxValue').textContent = formatarNumero(max);
        el('maxLocal').textContent = `Rua: ${maxLoc.split('-')[0]} - Prédio: ${maxLoc.split('-')[1]}`;
        el('avgValue').textContent = formatarNumero(Math.round(total / posicoes));
    } else {
        ['minValue', 'minLocal', 'maxValue', 'maxLocal', 'avgValue'].forEach(id => el(id).textContent = '-');
    }
    el('totalPositions').textContent = formatarNumero(posicoes);
    el('totalMovements').textContent = formatarNumero(total);
}

function mostrarDetalhes(rua, predio, lado, dadosFiltrados, mapaABC_Enderecos) {
    const container = document.getElementById('detalhesContainer');
    const dadosPredio = dadosFiltrados.filter(d => d.rua === rua && d.predio === predio);
    
    const pivot = {};
    const tipos = new Set();
    
    dadosPredio.forEach(item => {
        tipos.add(item.tipoMovimento);
        const end = `${item.nivel}-${item.apto}`;
        if (!pivot[end]) pivot[end] = { total: 0 };
        pivot[end][item.tipoMovimento] = (pivot[end][item.tipoMovimento] || 0) + 1;
        pivot[end].total++;
    });

    const headers = Array.from(tipos).sort();

    let html = `
        <h3>Detalhes: Rua ${rua} - Prédio ${predio} (${lado}) 
            <button id="btnFechar">Fechar X</button>
        </h3>
        <div class="tabela-detalhes-container">
            <table class="detalhes">
                <thead>
                    <tr>
                        <th>Local (N-Apto)</th>
                        <th>Curva</th>`;
    headers.forEach(h => html += `<th>${h}</th>`);
    html += `<th>Total</th></tr></thead><tbody>`;

    Object.entries(pivot).sort((a,b) => b[1].total - a[1].total).forEach(([end, vals]) => {
        const keyFull = `${rua}-${predio}-${end}`;
        const classe = mapaABC_Enderecos[keyFull] || '-';
        const cor = classe === 'A' ? '#e74c3c' : (classe === 'B' ? '#f39c12' : '#2ecc71');

        html += `<tr>
            <td>${end}</td>
            <td style="font-weight:bold; color:${cor}; text-align:center">${classe}</td>`;
        headers.forEach(h => html += `<td>${formatarNumero(vals[h] || 0)}</td>`);
        html += `<td><strong>${formatarNumero(vals.total)}</strong></td></tr>`;
    });
    
    if (Object.keys(pivot).length === 0) html += '<tr><td colspan="10">Nenhum dado encontrado.</td></tr>';
    
    html += '</tbody></table></div>';
    container.innerHTML = html;
    container.style.display = 'block';
    
    document.getElementById('btnFechar').onclick = fecharDetalhes;
    container.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function fecharDetalhes() {
    document.getElementById('detalhesContainer').style.display = 'none';
    if (predioSelecionado) {
        predioSelecionado.classList.remove('selecionado');
        predioSelecionado = null;
    }
}

// ===================================================================
// 8. INICIALIZAÇÃO
// ===================================================================
function atualizarFiltrosUI() {
    const sel = document.getElementById('tipoMovimento');
    const atual = sel.value;
    while (sel.options.length > 1) sel.remove(1);
    
    const tipos = new Set(dadosConsolidados.map(d => d.tipoMovimento).filter(Boolean));
    Array.from(tipos).sort().forEach(t => {
        const opt = document.createElement('option');
        opt.value = t; opt.textContent = t; sel.appendChild(opt);
    });
    sel.value = atual;
}

document.addEventListener('DOMContentLoaded', () => {
    if (typeof flatpickr !== 'undefined') {
        flatpickr("#dataInicial", { dateFormat: "d/m/Y", locale: "pt", onChange: (d) => filtrosAtivos.dataInicial = d[0] });
        flatpickr("#dataFinal", { dateFormat: "d/m/Y", locale: "pt", onChange: (d) => filtrosAtivos.dataFinal = d[0] });
    }

    document.getElementById('aplicarFiltros').addEventListener('click', () => {
        filtrosAtivos.tipoMovimento = document.getElementById('tipoMovimento').value;
        filtrosAtivos.picking = document.getElementById('picking').value;
        filtrosAtivos.curva = document.getElementById('curva').value;
        renderizarMapa();
    });

    document.getElementById('resetarFiltros').addEventListener('click', () => {
        document.getElementById('tipoMovimento').value = '';
        document.getElementById('picking').value = '';
        document.getElementById('curva').value = '';
        document.querySelectorAll('.flatpickr-input').forEach(i => { i._flatpickr.clear(); i.value = ''; });
        
        filtrosAtivos = { tipoMovimento: '', picking: '', curva: '', dataInicial: null, dataFinal: null };
        renderizarMapa(); 
    });

    document.querySelector('.container').addEventListener('click', (e) => {
        if (!e.target.closest('.predio') && !e.target.closest('.detalhes-container')) fecharDetalhes();
    });

    iniciarSistema();
});
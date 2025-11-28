// =========================================================================
// 1. CONFIGURAÇÃO & ESTRUTURA
// =========================================================================

const RUAS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 50];

const CONFIG = {
    arquivos: {
        enderecos: 'Endereços.xlsx',
        convocacoes: 'Convocações Matriz-Mapeamento.xlsx',
        abastecimento: 'Abastecimento-mapeamento.xlsx'
    },
    // IMPORTANTE: O sistema procura EXATAMENTE estes textos no Excel
    colsEnderecos: {
        rua: 'Rua', predio: 'Prédio', nivel: 'Nível', apto: 'Apartamento',
        tipoEndereco: 'Tipo de endereço' 
    },
    colsConvocacoes: {
        rua: 'Rua', predio: 'Prédio', nivel: 'Nível', apto: 'Apartamento',
        tipoMov: 'Tipo de movimento', 
        data: 'Data/hora inserção',
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

/**
 * Lógica Central de Agrupamento Visual.
 * Determina qual ID será desenhado no mapa.
 */
function obterVisualId(rua, predio, apto) {
    const r = parseInt(rua);
    const p = parseInt(predio);
    const a = parseInt(apto) || 0;

    // REGRA RUA 50: O Apartamento vira o ID Visual
    if (r === 50) {
        return a;
    }

    // REGRA RUA 1: Agrupamento 71-129 (Ímpar)
    if (r === 1 && p >= 71 && p <= 129 && p % 2 !== 0) {
        const base = 71;
        const passo = 8; 
        const grupo = Math.floor((p - base) / passo);
        return base + (grupo * passo);
    }

    return p; // Padrão
}

let mapaVisualValido = new Set(); 

function getEstruturaPredios() {
    const estrutura = {};
    mapaVisualValido = new Set();

    for (let rua of RUAS) {
        let prediosImpar = [];
        let prediosPar = [];
        
        if (rua === 50) {
            // === CONFIGURAÇÃO RUA 50 (Visualizacao por Apartamento) ===
            for (let i = 101; i <= 112; i++) prediosImpar.push(i);
            for (let i = 201; i <= 217; i++) prediosPar.push(i);

        } else if (rua === 1) {
            // Rua 1 Normal
            for (let i = 41; i <= 69; i += 2) prediosImpar.push(i);
            // Rua 1 Agrupada
            for (let i = 71; i <= 129; i += 8) prediosImpar.push(i);
            // Rua 1 Par
            prediosPar = Array.from({ length: 5 }, (_, i) => 48 + i * 2);
            prediosPar.push(70);
            
        } else if (rua === 2) {
            for (let i = 3; i <= 46; i += 2) prediosImpar.push(i);
            for (let i = 2; i <= 46; i += 2) prediosPar.push(i);
        } else {
            for (let i = 1; i <= 46; i += 2) prediosImpar.push(i);
            for (let i = 2; i <= 46; i += 2) prediosPar.push(i);
        }
        
        estrutura[rua] = { Impar: prediosImpar, Par: prediosPar };

        [...prediosImpar, ...prediosPar].forEach(idVisual => {
            mapaVisualValido.add(`${rua}-${idVisual}`);
        });
    }
    return estrutura;
}

getEstruturaPredios(); 

let dadosConsolidados = [];
let filtrosAtivos = { tipoMovimento: '', picking: '', curva: '', dataInicial: null, dataFinal: null };
let predioSelecionado = null; 

// ===================================================================
// 4. CARREGAMENTO INTELIGENTE (RESOLUÇÃO DE ERROS DE CABEÇALHO)
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
    if (!wb) {
        console.error(`[${nomeArquivo}] ERRO: Arquivo não foi carregado (wb é nulo).`);
        return [];
    }

    console.log(`%c[${nomeArquivo}] Iniciando Varredura Inteligente... Procurando: "${colChave}"`, "color: blue; font-weight: bold");

    for (const nomeAba of wb.SheetNames) {
        const sheet = wb.Sheets[nomeAba];
        // Lê as primeiras 20 linhas como matriz (array de arrays) para inspecionar
        const dadosBrutos = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 0, defval: "" });

        if (dadosBrutos && dadosBrutos.length > 0) {
            let linhaCabecalho = -1;
            
            // Varre as primeiras 50 linhas
            for (let i = 0; i < Math.min(dadosBrutos.length, 50); i++) {
                const linha = dadosBrutos[i];
                // Verifica se a linha contém a coluna chave (com normalização)
                const colChaveNormalizada = String(colChave).trim().toLowerCase();
                const linhaNormalizada = linha.map(c => String(c).trim().toLowerCase());
                if (linha && Array.isArray(linha) && linhaNormalizada.includes(colChaveNormalizada)) {
                    linhaCabecalho = i;
                    break;
                }
            }

            if (linhaCabecalho !== -1) {
                console.log(`%c[${nomeArquivo}] ✅ Cabeçalho encontrado na LINHA ${linhaCabecalho + 1} da aba "${nomeAba}".`, "color: green");
                
                // Lê novamente a partir da linha correta
                const dadosFinais = XLSX.utils.sheet_to_json(sheet, { range: linhaCabecalho, defval: "" });
                return dadosFinais;
            } else {
                // LOG DE DEPURAÇÃO PARA O USUÁRIO VER O QUE O SCRIPT ESTÁ LENDO
                console.warn(`[${nomeArquivo}] ⚠️ Aba "${nomeAba}": Coluna "${colChave}" não encontrada nas primeiras 20 linhas.`);
                console.log(`%c[${nomeArquivo}] Conteúdo das primeiras 3 linhas desta aba (para conferência):`, "color: #777");
                console.table(dadosBrutos.slice(0, 3)); 
            }
        }
    }
    
    console.error(`[${nomeArquivo}] ❌ ERRO FATAL: Não encontrei a coluna "${colChave}" em nenhuma aba.`);
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
        if (!rows || rows.length === 0) return;

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
                    apto: parseInt(a) || 0, 
                    nivel: n,
                    tipoMovimento: row[cols.tipoMov] || "Outros",
                    data: dataProc,
                    picking: info.picking
                });
            }
        });
    };

    // Processamento com logs de sucesso
    const dadosConv = buscarDadosValidos(wbConv, "Convocacoes", CONFIG.colsConvocacoes.rua);
    if (dadosConv.length > 0) {
        processarTabela(dadosConv, CONFIG.colsConvocacoes);
        console.log(`[Convocações] ${dadosConv.length} linhas processadas com sucesso.`);
    }

    const dadosAbast = buscarDadosValidos(wbAbast, "Abastecimento", CONFIG.colsAbastecimento.rua);
    if (dadosAbast.length > 0) {
        processarTabela(dadosAbast, CONFIG.colsAbastecimento);
        console.log(`[Abastecimento] ${dadosAbast.length} linhas processadas com sucesso.`);
    }
    
    atualizarFiltrosUI();
    loading.classList.remove('active');
    renderizarMapa();
    
    setTimeout(verificarDiscrepancias, 1000);
}

// ===================================================================
// 5. FILTRAGEM E CÁLCULOS ABC
// ===================================================================

function getDadosBasicos() {
    return dadosConsolidados.filter(item => {
        const idVisual = obterVisualId(item.rua, item.predio, item.apto);
        const keyVisual = `${item.rua}-${idVisual}`;
        
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

function calcularClassificacaoABC_Visual(dados) {
    const contagem = {};
    let totalMovimentos = 0;

    dados.forEach(item => {
        const idVisual = obterVisualId(item.rua, item.predio, item.apto);
        const key = `${item.rua}-${idVisual}`; 
        contagem[key] = (contagem[key] || 0) + 1;
        totalMovimentos++;
    });

    return classificarPorPareto(contagem, totalMovimentos);
}

function classificarPorPareto(contagemObj, totalMovimentos) {
    const gruposPorQtd = {};
    Object.entries(contagemObj).forEach(([key, qtd]) => {
        if (!gruposPorQtd[qtd]) gruposPorQtd[qtd] = [];
        gruposPorQtd[qtd].push(key);
    });

    const quantidadesOrdenadas = Object.keys(gruposPorQtd).map(Number).sort((a, b) => b - a);

    const mapaClassificacao = {}; 
    let acumulado = 0;

    quantidadesOrdenadas.forEach(qtd => {
        const listaItens = gruposPorQtd[qtd];
        const volumeDoGrupo = qtd * listaItens.length;
        
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

    const mapaABC_Enderecos = calcularClassificacaoABC_Endereco(dadosBase);
    const mapaABC_Visuais = calcularClassificacaoABC_Visual(dadosBase);

    const dadosFinais = dadosBase.filter(item => {
        if (!filtrosAtivos.curva) return true; 
        
        const idVisual = obterVisualId(item.rua, item.predio, item.apto);
        const keyVisual = `${item.rua}-${idVisual}`;
        const classeVisual = mapaABC_Visuais[keyVisual] || 'C';
        
        return classeVisual === filtrosAtivos.curva;
    });

    const contagensPorVisual = {};
    let totalMovimentos = 0;
    let posicoesComDados = 0;

    for (const item of dadosFinais) {
        const idVisual = obterVisualId(item.rua, item.predio, item.apto);
        const key = `${item.rua}-${idVisual}`;
        
        if (!contagensPorVisual[key]) {
            contagensPorVisual[key] = 0;
            posicoesComDados++;
        }
        contagensPorVisual[key]++;
        totalMovimentos++;
    }

    let minVal = Infinity, maxVal = 0;
    let minLocal = '-', maxLocal = '-';
    
    const entradas = Object.entries(contagensPorVisual);
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
        titulo.textContent = rua === 50 ? `Rua ${rua} (Aptos)` : `Rua ${rua}`;
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

                lista.forEach(idVisual => {
                    const key = `${rua}-${idVisual}`;
                    const qtd = contagensPorVisual[key] || 0;
                    const el = criarElementoPredio(idVisual, qtd, minVal, maxVal, rua, nomeLado, dadosFinais, mapaABC_Enderecos);
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

function criarElementoPredio(idVisual, contagem, min, max, rua, lado, dadosFiltrados, mapaABC_Enderecos) {
    const div = document.createElement('div');
    div.className = 'predio';
    
    // 1. Lógica visual: Apenas Rua 4, prédios 25 ao 44 (Linha Azul)
    const isPAR = (rua === 4 && idVisual >= 25 && idVisual <= 44);
    if (isPAR) div.classList.add('par');

    if (contagem === 0) {
        div.classList.add('sem-dados');
        div.textContent = '-';
    } else {
        div.style.backgroundColor = calcularCor(contagem, min, max);
        div.textContent = contagem > 9999 ? (contagem/1000).toFixed(0) + 'k' : formatarNumero(contagem);

        // --- LÓGICA DO TOOLTIP GLOBAL (Sem "filhos" aqui) ---
        const tooltipGlobal = document.getElementById('tooltip-global');

        div.addEventListener('mouseenter', () => {
            // 1. Preparar Conteúdo
            let cA = 0, cB = 0, cC = 0;
            const unicos = new Set();
            
            dadosFiltrados.filter(d => d.rua === rua && obterVisualId(d.rua, d.predio, d.apto) === idVisual).forEach(d => {
                const keyFull = `${d.rua}-${d.predio}-${d.nivel}-${d.apto}`;
                if(!unicos.has(keyFull)) {
                    unicos.add(keyFull);
                    const classe = mapaABC_Enderecos[keyFull] || 'C';
                    if(classe === 'A') cA++; else if(classe === 'B') cB++; else cC++;
                }
            });

            let textoTitulo = '';
            if (rua === 50) textoTitulo = `Apto ${idVisual}`;
            else {
                textoTitulo = `Prédio ${idVisual}`;
                if (isPAR) textoTitulo += ' (P.A.R)';
            }

            tooltipGlobal.innerHTML = `
                <strong>${textoTitulo}</strong><br>
                Clique para detalhes.<br>
                <hr style="margin:4px 0; opacity:0.5; border-color: #aaa">
                Posições A: ${cA}<br>
                Posições B: ${cB}<br>
                Posições C: ${cC}
            `;

            // 2. Calcular Posição Global (Fixed)
            tooltipGlobal.classList.add('visible'); // Mostra para medir tamanho
            
            const pRect = div.getBoundingClientRect(); // Retângulo do prédio
            const tRect = tooltipGlobal.getBoundingClientRect(); // Retângulo do Tooltip
            const vW = window.innerWidth;

            // Define Top vs Bottom (Prioridade: Topo)
            let top = pRect.top - tRect.height - 10;
            let clsPos = 'pos-top';

            // Se não couber em cima, joga pra baixo
            if (pRect.top < tRect.height + 15) {
                top = pRect.bottom + 10;
                clsPos = 'pos-bottom';
            }

            // Centraliza horizontalmente
            let left = pRect.left + (pRect.width / 2) - (tRect.width / 2);

            // Clamp (Evita sair pelas laterais)
            let arrowOffset = 0;
            if (left < 5) {
                arrowOffset = left - 5; // Negativo (estourou esquerda)
                left = 5;
            } else if (left + tRect.width > vW - 5) {
                let newLeft = vW - tRect.width - 5;
                arrowOffset = left - newLeft; // Positivo (estourou direita)
                left = newLeft;
            }
            
            // Correção da Seta (Seta move-se para acompanhar o alvo)
            const targetCenterX = pRect.left + (pRect.width / 2);
            const tooltipCenterX = left + (tRect.width / 2);
            const finalOffset = targetCenterX - tooltipCenterX;

            // Aplica estilos
            tooltipGlobal.style.top = `${top}px`;
            tooltipGlobal.style.left = `${left}px`;
            tooltipGlobal.className = `visible ${clsPos}`;
            tooltipGlobal.style.setProperty('--arrow-offset', `${finalOffset}px`);
        });

        div.addEventListener('mouseleave', () => {
            tooltipGlobal.classList.remove('visible');
        });

        div.addEventListener('click', (e) => {
            e.stopPropagation(); 
            if (predioSelecionado) predioSelecionado.classList.remove('selecionado');
            predioSelecionado = div;
            predioSelecionado.classList.add('selecionado');
            mostrarDetalhes(rua, idVisual, lado, dadosFiltrados, mapaABC_Enderecos);
        });
    }
    return div;
}

// ===================================================================
// 7. PAINÉIS E DETALHES
// ===================================================================
function actualizarPainelEstatisticas(min, max, minLoc, maxLoc, total, posicoes) {
    const el = (id) => document.getElementById(id);
    
    const formatLoc = (locStr) => {
        if(locStr === '-') return '-';
        const parts = locStr.split('-');
        const r = parseInt(parts[0]);
        const label = r === 50 ? 'Apto' : 'Prédio';
        return `Rua ${r} - ${label} ${parts[1]}`;
    };

    if (posicoes > 0) {
        el('minValue').textContent = formatarNumero(min);
        el('minLocal').textContent = formatLoc(minLoc);
        el('maxValue').textContent = formatarNumero(max);
        el('maxLocal').textContent = formatLoc(maxLoc);
        el('avgValue').textContent = formatarNumero(Math.round(total / posicoes));
    } else {
        ['minValue', 'minLocal', 'maxValue', 'maxLocal', 'avgValue'].forEach(id => el(id).textContent = '-');
    }
    el('totalPositions').textContent = formatarNumero(posicoes);
    el('totalMovements').textContent = formatarNumero(total);
}

function mostrarDetalhes(rua, idVisual, lado, dadosFiltrados, mapaABC_Enderecos) {
    const container = document.getElementById('detalhesContainer');
    
    const dadosBloco = dadosFiltrados.filter(d => d.rua === rua && obterVisualId(d.rua, d.predio, d.apto) === idVisual);
    
    const pivot = {};
    const tipos = new Set();
    
    dadosBloco.forEach(item => {
        tipos.add(item.tipoMovimento);
        let labelLinha;
        if (rua === 50) {
            labelLinha = `Físico P${item.predio} - N${item.nivel}`;
        } else {
            labelLinha = `P${item.predio} - ${item.nivel}-${item.apto}`;
        }

        if (!pivot[labelLinha]) pivot[labelLinha] = { total: 0, predioReal: item.predio, nivel: item.nivel, apto: item.apto };
        pivot[labelLinha][item.tipoMovimento] = (pivot[labelLinha][item.tipoMovimento] || 0) + 1;
        pivot[labelLinha].total++;
    });

    const headers = Array.from(tipos).sort();
    const tituloTipo = rua === 50 ? 'Apartamento' : 'Prédio';

    let html = `
        <h3>Detalhes: Rua ${rua} - ${tituloTipo} ${idVisual} (${lado})
            <button id="btnFechar">Fechar X</button>
        </h3>
        <div class="tabela-detalhes-container">
            <table class="detalhes">
                <thead>
                    <tr>
                        <th>Local</th>
                        <th>Curva</th>`;
    headers.forEach(h => html += `<th>${h}</th>`);
    html += `<th>Total</th></tr></thead><tbody>`;

    Object.entries(pivot).sort((a,b) => b[1].total - a[1].total).forEach(([label, vals]) => {
        const keyFull = `${rua}-${vals.predioReal}-${vals.nivel}-${vals.apto}`;
        const classe = mapaABC_Enderecos[keyFull] || '-';
        const cor = classe === 'A' ? '#e74c3c' : (classe === 'B' ? '#f39c12' : '#2ecc71');

        html += `<tr>
            <td>${label}</td>
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
// 8. INICIALIZAÇÃO E EVENTOS
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

// ===================================================================
// 9. AUDITORIA DE DADOS
// ===================================================================
function verificarDiscrepancias() {
    console.log("%c=== AUDITORIA DE DADOS ===", "background: #222; color: #bada55; font-size: 14px");
    
    const totalExcel = dadosConsolidados.length;
    let totalNoMapa = 0;
    
    const dadosNoMapa = dadosConsolidados.filter(item => {
        const idVisual = obterVisualId(item.rua, item.predio, item.apto);
        const keyVisual = `${item.rua}-${idVisual}`;
        return mapaVisualValido.has(keyVisual);
    });
    totalNoMapa = dadosNoMapa.length;

    console.log(`Total Excel:    ${formatarNumero(totalExcel)}`);
    console.log(`Total no Mapa:  ${formatarNumero(totalNoMapa)}`);
    console.log(`Ignorados:      ${formatarNumero(totalExcel - totalNoMapa)}`);

    if (totalExcel === totalNoMapa) {
        console.log("✅ Dados 100% conciliados.");
        return;
    }

    const ignorados = {};
    dadosConsolidados.forEach(item => {
        const idVisual = obterVisualId(item.rua, item.predio, item.apto);
        const keyVisual = `${item.rua}-${idVisual}`;
        
        if (!mapaVisualValido.has(keyVisual)) {
            let local;
            if (item.rua === 50) local = `Rua 50 - Apto ${item.apto} (Range 101-112 ou 201-217?)`;
            else local = `Rua ${item.rua} - Prédio ${item.predio}`;
            
            ignorados[local] = (ignorados[local] || 0) + 1;
        }
    });

    console.log("\n⚠️ Principais Locais Ignorados:");
    console.table(
        Object.entries(ignorados)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 20)
            .map(([local, qtd]) => ({ Local: local, Quantidade: qtd }))
    );
}
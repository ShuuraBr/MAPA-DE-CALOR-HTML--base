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
    // ATENÇÃO: Os nomes aqui devem ser parte do nome que está no cabeçalho do Excel
    colsEnderecos: {
        rua: 'Rua', predio: 'Prédio', nivel: 'Nível', apto: 'Apartamento',
        tipoEndereco: 'Tipo de endereço' 
    },
    colsConvocacoes: {
        rua: 'Rua', predio: 'Prédio', nivel: 'Nível', apto: 'Apartamento',
        data: 'Data/hora inserção', // O código tentará achar colunas parecidas se não encontrar essa
        tipoMov: 'Tipo de movimento'
    },
    colsAbastecimento: {
        rua: 'Rua', predio: 'Prédio', nivel: 'Nível', apto: 'Apartamento',
        tipoMov: 'Tipo da Transferência', 
        data: 'Data' // O código tentará achar colunas parecidas
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

// Função auxiliar para "limpar" textos para comparação (ignora acentos, espaços e maiúsculas)
const limparTexto = (t) => String(t || "").toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") 
    .replace(/[^a-z0-9]/g, ""); 

const processarData = (v) => {
    if (!v) return null;
    // Data Excel Serial (ex: 45200)
    if (typeof v === 'number') {
        return new Date(Math.round((v - 25569) * 86400 * 1000));
    }
    // String (ex: "25/10/2024")
    if (typeof v === 'string') {
        if (v.includes('/')) {
            const partes = v.split(' '); 
            const dataPartes = partes[0].split('/');
            if (dataPartes.length === 3) {
                return new Date(dataPartes[2], dataPartes[1] - 1, dataPartes[0]);
            }
        }
    }
    // ISO ou outra
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
};

// ===================================================================
// 3. ESTRUTURA DO ARMAZÉM
// ===================================================================

function obterVisualId(rua, predio, apto) {
    const r = parseInt(rua);
    const p = parseInt(predio);
    const a = parseInt(apto) || 0;

    if (r === 50) return a;

    if (r === 1 && p >= 71 && p <= 129 && p % 2 !== 0) {
        const base = 71;
        const passo = 8; 
        const grupo = Math.floor((p - base) / passo);
        return base + (grupo * passo);
    }

    return p; 
}

let mapaVisualValido = new Set(); 

function getEstruturaPredios() {
    const estrutura = {};
    mapaVisualValido = new Set();

    for (let rua of RUAS) {
        let prediosImpar = [];
        let prediosPar = [];
        
        if (rua === 50) {
            for (let i = 101; i <= 112; i++) prediosImpar.push(i);
            for (let i = 201; i <= 217; i++) prediosPar.push(i);
        } else if (rua === 1) {
            for (let i = 41; i <= 69; i += 2) prediosImpar.push(i);
            for (let i = 71; i <= 129; i += 8) prediosImpar.push(i);
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
// 4. CARREGAMENTO E BUSCA ROBUSTA (AQUI ESTÁ A CORREÇÃO PRINCIPAL)
// ===================================================================
async function carregarExcel(url) {
    try {
        const res = await fetch(url + `?t=${new Date().getTime()}`);
        if (!res.ok) throw new Error(`Erro ${res.status}`);
        const ab = await res.arrayBuffer();
        return XLSX.read(ab, { type: 'array' });
    } catch (e) { console.error(`Falha ao carregar ${url}`, e); return null; }
}

function buscarDadosValidos(wb, nomeArquivo, colChaveOuLista) {
    if (!wb) {
        console.error(`[${nomeArquivo}] ERRO: Arquivo não foi carregado corretamente.`);
        return [];
    }

    const candidatos = Array.isArray(colChaveOuLista) ? colChaveOuLista : [colChaveOuLista];
    const candidatosLimpos = candidatos.map(c => limparTexto(c));

    console.log(`[${nomeArquivo}] Procurando colunas:`, candidatos);

    // Itera sobre TODAS as abas até achar algo
    for (const nomeAba of wb.SheetNames) {
        const sheet = wb.Sheets[nomeAba];
        
        // Pega TODO o conteúdo da aba como uma matriz (linhas x colunas)
        // defval: "" garante que células vazias não quebrem o array
        const dadosBrutos = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

        if (dadosBrutos && dadosBrutos.length > 0) {
            // Varre as primeiras 50 linhas procurando o cabeçalho
            for (let i = 0; i < Math.min(dadosBrutos.length, 50); i++) {
                const linha = dadosBrutos[i];
                
                if (linha && Array.isArray(linha)) {
                    const linhaLimpa = linha.map(c => limparTexto(c));
                    
                    // Verifica se ALGUM dos candidatos (ex: "rua") está nesta linha
                    const encontrou = candidatosLimpos.find(c => linhaLimpa.includes(c));

                    if (encontrou) {
                        const indexReal = linhaLimpa.indexOf(encontrou);
                        const nomeReal = linha[indexReal];
                        
                        console.log(`%c[${nomeArquivo}] ✅ Cabeçalho encontrado na linha ${i + 1}: "${nomeReal}"`, "color: green");

                        // Ajuste dinâmico para Convocações (Data)
                        if (nomeArquivo === "Convocacoes") {
                            // Tenta achar a coluna de data nesta mesma linha
                            linha.forEach(col => {
                                const txt = limparTexto(col);
                                // Se contiver 'data', 'hora' ou 'inclusao', assume que é a data
                                if (txt.includes('data') || txt.includes('hora') || txt.includes('inclusao')) {
                                    console.log(`[${nomeArquivo}] Coluna de Data identificada: "${col}"`);
                                    CONFIG.colsConvocacoes.data = col;
                                }
                            });
                        }
                        
                        // Ajuste dinâmico para Abastecimento (Data)
                        if (nomeArquivo === "Abastecimento") {
                            linha.forEach(col => {
                                const txt = limparTexto(col);
                                if (txt.includes('data') || txt.includes('cadastro')) {
                                    console.log(`[${nomeArquivo}] Coluna de Data identificada: "${col}"`);
                                    CONFIG.colsAbastecimento.data = col;
                                }
                            });
                        }

                        // Retorna os dados usando a linha encontrada como cabeçalho (range: i)
                        return XLSX.utils.sheet_to_json(sheet, { range: i, defval: "" });
                    }
                }
            }
        }
    }
    
    // Se chegou aqui, falhou
    console.warn(`[${nomeArquivo}] ⚠️ Nenhuma coluna válida encontrada nas primeiras 50 linhas de nenhuma aba.`);
    
    // Diagnóstico final para o usuário
    const primeiraAba = wb.Sheets[wb.SheetNames[0]];
    const dadosDiagnostico = XLSX.utils.sheet_to_json(primeiraAba, { header: 1 }).slice(0, 5);
    
    if (dadosDiagnostico.length === 0) {
        console.error(`[${nomeArquivo}] O arquivo parece vazio ou em formato inválido (XML/HTML). Tente abrir no Excel e 'Salvar Como' .xlsx`);
    } else {
        console.table(dadosDiagnostico);
        console.log(`[${nomeArquivo}] Verifique a tabela acima. Seus cabeçalhos estão aí? Atualize o CONFIG.`);
    }

    return [];
}

async function iniciarSistema() {
    const loading = document.getElementById('loading');
    const erroDiv = document.getElementById('errorMessage');
    loading.classList.add('active');
    
    getEstruturaPredios();

    // Carrega os 3 arquivos
    const [wbEnd, wbConv, wbAbast] = await Promise.all([
        carregarExcel(CONFIG.arquivos.enderecos),
        carregarExcel(CONFIG.arquivos.convocacoes),
        carregarExcel(CONFIG.arquivos.abastecimento)
    ]);

    // Verifica Endereços (Obrigatório)
    if (!wbEnd) {
        erroDiv.textContent = "Erro Crítico: Arquivo 'Endereços.xlsx' não carregou.";
        erroDiv.classList.add('active');
        loading.classList.remove('active');
        return;
    }

    // 1. Processar Endereços
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

    // Função genérica de processamento
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

    // 2. Processar Convocações
    const dadosConv = buscarDadosValidos(wbConv, "Convocacoes", CONFIG.colsConvocacoes.rua);
    if (dadosConv.length > 0) {
        processarTabela(dadosConv, CONFIG.colsConvocacoes);
        console.log(`[Convocações] ${dadosConv.length} linhas importadas.`);
    } else {
        console.error("ERRO: Não foi possível ler dados de Convocações.");
    }

    // 3. Processar Abastecimento
    const dadosAbast = buscarDadosValidos(wbAbast, "Abastecimento", CONFIG.colsAbastecimento.rua);
    if (dadosAbast.length > 0) {
        processarTabela(dadosAbast, CONFIG.colsAbastecimento);
        console.log(`[Abastecimento] ${dadosAbast.length} linhas importadas.`);
    }
    
    atualizarFiltrosUI();
    loading.classList.remove('active');
    renderizarMapa();
    
    // Verifica se realmente importou algo
    setTimeout(() => {
        if (dadosConsolidados.length === 0) {
            alert("ATENÇÃO: Nenhum dado foi processado. Verifique o Console (F12) para ver os detalhes do erro.");
        } else {
            verificarDiscrepancias();
        }
    }, 1000);
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
    
    const isPAR = (rua === 4 && idVisual >= 25 && idVisual <= 44);
    if (isPAR) div.classList.add('par');

    if (contagem === 0) {
        div.classList.add('sem-dados');
        div.textContent = '-';
    } else {
        div.style.backgroundColor = calcularCor(contagem, min, max);
        div.textContent = contagem > 9999 ? (contagem/1000).toFixed(0) + 'k' : formatarNumero(contagem);

        const tooltipGlobal = document.getElementById('tooltip-global');

        div.addEventListener('mouseenter', () => {
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

            tooltipGlobal.classList.add('visible'); 
            
            const pRect = div.getBoundingClientRect(); 
            const tRect = tooltipGlobal.getBoundingClientRect(); 
            const vW = window.innerWidth;

            let top = pRect.top - tRect.height - 10;
            let clsPos = 'pos-top';

            if (pRect.top < tRect.height + 15) {
                top = pRect.bottom + 10;
                clsPos = 'pos-bottom';
            }

            let left = pRect.left + (pRect.width / 2) - (tRect.width / 2);

            let arrowOffset = 0;
            if (left < 5) {
                arrowOffset = left - 5; 
                left = 5;
            } else if (left + tRect.width > vW - 5) {
                let newLeft = vW - tRect.width - 5;
                arrowOffset = left - newLeft; 
                left = newLeft;
            }
            
            const targetCenterX = pRect.left + (pRect.width / 2);
            const tooltipCenterX = left + (tRect.width / 2);
            const finalOffset = targetCenterX - tooltipCenterX;

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

    if (totalExcel !== totalNoMapa) {
        console.log("\n⚠️ Principais Locais Ignorados:");
        console.table(
            Object.entries(ignorados)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 20)
                .map(([local, qtd]) => ({ Local: local, Quantidade: qtd }))
        );
    }
}
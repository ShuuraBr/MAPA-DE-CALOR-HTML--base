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
        rua: 'Endereço de WMS - Parte 1',
        predio: 'Endereço de WMS - Parte 2',
        nivel: 'Endereço de WMS - Parte 3',
        apto: 'Endereço de WMS - Parte 4',
        tipoMov: 'Tipo de movimento',
        data: 'Data/hora inserção'
    },
    colsAbastecimento: {
        rua: 'Rua',
        predio: 'Prédio',
        nivel: 'Nível',
        apto: 'Apartamento',
        tipoMov: 'Tipo da Transferência',
        data: 'Data/hora cadastro'
    },
    cores: {
        vazio: '#1f2933',
        muitoBaixo: '#304FFE',
        baixo: '#009688',
        medio: '#FFEB3B',
        alto: '#FF9800',
        muitoAlto: '#D32F2F',
        borda: '#111827',
        texto: '#E5E7EB',
        destaque: '#FACC15',
        predioSemEndereco: '#4B5563',
        ruaSelecionada: '#F97316'
    }
};

const estruturaPredios = {};
const mapaVisualValido = new Set();

// =========================================================================
// 2. UTILITÁRIOS DE DATA, NORMALIZAÇÃO, CORES
// =========================================================================

function normalizarEndereco(value) {
    if (value === null || value === undefined) return null;
    if (typeof value === 'number') return value;
    if (typeof value === 'string') {
        const limpo = value.trim().replace(/[^\d]/g, '');
        const num = parseInt(limpo, 10);
        return isNaN(num) ? null : num;
    }
    return null;
}

function parseDataBr(dataStr) {
    if (!dataStr) return null;
    if (dataStr instanceof Date && !isNaN(dataStr)) return dataStr;

    if (typeof dataStr === 'number') {
        const utcDays = Math.floor(dataStr - 25569);
        const utcValue = utcDays * 86400;
        return new Date(utcValue * 1000);
    }

    if (typeof dataStr === 'string') {
        const partesDataHora = dataStr.trim().split(' ');
        const partesData = partesDataHora[0]?.split('/') || [];
        if (partesData.length === 3) {
            const [dia, mes, ano] = partesData.map(v => parseInt(v, 10));
            if (!isNaN(dia) && !isNaN(mes) && !isNaN(ano)) {
                const horasPartes = (partesDataHora[1] || '00:00:00').split(':').map(v => parseInt(v, 10));
                const [h = 0, m = 0, s = 0] = horasPartes;
                return new Date(ano, mes - 1, dia, h, m, s);
            }
        }
        const iso = new Date(dataStr);
        return isNaN(iso) ? null : iso;
    }
    return null;
}

function dataDentroIntervalo(data, inicio, fim) {
    if (!data) return false;
    const t = data.getTime();
    if (inicio && t < inicio.getTime()) return false;
    if (fim && t > fim.getTime()) return false;
    return true;
}

function formatarDataCurta(data) {
    if (!(data instanceof Date) || isNaN(data)) return '';
    return data.toLocaleDateString('pt-BR');
}

function obterClassificacaoNivel(nivel) {
    if (!nivel) return '';
    const n = String(nivel).trim().toUpperCase();
    if (n.includes('P')) return 'picking';
    return 'armazenagem';
}

function obterCurvaPorProduto(produto) {
    return 'A';
}

function gerarGradienteCor(valor, minimo, maximo) {
    if (valor === 0 || !isFinite(valor)) return CONFIG.cores.vazio;
    const corMinimo = CONFIG.cores.muitoBaixo;
    const corBaixo = CONFIG.cores.baixo;
    const corMedio = CONFIG.cores.medio;
    const corAlto = CONFIG.cores.alto;
    const corMaximo = CONFIG.cores.muitoAlto;

    if (maximo === minimo) return corMedio;
    const normalized = (valor - minimo) / (maximo - minimo);
    if (normalized <= 0.25) return interpolarCor(corMinimo, corBaixo, normalized / 0.25);
    else if (normalized <= 0.50) return interpolarCor(corBaixo, corMedio, (normalized - 0.25) / 0.25);
    else if (normalized <= 0.75) return interpolarCor(corMedio, corAlto, (normalized - 0.50) / 0.25);
    else return interpolarCor(corAlto, corMaximo, (normalized - 0.75) / 0.25);
}

function hexParaRgb(hex) {
    const limp = hex.replace('#', '');
    const bigint = parseInt(limp, 16);
    return { r: (bigint >> 16) & 255, g: (bigint >> 8) & 255, b: bigint & 255 };
}

function rgbParaHex(r, g, b) {
    return '#' + [r, g, b].map(x => {
        const he = x.toString(16);
        return he.length === 1 ? '0' + he : he;
    }).join('');
}

function interpolarCor(cor1, cor2, t) {
    const c1 = hexParaRgb(cor1);
    const c2 = hexParaRgb(cor2);
    const r = Math.round(c1.r + (c2.r - c1.r) * t);
    const g = Math.round(c1.g + (c2.g - c1.g) * t);
    const b = Math.round(c1.b + (c2.b - c1.b) * t);
    return rgbParaHex(r, g, b);
}

// =========================================================================
// 3. ESTRUTURA DE PREDIOS / RUAS
// =========================================================================

function getEstruturaPredios() {
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
        
        estruturaPredios[rua] = { Impar: prediosImpar, Par: prediosPar };

        [...prediosImpar, ...prediosPar].forEach(predio => {
            mapaVisualValido.add(`${rua}-${predio}`);
        });
    }
    return estruturaPredios;
}

getEstruturaPredios(); 

let dadosConsolidados = [];
let filtrosAtivos = { tipoMovimento: '', picking: '', curva: '', dataInicial: null, dataFinal: null };
const enderecosMap = new Map();

// =========================================================================
// 4. FUNÇÕES DE LEITURA DO EXCEL
// =========================================================================

async function carregarExcel(nomeArquivo) {
    try {
        const response = await fetch(nomeArquivo);
        if (!response.ok) {
            console.error(`Erro ao carregar arquivo ${nomeArquivo}:`, response.status, response.statusText);
            return null;
        }
        const arrayBuffer = await response.arrayBuffer();
        return XLSX.read(arrayBuffer, { type: 'array' });
    } catch (e) {
        console.error('Erro ao carregar Excel:', e);
        return null;
    }
}

// ---- NOVO: normalização de cabeçalho + buscarDadosValidos mais tolerante ----

function normalizaTextoCabecalho(t) {
    return String(t || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, ' '); // colapsa espaços em branco
}

function buscarDadosValidos(wb, nomeArquivo, colChaveEsperada) {
    if (!wb) return [];
    const chaveNorm = normalizaTextoCabecalho(colChaveEsperada);

    for (const nomeAba of wb.SheetNames) {
        const sheet = wb.Sheets[nomeAba];
        const dados = XLSX.utils.sheet_to_json(sheet, { defval: null });

        if (!dados.length) continue;

        const colunas = Object.keys(dados[0] || {});
        console.log(`[${nomeArquivo}] Aba: ${nomeAba} | colunas detectadas:`, colunas);

        // Tenta encontrar uma coluna com nome equivalente à chave esperada
        let colunaReal = colunas.find(c => normalizaTextoCabecalho(c) === chaveNorm);

        if (!colunaReal) continue;

        console.log(
            `[${nomeArquivo}] Usando coluna "${colunaReal}" como "${colChaveEsperada}" com ${dados.length} linhas`
        );

        // Se o nome real for diferente do esperado, espelha o valor na chave esperada
        if (colunaReal !== colChaveEsperada) {
            dados.forEach(row => {
                if (row[colChaveEsperada] === undefined) {
                    row[colChaveEsperada] = row[colunaReal];
                }
            });
        }

        return dados;
    }

    console.warn(
        `[${nomeArquivo}] Nenhuma aba com coluna equivalente a "${colChaveEsperada}" encontrada`
    );
    return [];
}

// =========================================================================
// 5. PROCESSAMENTO DOS DADOS
// =========================================================================

function montarMapaEnderecos(dadosEnderecos) {
    enderecosMap.clear();

    dadosEnderecos.forEach(linha => {
        const rua = normalizarEndereco(linha[CONFIG.colsEnderecos.rua]);
        const predio = normalizarEndereco(linha[CONFIG.colsEnderecos.predio]);
        const nivel = linha[CONFIG.colsEnderecos.nivel];
        const apto = linha[CONFIG.colsEnderecos.apto];
        const tipoEndereco = (linha[CONFIG.colsEnderecos.tipoEndereco] || '').toString().toUpperCase();

        if (rua == null || predio == null || !nivel || !apto) return;

        const key = `${rua}-${predio}-${nivel}-${apto}`;
        enderecosMap.set(key, {
            rua,
            predio,
            nivel,
            apto,
            tipoEndereco,
            tipoMapa: obterClassificacaoNivel(nivel),
            curva: obterCurvaPorProduto(null)
        });
    });

    console.log('Mapa de endereços montado. Total:', enderecosMap.size);
}

function consolidarMovimentos(dadosConvocacoes, dadosAbastecimento) {
    const mapaConsolidado = new Map();

    function adicionarMovimento(linha, tipoFonte) {
        let rua, predio, nivel, apto, tipoMov, dataRaw;

        if (tipoFonte === 'convocacao') {
            rua = normalizarEndereco(linha[CONFIG.colsConvocacoes.rua]);
            predio = normalizarEndereco(linha[CONFIG.colsConvocacoes.predio]);
            nivel = linha[CONFIG.colsConvocacoes.nivel];
            apto = linha[CONFIG.colsConvocacoes.apto];
            tipoMov = (linha[CONFIG.colsConvocacoes.tipoMov] || '').toString();
            dataRaw = linha[CONFIG.colsConvocacoes.data];
        } else {
            rua = normalizarEndereco(linha[CONFIG.colsAbastecimento.rua]);
            predio = normalizarEndereco(linha[CONFIG.colsAbastecimento.predio]);
            nivel = linha[CONFIG.colsAbastecimento.nivel];
            apto = linha[CONFIG.colsAbastecimento.apto];
            tipoMov = (linha[CONFIG.colsAbastecimento.tipoMov] || '').toString();
            dataRaw = linha[CONFIG.colsAbastecimento.data];
        }

        if (rua == null || predio == null || !nivel || !apto) return;

        const dataMov = parseDataBr(dataRaw);
        if (!dataMov) return;

        const keyEndereco = `${rua}-${predio}-${nivel}-${apto}`;
        const metaEndereco = enderecosMap.get(keyEndereco);

        const tipoMapa = metaEndereco ? metaEndereco.tipoMapa : obterClassificacaoNivel(nivel);
        const curva = metaEndereco ? metaEndereco.curva : obterCurvaPorProduto(null);

        const keyPredio = `${rua}-${predio}`;
        let registro = mapaConsolidado.get(keyPredio);

        if (!registro) {
            registro = {
                rua,
                predio,
                totalMovimentos: 0,
                movimentos: [],
                contagemPorTipo: {},
                contagemPorFonte: { convocacao: 0, abastecimento: 0 },
                tipoMapaPredominante: tipoMapa,
                curvaPredominante: curva
            };
            mapaConsolidado.set(keyPredio, registro);
        }

        const mov = {
            rua,
            predio,
            nivel,
            apto,
            tipoMov,
            data: dataMov,
            fonte: tipoFonte,
            tipoMapa,
            curva
        };

        registro.movimentos.push(mov);
        registro.totalMovimentos++;
        registro.contagemPorTipo[tipoMov] = (registro.contagemPorTipo[tipoMov] || 0) + 1;
        registro.contagemPorFonte[tipoFonte] += 1;
    }

    dadosConvocacoes.forEach(l => adicionarMovimento(l, 'convocacao'));
    dadosAbastecimento.forEach(l => adicionarMovimento(l, 'abastecimento'));

    const resultado = Array.from(mapaConsolidado.values());
    console.log('Consolidação concluída. Total prédios:', resultado.length);
    return resultado;
}

function aplicarFiltros(dados) {
    return dados.filter(reg => {
        if (filtrosAtivos.tipoMovimento) {
            const tipos = Object.keys(reg.contagemPorTipo || {});
            if (!tipos.includes(filtrosAtivos.tipoMovimento)) return false;
        }

        if (filtrosAtivos.picking) {
            const esperado = filtrosAtivos.picking;
            if (reg.tipoMapaPredominante !== esperado) {
                let encontrou = false;
                for (const mov of reg.movimentos) {
                    if (mov.tipoMapa === esperado) {
                        encontrou = true;
                        break;
                    }
                }
                if (!encontrou) return false;
            }
        }

        if (filtrosAtivos.curva) {
            const esperadoCurva = filtrosAtivos.curva;
            if (reg.curvaPredominante !== esperadoCurva) {
                let encontrou = false;
                for (const mov of reg.movimentos) {
                    if (mov.curva === esperadoCurva) {
                        encontrou = true;
                        break;
                    }
                }
                if (!encontrou) return false;
            }
        }

        if (filtrosAtivos.dataInicial || filtrosAtivos.dataFinal) {
            const temDentro = reg.movimentos.some(m =>
                dataDentroIntervalo(m.data, filtrosAtivos.dataInicial, filtrosAtivos.dataFinal)
            );
            if (!temDentro) return false;
        }

        return true;
    });
}

// =========================================================================
// 6. RENDERIZAÇÃO DO MAPA
// =========================================================================

function calcularMinMaxMovimentos(dados) {
    if (!dados.length) return { min: 0, max: 0 };
    let min = Infinity;
    let max = -Infinity;
    dados.forEach(reg => {
        if (reg.totalMovimentos < min) min = reg.totalMovimentos;
        if (reg.totalMovimentos > max) max = reg.totalMovimentos;
    });
    if (!isFinite(min) || !isFinite(max)) return { min: 0, max: 0 };
    return { min, max };
}

function renderizarMapa() {
    const container = document.querySelector('.container');
    container.innerHTML = '';

    const dadosFiltrados = aplicarFiltros(dadosConsolidados);
    const { min, max } = calcularMinMaxMovimentos(dadosFiltrados);

    const mapaPorRua = new Map();
    dadosFiltrados.forEach(reg => {
        const keyRua = reg.rua;
        if (!mapaPorRua.has(keyRua)) {
            mapaPorRua.set(keyRua, []);
        }
        mapaPorRua.get(keyRua).push(reg);
    });

    RUAS.forEach(rua => {
        const sec = document.createElement('section');
        sec.classList.add('rua-section');

        const tituloRua = document.createElement('h2');
        tituloRua.textContent = `RUA ${rua}`;
        tituloRua.classList.add('titulo-rua');

        sec.appendChild(tituloRua);

        const grid = document.createElement('div');
        grid.classList.add('grid-rua');

        const prediosRua = estruturaPredios[rua];
        if (!prediosRua) return;

        [...prediosRua.Impar, ...prediosRua.Par].forEach(predio => {
            const keyPredio = `${rua}-${predio}`;
            const dadosPredio = dadosFiltrados.find(r => r.rua === rua && r.predio === predio);

            const card = document.createElement('div');
            card.classList.add('predio-card');

            const titulo = document.createElement('div');
            titulo.classList.add('predio-titulo');
            titulo.textContent = `Prédio ${predio}`;
            card.appendChild(titulo);

            const bloco = document.createElement('div');
            bloco.classList.add('predio');

            if (!dadosPredio) {
                bloco.style.backgroundColor = CONFIG.cores.vazio;
                bloco.classList.add('predio-vazio');
            } else {
                const cor = gerarGradienteCor(dadosPredio.totalMovimentos, min, max);
                bloco.style.backgroundColor = cor;

                bloco.addEventListener('click', () => {
                    exibirDetalhesPredio(dadosPredio);
                });
            }

            const qnt = document.createElement('span');
            qnt.classList.add('predio-quantidade');
            qnt.textContent = dadosPredio ? dadosPredio.totalMovimentos : '-';

            bloco.appendChild(qnt);
            card.appendChild(bloco);

            grid.appendChild(card);
        });

        sec.appendChild(grid);
        container.appendChild(sec);
    });
}

function exibirDetalhesPredio(registro) {
    const detalhesContainer = document.querySelector('.detalhes-container');
    const detalhesTitulo = detalhesContainer.querySelector('.detalhes-titulo');
    const detalhesBody = detalhesContainer.querySelector('.detalhes-body');

    detalhesTitulo.textContent = `Detalhes - Rua ${registro.rua}, Prédio ${registro.predio}`;

    detalhesBody.innerHTML = '';

    const infoResumo = document.createElement('div');
    infoResumo.classList.add('detalhes-resumo');

    const total = document.createElement('p');
    total.textContent = `Total de movimentos: ${registro.totalMovimentos}`;

    const tipos = document.createElement('p');
    const tiposStr = Object.entries(registro.contagemPorTipo || {})
        .map(([tipo, qt]) => `${tipo}: ${qt}`)
        .join(' | ');
    tipos.textContent = `Por tipo: ${tiposStr || 'N/A'}`;

    const fontes = document.createElement('p');
    const fontesStr = Object.entries(registro.contagemPorFonte || {})
        .map(([fonte, qt]) => `${fonte}: ${qt}`)
        .join(' | ');
    fontes.textContent = `Por fonte: ${fontesStr || 'N/A'}`;

    const tipoMapaP = document.createElement('p');
    tipoMapaP.textContent = `Tipo predominante (picking/armazenagem): ${registro.tipoMapaPredominante}`;

    const curvaP = document.createElement('p');
    curvaP.textContent = `Curva predominante: ${registro.curvaPredominante}`;

    infoResumo.appendChild(total);
    infoResumo.appendChild(tipos);
    infoResumo.appendChild(fontes);
    infoResumo.appendChild(tipoMapaP);
    infoResumo.appendChild(curvaP);

    detalhesBody.appendChild(infoResumo);

    const tabela = document.createElement('table');
    tabela.classList.add('tabela-detalhes');

    const thead = document.createElement('thead');
    thead.innerHTML = `
        <tr>
            <th>Nível</th>
            <th>Apartamento</th>
            <th>Tipo de movimento</th>
            <th>Fonte</th>
            <th>Data</th>
            <th>Picking/Armazenagem</th>
            <th>Curva</th>
        </tr>
    `;
    tabela.appendChild(thead);

    const tbody = document.createElement('tbody');

    const movimentosOrdenados = [...registro.movimentos].sort((a, b) => {
        const ta = a.data?.getTime?.() || 0;
        const tb = b.data?.getTime?.() || 0;
        return tb - ta;
    });

    movimentosOrdenados.forEach(mov => {
        const tr = document.createElement('tr');

        const tdNivel = document.createElement('td');
        tdNivel.textContent = mov.nivel || '';

        const tdApto = document.createElement('td');
        tdApto.textContent = mov.apto || '';

        const tdTipoMov = document.createElement('td');
        tdTipoMov.textContent = mov.tipoMov || '';

        const tdFonte = document.createElement('td');
        tdFonte.textContent = mov.fonte || '';

        const tdData = document.createElement('td');
        tdData.textContent = mov.data ? formatarDataCurta(mov.data) : '';

        const tdTipoMapa = document.createElement('td');
        tdTipoMapa.textContent = mov.tipoMapa || '';

        const tdCurva = document.createElement('td');
        tdCurva.textContent = mov.curva || '';

        tr.appendChild(tdNivel);
        tr.appendChild(tdApto);
        tr.appendChild(tdTipoMov);
        tr.appendChild(tdFonte);
        tr.appendChild(tdData);
        tr.appendChild(tdTipoMapa);
        tr.appendChild(tdCurva);

        tbody.appendChild(tr);
    });

    tabela.appendChild(tbody);
    detalhesBody.appendChild(tabela);

    detalhesContainer.classList.add('ativo');
}

function fecharDetalhes() {
    const detalhesContainer = document.querySelector('.detalhes-container');
    detalhesContainer.classList.remove('ativo');
}

// =========================================================================
// 7. INICIALIZAÇÃO E EVENTOS
// =========================================================================

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

    try {
        const dadosEnderecos = buscarDadosValidos(wbEnd, CONFIG.arquivos.enderecos, CONFIG.colsEnderecos.rua);
        const dadosConvocacoes = buscarDadosValidos(wbConv, CONFIG.arquivos.convocacoes, CONFIG.colsConvocacoes.rua);
        const dadosAbastecimento = buscarDadosValidos(wbAbast, CONFIG.arquivos.abastecimento, CONFIG.colsAbastecimento.rua);

        montarMapaEnderecos(dadosEnderecos);
        dadosConsolidados = consolidarMovimentos(dadosConvocacoes, dadosAbastecimento);

        if (!dadosConsolidados.length) {
            erroDiv.textContent = 'Não foram encontrados dados consolidados a partir das planilhas.';
            erroDiv.classList.add('ativo');
        } else {
            erroDiv.classList.remove('ativo');
        }

        renderizarMapa();
    } catch (e) {
        console.error('Erro no processamento geral:', e);
        erroDiv.textContent = 'Erro ao processar as informações das planilhas. Verifique o console para detalhes.';
        erroDiv.classList.add('ativo');
    } finally {
        loading.classList.remove('active');
    }
}

// =========================================================================
// 8. CONTROLES DE FILTRO (UI)
// =========================================================================

document.addEventListener('DOMContentLoaded', () => {
    flatpickr('.datepicker', {
        dateFormat: 'd/m/Y',
        locale: 'pt'
    });

    const tipoMovimentoSelect = document.getElementById('tipoMovimento');
    const pickingSelect = document.getElementById('picking');
    const curvaSelect = document.getElementById('curva');
    const dataInicioInput = document.getElementById('dataInicio');
    const dataFimInput = document.getElementById('dataFim');
    const btnAplicarFiltros = document.getElementById('btnAplicarFiltros');
    const btnLimparFiltros = document.getElementById('btnLimparFiltros');

    btnAplicarFiltros.addEventListener('click', () => {
        filtrosAtivos.tipoMovimento = tipoMovimentoSelect.value || '';
        filtrosAtivos.picking = pickingSelect.value || '';
        filtrosAtivos.curva = curvaSelect.value || '';

        const dInicio = dataInicioInput.value ? parseDataBr(dataInicioInput.value) : null;
        const dFim = dataFimInput.value ? parseDataBr(dataFimInput.value) : null;

        filtrosAtivos.dataInicial = dInicio;
        filtrosAtivos.dataFinal = dFim;

        renderizarMapa();
    });

    btnLimparFiltros.addEventListener('click', () => {
        tipoMovimentoSelect.value = '';
        pickingSelect.value = '';
        curvaSelect.value = '';
        dataInicioInput.value = '';
        dataFimInput.value = '';

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

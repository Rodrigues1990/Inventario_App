// ═══════════════════════════════════════════════════════════════
//  INVENTÁRIO - Google Apps Script
//  ⚠️  ALTERE APENAS AS 3 LINHAS ABAIXO PARA CADA NOVA LOJA
// ═══════════════════════════════════════════════════════════════

const PREFIXO     = 'OT';        // Sigla da loja (2-3 letras, maiúsculas)
const TOTAL_LOTES = 999;          // Quantidade total de lotes desta loja
const NOME_LOJA   = 'OUTLET';     // Nome completo da loja (aparece nos relatórios)

// ─── NÃO ALTERE NADA ABAIXO DESTA LINHA ───────────────────────

const SPREADSHEET_ID = '';

const ABA_LOTES      = 'Lotes';
const ABA_CODIGOS    = 'Códigos';
const ABA_PROGRESSO  = 'Progresso';
const ABA_KPIS       = 'KPIs';
const ABA_CONFIG     = 'Config';
const ABA_RAW        = 'Raw';
const ABA_ALERTAS    = 'Alertas';
const ABA_HISTORICO  = 'Histórico';

const COR_LIDO      = '#34a853';
const COR_FALTANDO  = '#ea4335';
const COR_DUPLICADO = '#f9ab00';
const COR_SUSPEITO  = '#ff6d00';
const COR_CABECALHO = '#1a73e8';
const COR_RESUMO_BG = '#e8f0fe';

// ══════════════════════════════════════════════════════════════
//  RECEBE DADOS DO APP
//  ⚡ OTIMIZADO: só grava dados — não reconstrói abas visuais
//  Isso evita timeout quando vários funcionários enviam ao mesmo tempo
// ══════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const dados     = JSON.parse(e.postData.contents);
    const ss        = obterPlanilha();
    garantirAbas(ss);

    const loja      = dados.loja      || '';
    const nome      = dados.nome      || '';
    const lote      = dados.lote      || '';
    const codigos   = dados.codigos   || [];
    const timestamp = dados.timestamp || new Date().toLocaleString('pt-BR');

    // Grava dados — rápido, não trava
    ss.getSheetByName(ABA_LOTES).appendRow([timestamp, loja, lote, nome, codigos.length, codigos.join(' | ')]);
    const abaCodigos = ss.getSheetByName(ABA_CODIGOS);
    codigos.forEach((cod, idx) => abaCodigos.appendRow([timestamp, loja, lote, nome, idx + 1, cod]));
    ss.getSheetByName(ABA_RAW).appendRow([timestamp, loja, lote, nome, JSON.stringify(codigos)]);

    // Retorna OK imediatamente — abas visuais atualizam via trigger ou manualmente
    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput('ERRO: ' + err.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

function doGet(e) {
  // Se chamado com ?dados=1, retorna JSON com KPIs para o dashboard
  if (e && e.parameter && e.parameter.dados === '1') {
    try {
      const ss = obterPlanilha();
      garantirAbas(ss);
      const {
        lidos, faltandoLotes, pctLotes,
        totalProdutos, minsTrabalhados,
        diasComDados, mapaUsers,
        primeiroTsGlobal, ultimoTsGlobal
      } = lerDados(ss);
      const meta = lerMeta(ss);

      // Ritmo atual (últimos 30 min para ser mais preciso)
      const agora = new Date();
      const pctProd = meta > 0 ? Math.min((totalProdutos / meta) * 100, 100) : 0;
      const ritmoLotes = minsTrabalhados > 0 ? Math.round(lidos / (minsTrabalhados / 60)) : 0;
      const ritmoProd  = minsTrabalhados > 0 ? Math.round(totalProdutos / (minsTrabalhados / 60)) : 0;

      // Previsão de conclusão por produtos
      let previsaoMin = null;
      if (meta > 0 && totalProdutos > 0 && minsTrabalhados > 0 && totalProdutos < meta) {
        previsaoMin = Math.round((meta - totalProdutos) / (totalProdutos / minsTrabalhados));
      }

      // Por funcionário — simplificado
      const funcionarios = Object.entries(mapaUsers)
        .sort((a, b) => b[1].produtos - a[1].produtos)
        .map(([nome, info]) => ({
          nome,
          lotes: info.lotes,
          produtos: info.produtos,
          pct: lidos > 0 ? ((info.lotes / lidos) * 100).toFixed(1) : '0.0'
        }));

      const payload = {
        loja: NOME_LOJA,
        prefixo: PREFIXO,
        totalLotes: TOTAL_LOTES,
        metaProdutos: meta,
        lidos,
        faltandoLotes,
        pctLotes,
        totalProdutos,
        pctProdutos: pctProd.toFixed(1),
        minsTrabalhados,
        ritmoLotes,
        ritmoProd,
        previsaoMin,
        diasCount: diasComDados.length,
        inicio: primeiroTsGlobal ? primeiroTsGlobal.toLocaleString('pt-BR') : null,
        ultimaAtualizacao: agora.toLocaleString('pt-BR'),
        funcionarios
      };

      return ContentService
        .createTextOutput(JSON.stringify(payload))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ erro: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput('OK - Inventário ' + NOME_LOJA + ' online').setMimeType(ContentService.MimeType.TEXT);
}

// ══════════════════════════════════════════════════════════════
//  LEITURA CENTRALIZADA
// ══════════════════════════════════════════════════════════════
function lerDados(ss) {
  const abaLotes = ss.getSheetByName(ABA_LOTES);
  if (!abaLotes || abaLotes.getLastRow() < 2) {
    return {
      mapaLotes: {}, contagemLotes: {}, mapaUsers: {}, mapaDias: {}, diasComDados: [],
      lidos: 0, faltandoLotes: TOTAL_LOTES, pctLotes: '0.0',
      totalProdutos: 0, minsTrabalhados: 0, primeiroTsGlobal: null, ultimoTsGlobal: null
    };
  }

  const dados    = abaLotes.getDataRange().getValues();
  const mapaLotes  = {};
  const todosBruto = [];
  const mapaUsers  = {};
  const mapaDias   = {};

  for (let i = 1; i < dados.length; i++) {
    const tsRaw = dados[i][0];
    const lote  = String(dados[i][2]).trim().toUpperCase();
    const nome  = String(dados[i][3]).trim();
    const qtd   = Number(dados[i][4]) || 0;

    if (!lote.startsWith(PREFIXO)) continue;

    const tsDate   = parsarTimestamp(tsRaw);
    const loteNorm = lote.replace(/[\s-]/g, '');

    todosBruto.push({ loteNorm, nome, qtd, tsDate, tsRaw });
    mapaLotes[loteNorm] = { ts: tsRaw, tsDate, nome, qtd, loteOriginal: lote };

    // Por usuário
    if (!mapaUsers[nome]) mapaUsers[nome] = { lotes: 0, produtos: 0, diasAtivos: {} };
    mapaUsers[nome].lotes++;
    mapaUsers[nome].produtos += qtd;
    if (tsDate) {
      const d = tsDate.toLocaleDateString('pt-BR');
      if (!mapaUsers[nome].diasAtivos[d]) mapaUsers[nome].diasAtivos[d] = { primeiro: tsDate, ultimo: tsDate };
      else {
        if (tsDate < mapaUsers[nome].diasAtivos[d].primeiro) mapaUsers[nome].diasAtivos[d].primeiro = tsDate;
        if (tsDate > mapaUsers[nome].diasAtivos[d].ultimo)   mapaUsers[nome].diasAtivos[d].ultimo   = tsDate;
      }
    }

    // Por dia
    if (tsDate) {
      const d = tsDate.toLocaleDateString('pt-BR');
      if (!mapaDias[d]) mapaDias[d] = { lotes: 0, produtos: 0, primeiro: tsDate, ultimo: tsDate };
      mapaDias[d].lotes++;
      mapaDias[d].produtos += qtd;
      if (tsDate < mapaDias[d].primeiro) mapaDias[d].primeiro = tsDate;
      if (tsDate > mapaDias[d].ultimo)   mapaDias[d].ultimo   = tsDate;
    }
  }

  const contagemLotes = {};
  todosBruto.forEach(r => {
    if (!contagemLotes[r.loteNorm]) contagemLotes[r.loteNorm] = [];
    contagemLotes[r.loteNorm].push(r);
  });

  const lidos         = Object.keys(mapaLotes).length;
  const faltandoLotes = TOTAL_LOTES - lidos;
  const pctLotes      = ((lidos / TOTAL_LOTES) * 100).toFixed(1);
  const totalProdutos = Object.values(mapaUsers).reduce((s, u) => s + u.produtos, 0);

  let minsTrabalhados = 0;
  let diasComDados    = [];
  Object.entries(mapaDias).forEach(([diaStr, d]) => {
    const minsDia = Math.floor((d.ultimo - d.primeiro) / 60000);
    minsTrabalhados += minsDia;
    diasComDados.push({ diaStr, ...d, minsDia });
  });
  diasComDados.sort((a, b) => a.primeiro - b.primeiro);

  return {
    mapaLotes, contagemLotes, mapaUsers, mapaDias, diasComDados,
    lidos, faltandoLotes, pctLotes,
    totalProdutos, minsTrabalhados,
    primeiroTsGlobal: diasComDados.length > 0 ? diasComDados[0].primeiro : null,
    ultimoTsGlobal:   diasComDados.length > 0 ? diasComDados[diasComDados.length - 1].ultimo : null
  };
}

// ══════════════════════════════════════════════════════════════
//  ABA PROGRESSO
// ══════════════════════════════════════════════════════════════
function atualizarProgresso(ss) {
  if (!ss) ss = obterPlanilha();
  const aba    = ss.getSheetByName(ABA_PROGRESSO);
  const dados  = lerDados(ss);
  const secoes = lerSecoes(ss);
  const { mapaLotes, contagemLotes, lidos, faltandoLotes, pctLotes } = dados;

  aba.clearContents();
  aba.clearFormats();

  aba.getRange('A1:G1').merge()
    .setValue('PROGRESSO DO INVENTÁRIO — ' + NOME_LOJA)
    .setBackground(COR_CABECALHO).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  const resumo = [
    ['Atualizado em:',  new Date().toLocaleString('pt-BR'), '', '', '', '', ''],
    ['Total de lotes:', TOTAL_LOTES,   '', '', '', '', ''],
    ['Lotes lidos:',    lidos,         '', '', '', '', ''],
    ['Lotes faltando:', faltandoLotes, '', '', '', '', ''],
    ['Conclusão:',      pctLotes + '%', barraTexto(Number(pctLotes), 20), '', '', '', ''],
  ];
  aba.getRange(2, 1, 5, 7).setValues(resumo).setBackground(COR_RESUMO_BG);
  aba.getRange('A2:A6').setFontWeight('bold').setFontColor('#444');
  aba.getRange('B4').setBackground('#d9ead3').setFontColor('#1e4620').setFontWeight('bold');
  aba.getRange('B5').setBackground('#fce8e6').setFontColor('#b31412').setFontWeight('bold');
  aba.getRange('B6').setFontWeight('bold').setFontSize(13);
  aba.getRange('C6:G6').merge().setFontFamily('Courier New').setFontColor(COR_LIDO);
  aba.getRange('A1:G6').setBorder(true,true,true,true,true,true,'#c5cae9',SpreadsheetApp.BorderStyle.SOLID);

  let L = 8;
  aba.getRange(L, 1, 1, 7).merge().setValue('LEGENDA DE CORES')
    .setBackground('#455a64').setFontColor('#fff').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  L++;
  [[COR_LIDO,'#fff','LIDO — bipado normalmente'],
   [COR_FALTANDO,'#fff','FALTANDO — ainda não lido'],
   [COR_DUPLICADO,'#000','DUPLICADO — bipado mais de uma vez (verificar)'],
   [COR_SUSPEITO,'#fff','SUSPEITO — lido com 0 produtos (verificar)']
  ].forEach(([bg, fc, txt]) => {
    aba.getRange(L,1,1,7).merge().setValue(txt).setBackground(bg).setFontColor(fc).setFontSize(10);
    L++;
  });
  L++;

  const CAB = L;
  aba.getRange(CAB, 1, 1, 7)
    .setValues([['Lote','Status','Funcionário','Data','Hora','Qtd Itens','Seção']])
    .setBackground(COR_CABECALHO).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  aba.setFrozenRows(CAB);
  L++;

  const linhas = [], bgs = [], fcs = [];
  for (let n = 1; n <= TOTAL_LOTES; n++) {
    const chave  = PREFIXO + String(n).padStart(4, '0');
    const exibir = PREFIXO + ' ' + String(n).padStart(4, '0');
    const info   = mapaLotes[chave];
    const secao  = secaoDoLote(n, secoes);
    const vezes  = contagemLotes[chave] ? contagemLotes[chave].length : 0;

    if (info) {
      let dataStr = '', horaStr = '';
      if (info.tsDate) {
        dataStr = info.tsDate.toLocaleDateString('pt-BR');
        horaStr = info.tsDate.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
      }
      const status = vezes > 1 ? 'DUPLICADO (' + vezes + 'x)' : (info.qtd === 0 ? 'SUSPEITO' : 'LIDO');
      const cor    = vezes > 1 ? COR_DUPLICADO : (info.qtd === 0 ? COR_SUSPEITO : COR_LIDO);
      const fc     = vezes > 1 ? '#000000' : '#ffffff';
      linhas.push([exibir, status, info.nome, dataStr, horaStr, info.qtd, secao]);
      bgs.push(Array(7).fill(cor));
      fcs.push(Array(7).fill(fc));
    } else {
      linhas.push([exibir, 'FALTANDO', '', '', '', '', secao]);
      bgs.push(Array(7).fill(COR_FALTANDO));
      fcs.push(Array(7).fill('#ffffff'));
    }
  }

  const range = aba.getRange(L, 1, linhas.length, 7);
  range.setValues(linhas).setBackgrounds(bgs).setFontColors(fcs);
  aba.getRange(L, 2, linhas.length, 1).setFontWeight('bold');
  [110,120,150,110,90,80,160].forEach((w,i) => aba.setColumnWidth(i+1, w));
}

// ══════════════════════════════════════════════════════════════
//  ABA ALERTAS
// ══════════════════════════════════════════════════════════════
function atualizarAlertas(ss) {
  if (!ss) ss = obterPlanilha();
  const aba = ss.getSheetByName(ABA_ALERTAS);
  const { contagemLotes, mapaLotes } = lerDados(ss);

  aba.clearContents();
  aba.clearFormats();

  let L = 1;
  aba.getRange(L,1,1,6).merge()
    .setValue('ALERTAS — ' + NOME_LOJA)
    .setBackground('#b71c1c').setFontColor('#ffffff').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  L++;
  aba.getRange(L,1).setValue('Atualizado em: ' + new Date().toLocaleString('pt-BR')).setFontColor('#888').setFontSize(10);
  L += 2;

  const duplicados = Object.entries(contagemLotes).filter(([, v]) => v.length > 1);
  aba.getRange(L,1,1,6).merge()
    .setValue('⚠ LOTES DUPLICADOS (' + duplicados.length + ')')
    .setBackground(COR_DUPLICADO).setFontColor('#000').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  L++;
  if (duplicados.length > 0) {
    aba.getRange(L,1,1,6).setValues([['Lote','Leituras','Funcionários','Horários','Qtd 1ª','Qtd última']])
      .setBackground('#fff8e1').setFontWeight('bold');
    L++;
    duplicados.forEach(([loteNorm, leituras]) => {
      const exibir   = PREFIXO + ' ' + loteNorm.slice(PREFIXO.length);
      const funcs    = [...new Set(leituras.map(r => r.nome))].join(', ');
      const horarios = leituras.map(r => r.tsDate ? r.tsDate.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'}) : '—').join(' / ');
      aba.getRange(L,1,1,6).setValues([[exibir, leituras.length, funcs, horarios, leituras[0].qtd, leituras[leituras.length-1].qtd]])
        .setBackground('#fffde7').setFontColor('#000');
      L++;
    });
  } else {
    aba.getRange(L,1).setValue('Nenhum lote duplicado ✓').setFontColor('#34a853').setFontWeight('bold');
    L++;
  }
  L++;

  const suspeitos = Object.entries(mapaLotes).filter(([, info]) => info.qtd === 0);
  aba.getRange(L,1,1,6).merge()
    .setValue('⚠ LOTES COM 0 PRODUTOS (' + suspeitos.length + ')')
    .setBackground(COR_SUSPEITO).setFontColor('#fff').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  L++;
  if (suspeitos.length > 0) {
    aba.getRange(L,1,1,3).setValues([['Lote','Funcionário','Data/Hora']]).setBackground('#fff3e0').setFontWeight('bold');
    L++;
    suspeitos.forEach(([loteNorm, info]) => {
      aba.getRange(L,1,1,3).setValues([[PREFIXO + ' ' + loteNorm.slice(PREFIXO.length), info.nome, info.ts]])
        .setBackground('#fff8f0').setFontColor('#000');
      L++;
    });
  } else {
    aba.getRange(L,1).setValue('Nenhum lote suspeito ✓').setFontColor('#34a853').setFontWeight('bold');
  }
  [120,80,200,200,100,100].forEach((w,i) => aba.setColumnWidth(i+1, w));
}

// ══════════════════════════════════════════════════════════════
//  ABA KPIs
// ══════════════════════════════════════════════════════════════
function atualizarKPIs(ss) {
  if (!ss) ss = obterPlanilha();
  const aba    = ss.getSheetByName(ABA_KPIS);
  const meta   = lerMeta(ss);
  const secoes = lerSecoes(ss);
  const {
    mapaLotes, mapaUsers, diasComDados,
    lidos, faltandoLotes, pctLotes,
    totalProdutos, minsTrabalhados,
    primeiroTsGlobal, ultimoTsGlobal
  } = lerDados(ss);

  aba.clearContents();
  aba.clearFormats();
  let L = 1;

  aba.getRange(L,1,1,8).merge()
    .setValue('KPIs — ' + NOME_LOJA)
    .setBackground(COR_CABECALHO).setFontColor('#ffffff').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  L++;

  // ── PROGRESSO DE PRODUTOS ──
  aba.getRange(L,1,1,8).merge().setValue('PROGRESSO DE PRODUTOS')
    .setBackground('#1b5e20').setFontColor('#fff').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  L++;

  const pctProd    = meta > 0 ? Math.min((totalProdutos/meta)*100,100) : 0;
  const prodFalt   = meta > 0 ? Math.max(meta-totalProdutos,0) : 0;
  const ritmoPhStr = minsTrabalhados > 0 ? Math.round(totalProdutos/(minsTrabalhados/60))+' prod/h' : '—';
  let previsao = '—';
  if (meta > 0 && totalProdutos > 0 && minsTrabalhados > 0 && prodFalt > 0)
    previsao = fmtMin(Math.round(prodFalt/(totalProdutos/minsTrabalhados))) + ' restantes';
  else if (meta > 0 && totalProdutos >= meta) previsao = '✓ Meta atingida!';

  const bp = [
    ['Meta:', meta>0?meta.toLocaleString('pt-BR'):'Não definida', 'Lidos:', totalProdutos.toLocaleString('pt-BR'), 'Faltando:', prodFalt>0?prodFalt.toLocaleString('pt-BR'):'—','',''],
    ['Conclusão:', (meta>0?pctProd.toFixed(1)+'%':'—'), 'Ritmo:', ritmoPhStr, 'Previsão:', previsao,'',''],
    ['Barra:', barraTexto(pctProd,30),'','','','','',''],
  ];
  aba.getRange(L,1,bp.length,8).setValues(bp).setBackground('#e8f5e9');
  aba.getRange(L,1,bp.length,1).setFontWeight('bold').setFontColor('#1b5e20');
  aba.getRange(L,3,bp.length,1).setFontWeight('bold').setFontColor('#1b5e20');
  aba.getRange(L,5,bp.length,1).setFontWeight('bold').setFontColor('#1b5e20');
  aba.getRange(L,4).setBackground('#d9ead3').setFontColor('#1e4620').setFontWeight('bold');
  aba.getRange(L,6).setBackground('#fce8e6').setFontColor('#b31412').setFontWeight('bold');
  aba.getRange(L+1,2).setFontWeight('bold').setFontSize(13);
  aba.getRange(L+2,2,1,7).merge().setFontFamily('Courier New').setFontColor(COR_LIDO).setFontSize(12);
  L += bp.length + 1;

  // ── RESUMO GERAL ──
  aba.getRange(L,1,1,8).merge().setValue('RESUMO GERAL')
    .setBackground('#1565c0').setFontColor('#fff').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  L++;
  const numDias = diasComDados.length;
  const br = [
    ['Início:', primeiroTsGlobal?primeiroTsGlobal.toLocaleString('pt-BR'):'—', 'Fim:', ultimoTsGlobal?ultimoTsGlobal.toLocaleString('pt-BR'):'—','','','',''],
    ['Dias:', numDias, 'Horas efetivas:', fmtMin(minsTrabalhados), 'Média/dia:', numDias>0?fmtMin(Math.round(minsTrabalhados/numDias)):'—','',''],
    ['Lotes lidos:', lidos, 'Produtos:', totalProdutos.toLocaleString('pt-BR'), 'Ritmo lotes:', minsTrabalhados>0?Math.round(lidos/(minsTrabalhados/60))+' /h':'—','',''],
  ];
  aba.getRange(L,1,br.length,8).setValues(br).setBackground(COR_RESUMO_BG);
  [1,3,5].forEach(c => aba.getRange(L,c,br.length,1).setFontWeight('bold').setFontColor('#1a237e'));
  L += br.length + 1;

  // ── POR SEÇÃO ──
  if (secoes.length > 0) {
    aba.getRange(L,1,1,8).merge().setValue('PROGRESSO POR SEÇÃO')
      .setBackground('#4e342e').setFontColor('#fff').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
    L++;
    aba.getRange(L,1,1,8).setValues([['Seção','Total Lotes','Lidos','Faltando','%','Produtos','Barra','']])
      .setBackground('#6d4c41').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
    L++;
    secoes.forEach((sec, idx) => {
      let tL=0, lL=0, pL=0;
      for (let n=sec.de; n<=sec.ate; n++) {
        tL++;
        const ch = PREFIXO + String(n).padStart(4,'0');
        if (mapaLotes[ch]) { lL++; pL += mapaLotes[ch].qtd; }
      }
      const pSec = tL>0?((lL/tL)*100).toFixed(1):'0.0';
      aba.getRange(L,1,1,8).setValues([[sec.nome,tL,lL,tL-lL,pSec+'%',pL,barraTexto(Number(pSec),15),'']])
        .setBackground(idx%2===0?'#efebe9':'#ffffff').setFontColor('#222');
      aba.getRange(L,1).setFontWeight('bold');
      aba.getRange(L,7).setFontFamily('Courier New').setFontColor(lL===tL?COR_LIDO:COR_FALTANDO);
      L++;
    });
    L++;
  }

  // ── POR DIA ──
  aba.getRange(L,1,1,8).merge().setValue('DESEMPENHO POR DIA')
    .setBackground('#e65100').setFontColor('#fff').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  L++;
  aba.getRange(L,1,1,8).setValues([['Data','Início','Fim','Horas ativas','Lotes','Produtos','Lotes/h','Prods/h']])
    .setBackground('#bf360c').setFontColor('#fff').setFontWeight('bold');
  L++;
  diasComDados.forEach((d, idx) => {
    const lh = d.minsDia>0?Math.round(d.lotes/(d.minsDia/60)):0;
    const ph = d.minsDia>0?Math.round(d.produtos/(d.minsDia/60)):0;
    aba.getRange(L,1,1,8).setValues([[d.diaStr,
      d.primeiro.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'}),
      d.ultimo.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'}),
      fmtMin(d.minsDia),d.lotes,d.produtos,lh,ph]])
      .setBackground(idx%2===0?'#fff3e0':'#ffffff').setFontColor('#222');
    aba.getRange(L,1).setFontWeight('bold');
    L++;
  });
  const tlh = minsTrabalhados>0?Math.round(lidos/(minsTrabalhados/60)):0;
  const tph = minsTrabalhados>0?Math.round(totalProdutos/(minsTrabalhados/60)):0;
  aba.getRange(L,1,1,8).setValues([['TOTAL','','',fmtMin(minsTrabalhados),lidos,totalProdutos,tlh,tph]])
    .setBackground('#e65100').setFontColor('#fff').setFontWeight('bold');
  L += 2;

  // ── POR FUNCIONÁRIO ──
  aba.getRange(L,1,1,8).merge().setValue('DESEMPENHO POR FUNCIONÁRIO')
    .setBackground('#4a148c').setFontColor('#fff').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  L++;
  aba.getRange(L,1,1,8).setValues([['Funcionário','Dias','Horas efetivas','Lotes','% lotes','Produtos','Lotes/h','Prods/h']])
    .setBackground('#7b1fa2').setFontColor('#fff').setFontWeight('bold');
  L++;
  const total_lidos_n = lidos > 0 ? lidos : 1;
  Object.entries(mapaUsers).sort((a,b) => b[1].produtos-a[1].produtos).forEach(([nomeU, info], idx) => {
    let minsEf = 0;
    Object.values(info.diasAtivos).forEach(d => { minsEf += Math.floor((d.ultimo-d.primeiro)/60000); });
    const dias    = Object.keys(info.diasAtivos).length;
    const pctU    = ((info.lotes/total_lidos_n)*100).toFixed(1)+'%';
    const lhU     = minsEf>0?Math.round(info.lotes/(minsEf/60)):'—';
    const phU     = minsEf>0?Math.round(info.produtos/(minsEf/60)):'—';
    aba.getRange(L,1,1,8).setValues([[nomeU,dias,fmtMin(minsEf),info.lotes,pctU,info.produtos,lhU,phU]])
      .setBackground(idx%2===0?'#f3e5f5':'#ffffff').setFontColor('#222');
    aba.getRange(L,1).setFontWeight('bold');
    L++;
  });
  aba.getRange(L,1,1,8).setValues([['TOTAL',numDias,fmtMin(minsTrabalhados),lidos,'100%',totalProdutos,tlh,tph]])
    .setBackground('#4a148c').setFontColor('#fff').setFontWeight('bold');
  L += 2;

  [150,80,120,100,80,120,90,100].forEach((w,i) => aba.setColumnWidth(i+1, w));
  aba.getRange(L,1).setValue('Atualizado em: '+new Date().toLocaleString('pt-BR')).setFontColor('#888').setFontSize(10);
}

// ══════════════════════════════════════════════════════════════
//  ABA HISTÓRICO
// ══════════════════════════════════════════════════════════════
function salvarHistorico(ss) {
  if (!ss) ss = obterPlanilha();
  const aba = ss.getSheetByName(ABA_HISTORICO);
  const { lidos, totalProdutos, minsTrabalhados, diasComDados, primeiroTsGlobal, ultimoTsGlobal } = lerDados(ss);
  const meta = lerMeta(ss);
  aba.appendRow([
    new Date().toLocaleString('pt-BR'), NOME_LOJA, PREFIXO, TOTAL_LOTES, meta,
    lidos, TOTAL_LOTES-lidos, ((lidos/TOTAL_LOTES)*100).toFixed(1)+'%',
    totalProdutos, meta>0?((totalProdutos/meta)*100).toFixed(1)+'%':'—',
    diasComDados.length, fmtMin(minsTrabalhados),
    primeiroTsGlobal?primeiroTsGlobal.toLocaleString('pt-BR'):'—',
    ultimoTsGlobal?ultimoTsGlobal.toLocaleString('pt-BR'):'—',
  ]);
}

// ══════════════════════════════════════════════════════════════
//  BACKUP HORÁRIO (trigger automático)
// ══════════════════════════════════════════════════════════════
function salvarBackupHorario() {
  const ss     = obterPlanilha();
  const abaRaw = ss.getSheetByName(ABA_RAW);
  if (!abaRaw) return;

  const agora      = new Date();
  const nomeBackup = 'BACKUP_' + Utilities.formatDate(agora, Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  if (ss.getSheetByName(nomeBackup)) return;

  const backup = abaRaw.copyTo(ss);
  backup.setName(nomeBackup);
  backup.hideSheet();

  // Apaga backups com mais de 48h
  const limite = new Date(agora.getTime() - 48*60*60*1000);
  ss.getSheets().forEach(aba => {
    if (!aba.getName().startsWith('BACKUP_')) return;
    const m = aba.getName().match(/BACKUP_(\d{8})_(\d{4})/);
    if (!m) return;
    const ts = new Date(parseInt(m[1].substr(0,4)),parseInt(m[1].substr(4,2))-1,parseInt(m[1].substr(6,2)),parseInt(m[2].substr(0,2)),parseInt(m[2].substr(2,2)));
    if (ts < limite) ss.deleteSheet(aba);
  });
  Logger.log('Backup salvo: ' + nomeBackup);
}

// ══════════════════════════════════════════════════════════════
//  ABA CONFIG
// ══════════════════════════════════════════════════════════════
function criarAbaConfig(ss) {
  if (ss.getSheetByName(ABA_CONFIG)) return;
  const aba = ss.insertSheet(ABA_CONFIG, 2);

  aba.getRange('A1:D1').merge()
    .setValue('CONFIGURAÇÕES — ' + NOME_LOJA)
    .setBackground(COR_CABECALHO).setFontColor('#fff').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');

  const config = [
    ['Loja:',               NOME_LOJA,   '', ''],
    ['Prefixo dos lotes:',  PREFIXO,     '', ''],
    ['Total de lotes:',     TOTAL_LOTES, '', ''],
    ['Meta de produtos:',   0,           '', '← EDITE AQUI: quantidade esperada de produtos em loja'],
  ];
  aba.getRange(2,1,config.length,4).setValues(config);
  aba.getRange('B5').setBackground('#fffde7').setFontSize(14).setFontWeight('bold').setFontColor('#e65100').setHorizontalAlignment('center');
  aba.getRange(2,1,config.length,1).setFontWeight('bold').setFontColor('#444');
  aba.getRange('D5').setFontColor('#888').setFontSize(10);

  let L = config.length + 3;
  aba.getRange(L,1,1,4).merge()
    .setValue('MAPA DE SEÇÕES — edite conforme a divisão física da loja')
    .setBackground('#4e342e').setFontColor('#fff').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  L++;
  aba.getRange(L,1,1,4).setValues([['Nome da Seção','Lote inicial (número)','Lote final (número)','']])
    .setBackground('#6d4c41').setFontColor('#fff').setFontWeight('bold');
  L++;
  const exemplos = [
    ['Calçados Masculino', 1,   150, ''],
    ['Calçados Feminino',  151, 300, ''],
    ['Moda Feminina',      301, 500, ''],
    ['Moda Masculina',     501, 700, ''],
    ['Infantil',           701, 850, ''],
    ['Acessórios',         851, 999, ''],
  ];
  aba.getRange(L,1,exemplos.length,4).setValues(exemplos).setBackground('#efebe9');
  aba.getRange(L,1,exemplos.length,1).setFontWeight('bold');
  [180,140,140,340].forEach((w,i) => aba.setColumnWidth(i+1, w));
}

// ══════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════
function lerMeta(ss) {
  const aba = ss.getSheetByName(ABA_CONFIG);
  if (!aba) return 0;
  return Number(aba.getRange('B5').getValue()) || 0;
}

function lerSecoes(ss) {
  const aba = ss.getSheetByName(ABA_CONFIG);
  if (!aba) return [];
  const dados = aba.getDataRange().getValues();
  const secoes = [];
  let ok = false;
  for (let i = 0; i < dados.length; i++) {
    if (String(dados[i][0]).trim() === 'Nome da Seção') { ok = true; continue; }
    if (!ok) continue;
    const nome = String(dados[i][0]).trim();
    const de   = Number(dados[i][1]);
    const ate  = Number(dados[i][2]);
    if (nome && de > 0 && ate > 0) secoes.push({ nome, de, ate });
  }
  return secoes;
}

function secaoDoLote(n, secoes) {
  for (const s of secoes) { if (n >= s.de && n <= s.ate) return s.nome; }
  return '—';
}

function fmtMin(min) {
  if (!min || min <= 0) return '0min';
  return (Math.floor(min/60)>0 ? Math.floor(min/60)+'h ' : '') + (min%60) + 'min';
}

function barraTexto(pct, tam) {
  const f = Math.max(0,Math.min(Math.round((pct/100)*tam),tam));
  return '█'.repeat(f)+'░'.repeat(tam-f);
}

function parsarTimestamp(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const s = String(val).replace(',','').trim();
  const m = s.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):?(\d{2})?/);
  if (m) return new Date(parseInt(m[3]),parseInt(m[2])-1,parseInt(m[1]),parseInt(m[4]),parseInt(m[5]),parseInt(m[6]||0));
  const d = new Date(val);
  return isNaN(d) ? null : d;
}

function garantirAbas(ss) {
  criarAba(ss, ABA_LOTES,   ['Data/Hora','Loja','Lote','Funcionário','Qtd Códigos','Códigos'],  [160,80,110,150,100,400]);
  criarAba(ss, ABA_CODIGOS, ['Data/Hora','Loja','Lote','Funcionário','Seq','Código de Barras'],[160,80,110,150,60,200]);
  criarAba(ss, ABA_RAW,     ['Data/Hora','Loja','Lote','Funcionário','JSON'],                  [160,80,110,150,400]);
  criarAbaConfig(ss);
  if (!ss.getSheetByName(ABA_PROGRESSO)) ss.insertSheet(ABA_PROGRESSO, 0);
  if (!ss.getSheetByName(ABA_KPIS))      ss.insertSheet(ABA_KPIS, 1);
  if (!ss.getSheetByName(ABA_ALERTAS))   ss.insertSheet(ABA_ALERTAS, 2);
  criarAbaHistorico(ss);
}

function criarAbaHistorico(ss) {
  if (ss.getSheetByName(ABA_HISTORICO)) return;
  const aba = ss.insertSheet(ABA_HISTORICO);
  const cab = ['Salvo em','Loja','Prefixo','Total Lotes','Meta','Lotes Lidos','Faltando','% Lotes','Produtos','% Produtos','Dias','Horas Efetivas','Início','Fim'];
  aba.appendRow(cab);
  aba.getRange(1,1,1,cab.length).setBackground(COR_CABECALHO).setFontColor('#fff').setFontWeight('bold').setFontSize(11);
  aba.setFrozenRows(1);
  cab.forEach((_,i) => aba.setColumnWidth(i+1, 130));
}

function criarAba(ss, nome, cabecalho, larguras) {
  if (ss.getSheetByName(nome)) return;
  const aba = ss.insertSheet(nome);
  aba.appendRow(cabecalho);
  aba.getRange(1,1,1,cabecalho.length).setBackground(COR_CABECALHO).setFontColor('#fff').setFontWeight('bold').setFontSize(11);
  aba.setFrozenRows(1);
  larguras.forEach((w,i) => aba.setColumnWidth(i+1, w));
}

function obterPlanilha() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

// ══════════════════════════════════════════════════════════════
//  FUNÇÕES MANUAIS — execute no editor do Apps Script
// ══════════════════════════════════════════════════════════════

// Atualiza Progresso + KPIs + Alertas com os dados atuais
function forcarAtualizacaoProgresso() {
  const ss = obterPlanilha();
  garantirAbas(ss);
  atualizarProgresso(ss);
  atualizarKPIs(ss);
  atualizarAlertas(ss);
  Logger.log('Tudo atualizado! ' + new Date().toLocaleString('pt-BR'));
}

// Rodar ao FINAL do inventário: salva no Histórico + backup final
function fecharInventario() {
  const ss = obterPlanilha();
  garantirAbas(ss);
  atualizarProgresso(ss);
  atualizarKPIs(ss);
  atualizarAlertas(ss);
  salvarHistorico(ss);
  salvarBackupHorario();
  Logger.log('Inventário fechado! ' + new Date().toLocaleString('pt-BR'));
}

// Rodar UMA vez por planilha para ativar backup automático a cada hora
function criarTriggerBackup() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'salvarBackupHorario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('salvarBackupHorario').timeBased().everyHours(1).create();
  Logger.log('Trigger de backup horário criado!');
}

// Rodar UMA vez por planilha para atualizar abas a cada 5 minutos automaticamente
function criarTriggerAtualizacao() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'forcarAtualizacaoProgresso') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('forcarAtualizacaoProgresso').timeBased().everyMinutes(5).create();
  Logger.log('Trigger de atualização a cada 5 minutos criado!');
}

// ══════════════════════════════════════════════════════════════
//  LIMPAR DUPLICATAS
//  Remove entradas duplicadas das abas Raw, Lotes e Códigos.
//  Para cada lote duplicado, mantém apenas a entrada com MAIS produtos.
//
//  ⚠️  ANTES DE RODAR: baixe a planilha (Arquivo → Baixar → Excel)
//  como backup. Após confirmar que está tudo ok, pode descartar.
// ══════════════════════════════════════════════════════════════
function limparDuplicatas() {
  const ss  = obterPlanilha();
  const log = [];

  function norm(s) {
    if (!s) return '';
    return String(s).trim().toUpperCase().replace(/[\s-]/g, '');
  }

  // ── 1. RAW ─────────────────────────────────────────────────
  const abaRaw = ss.getSheetByName(ABA_RAW);
  if (abaRaw && abaRaw.getLastRow() > 1) {
    const dados = abaRaw.getRange(2, 1, abaRaw.getLastRow() - 1, 5).getValues();

    // Para cada lote, guarda o índice da linha com JSON mais longo (mais produtos)
    const mapa = {};
    dados.forEach((row, idx) => {
      const k = norm(row[2]);
      if (!k) return;
      const tam = row[4] ? String(row[4]).length : 0;
      if (!mapa[k] || tam > mapa[k].tam) mapa[k] = { idx, tam };
    });

    const melhores = new Set(Object.values(mapa).map(v => v.idx));
    const deletar  = dados.map((_, i) => i).filter(i => !melhores.has(i)).map(i => i + 2);
    deletar.sort((a, b) => b - a).forEach(li => abaRaw.deleteRow(li));
    log.push('Raw: ' + deletar.length + ' linhas removidas. ' + Object.keys(mapa).length + ' lotes únicos.');
  }

  // ── 2. LOTES ───────────────────────────────────────────────
  const abaLotes = ss.getSheetByName(ABA_LOTES);
  if (abaLotes && abaLotes.getLastRow() > 1) {
    const dados = abaLotes.getRange(2, 1, abaLotes.getLastRow() - 1, 6).getValues();

    const mapa = {};
    dados.forEach((row, idx) => {
      const k = norm(row[2]);
      if (!k) return;
      const qtd = Number(row[4]) || 0;
      if (!mapa[k] || qtd > mapa[k].qtd) mapa[k] = { idx, qtd };
    });

    const melhores = new Set(Object.values(mapa).map(v => v.idx));
    const deletar  = dados.map((_, i) => i).filter(i => !melhores.has(i)).map(i => i + 2);
    deletar.sort((a, b) => b - a).forEach(li => abaLotes.deleteRow(li));
    log.push('Lotes: ' + deletar.length + ' linhas removidas. ' + Object.keys(mapa).length + ' lotes únicos.');
  }

  // ── 3. CÓDIGOS ─────────────────────────────────────────────
  // Agrupa por lote + timestamp. Para cada lote, mantém o grupo
  // (leitura) que tem a maior sequência — ou seja, mais produtos.
  const abaCod = ss.getSheetByName(ABA_CODIGOS);
  if (abaCod && abaCod.getLastRow() > 1) {
    const dados = abaCod.getRange(2, 1, abaCod.getLastRow() - 1, 6).getValues();

    // Calcula seq máxima por lote+ts
    const grupos = {}; // loteNorm -> { ts -> maxSeq }
    dados.forEach(row => {
      const k  = norm(row[2]);
      const ts = String(row[0]);
      const sq = Number(row[4]) || 0;
      if (!k) return;
      if (!grupos[k]) grupos[k] = {};
      if (!grupos[k][ts] || sq > grupos[k][ts]) grupos[k][ts] = sq;
    });

    // Melhor ts de cada lote = o que tem maior maxSeq
    const melhorTs = {};
    Object.entries(grupos).forEach(([k, tss]) => {
      melhorTs[k] = Object.entries(tss).sort((a, b) => b[1] - a[1])[0][0];
    });

    // Deleta linhas cujo ts não é o melhor para aquele lote
    const deletar = [];
    dados.forEach((row, idx) => {
      const k  = norm(row[2]);
      const ts = String(row[0]);
      if (!k) return;
      if (ts !== melhorTs[k]) deletar.push(idx + 2);
    });

    deletar.sort((a, b) => b - a).forEach(li => abaCod.deleteRow(li));
    log.push('Códigos: ' + deletar.length + ' linhas removidas.');
  }

  // ── 4. ATUALIZA VISUAIS ────────────────────────────────────
  atualizarProgresso(ss);
  atualizarKPIs(ss);
  atualizarAlertas(ss);

  log.forEach(l => Logger.log(l));
  Logger.log('✓ Limpeza concluída! ' + new Date().toLocaleString('pt-BR'));
}

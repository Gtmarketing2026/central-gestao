/* Cálculo do retrato de Negócio (menu Negócio → Dash de venda).
   Fica FORA do index.html porque roda em dois lugares: aqui na tela, quando alguém sobe a base à mão,
   e no agendador que atualiza sozinha a base de vendas do OneDrive (~/.central-gestao/negocio_sync.js).
   Copiar o cálculo para os dois faria os números divergirem com o tempo — então é um arquivo só.
   Nada aqui pode tocar em tela (document/window): o agendador roda no Node, sem navegador.
   Depende de XLSX (SheetJS), que na tela vem do CDN e no Node vem do pacote xlsx. */
if (typeof XLSX === 'undefined' && typeof require !== 'undefined') globalThis.XLSX = require('xlsx');

const NEG_DUR_PADRAO=6; // meses, quando o produto não diz — só afeta produto fora do padrão
function _negMeses(prod){
  const p=String(prod||'').toLowerCase();
  if(/vital[íi]cio|lifetime/.test(p))return -1;      // -1 = não expira
  if(/cr[ée]dito/.test(p))return 0;                   // crédito não dá prazo (conta só no LTV)
  const m=p.match(/(\d+)\s*(?:\+\s*\d+\s*)?m[eê]s(?:es)?/); // pega "12 meses" e também "12+2 meses"
  if(m)return parseInt(m[1],10)+((p.match(/(\d+)\s*\+\s*(\d+)\s*m[eê]s/)||[])[2]?parseInt(p.match(/(\d+)\s*\+\s*(\d+)\s*m[eê]s/)[2],10):0);
  return null;                                        // sem prazo: fica separado
}
function _negNum(v){
  if(typeof v==='number')return v;
  const s=String(v||'').replace(/[^\d,.-]/g,'').replace(/\.(?=\d{3}\b)/g,'').replace(',','.');
  const n=parseFloat(s);return isNaN(n)?0:n;
}
/* Status que representam dinheiro que entrou de verdade. O porquê de cada um está no comentário
   de ORDER_STATUS_GROUPS. "Estornado Parcialmente" fica aqui porque a maior parte do valor ficou. */
const NEG_VENDA=new Set(['pago','disponível','disponivel','finalizado','debitado','estornado parcialmente','aguardando estorno']);
const NEG_DEVOLVIDO=new Set(['estornado','chargeback']); // foi aluno, mas o dinheiro voltou
/* "Liberado Automaticamente" é acesso sem pagamento (R$ 0,00 em 100% das linhas) e tem dois sentidos
   bem diferentes: a maior parte é INSCRIÇÃO em processo seletivo, seleção de corretores ou simulado
   gratuito — essa gente nunca foi aluna. O resto é acesso de verdade a um curso (bolsista, parceiro),
   que é aluno sim, só que sem receita. A regra testa a NEGAÇÃO primeiro: é cortesia o que não for
   um desses processos. */
function _negCortesia(prod){
  return !/processo\s+seletivo|bolsas?\s+de\s+estudo|corretor|sistema\s+seriado|ssa\s*\/?\s*upe|^\s*teste\s*$/i.test(String(prod||''));
}
/* A base não tem coluna de estado — tem CEP. As faixas dos Correios são fixas e resolvem a UF sem
   depender de consulta externa. Cuidado com as faixas quebradas: DF e GO se intercalam, e AM aparece
   em dois pedaços com RR no meio. Por isso a lista é percorrida em ordem, do menor pro maior. */
const NEG_CEP_UF=[[19999,'SP'],[28999,'RJ'],[29999,'ES'],[39999,'MG'],[48999,'BA'],[49999,'SE'],
  [56999,'PE'],[57999,'AL'],[58999,'PB'],[59999,'RN'],[63999,'CE'],[64999,'PI'],[65999,'MA'],
  [68899,'PA'],[68999,'AP'],[69299,'AM'],[69399,'RR'],[69899,'AM'],[69999,'AC'],[72799,'DF'],
  [72999,'GO'],[73699,'DF'],[76799,'GO'],[76999,'RO'],[77999,'TO'],[78899,'MT'],[78999,'RO'],
  [79999,'MS'],[87999,'PR'],[89999,'SC'],[99999,'RS']];
function _negUF(cep){
  let d=String(cep||'').replace(/\D/g,'');
  // CEP guardado como número perde o zero à esquerda: 01310-100 vira 1310100. Sem devolver o zero,
  // um CEP de São Paulo (09xxx) viraria 9xxxxxx e cairia na faixa do Rio Grande do Sul.
  if(d.length===7)d='0'+d;
  if(d.length<5)return '';
  const p=parseInt(d.slice(0,5),10);
  if(!(p>=1000))return '';                       // CEP começa em 01000; abaixo disso é lixo de cadastro
  for(const [ate,uf] of NEG_CEP_UF)if(p<=ate)return uf;
  return '';
}
/* Segunda fonte de estado: o DDD do telefone. Não é tão preciso quanto o CEP (gente que muda de
   estado costuma manter o número), mas cobre quem comprou sem informar CEP — e é melhor do que
   deixar a venda fora do mapa. A tela mostra quanto veio de cada fonte. */
const NEG_DDD_UF={11:'SP',12:'SP',13:'SP',14:'SP',15:'SP',16:'SP',17:'SP',18:'SP',19:'SP',
  21:'RJ',22:'RJ',24:'RJ',27:'ES',28:'ES',31:'MG',32:'MG',33:'MG',34:'MG',35:'MG',37:'MG',38:'MG',
  41:'PR',42:'PR',43:'PR',44:'PR',45:'PR',46:'PR',47:'SC',48:'SC',49:'SC',
  51:'RS',53:'RS',54:'RS',55:'RS',61:'DF',62:'GO',64:'GO',63:'TO',65:'MT',66:'MT',67:'MS',
  68:'AC',69:'RO',71:'BA',73:'BA',74:'BA',75:'BA',77:'BA',79:'SE',81:'PE',87:'PE',82:'AL',83:'PB',
  84:'RN',85:'CE',88:'CE',86:'PI',89:'PI',91:'PA',93:'PA',94:'PA',92:'AM',97:'AM',95:'RR',96:'AP',
  98:'MA',99:'MA'};
function _negUFtel(...tels){
  for(const t of tels){
    let d=String(t||'').replace(/\D/g,'');
    if(d.startsWith('55')&&d.length>=12)d=d.slice(2);   // +55 do formato internacional
    if(d.length<10)continue;
    const uf=NEG_DDD_UF[+d.slice(0,2)];
    if(uf)return uf;
  }
  return '';
}
/* O nome do produto carrega DUAS informações: o produto em si e a turma em que ele foi vendido.
   "Mega Combo (12 meses de acesso) - 2026.1" e "... - 2025.1" são o MESMO produto, em turmas
   diferentes — 2026.1 é janeiro de 2026, 2023.10 é outubro de 2023. Sem separar os dois, o mesmo
   produto aparece uma vez por turma no ranking e nunca se vê quanto ele vale de verdade.
   A turma é diferente da data da compra: alguém pode comprar a turma 2026.1 em março. Por isso ela
   vira filtro próprio, e não um recorte de data.
   O que NÃO é referência de turma ("Antigo", "II", "INTENSIVO", "Acesso Parceiros") continua no nome:
   são ofertas diferentes, não a mesma coisa em outro mês. */
function _negNorm2(t){return String(t||'').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g,'')}
function _negProdRef(prod){
  const s=String(prod||'').replace(/\s+/g,' ').trim();
  const m=s.match(/[\s-]*(\d{4})(?:\.(\d{1,2}))?\s*$/);
  let base=s,ref='',ano=0,mes=0;
  if(m&&+m[1]>=2018&&+m[1]<=2100){
    base=s.slice(0,m.index);
    ano=+m[1];mes=m[2]?+m[2]:0;ref=ano+(mes?'.'+mes:'');
  }
  // o mesmo produto aparece como "Mega Combo - (12 meses…)" e "Mega Combo  (12 meses…)" conforme o ano
  base=base.replace(/\s*-\s*\(/,' (').replace(/\s+/g,' ').replace(/[\s-]+$/,'').trim();
  return {base:base||s,ref,ano,mes};
}
const NEG_MES_NOME=['','jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
function _negRefLabel(r){
  const [a,m]=String(r||'').split('.');
  return m?`${NEG_MES_NOME[+m]||m}/${a}`:a||'sem turma';
}
/* ---- Cupom = parceiro ----
   O cupom carrega DUAS coisas, como o nome do produto carrega a turma: quem indicou e o desconto.
   Na base real: ILKA e ILKA10, GABYNAMED com 10/55/60/65, SARAH com 55/65/70. O ranking por parceiro
   junta essas variacoes; o desconto vira detalhe dentro do parceiro.
   Cuidado: o sufixo de desconto tem 2 digitos. "CAMILA1000" e "WELLINGTON1000" NAO sao desconto de
   1000 - o 1000 vem do nome do produto (Combo 1000), entao 3+ digitos ficam no nome. */
function _cupomParceiro(c){
  const u=String(c||'').trim().toUpperCase().replace(/\s+/g,' ');
  if(!u)return {base:'',desc:''};
  const m=u.match(/^(.*?[A-Z])(\d{2})$/);
  if(m&&m[1].length>=2)return {base:m[1],desc:m[2]};  // "FP55" = FP com 55% — nome curto tambem e parceiro
  return {base:u,desc:''};
}
function _negData(v){
  if(v instanceof Date)return v;
  if(typeof v==='number'){const d=XLSX.SSF.parse_date_code(v);return d?new Date(d.y,d.m-1,d.d):null}
  const m=String(v||'').match(/(\d{2})\/(\d{2})\/(\d{4})/);return m?new Date(+m[3],+m[2]-1,+m[1]):null;
}
// separado do upload pra dar pra testar o cálculo sem depender de banco nem de sessão
async function negocioProcessar(arquivos,aviso,emAberto,extraLinhas){
  const arqs=Array.isArray(arquivos)?arquivos:[arquivos];
  const setMsg=t=>{if(aviso)aviso(t)};
  /* FATURADO x RECEBIDO. A base lança a venda pelo valor CHEIO na data em que foi feita — inclusive a
     recorrência, que na verdade entra mês a mês. O relatório de inadimplência diz, por pedido, quanto
     ainda não entrou. Cruzando pelo Id (que casa 100%), cada venda passa a ter os dois números:
     o que foi vendido e o que de fato caiu. Sem o relatório subido, recebido = faturado. */
  const _aberto=new Map(),_casados=new Set(),_abMesTot={},_abMesCasado={};
  for(const x of (emAberto||[]))if(x&&x.i){
    _aberto.set(String(x.i),{a:+x.a||0,m:String(x.m||'')});
    const mm=String(x.m||'');if(mm)_abMesTot[mm]=(_abMesTot[mm]||0)+(+x.a||0);
  }
  {
    /* Aceita mais de um arquivo porque o histórico veio partido: 2020 e 2021 saíram da plataforma
       em relatórios separados e a base grande começa em 2022. O Id do pedido é único na base inteira
       (conferido em 435.827 linhas: 435.827 ids distintos), então é ele que impede linha repetida se
       dois arquivos se sobrepuserem. */
    const linhas=[],vistos=new Set();let repetidas=0;
    for(let i=0;i<arqs.length;i++){
      setMsg(`Lendo ${arqs[i].name}${arqs.length>1?` (${i+1} de ${arqs.length})`:''}…`);
      const wb=XLSX.read(await arqs[i].arrayBuffer(),{type:'array',cellDates:false});
      const aba=wb.SheetNames.find(n=>/venda/i.test(n))||wb.SheetNames[0];
      for(const l of XLSX.utils.sheet_to_json(wb.Sheets[aba],{defval:''})){
        const id=String(l.Id||l.id||l.ID||'').trim();
        if(id){if(vistos.has(id)){repetidas++;continue}vistos.add(id)}
        l.__id=id;
        linhas.push(l);
      }
    }
    // histórico guardado (2020/2021 etc.): junta aqui pra não precisar reenviar toda vez que a base
    // atual é atualizada. Mesmo dedup por Id do arquivo do dia — se algum pedido histórico também
    // estiver no arquivo atual, o atual (mais recente) prevalece.
    for(const l of (extraLinhas||[])){
      const id=String(l.__id||l.Id||l.id||l.ID||'').trim();
      if(id){if(vistos.has(id)){repetidas++;continue}vistos.add(id)}
      l.__id=id;
      linhas.push(l);
    }
    if(!linhas.length)throw new Error('planilha vazia ou sem cabeçalho');
    setMsg('Processando as linhas…');
    const col=(l,...nomes)=>{for(const n of nomes){const k=Object.keys(l).find(x=>x.toLowerCase().trim()===n);if(k)return l[k]}return ''};
    const alunos={},mesMap={},diaMap={},prodMap={};let creditos=0;
    const cortesiaK={},seletivoK={},devolK={};let devN=0,devVal=0;
    const ufMap={},ufMesMap={};let semUF=0,ufPorCep=0,ufPorDdd=0;
    let recebidoTot=0,naoRecebido=0,vendasComAberto=0;
    /* Parceiros: o cupom é a chave. Preciso passar por TODAS as linhas, não só pelas vendas — o valor
       do módulo está justamente em ver cancelamento, estorno e pedido não pago por parceiro. */
    const cupMap={},cupDia={},cupProd={},cupUf={},cupMes={};
    const refMap={},prodRefMap={},cuboMap={};
    const devMes={},gratMes={},prodMesMap={};
    // acesso ao longo do tempo: guardo os intervalos pra depois montar "ativos ao fim de cada mês"
    const mesIdx=d=>d.getFullYear()*12+d.getMonth();
    const idxMes=i=>`${Math.floor(i/12)}-${String(i%12+1).padStart(2,'0')}`;
    const deltaAtivos={},expMes={};
    for(const l of linhas){
      const st=String(col(l,'status')).trim().toLowerCase();
      const dt=_negData(col(l,'data'));if(!dt)continue;
      const mail=String(col(l,'email','e-mail')).trim().toLowerCase();
      const cpf=String(col(l,'cpf')).replace(/\D/g,'');
      const chave=cpf.length>=11?'c:'+cpf:(mail?'e:'+mail:null);if(!chave)continue;
      const prod=String(col(l,'produto')).replace(/\s+/g,' ').trim();
      const val=_negNum(col(l,'total','valor'));
      /* Guardo o cupom CRU. Juntar ILKA com ILKA10 é decisão de leitura, não de gravação: assim a tela
         pode desligar o agrupamento sem precisar processar a base de novo. */
      const _cpRaw=String(col(l,'cupom')||'').trim().toUpperCase().replace(/\s+/g,' ');
      const _cp={base:_cpRaw};
      let _cu=null;
      if(_cp.base){
        _cu=cupMap[_cp.base]=cupMap[_cp.base]||{cupom:_cp.base,vendas:0,receita:0,recebido:0,aberto:0,
          alunos:{},cancel:0,cancel_val:0,estorno:0,estorno_val:0,aguard:0,aguard_val:0,naoautoriz:0,pri:'',ult:''};
        if(/^estornado$|^chargeback$|^estornado parcialmente$/.test(st)){_cu.estorno++;_cu.estorno_val+=val}
        else if(st==='cancelado'){_cu.cancel++;_cu.cancel_val+=val}
        else if(st==='aguardando pagamento'||st==='pagamento pendente'){_cu.aguard++;_cu.aguard_val+=val}
        else if(st==='não autorizada'||st==='nao autorizada')_cu.naoautoriz++;
      }
      if(st==='liberado automaticamente'){
        const ehCort=_negCortesia(prod);(ehCort?cortesiaK:seletivoK)[chave]=1;
        const cm=dt.toISOString().slice(0,7);
        (gratMes[cm]=gratMes[cm]||{mes:cm,cortesia:0,seletivo:0});gratMes[cm][ehCort?'cortesia':'seletivo']++;
        continue;
      }
      if(NEG_DEVOLVIDO.has(st)){
        devolK[chave]=1;devN++;devVal+=val;
        const dm=dt.toISOString().slice(0,7);
        (devMes[dm]=devMes[dm]||{mes:dm,compras:0,valor:0});devMes[dm].compras++;devMes[dm].valor+=val;
        continue;
      }
      if(!NEG_VENDA.has(st))continue;
      const meses=_negMeses(prod);if(meses===0)creditos++;
      const mk=dt.toISOString().slice(0,7),dk=dt.toISOString().slice(0,10);
      // o que ainda não entrou desta venda (0 quando o pedido não está no relatório de inadimplência)
      const _ab=_aberto.get(l.__id||'');
      const emAb=_ab?_ab.a:0;
      const rec=Math.max(0,val-emAb);
      if(emAb>0){
        const usado=Math.min(emAb,val);
        naoRecebido+=usado;vendasComAberto++;_casados.add(l.__id);
        if(_ab.m)_abMesCasado[_ab.m]=(_abMesCasado[_ab.m]||0)+usado;
      }
      recebidoTot+=rec;
      (mesMap[mk]=mesMap[mk]||{mes:mk,vendas:0,receita:0,recebido:0,novos:0});mesMap[mk].vendas++;mesMap[mk].receita+=val;mesMap[mk].recebido+=rec;
      // série DIÁRIA: é o que permite o calendário com intervalo livre e o comparativo de período
      (diaMap[dk]=diaMap[dk]||{dia:dk,vendas:0,receita:0,recebido:0,novos:0});diaMap[dk].vendas++;diaMap[dk].receita+=val;diaMap[dk].recebido+=rec;
      // estado: CEP primeiro (é o endereço declarado); DDD do telefone como segunda fonte
      let uf=_negUF(col(l,'cep'));
      if(uf)ufPorCep++;
      else{uf=_negUFtel(col(l,'celular'),col(l,'fone'),col(l,'telefone'));if(uf)ufPorDdd++;}
      if(_cu){
        const dkc=dt.toISOString().slice(0,10),mkc=dt.toISOString().slice(0,7);
        _cu.vendas++;_cu.receita+=val;_cu.recebido+=rec;_cu.aberto+=Math.min(emAb,val);_cu.alunos[chave]=1;
        if(!_cu.pri||dkc<_cu.pri)_cu.pri=dkc;if(!_cu.ult||dkc>_cu.ult)_cu.ult=dkc;
        const kd=_cp.base+'|'+dkc;(cupDia[kd]=cupDia[kd]||{cupom:_cp.base,dia:dkc,vendas:0,receita:0});cupDia[kd].vendas++;cupDia[kd].receita+=val;
        const km=_cp.base+'|'+mkc;(cupMes[km]=cupMes[km]||{cupom:_cp.base,mes:mkc,vendas:0,receita:0});cupMes[km].vendas++;cupMes[km].receita+=val;
        if(uf){const ku=_cp.base+'|'+uf;(cupUf[ku]=cupUf[ku]||{cupom:_cp.base,uf,vendas:0,receita:0});cupUf[ku].vendas++;cupUf[ku].receita+=val}
        if(prod){const pb=_negProdRef(prod).base,kp=_cp.base+'|'+pb;
          (cupProd[kp]=cupProd[kp]||{cupom:_cp.base,produto:pb,vendas:0,receita:0});cupProd[kp].vendas++;cupProd[kp].receita+=val}
      }
      if(prod){
        const pr=_negProdRef(prod);
        (prodMap[pr.base]=prodMap[pr.base]||{produto:pr.base,vendas:0,receita:0});prodMap[pr.base].vendas++;prodMap[pr.base].receita+=val;
        const pmk=pr.base+'|'+mk;
        (prodMesMap[pmk]=prodMesMap[pmk]||{produto:pr.base,mes:mk,vendas:0,receita:0});
        prodMesMap[pmk].vendas++;prodMesMap[pmk].receita+=val;
        const rk=pr.ref||'—';
        (refMap[rk]=refMap[rk]||{ref:rk,ano:pr.ano,mes:pr.mes,vendas:0,receita:0});refMap[rk].vendas++;refMap[rk].receita+=val;
        const prk=pr.base+''+rk;
        (prodRefMap[prk]=prodRefMap[prk]||{produto:pr.base,ref:rk,ano:pr.ano,mes:pr.mes,vendas:0,receita:0});
        prodRefMap[prk].vendas++;prodRefMap[prk].receita+=val;
        // cubo produto × estado × turma: é o que deixa filtrar pelos dois juntos sem inventar número
        const ck=prk+'|'+(uf||'');
        (cuboMap[ck]=cuboMap[ck]||{produto:pr.base,uf:uf||'',ano:pr.ano,mes:pr.mes,vendas:0,receita:0});
        cuboMap[ck].vendas++;cuboMap[ck].receita+=val;
      }
      if(uf){
        (ufMap[uf]=ufMap[uf]||{uf,vendas:0,receita:0,_al:{}});ufMap[uf].vendas++;ufMap[uf].receita+=val;ufMap[uf]._al[chave]=1;
        const uk=uf+'|'+mk;(ufMesMap[uk]=ufMesMap[uk]||{uf,mes:mk,vendas:0,receita:0,novos:0});
        ufMesMap[uk].vendas++;ufMesMap[uk].receita+=val;
      }else semUF++;
      const a=alunos[chave]=alunos[chave]||{n:0,tot:0,pri:dt,ult:dt,ivs:[],exp:null,vital:false,semPrazo:false,anos:{},uf};
      a.n++;a.tot+=val;if(dt<a.pri){a.pri=dt;a.uf=uf}if(dt>a.ult)a.ult=dt;
      (a.anos[dt.getFullYear()]=a.anos[dt.getFullYear()]||{compras:0,receita:0}),a.anos[dt.getFullYear()].compras++,a.anos[dt.getFullYear()].receita+=val;
      if(meses===-1){a.vital=true}
      else if(meses>0){const fim=new Date(dt);fim.setMonth(fim.getMonth()+meses);a.ivs.push([dt,fim]);if(!a.exp||fim>a.exp)a.exp=fim}
      else if(meses===null){a.semPrazo=true}
    }
    const hoje=new Date();let ativos=0,expirados=0,exp30=0,exp60=0,semPrazo=0;
    const permT=[],permR=[],relac=[];
    /* SAFRA = mês da PRIMEIRA compra do aluno. É o que permite responder "LTV e permanência de quem
       entrou neste período" — diferente do LTV do topo, que é da base inteira. Guardo a receita da
       VIDA INTEIRA de cada aluno na safra em que ele entrou; recortar a receita pelo período daria
       um número menor e enganoso, porque a maior parte do LTV acontece depois da entrada. */
    const safraMap={};
    for(const k in alunos){
      const a=alunos[k];
      if(a.ivs.length){
        const ivs=a.ivs.sort((x,y)=>x[0]-y[0]);let tot=0,ini=ivs[0][0],fim=ivs[0][1];
        const unidos=[];
        for(let j=1;j<ivs.length;j++){if(ivs[j][0]<=fim){if(ivs[j][1]>fim)fim=ivs[j][1]}else{tot+=fim-ini;unidos.push([ini,fim]);ini=ivs[j][0];fim=ivs[j][1]}}
        tot+=fim-ini;unidos.push([ini,fim]);
        const m=tot/(1000*60*60*24*30.44);permT.push(m);if(a.n>1)permR.push(m);a._perm=m;
        // cada intervalo de acesso marca entrada e saída: a soma acumulada vira "ativos ao fim do mês"
        for(const [i0,i1] of unidos){
          const a0=mesIdx(i0),a1=mesIdx(i1);
          deltaAtivos[a0]=(deltaAtivos[a0]||0)+1;deltaAtivos[a1+1]=(deltaAtivos[a1+1]||0)-1;
        }
      }
      if(a.vital){const v0=mesIdx(a.pri);deltaAtivos[v0]=(deltaAtivos[v0]||0)+1;}
      if(!a.vital&&a.exp){const e=idxMes(mesIdx(a.exp));(expMes[e]=expMes[e]||{mes:e,qtd:0}).qtd++;}
      if(a.n>1)relac.push((a.ult-a.pri)/(1000*60*60*24*30.44));
      if(a.vital)ativos++;
      else if(a.exp){if(a.exp>=hoje){ativos++;const dd=(a.exp-hoje)/(1000*60*60*24);if(dd<=30)exp30++;else if(dd<=60)exp60++}else expirados++}
      else semPrazo++;
      const pk=a.pri.toISOString().slice(0,7);if(mesMap[pk])mesMap[pk].novos++;
      (safraMap[pk]=safraMap[pk]||{mes:pk,alunos:0,receita:0,perm_soma:0,perm_n:0,recompradores:0});
      const _sf=safraMap[pk];_sf.alunos++;_sf.receita+=a.tot;if(a.n>1)_sf.recompradores++;
      if(a._perm!=null){_sf.perm_soma+=a._perm;_sf.perm_n++}
      const pd=a.pri.toISOString().slice(0,10);if(diaMap[pd])diaMap[pd].novos++;
      // aluno novo entra no estado da PRIMEIRA compra: é de lá que ele veio
      if(a.uf&&ufMesMap[a.uf+'|'+pk])ufMesMap[a.uf+'|'+pk].novos++;
    }
    const med=a=>a.length?a.reduce((s,x)=>s+x,0)/a.length:0;
    const chaves=Object.keys(alunos),rec=chaves.filter(k=>alunos[k].n>1);
    // quem SÓ aparece fora da compra: cortesia sem nunca ter pago, inscrito em seletivo, e quem
    // só teve compra devolvida. São contagens de pessoa, não de linha.
    const comprou=new Set(chaves);
    const soCortesia=Object.keys(cortesiaK).filter(k=>!comprou.has(k)).length;
    const soSeletivo=Object.keys(seletivoK).filter(k=>!comprou.has(k)&&!cortesiaK[k]).length;
    const soDevolvido=Object.keys(devolK).filter(k=>!comprou.has(k)).length;
    const fat=chaves.reduce((s,k)=>s+alunos[k].tot,0);
    const dados={gerado_em:new Date().toISOString(),
      periodo:{de:Object.keys(mesMap).sort()[0]+'-01',ate:hoje.toISOString().slice(0,10)},
      alunos:chaves.length,compras:chaves.reduce((s,k)=>s+alunos[k].n,0),faturamento:+fat.toFixed(2),
      recebido:+recebidoTot.toFixed(2),nao_recebido:+naoRecebido.toFixed(2),vendas_com_aberto:vendasComAberto,
      tem_recebido:(emAberto||[]).length>0,
      /* O relatorio de inadimplencia é maior do que o pedaço que afeta o faturamento: a maior parte dele
         são pedidos que NUNCA entraram como venda (aguardando pagamento). Guardar o total e o quanto
         casou com venda contada é o que permite a tela explicar por que os dois números diferem. */
      aberto_relatorio:+(emAberto||[]).reduce((t,x)=>t+(+x.a||0),0).toFixed(2),
      aberto_em_vendas:+naoRecebido.toFixed(2),
      aberto_fora_do_faturamento:+Math.max(0,(emAberto||[]).reduce((t,x)=>t+(+x.a||0),0)-naoRecebido).toFixed(2),
      contratos_relatorio:(emAberto||[]).length,contratos_casados:_casados.size,
      // mesmo cruzamento quebrado pelo mês de criação do contrato, pra tela de inadimplência filtrar
      aberto_por_mes:Object.keys(_abMesTot).sort().map(m=>({mes:m,total:+_abMesTot[m].toFixed(2),
        em_vendas:+(_abMesCasado[m]||0).toFixed(2),fora:+Math.max(0,_abMesTot[m]-(_abMesCasado[m]||0)).toFixed(2)})),
      ltv_medio:+(fat/chaves.length).toFixed(2),
      ltv_recomprador:+(rec.reduce((s,k)=>s+alunos[k].tot,0)/(rec.length||1)).toFixed(2),
      recompradores:rec.length,pct_recompra:+(rec.length*100/chaves.length).toFixed(1),
      permanencia_media:+med(permT).toFixed(1),permanencia_recomprador:+med(permR).toFixed(1),
      relacionamento_medio:+med(relac).toFixed(1),
      base:{ativos,expirados,expira_30d:exp30,expira_60d:exp60,sem_prazo:semPrazo},
      compras_por_credito:creditos,
      arquivos:arqs.map(a=>a.name),linhas_lidas:linhas.length,linhas_repetidas:repetidas,
      // estado sai do CEP. por_estado é a vida inteira (tem aluno único); por_estado_mes é o que
      // permite cruzar estado com período — por isso o recorte por estado é mensal, não diário.
      por_estado:Object.values(ufMap).map(x=>({uf:x.uf,vendas:x.vendas,receita:+x.receita.toFixed(2),alunos:Object.keys(x._al).length}))
        .sort((a,b)=>b.receita-a.receita),
      por_estado_mes:Object.values(ufMesMap).map(x=>({...x,receita:+x.receita.toFixed(2)})).sort((a,b)=>a.mes<b.mes?-1:1),
      vendas_sem_estado:semUF,uf_por_cep:ufPorCep,uf_por_ddd:ufPorDdd,
      // turmas: o sufixo do nome do produto ("- 2026.1" = janeiro/2026). Vira filtro, não recorte de data
      turmas:Object.values(refMap).map(x=>({...x,receita:+x.receita.toFixed(2)}))
        .sort((a,b)=>(b.ano-a.ano)||(b.mes-a.mes)),
      /* Séries mensais: são elas que deixam o filtro de datas valer para a tela inteira, e não só
         para o bloco de período. Ativos ao fim do mês sai da soma acumulada das entradas e saídas
         de acesso; expirados por mês, do vencimento de cada aluno. */
      ativos_mes:(()=>{
        const idxs=Object.keys(deltaAtivos).map(Number).sort((a,b)=>a-b);
        if(!idxs.length)return [];
        const fim=Math.max(idxs[idxs.length-1],mesIdx(hoje)+24);
        const out=[];let acc=0;
        for(let i=idxs[0];i<=fim;i++){acc+=deltaAtivos[i]||0;out.push({mes:idxMes(i),ativos:acc})}
        return out;
      })(),
      expira_mes:Object.values(expMes).sort((a,b)=>a.mes<b.mes?-1:1),
      devolucoes_mes:Object.values(devMes).map(x=>({...x,valor:+x.valor.toFixed(2)})).sort((a,b)=>a.mes<b.mes?-1:1),
      gratuitos_mes:Object.values(gratMes).sort((a,b)=>a.mes<b.mes?-1:1),
      devolucoes:{compras:devN,valor:+devVal.toFixed(2),pessoas:Object.keys(devolK).length,so_devolvido:soDevolvido},
      cortesia:{pessoas:Object.keys(cortesiaK).length,sem_compra:soCortesia},
      seletivo:{pessoas:soSeletivo},
      // recompra ANO A ANO: dos que compraram naquele ano, quantos já eram clientes (compraram em algum ano anterior)
      recompra_ano:(()=>{
        const anos={};
        for(const k in alunos){const a=alunos[k],ini=a.pri.getFullYear();
          for(const y of Object.keys(a.anos)){const ano=+y;
            (anos[ano]=anos[ano]||{ano,compradores:0,recompradores:0,novos:0,receita:0,compras:0});
            anos[ano].compradores++;anos[ano].receita+=a.anos[y].receita;anos[ano].compras+=a.anos[y].compras;
            if(ini<ano)anos[ano].recompradores++;else anos[ano].novos++;}}
        return Object.values(anos).sort((a,b)=>a.ano-b.ano).map(x=>({...x,receita:+x.receita.toFixed(2),
          pct_recompra:+(x.compradores?x.recompradores*100/x.compradores:0).toFixed(1),
          ticket:+(x.compradores?x.receita/x.compradores:0).toFixed(2)}));
      })(),
      // safra: LTV e permanência de quem ENTROU em cada mês, medidos na vida inteira dele
      safra_mes:Object.values(safraMap).sort((a,b)=>a.mes<b.mes?-1:1).map(x=>({...x,receita:+x.receita.toFixed(2),perm_soma:+x.perm_soma.toFixed(2)})),
      serie_mensal:Object.values(mesMap).sort((a,b)=>a.mes<b.mes?-1:1).map(m=>({...m,receita:+m.receita.toFixed(2)})),
      serie_diaria:Object.values(diaMap).sort((a,b)=>a.dia<b.dia?-1:1).map(x=>({...x,receita:+x.receita.toFixed(2)})),
      /* produto SEM a turma no nome: é isso que responde "quanto esse produto vale".
         Guarda TODOS: cortar nos 30 maiores fazia produto novo (que ainda vende pouco) nunca aparecer
         — foi o que aconteceu com "Até Passar" e "História da Arte". O gráfico mostra os 15 primeiros,
         mas a lista e o seletor precisam do catálogo inteiro. */
      top_produtos:Object.values(prodMap).sort((a,b)=>b.receita-a.receita).map(p=>({...p,receita:+p.receita.toFixed(2)})),
      produtos_distintos:Object.keys(prodMap).length};
    // produto × turma só dos que aparecem no ranking — é o que sustenta o filtro de turma
    dados.por_produto_turma=Object.values(prodRefMap)
      .map(x=>({...x,receita:+x.receita.toFixed(2)})).sort((a,b)=>b.receita-a.receita);
    /* Cubo produto × estado × turma, só dos produtos do ranking. O nome do produto vira índice num
       dicionário porque repetir a string em milhares de linhas engordaria o retrato à toa. */
    const idx={},nomes=[];
    const linhasCubo=Object.values(cuboMap).map(x=>{
      if(idx[x.produto]==null){idx[x.produto]=nomes.length;nomes.push(x.produto)}
      return {p:idx[x.produto],uf:x.uf,ano:x.ano,mes:x.mes,v:x.vendas,r:+x.receita.toFixed(2)};
    });
    dados.cubo_produto={produtos:nomes,linhas:linhasCubo};
    /* Parceiros no retrato. TODO cupom leva os recortes de dia, produto e estado — o gráfico e os
       filtros precisam valer para qualquer um. Para caber, nome de cupom e de produto viram índice
       num dicionário e as chaves são curtas (c/d/m/v/r): a mesma string de 12 caracteres se repetiria
       dezenas de milhares de vezes. */
    const cupLista=Object.values(cupMap).map(c=>({
      cupom:c.cupom,vendas:c.vendas,receita:+c.receita.toFixed(2),recebido:+c.recebido.toFixed(2),
      aberto:+c.aberto.toFixed(2),alunos:Object.keys(c.alunos).length,
      cancel:c.cancel,cancel_val:+c.cancel_val.toFixed(2),
      estorno:c.estorno,estorno_val:+c.estorno_val.toFixed(2),
      aguard:c.aguard,aguard_val:+c.aguard_val.toFixed(2),naoautoriz:c.naoautoriz,
      pri:c.pri,ult:c.ult
    })).sort((a,b)=>b.receita-a.receita);
    dados.parceiros=cupLista;
    const cupNomes=[],cupIdx={},prodNomes=[],prodIdx={};
    const ci=c=>{if(cupIdx[c]==null){cupIdx[c]=cupNomes.length;cupNomes.push(c)}return cupIdx[c]};
    const pi=x=>{if(prodIdx[x]==null){prodIdx[x]=prodNomes.length;prodNomes.push(x)}return prodIdx[x]};
    dados.parceiro_dia=Object.values(cupDia).map(x=>({c:ci(x.cupom),d:x.dia,v:x.vendas,r:+x.receita.toFixed(2)}));
    dados.parceiro_mes=Object.values(cupMes).map(x=>({c:ci(x.cupom),m:x.mes,v:x.vendas,r:+x.receita.toFixed(2)}));
    dados.parceiro_produto=Object.values(cupProd).map(x=>({c:ci(x.cupom),p:pi(x.produto),v:x.vendas,r:+x.receita.toFixed(2)}));
    dados.parceiro_uf=Object.values(cupUf).map(x=>({c:ci(x.cupom),uf:x.uf,v:x.vendas,r:+x.receita.toFixed(2)}));
    dados.parceiros_idx=cupNomes;
    dados.parceiro_produtos_idx=prodNomes;
    // produto × mês: é o que faz o ranking de produtos respeitar o filtro de datas
    dados.por_produto_mes=Object.values(prodMesMap).map(x=>({...x,receita:+x.receita.toFixed(2)}));
    return dados;
  }
}

// deixa o agendador (Node) pegar as funções; na tela, 'module' não existe e esta linha é ignorada.
if (typeof module !== 'undefined' && module.exports) module.exports = { negocioProcessar, _negProdRef, _negUF, _negUFtel };

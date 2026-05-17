// data.js — Dados estáticos + parser dinâmico de Excel (SheetJS)

// ============================================================
// DADOS INICIAIS (planilha BASE PARA PPT_ATUAL.xlsx)
// ============================================================
let DADOS = {
  capacity: {
    operacoes: ['ALARES','EDUCACIONAL','M.DIAS BRANCO','AGORACRED'],
    meses: ['Jan','Fev','Mar','Abr'],
    valores: [[7,7,12,17],[5,5,1,1],[1,1,1,1],[3,3,3,3]],
    totais: [16,16,17,22]
  },
  funil: {
    abrilGeral: [
      {op:'AGORACRED',     disc:260960, alo:8346,  alo_pct:0.0320, cpc:1353, cpc_pct:0.1621, prom:274,  prom_pct:0.2025},
      {op:'ALARES VARIÁVEL',disc:1066345,alo:14016, alo_pct:0.0131, cpc:6144, cpc_pct:0.4384, prom:1257, prom_pct:0.2046},
      {op:'ALARES PA FIXA',disc:403852, alo:22191, alo_pct:0.0549, cpc:6166, cpc_pct:0.2779, prom:3040, prom_pct:0.4930},
      {op:'M.DIAS',        disc:118135, alo:1897,  alo_pct:0.0161, cpc:463,  cpc_pct:0.2441, prom:40,   prom_pct:0.0864},
    ],
    maio9Dias: [
      {op:'ALARES VARIÁVEL',disc:608040, alo:5429,  alo_pct:0.0089, cpc:2284, cpc_pct:0.4207, prom:484,  prom_pct:0.2119},
      {op:'ALARES PA FIXA',disc:201566, alo:15458, alo_pct:0.0767, cpc:4536, cpc_pct:0.2934, prom:2034, prom_pct:0.4484},
      {op:'M.DIAS',        disc:56121,  alo:963,   alo_pct:0.0172, cpc:298,  cpc_pct:0.3094, prom:23,   prom_pct:0.0772},
    ],
    abril9Dias: [
      {op:'ALARES VARIÁVEL',disc:627763, alo:5884,  alo_pct:0.0094, cpc:2494, cpc_pct:0.4239, prom:456,  prom_pct:0.1828},
      {op:'ALARES PA FIXA',disc:162912, alo:9385,  alo_pct:0.0576, cpc:2692, cpc_pct:0.2868, prom:1400, prom_pct:0.5201},
      {op:'M.DIAS',        disc:52888,  alo:916,   alo_pct:0.0173, cpc:261,  cpc_pct:0.2849, prom:15,   prom_pct:0.0575},
    ]
  },
  quartil: {
    tarde: {
      quartis:['1º Quartil','2º Quartil','3º Quartil','4º Quartil'],
      dispersao:[1.00,0.50,0.3031,0.17],
      mediapromessas:[494,318,193,109]
    },
    manha: {
      quartis:['1º Quartil','2º Quartil','3º Quartil','4º Quartil'],
      dispersao:[1.00,0.55,0.48,0.3325],
      mediapromessas:[471,283,246,35]
    }
  },
  acoesDigitais: {
    whatsapp: {
      operacoes:['ALARES','EQUATORIAL','M.DIAS','COPEL','ENEL SP','ENEL CE','ENEL RJ'],
      marco:[0,1,42,173,34,45,33], abril:[4,0,17,200,105,42,31], maio:[4,1,13,150,41,3,26]
    },
    chatbot: {
      operacoes:['ENEL SP','ENEL CE','ENEL RJ'],
      marco:[13,60,21], abril:[11,33,24], maio:[17,26,11]
    }
  },
  planoAcao: [
    {oportunidade:'Promessas',          motivo:'Feedbacks e acompanhamento',           responsavel:'Operação'},
    {oportunidade:'Nota Qualidade',      motivo:'Auditoria e feedbacks',                responsavel:'Qualidade'},
    {oportunidade:'Capacity',            motivo:'Operadores em curva de aprendizado',   responsavel:'Operação'},
    {oportunidade:'Divisão PA Fixa/Var.',motivo:'Quartil unificado',                    responsavel:'Operação / INT'},
  ],
  dadosGerais: {
    indicadores: ['ABS (meta 2%)','Qualidade (meta 95%)','Pausa (meta 16%)','Tempo Logado','Turnover','MC'],
    meses: ['Janeiro','Fevereiro','Março','Abril','Maio'],
    valores: [
      [0.0235,0.0128,3.92,5.12,0.0236],
      [99.28,98.67,99.04,0.93,0.9972],
      [0.13,0.14,0.14,0.15,0.15],
      [1,1,1,1,1],
      [0,0,0,0,0],
      [0.706,0.084,-0.43,null,null],
    ],
    metas: [0.02, 0.95, 0.16, null, null, null]
  },
  digital: [],
  mapaDigital: {
    canais: ['SMS','Email','WhatsApp','AGV','URA','Portal','Chatbot'],
    operacoes: [
      { op:'EDUCAÇÃO',    canais:[false,false,true, false,false,false,false] },
      { op:'M DIAS',      canais:[false,false,true, false,false,false,false] },
      { op:'AGORACRED',   canais:[true, false,false,false,false,false,false] },
      { op:'ALARES',      canais:[true, true, true, true, true, false,false] },
      { op:'TIM',         canais:[true, false,false,true, true, false,false] },
      { op:'COPEL',       canais:[true, true, true, true, true, false,false] },
      { op:'EQUATORIAL',  canais:[true, true, true, true, true, false,false] },
      { op:'SABESP',      canais:[true, true, false,false,false,true, false] },
      { op:'ENEL CE',     canais:[true, true, true, true, true, false,true ] },
      { op:'ENEL RJ',     canais:[true, true, true, true, true, false,true ] },
      { op:'ENEL SP',     canais:[true, true, true, true, true, false,true ] },
    ]
  },
  periodo: 'Abril / Maio 2026'
};

// ============================================================
// CONFIGURAÇÃO DO FIREBASE (Sincronização em Nuvem Real-Time)
// ============================================================
const firebaseConfig = {
  apiKey: "AIzaSyDfJc_gjYOBdnGRSw9q71TzBttXAi-07cw",
  authDomain: "rr2026-127c2.firebaseapp.com",
  projectId: "rr2026-127c2",
  storageBucket: "rr2026-127c2.firebasestorage.app",
  messagingSenderId: "306089663556",
  appId: "1:306089663556:web:57d224233e1e06ebe33011",
  databaseURL: "https://rr2026-127c2-default-rtdb.firebaseio.com"
};

// Inicializa o Firebase
if (!firebase.apps.length) {
  firebase.initializeApp(firebaseConfig);
}
const db = firebase.database();

// Tenta ler os dados do Firebase
db.ref('dashboard/DADOS').on('value', (snapshot) => {
  const cloudData = snapshot.val();
  if (cloudData) {
    DADOS = cloudData;
    // Se a função renderAll estiver disponível (definida no index.html), nós a chamamos para atualizar a tela
    if (typeof window.renderAll === 'function') {
      window.renderAll();
    }
  }
});

// ============================================================
// PARSER DE EXCEL (SheetJS) — atualiza DADOS dinamicamente
// ============================================================
async function lerExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, {type:'array'});
        const toArr = (nome) => {
          if (!wb.SheetNames.includes(nome)) return null;
          return XLSX.utils.sheet_to_json(wb.Sheets[nome], {header:1, defval:null});
        };

        // CAPACITY
        const cap = toArr('CAPACITY');
        if (cap) {
          const ops=[], meses=[], vals=[], totais=[];
          const header = cap[0];
          for(let i=1;i<header.length;i++) if(header[i]) meses.push(String(header[i]));
          for(let r=1;r<cap.length;r++){
            const row=cap[r]; if(!row[0]) continue;
            ops.push(String(row[0]));
            const rv=[]; let tot=0;
            for(let i=1;i<=meses.length;i++){ const v=Number(row[i]||0); rv.push(v); tot+=v; }
            vals.push(rv);
          }
          for(let i=0;i<meses.length;i++) totais.push(vals.reduce((s,r)=>s+(r[i]||0),0));
          if(ops.length) DADOS.capacity={operacoes:ops,meses,valores:vals,totais};
        }

        // FUNIL
        const fun = toArr('FUNIL');
        if (fun) {
          const parseFunil=(raw, startRow)=>{
            const result=[];
            for(let r=startRow;r<Math.min(startRow+8,raw.length);r++){
              const row=raw[r]; if(!row||!row[0]) continue;
              result.push({
                op:String(row[0]),
                disc:Number(row[1]||0), alo:Number(row[2]||0), alo_pct:Number(row[3]||0),
                cpc:Number(row[4]||0), cpc_pct:Number(row[5]||0),
                prom:Number(row[6]||0), prom_pct:Number(row[7]||0)
              });
            }
            return result;
          };
          // Detecta blocos pelo header
          let abrilRow=1, maio9Row=10, abril9Row=20;
          for(let r=0;r<fun.length;r++){
            const cell=String(fun[r]?.[0]||'').toLowerCase();
            if(cell.includes('geral')) abrilRow=r+1;
            if(cell.includes('maio')&&cell.includes('9')) maio9Row=r+1;
            if(cell.includes('abril')&&cell.includes('9')) abril9Row=r+1;
          }
          const ag=parseFunil(fun,abrilRow); if(ag.length) DADOS.funil.abrilGeral=ag;
          const m9=parseFunil(fun,maio9Row); if(m9.length) DADOS.funil.maio9Dias=m9;
          const a9=parseFunil(fun,abril9Row); if(a9.length) DADOS.funil.abril9Dias=a9;
        }

        // QUARTIL2026
        const qraw = toArr('QUARTIL2026');
        if (qraw) {
          const parseQ=(rows,startRow)=>{
            const q={quartis:[],dispersao:[],mediapromessas:[]};
            const hr=rows[startRow]||[];
            for(let c=1;c<hr.length;c++) if(hr[c]) q.quartis.push(String(hr[c]));
            for(let r=startRow+1;r<startRow+5&&r<rows.length;r++){
              const row=rows[r]; if(!row) continue;
              const lbl=String(row[0]||'').toLowerCase();
              if(lbl.includes('dispers')) for(let c=1;c<=q.quartis.length;c++) q.dispersao.push(Number(row[c]||0));
              if(lbl.includes('prom')) for(let c=1;c<=q.quartis.length;c++) q.mediapromessas.push(Number(row[c]||0));
            }
            return q;
          };
          let tardeRow=0,manhaRow=0;
          for(let r=0;r<qraw.length;r++){
            const c=String(qraw[r]?.[0]||'').toLowerCase();
            if(c.includes('tarde')) tardeRow=r;
            if(c.includes('manh')) manhaRow=r;
          }
          const qt=parseQ(qraw,tardeRow); if(qt.quartis.length) DADOS.quartil.tarde=qt;
          const qm=parseQ(qraw,manhaRow); if(qm.quartis.length) DADOS.quartil.manha=qm;
        }

        // PLANO DE AÇÃO
        const praw = toArr('PLANO DE AÇÃO');
        if(praw){
          const pl=[];
          for(let r=1;r<praw.length;r++){
            const row=praw[r]; if(!row||!row[0]) continue;
            pl.push({oportunidade:String(row[0]),motivo:String(row[1]||''),responsavel:String(row[2]||'')});
          }
          if(pl.length) DADOS.planoAcao=pl;
        }

        // AÇÕES DIGITAIS
        const adraw = toArr('AÇÕES DIGITAIS');
        if(adraw && adraw.length>2){
          const header=adraw[0]||[];
          const ops=[], mar=[], abr=[], mai=[];
          for(let r=1;r<adraw.length;r++){
            const row=adraw[r]; if(!row||!row[0]) continue;
            ops.push(String(row[0]));
            mar.push(Number(row[1]||0)); abr.push(Number(row[2]||0)); mai.push(Number(row[3]||0));
          }
          if(ops.length) DADOS.acoesDigitais.whatsapp={operacoes:ops,marco:mar,abril:abr,maio:mai};
        }

        // DADOS GERAIS
        const dgraw = toArr('DADOS GERAIS');
        if(dgraw && dgraw.length>1){
          const header=dgraw[0]||[];
          const meses=[]; for(let c=1;c<header.length;c++) if(header[c]) meses.push(String(header[c]));
          const inds=[], vals=[], metas=[];
          for(let r=1;r<dgraw.length;r++){
            const row=dgraw[r]; if(!row||!row[0]) continue;
            inds.push(String(row[0]));
            const rv=[]; for(let c=1;c<=meses.length;c++) rv.push(row[c]!==undefined&&row[c]!==null?Number(row[c]):null);
            vals.push(rv);
            // extrai meta do nome se houver (ex: "ABS- 2%" ou "meta 10")
            let metaVal = null;
            const text = String(row[0]);
            // Tenta achar "meta X%" ou "X%"
            let m = text.match(/(-?\d+(?:\.\d+)?)\s*%/);
            if (m) {
              metaVal = Number(m[1]) / 100;
            } else {
              // Tenta achar "meta X"
              m = text.match(/meta\s*(-?\d+(?:\.\d+)?)/i);
              if (m) metaVal = Number(m[1]);
            }
            metas.push(metaVal);
          }
          if(inds.length) DADOS.dadosGerais={indicadores:inds,meses,valores:vals,metas};
        }

        // DIGITAL
        const digiraw = toArr('DIGITAL');
        if(digiraw) {
          const rows=[]; 
          for(const r of digiraw) if(r&&r.some(v=>v!==null&&v!=='')) rows.push(r);
          DADOS.digital=rows;
        }

        // Salvar no Firebase para manter os dados no refresh e propagar para todos
        try {
          db.ref('dashboard/DADOS').set(DADOS);
          localStorage.setItem('MF_DADOS_2026', JSON.stringify(DADOS)); // fallback opcional
        } catch (e) {
          console.error("Erro ao salvar no Firebase/localStorage", e);
        }

        resolve(true);
      } catch(err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}

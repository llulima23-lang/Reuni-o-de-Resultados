// export-ppt.js — Geração do PPT visual idêntico ao sistema
// Depende de: ppt-helpers.js

async function exportarPPT() {
  const btn = document.getElementById('btn-ppt');
  btn.disabled = true;
  const pptx = new PptxGenJS(); pptx.layout='LAYOUT_16x9';
  const PAL = [G.dk, G.md, G.gr, G.lm, G.or];

  // ===== SLIDE 1 - CAPA =====
  btn.innerHTML='⏳ Capa...';
  {
    const s=pptx.addSlide();
    s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:'0D3321'}});
    s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.12,h:'100%',fill:{color:G.lm}});
    s.addShape(pptx.ShapeType.ellipse,{x:7.2,y:-0.8,w:3.2,h:3.2,fill:{color:G.md,transparency:85},line:{color:G.md,transparency:85}});
    s.addShape(pptx.ShapeType.ellipse,{x:8.3,y:3.0,w:2.1,h:2.1,fill:{color:G.lm,transparency:90},line:{color:G.lm,transparency:90}});
    if(typeof LOGO_B64 !== 'undefined') {
      s.addShape(pptx.ShapeType.roundRect,{x:0.45,y:0.45,w:2.6,h:0.8,rectRadius:0.1,fill:{color:G.wh}});
      s.addImage({data:LOGO_B64, x:0.65, y:0.55, w:2.2, h:0.6, sizing: {type: 'contain', w: 2.2, h: 0.6}});
    }
    s.addText('REUNIÃO DE RESULTADOS',{x:0.45,y:1.85,w:7,h:0.3,fontSize:10,color:G.lm,bold:true,fontFace:FONT,charSpacing:2});
    s.addText('MONITORAMENTO\nOPERACIONAL · MF 2026',{x:0.45,y:2.15,w:8,h:1.4,fontSize:38,bold:true,color:G.wh,fontFace:FONT,lineSpacing:40});
    s.addText('Período de Referência: ' + (DADOS.periodo || 'Abril / Maio 2026') + '  |  Operações Variadas',{x:0.45,y:3.8,w:8,h:0.3,fontSize:14,color:'90c8a0',fontFace:FONT});
  }

  // ===== SLIDE 2 - DADOS GERAIS =====
  if (DADOS.dadosGerais && DADOS.dadosGerais.indicadores) {
    btn.innerHTML='⏳ Dados Gerais...';
    const s=pptx.addSlide();
    hdr(pptx,s,'DADOS GERAIS — Indicadores Operacionais','Visão consolidada por mês');
    const d = DADOS.dadosGerais;
    const fmt = (v, ind) => {
      if (v === null || v === undefined) return '—';
      const str = String(ind).toLowerCase();
      if (str.includes('abs') || str.includes('pausa') || str.includes('mc') || str.includes('qual') || str.includes('turn') || str.includes('tempo')) {
        if (Math.abs(v) <= 2) return (v*100).toFixed(1).replace('.', ',') + '%';
        return Number(v).toFixed(1).replace('.', ',') + '%';
      }
      return v;
    };
    const head = [{text:'INDICADOR',options:{bold:true,fill:{color:G.dk},color:G.wh,fontSize:10,fontFace:FONT,align:'left'}}];
    d.meses.forEach(m => head.push({text:m,options:{bold:true,fill:{color:G.dk},color:G.wh,fontSize:10,fontFace:FONT,align:'center'}}));
    const rows = [head];
    d.indicadores.forEach((ind, i) => {
      const meta = d.metas?.[i];
      const row = [{text:ind,options:{bold:true,fill:{color:i%2===0?G.wh:'F4FBF6'},fontSize:10,fontFace:FONT}}];
      d.valores[i].forEach(v => {
        let clr = G.dk;
        if (v !== null && meta !== null && meta !== undefined) {
          const bad = (ind.toLowerCase().includes('abs') || ind.toLowerCase().includes('pausa') || ind.toLowerCase().includes('turn')) ? v > meta : v < meta;
          clr = bad ? G.re : G.gr;
        }
        row.push({text:String(fmt(v,ind)),options:{bold:(clr!==G.dk),color:clr,fill:{color:i%2===0?G.wh:'F4FBF6'},fontSize:11,fontFace:FONT,align:'center'}});
      });
      rows.push(row);
    });
    s.addTable(rows,{x:0.2,y:1.2,w:9.6,rowH:0.4,border:{type:'solid',pt:0.5,color:G.bd}});
    ins(pptx,s,'Verde = dentro da meta | Vermelho = fora da meta',4.8);
  }

  // ===== SLIDE 3 - CAPACITY (Tudo em 1) =====
  btn.innerHTML='⏳ Capacity...';
  {
    const d=DADOS.capacity; const s=pptx.addSlide();
    hdr(pptx,s,'CAPACITY — Gráficos e Tabelas','Evolução Mensal | Jan → Abr 2026');
    [['JAN','16',null,null],['FEV','16',null,null],['MAR','17','▲ +1','E8F5ED'],['ABR','22','▲ +5','E8F5ED']]
      .forEach(([l,v,b,bc],i)=> kpiCard(pptx,s,0.2+i*2.42,0.82,2.3,0.8,l,v,b,bc));
    const img1=await chartImg('bar',{labels:d.meses,datasets:d.operacoes.map((op,i)=>({label:op,data:d.valores[i],backgroundColor:PAL[i%PAL.length],borderRadius:5}))},{scales:SCALE});
    const img2=await chartImg('line',{labels:d.meses,datasets:[{label:'Total Agentes',data:d.totais,borderColor:G.dk,backgroundColor:'rgba(27,94,56,0.12)',fill:true,tension:0.4,pointRadius:6,pointBackgroundColor:G.dk,borderWidth:3}]},{scales:SCALE,plugins:{legend:{display:false}}});
    s.addImage({data:img1,x:0.2,y:1.75,w:4.8,h:1.9});
    s.addImage({data:img2,x:5.1,y:1.75,w:4.8,h:1.9});
    
    // Tabela Compacta em baixo
    const rows=[tH(['OPERAÇÃO','JAN','FEV','MAR','ABR','TOTAL']),...d.operacoes.map((op,i)=>tR([op,...d.valores[i],d.valores[i].reduce((a,b)=>a+b,0)],i%2===1))];
    rows.push([{text:'TOTAL GERAL',options:{bold:true,fill:{color:'C8EBD0'},fontSize:9,fontFace:FONT}},...d.totais.map(v=>({text:String(v),options:{bold:true,fill:{color:'C8EBD0'},fontSize:9,fontFace:FONT}})),{text:String(d.totais.reduce((a,b)=>a+b,0)),options:{bold:true,fill:{color:'C8EBD0'},fontSize:9,fontFace:FONT}}]);
    s.addTable(rows,{x:0.2,y:3.75,w:9.6,rowH:0.35,border:{type:'solid',pt:0.5,color:G.bd}});
  }

  // ===== SLIDES 4 a N - FUNIL =====
  for(const [tabKey, tabLabel] of [['abrilGeral','Abril — Geral'],['maio9Dias','Maio — 9 Dias Úteis'],['abril9Dias','Abril — 9 Dias Úteis']]) {
    const fd=DADOS.funil[tabKey]; 
    if(!fd || !fd.length) continue;
    
    btn.innerHTML='⏳ Funil ' + tabLabel + '...';
    const s1=pptx.addSlide();
    hdr(pptx,s1,'FUNIL DE CONVERSÃO — ' + tabLabel,'Visão Gráfica e Tabela Detalhada');
    
    // Imagens dos Funis (Top)
    const funnels=await Promise.all(fd.map(r=>funnelImg(r,r.op,260,260)));
    const startX = (10 - (funnels.length * 2.2)) / 2; // Center horizontally
    funnels.forEach((img,i)=>{
      s1.addImage({data:img,x:startX + (i*2.2),y:0.8,w:2.0,h:1.8});
    });

    // Tabela (Bottom)
    const rows=[tH(['OPERAÇÃO','DISC.','ALÔ','ALÔ%','CPC','CPC/ALÔ%','PROM.','PROM/CPC%']),...fd.map((r,i)=>tR([r.op,r.disc.toLocaleString('pt-BR'),r.alo.toLocaleString('pt-BR'),(r.alo_pct*100).toFixed(1)+'%',r.cpc.toLocaleString('pt-BR'),(r.cpc_pct*100).toFixed(1)+'%',r.prom.toLocaleString('pt-BR'),(r.prom_pct*100).toFixed(1)+'%'],i%2===1))];
    s1.addTable(rows,{x:0.2,y:2.8,w:9.6,rowH:0.35,border:{type:'solid',pt:0.5,color:G.bd}});
    
    const insights={abrilGeral:'ALARES PA FIXA lidera: 49,3% de conversão CPC→Promessa. AGORACRED concentra maior volume de discagens.',abril9Dias:'Nos 9 dias úteis de Abril, ALARES PA FIXA atingiu 52% Promessa/CPC.',maio9Dias:'ALARES PA FIXA em 9 dias de Maio já supera Abril no mesmo período (+45%).'};
    ins(pptx,s1,insights[tabKey] || 'Evolução consistente nas taxas de conversão operacionais.',4.9);
  }

  // ===== SLIDE - QUARTIL =====
  btn.innerHTML='⏳ Quartil...';
  {
    const s=pptx.addSlide();
    hdr(pptx,s,'QUARTIL 2026 — Desempenho por Turno','Dispersão e Média de Promessas | Manhã e Tarde');
    const qt=DADOS.quartil.tarde, qm=DADOS.quartil.manha;
    
    quartilCard(pptx,s,0.2,0.82,4.8,'🕓 Turno TARDE','tarde',qt.quartis,qt.mediapromessas,qt.dispersao,'tarde');
    quartilCard(pptx,s,5.1,0.82,4.8,'🌅 Turno MANHÃ','manha',qm.quartis,qm.mediapromessas,qm.dispersao,'manha');
    
    // Adicionar Estratégias do 4º Quartil na parte inferior em vez de gráficos (igual no sistema)
    s.addShape(pptx.ShapeType.rect,{x:0.2,y:3.5,w:9.6,h:0.3,fill:{color:G.dk}});
    s.addText('🎯 ESTRATÉGIAS — EQUIPE 4º QUARTIL MANHÃ (Queda de 246 para 35 promessas)',{x:0.2,y:3.5,w:9.6,h:0.3,fontSize:9,bold:true,color:G.wh,fontFace:FONT,align:'center',valign:'middle'});
    
    [['🎧 Escuta & Calibração','Revisar gravações 4º Q vs 1º Q. Feedback individual semanal.'],
     ['🤝 Mentoring Cruzado','Parear cada op. do 4º Q com top performer. 2 sessões/semana.'],
     ['📈 Micro-metas Diárias','Meta escalonada: 3→5→8 promessas/dia. Celebrar cada avanço.'],
     ['🎭 Role-Play Diário','15 min de simulação antes do turno. Treinar objeções.']
    ].forEach(([t,d],i)=>{
      const x=0.2+i*2.42;
      s.addShape(pptx.ShapeType.roundRect,{x,y:3.85,w:2.34,h:1.0,rectRadius:0.08,fill:{color:G.bg},line:{color:G.gr,pt:1}});
      s.addShape(pptx.ShapeType.rect,{x,y:3.85,w:2.34,h:0.25,fill:{color:G.md}});
      s.addText(t,{x:x+0.05,y:3.85,w:2.24,h:0.25,fontSize:8,bold:true,color:G.wh,fontFace:FONT,valign:'middle'});
      s.addText(d,{x:x+0.07,y:4.15,w:2.2,h:0.65,fontSize:8,color:G.dk,fontFace:FONT,wrap:true,valign:'top'});
    });
  }

  // ===== SLIDE - AÇÕES DIGITAIS (Gráficos) =====
  btn.innerHTML='⏳ Ações Digitais 1...';
  {
    const wp=DADOS.acoesDigitais.whatsapp, cb=DADOS.acoesDigitais.chatbot;
    const s=pptx.addSlide();
    hdr(pptx,s,'AÇÕES DIGITAIS — WhatsApp & Chatbot','Evolução Gráfica | Março → Maio 2026');
    kpiCard(pptx,s,0.2,0.82,4.75,0.9,'COPEL — WhatsApp Abril','200','▲ Maior vol.','E8F5ED');
    kpiCard(pptx,s,5.1,0.82,4.75,0.9,'ENEL CE — Chatbot Abril','33','▼ Queda','FDECEA');
    const img1=await chartImg('bar',{labels:wp.operacoes,datasets:[{label:'Março',data:wp.marco,backgroundColor:'rgba(141,198,66,0.45)',borderRadius:4},{label:'Abril',data:wp.abril,backgroundColor:G.gr,borderRadius:4},{label:'Maio',data:wp.maio,backgroundColor:G.dk,borderRadius:4}]},{scales:SCALE});
    const img2=await chartImg('line',{labels:['Março','Abril','Maio'],datasets:cb.operacoes.map((op,i)=>({label:op,data:[cb.marco[i],cb.abril[i],cb.maio[i]],borderColor:PAL[i%PAL.length],backgroundColor:'transparent',pointRadius:7,pointBackgroundColor:PAL[i%PAL.length],borderWidth:3,tension:0.4}))},{scales:SCALE});
    s.addImage({data:img1,x:0.2,y:1.9,w:4.8,h:2.4});
    s.addImage({data:img2,x:5.1,y:1.9,w:4.8,h:2.4});
    ins(pptx,s,'COPEL lidera WhatsApp em Abril (200). ENEL CE queda Chatbot 60→26. ALARES/EQUATORIAL baixo uso digital — oportunidade.',4.7);
  }

  // ===== SLIDE - AÇÕES DIGITAIS (Tabelas) =====
  btn.innerHTML='⏳ Ações Digitais 2...';
  {
    const wp=DADOS.acoesDigitais.whatsapp, cb=DADOS.acoesDigitais.chatbot;
    const s=pptx.addSlide();
    hdr(pptx,s,'AÇÕES DIGITAIS — Tabelas Detalhadas','Envios por Operação | Março → Maio 2026');
    const wR=[tH(['OPERAÇÃO','MAR','ABR','MAI']),...wp.operacoes.map((op,i)=>tR([op,wp.marco[i],wp.abril[i],wp.maio[i]],i%2===1))];
    const cR=[tH(['OPERAÇÃO','MAR','ABR','MAI']),...cb.operacoes.map((op,i)=>tR([op,cb.marco[i],cb.abril[i],cb.maio[i]],i%2===1))];
    s.addText('WhatsApp',{x:0.2,y:0.9,w:4.6,h:0.3,fontSize:11,bold:true,color:G.dk,fontFace:FONT});
    s.addTable(wR,{x:0.2,y:1.3,w:4.6,border:{type:'solid',pt:0.5,color:G.bd}});
    s.addText('Chatbot',{x:5.2,y:0.9,w:4.6,h:0.3,fontSize:11,bold:true,color:G.dk,fontFace:FONT});
    s.addTable(cR,{x:5.2,y:1.3,w:4.6,border:{type:'solid',pt:0.5,color:G.bd}});
  }

  // ===== SLIDE - MAPA DIGITAL =====
  if (DADOS.mapaDigital) {
    btn.innerHTML='⏳ Mapa Digital...';
    const s=pptx.addSlide();
    hdr(pptx,s,'MAPA DE AÇÕES DIGITAIS','Canais digitais ativos por operação');
    const d=DADOS.mapaDigital;
    
    // Cards de Canais
    const cores = {SMS:'1B5E38',Email:'2E7D4F',WhatsApp:'25D366',AGV:'5FAD41',URA:'8DC642',Portal:'4BACC6',Chatbot:'F79646'};
    const totC = d.canais.map((_,ci)=>d.operacoes.filter(o=>o.canais[ci]).length);
    d.canais.forEach((c,i)=>{
      const pct=Math.round(totC[i]/d.operacoes.length*100);
      const x=0.15+i*1.38;
      s.addShape(pptx.ShapeType.roundRect,{x,y:0.8,w:1.3,h:1.2,rectRadius:0.08,fill:{color:G.wh},line:{color:G.bd,pt:1}});
      s.addShape(pptx.ShapeType.rect,{x,y:0.8,w:1.3,h:0.1,fill:{color:cores[c]||G.dk}});
      s.addText(c,{x,y:0.95,w:1.3,h:0.2,fontSize:8,bold:true,color:G.dk,align:'center'});
      s.addText(String(totC[i]),{x,y:1.15,w:1.3,h:0.4,fontSize:22,bold:true,color:cores[c]||G.dk,align:'center'});
      s.addText(pct+'% das ops',{x,y:1.6,w:1.3,h:0.2,fontSize:7,color:G.md,align:'center'});
    });

    const head = [{text:'OPERAÇÃO',options:{bold:true,fill:{color:G.dk},color:G.wh,fontSize:9,fontFace:FONT,align:'left'}}];
    d.canais.forEach(c => head.push({text:c,options:{bold:true,fill:{color:G.dk},color:G.wh,fontSize:8,fontFace:FONT,align:'center'}}));
    head.push({text:'TOTAL',options:{bold:true,fill:{color:G.dk},color:G.wh,fontSize:9,fontFace:FONT,align:'center'}});
    
    const rows = [head];
    d.operacoes.forEach((op, oi) => {
      const tot = op.canais.filter(Boolean).length;
      const row = [{text:op.op,options:{bold:true,fill:{color:oi%2===0?G.wh:'F4FBF6'},fontSize:9,fontFace:FONT,color:G.dk}}];
      op.canais.forEach(ativo => {
        row.push({text:ativo?'✓':'',options:{bold:true,fill:{color:oi%2===0?G.wh:'F4FBF6'},color:ativo?G.gr:'E0E0E0',fontSize:12,align:'center'}});
      });
      row.push({text:tot + '/' + d.canais.length,options:{bold:true,fill:{color:oi%2===0?G.wh:'F4FBF6'},fontSize:9,fontFace:FONT,align:'center',color:G.dk}});
      rows.push(row);
    });
    const tRow = [{text:'TOTAL CANAL',options:{bold:true,fill:{color:'C8EBD0'},fontSize:9,fontFace:FONT,color:G.dk}}];
    totC.forEach(t=> tRow.push({text:String(t),options:{bold:true,fill:{color:'C8EBD0'},fontSize:9,fontFace:FONT,color:G.dk,align:'center'}}));
    tRow.push({text:String(d.operacoes.length),options:{bold:true,fill:{color:'C8EBD0'},fontSize:9,fontFace:FONT,color:G.dk,align:'center'}});
    rows.push(tRow);
    s.addTable(rows,{x:0.2,y:2.2,w:9.6,rowH:0.28,border:{type:'solid',pt:0.5,color:G.bd}});
  }

  // ===== SLIDE - ENCERRAMENTO (Frase Corporativa) =====
  btn.innerHTML='⏳ Encerramento...';
  {
    const s=pptx.addSlide();
    s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:'0D3321'}});
    s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.12,h:'100%',fill:{color:G.lm}});
    s.addShape(pptx.ShapeType.ellipse,{x:7.2,y:-0.8,w:3.2,h:3.2,fill:{color:G.md,transparency:85},line:{color:G.md,transparency:85}});
    
    // Frase inspiracional
    s.addText('“Se você criar uma ótima experiência, os clientes \ncontam uns aos outros sobre isso. \nA palavra falada é muito poderosa.”',{x:1.0,y:2.0,w:8,h:1.5,fontSize:24,italic:true,bold:true,color:G.wh,fontFace:FONT,align:'center',lineSpacing:35});
    s.addText('— Jeff Bezos',{x:1.0,y:3.8,w:8,h:0.5,fontSize:18,bold:true,color:G.lm,fontFace:FONT,align:'center'});
  }

  btn.innerHTML='⏳ Salvando...';
  await pptx.writeFile({fileName:'MF_Resultados_Operacional_2026.pptx'});
  btn.disabled=false; btn.innerHTML='📥 Exportar PowerPoint (Nativo)';
}

async function exportarPPTClone() {
  const btn = document.getElementById('btn-ppt-clone');
  btn.disabled = true;
  btn.innerHTML = '📸 Fotografando...';
  
  const pptx = new PptxGenJS(); pptx.layout='LAYOUT_16x9';
  
  const sections = [
    { id: 'sec-capa' },
    { id: 'sec-dadosgerais' },
    { id: 'sec-capacity' },
    { id: 'sec-funil' },
    { id: 'sec-quartil' },
    { id: 'sec-digital' },
    { id: 'sec-mapadigital' }
  ];

  // Desabilita animações para evitar captura semi-transparente
  const style = document.createElement('style');
  style.innerHTML = '.section { animation: none !important; transition: none !important; opacity: 1 !important; }';
  document.head.appendChild(style);

  const activeNav = document.querySelector('.nav-item.active').getAttribute('data-sec');

  for(const sec of sections) {
    const el = document.getElementById(sec.id);
    if (!el) continue;
    
    // Altera classe active em vez de display: none
    document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
    el.classList.add('active');
    
    // Força largura fixa para que a proporção não mude com a janela do usuário
    const oldWidth = el.style.width;
    el.style.width = '1440px';
    
    // Aguarda um pequeno delay para garantir o paint completo sem animação
    await new Promise(r => setTimeout(r, 50));

    try {
      // Ajuste para não perder cores: usa o background padrao e log scale melhor
      const canvas = await html2canvas(el, { scale: 3, useCORS: true, backgroundColor: '#f5f7fa', logging: false });
      const imgData = canvas.toDataURL('image/jpeg', 0.9);
      
      const slide = pptx.addSlide();
      slide.addShape(pptx.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:'f5f7fa'}});
      slide.addImage({ data: imgData, x: 0.1, y: 0.1, w: 9.8, h: 5.425, sizing: { type: 'contain', w: 9.8, h: 5.425 } });
    } catch(err) {
      console.error(err);
    }
    
    el.style.width = oldWidth;
  }

  // ===== SLIDE - ENCERRAMENTO (Frase Corporativa) =====
  btn.innerHTML='⏳ Encerramento...';
  {
    const s=pptx.addSlide();
    s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:'0D3321'}});
    s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.12,h:'100%',fill:{color:G.lm}});
    s.addShape(pptx.ShapeType.ellipse,{x:7.2,y:-0.8,w:3.2,h:3.2,fill:{color:G.md,transparency:85},line:{color:G.md,transparency:85}});
    s.addText('“Se você criar uma ótima experiência, os clientes \ncontam uns aos outros sobre isso. \nA palavra falada é muito poderosa.”',{x:1.0,y:2.0,w:8,h:1.5,fontSize:24,italic:true,bold:true,color:G.wh,fontFace:FONT,align:'center',lineSpacing:35});
    s.addText('— Jeff Bezos',{x:1.0,y:3.8,w:8,h:0.5,fontSize:18,bold:true,color:G.lm,fontFace:FONT,align:'center'});
  }

  // Restore navigation and animations
  document.head.removeChild(style);
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  const activeEl = document.getElementById('sec-' + activeNav);
  if(activeEl) activeEl.classList.add('active');

  btn.innerHTML='⏳ Salvando...';
  await pptx.writeFile({fileName:'MF_Resultados_CloneExato.pptx'});
  btn.disabled = false;
  btn.innerHTML = '📸 Exportar PowerPoint (Idêntico)';
}

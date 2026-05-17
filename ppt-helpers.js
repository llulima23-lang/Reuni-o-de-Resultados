// ppt-helpers.js - funções auxiliares para geração do PPT

const G = {
  dk:'1B5E38', md:'2E7D4F', gr:'5FAD41', lm:'8DC642',
  re:'C0504D', or:'F79646', wh:'FFFFFF', bg:'E8F5ED', bd:'C5E0CC'
};

const FONT = 'Arial';
const SCALE = { x:{grid:{display:false}}, y:{grid:{color:'#e8eef5'}} };

// Cria imagem de gráfico em canvas temporário
async function chartImg(type, data, opts, W=900, H=420) {
  const c = document.createElement('canvas');
  c.width=W; c.height=H; c.style.cssText='position:fixed;left:-9999px;';
  document.body.appendChild(c);
  const ch = new Chart(c.getContext('2d'), {
    type, data,
    options: Object.assign({ animation:false, responsive:false,
      plugins:{ legend:{ position:'top', labels:{ font:{size:13}, boxWidth:14 } } }
    }, opts)
  });
  await new Promise(r=>setTimeout(r,400));
  const img = c.toDataURL('image/png');
  ch.destroy(); c.remove();
  return img;
}

// Desenha funil de uma operação
async function funnelImg(row, title, W=340, H=320) {
  const c = document.createElement('canvas');
  c.width=W; c.height=H; c.style.cssText='position:fixed;left:-9999px;';
  document.body.appendChild(c);
  const ctx = c.getContext('2d');
  ctx.fillStyle='#FFFFFF'; ctx.fillRect(0,0,W,H);

  // título
  ctx.fillStyle='#1B5E38'; ctx.font='bold 12px Arial';
  ctx.textAlign='center'; ctx.fillText(title, W/2-30, 18);

  const lvs = [
    {lbl:'DISCAGENS', val:row.disc, pctLbl:'100%',          clr:'#8DC642', w:0.92},
    {lbl:'ALÔ',       val:row.alo,  pctLbl:(row.alo_pct*100).toFixed(1)+'%', clr:'#5FAD41', w:0.72},
    {lbl:'CPC',       val:row.cpc,  pctLbl:(row.cpc_pct*100).toFixed(1)+'%', clr:'#2E7D4F', w:0.52},
    {lbl:'PROMESSA',  val:row.prom, pctLbl:(row.prom_pct*100).toFixed(1)+'%', clr:'#1B5E38', w:0.34},
  ];
  const fW=210, fX=20, sY=26, lH=(H-sY-10)/4;
  lvs.forEach((lv,i)=>{
    const wT=fW*lv.w, wB=fW*(lvs[i+1]?.w||lv.w*0.85);
    const xT=fX+(fW-wT)/2, xB=fX+(fW-wB)/2, y=sY+i*lH;
    ctx.beginPath();
    ctx.moveTo(xT,y+2); ctx.lineTo(xT+wT,y+2);
    ctx.lineTo(xB+wB,y+lH-3); ctx.lineTo(xB,y+lH-3);
    ctx.closePath(); ctx.fillStyle=lv.clr; ctx.fill();
    // valor
    ctx.fillStyle='#fff'; ctx.font='bold 11px Arial'; ctx.textAlign='center';
    ctx.fillText(lv.val.toLocaleString('pt-BR'), fX+fW/2, y+lH/2+4);
    // linha + label direita
    const lx=xT+wT+6, ly=y+lH/2;
    ctx.strokeStyle='#2E7D4F'; ctx.lineWidth=1;
    ctx.beginPath(); ctx.moveTo(xT+wT,ly); ctx.lineTo(lx+5,ly); ctx.stroke();
    ctx.fillStyle='#1B5E38'; ctx.font='bold 9px Arial'; ctx.textAlign='left';
    ctx.fillText(lv.lbl, lx+7, ly-2);
    ctx.fillStyle='#5A7A65'; ctx.font='9px Arial';
    ctx.fillText(lv.pctLbl, lx+7, ly+9);
  });
  const img=c.toDataURL('image/png'); c.remove();
  return img;
}

// Cabeçalho do slide
function hdr(pptx, s, title, sub) {
  s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:'100%',h:0.75,fill:{color:G.dk}});
  s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.06,h:'100%',fill:{color:G.lm}});
  s.addText(title,{x:0.25,y:0.05,w:7.5,h:0.42,fontSize:17,bold:true,color:G.wh,fontFace:FONT});
  if(sub) s.addText(sub,{x:0.25,y:0.47,w:7.5,h:0.22,fontSize:9,color:'90c8a0',fontFace:FONT});
  if(typeof LOGO_B64 !== 'undefined') {
    s.addImage({data:LOGO_B64, x:8.2, y:0.18, w:1.5, h:0.35});
  }
}

// Caixa de insight
function ins(pptx, s, txt, y) {
  s.addShape(pptx.ShapeType.roundRect,{x:0.2,y,w:9.6,h:0.6,rectRadius:0.07,fill:{color:G.bg},line:{color:G.gr,pt:1.5}});
  s.addText('💡 '+txt,{x:0.35,y:y+0.06,w:9.3,h:0.48,fontSize:8.5,color:G.dk,fontFace:FONT,wrap:true,valign:'top'});
}

// Linha de cabeçalho de tabela
function tH(cols) {
  return cols.map(c=>({text:c,options:{bold:true,fill:{color:G.dk},color:G.wh,fontSize:8.5,fontFace:FONT}}));
}
// Linha de dados de tabela
function tR(vals, even) {
  return vals.map(v=>({text:String(v??'-'),options:{fill:{color:even?'F4FBF6':G.wh},fontSize:8.5,fontFace:FONT}}));
}

// Card KPI (retângulo com label + valor + badge)
function kpiCard(pptx, s, x, y, w, h, lbl, val, badge, badgeClr) {
  s.addShape(pptx.ShapeType.roundRect,{x,y,w,h,rectRadius:0.1,fill:{color:G.wh},line:{color:G.bd,pt:1}});
  s.addText(lbl, {x:x+0.1,y:y+0.08,w:w-0.2,h:0.2,fontSize:7.5,color:'5A7A65',fontFace:FONT,bold:false});
  s.addText(val, {x:x+0.1,y:y+0.26,w:w-0.2,h:0.38,fontSize:22,bold:true,color:G.dk,fontFace:FONT});
  if(badge){
    s.addShape(pptx.ShapeType.roundRect,{x:x+0.1,y:y+h-0.28,w:0.55,h:0.2,rectRadius:0.05,fill:{color:badgeClr||'E8F5ED'}});
    s.addText(badge,{x:x+0.1,y:y+h-0.28,w:0.55,h:0.2,fontSize:7,bold:true,color:badgeClr==='FDECEA'?G.re:G.dk,fontFace:FONT,align:'center',valign:'middle'});
  }
}

// Card de Quartil
function quartilCard(pptx, s, x, y, w, title, turno, quartis, mediap, disp, cor) {
  s.addShape(pptx.ShapeType.roundRect,{x,y,w,h:2.6,rectRadius:0.1,fill:{color:G.wh},line:{color:G.bd,pt:1}});
  s.addText(title,{x:x+0.12,y:y+0.1,w:w-0.3,h:0.28,fontSize:10,bold:true,color:G.dk,fontFace:FONT});
  s.addShape(pptx.ShapeType.roundRect,{x:x+0.12+title.length*0.065+0.1,y:y+0.12,w:0.6,h:0.22,rectRadius:0.08,fill:{color:cor==='tarde'?'FFF3E0':'E3F5E8'}});
  s.addText(cor==='tarde'?'Tarde':'Manhã',{x:x+0.12+title.length*0.065+0.1,y:y+0.12,w:0.6,h:0.22,fontSize:7,bold:true,color:cor==='tarde'?'b35c00':'1a6632',fontFace:FONT,align:'center',valign:'middle'});
  // 4 quartis
  quartis.forEach((q,i)=>{
    const qx=x+0.12+i*(w-0.24)/4, qy=y+0.5;
    s.addText(q,{x:qx,y:qy,w:(w-0.24)/4,h:0.18,fontSize:7,color:'5A7A65',fontFace:FONT,align:'center',bold:true});
    s.addText(String(mediap[i]),{x:qx,y:qy+0.18,w:(w-0.24)/4,h:0.4,fontSize:20,bold:true,color:G.dk,fontFace:FONT,align:'center'});
    s.addText('Méd. Prom.',{x:qx,y:qy+0.58,w:(w-0.24)/4,h:0.16,fontSize:6.5,color:'5A7A65',fontFace:FONT,align:'center'});
  });
  s.addShape(pptx.ShapeType.line,{x:x+0.12,y:y+1.35,w:w-0.24,h:0,line:{color:G.bd,pt:0.5}});
  // dispersão
  s.addText('DISPERSÃO',{x:x+0.12,y:y+1.45,w:w-0.24,h:0.18,fontSize:7.5,bold:true,color:'5A7A65',fontFace:FONT,align:'center'});
  quartis.forEach((q,i)=>{
    const qx=x+0.12+i*(w-0.24)/4;
    s.addText((disp[i]*100).toFixed(0)+'%',{x:qx,y:y+1.63,w:(w-0.24)/4,h:0.32,fontSize:16,bold:true,color:cor==='tarde'?G.or:G.gr,fontFace:FONT,align:'center'});
  });
}

/* Testes do confronto apontado × pago, rodando o app de verdade no Chromium.
   Uso:  node testes/confronto.test.js
   Precisa do playwright (npm i -g playwright, ou o que já vier na máquina). */
const path=require("path"), fs=require("fs"), os=require("os");
let pw; try{ pw=require("playwright") }catch(e){
  pw=require(require("child_process").execSync("npm root -g").toString().trim()+"/playwright") }

/* ---------- .xlsx mínimo (zip sem compressão) ---------- */
const CRC=(()=>{const t=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xedb88320^(c>>>1):c>>>1;t[n]=c>>>0}return t})();
const crc32=b=>{let c=~0;for(const x of b)c=CRC[(c^x)&255]^(c>>>8);return(~c)>>>0};
function zip(arqs){
  const partes=[],central=[];let off=0;
  for(const [nome,txt] of Object.entries(arqs)){
    const n=Buffer.from(nome),d=Buffer.from(txt,"utf8"),c=crc32(d);
    const h=Buffer.alloc(30);h.writeUInt32LE(0x04034b50,0);h.writeUInt16LE(20,4);h.writeUInt32LE(c,14);
    h.writeUInt32LE(d.length,18);h.writeUInt32LE(d.length,22);h.writeUInt16LE(n.length,26);
    partes.push(h,n,d);
    const e=Buffer.alloc(46);e.writeUInt32LE(0x02014b50,0);e.writeUInt16LE(20,4);e.writeUInt16LE(20,6);
    e.writeUInt32LE(c,16);e.writeUInt32LE(d.length,20);e.writeUInt32LE(d.length,24);e.writeUInt16LE(n.length,28);
    e.writeUInt32LE(off,42);central.push(e,n);off+=30+n.length+d.length;
  }
  const cd=Buffer.concat(central),fim=Buffer.alloc(22);
  fim.writeUInt32LE(0x06054b50,0);fim.writeUInt16LE(Object.keys(arqs).length,8);fim.writeUInt16LE(Object.keys(arqs).length,10);
  fim.writeUInt32LE(cd.length,12);fim.writeUInt32LE(off,16);
  return Buffer.concat([...partes,cd,fim]);
}
/* fechamento simplificado: OBRA Nº / LOCAL, cabeçalho, INSTALAR, RETIRAR.
   comValor: "formula" (total =E*F, como o Excel grava), "nada" (só quantidade) */
/* opc.total: põe no rodapé "R$ | total" (como o fechamento real, ao lado de ENGENHARIA)
   opc.outraAba: itens de uma aba que fica em sheet1.xml mas vem DEPOIS no Excel */
function fechamento(obra,itens,comValor,opc={}){
  const S=[],si=t=>{let i=S.indexOf(t);if(i<0){S.push(t);i=S.length-1}return i};
  const aba=its=>{
  const s=(r,t)=>`<c r="${r}" t="s"><v>${si(t)}</v></c>`, n=(r,v)=>`<c r="${r}"><v>${v}</v></c>`;
  const rows=[s("A1","FECHAMENTO SIMPLIFICADO"),s("A2","OBRA Nº:")+'<c r="B2" s="1"/>'+s("C2",obra),
    s("A3","LOCAL:")+s("C3","RUA VITORIA · SAO BENTO"),
    s("B4","CÓDIGO")+s("D4","DESCRIÇÃO")+s("E4","QUANTIDADE")+s("F4","VALOR UNIT")+s("G4","VALOR TOTAL")];
  for(const op of ["I","R"]){
    rows.push(s("A"+(rows.length+1),op==="I"?"INSTALAR":"RETIRAR"));
    for(const it of its.filter(x=>x.op===op)){
      const r=rows.length+1;
      let c=n("B"+r,it.cod)+s("D"+r,it.desc)+n("E"+r,it.qtd);
      if(comValor==="formula") c+=n("F"+r,it.unit)+`<c r="G${r}" s="2"><f>E${r}*F${r}</f><v>${it.qtd*it.unit}</v></c>`;
      rows.push(c);
    }
  }
  if(opc.total){ const r=rows.length+2, t=its.reduce((a,x)=>a+x.qtd*x.unit,0);
    rows.push(""); rows.push(s("B"+r,"ENGENHARIA")); rows.push(s("A"+(r+1),"R$")+`<c r="B${r+1}" s="3"><f>SUM(G6:G${r-2})</f><v>${t}</v></c>`); }
  return `<?xml version="1.0"?><worksheet><sheetData>${rows.map((c,i)=>c?`<row r="${i+1}">${c}</row>`:"").join("")}</sheetData></worksheet>`;
  };
  const arqs={};
  if(opc.outraAba){
    arqs["xl/worksheets/sheet1.xml"]=aba(opc.outraAba); arqs["xl/worksheets/sheet2.xml"]=aba(itens);
    arqs["xl/workbook.xml"]=`<?xml version="1.0"?><workbook xmlns:r="r"><sheets><sheet name="FECHAMENTO" sheetId="2" r:id="rId2"/><sheet name="RASCUNHO ANTIGO" sheetId="1" r:id="rId1"/></sheets></workbook>`;
    arqs["xl/_rels/workbook.xml.rels"]=`<?xml version="1.0"?><Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>`;
  } else arqs["xl/worksheets/sheet1.xml"]=aba(itens);
  arqs["xl/sharedStrings.xml"]=`<?xml version="1.0"?><sst>${S.map(t=>`<si><t>${t}</t></si>`).join("")}</sst>`;
  return zip(arqs);
}

/* ---------- fórmula do app de obras pagas (calcularConfronto), reescrita à parte ---------- */
function referencia(apont,pagos,outrasMedicoes){
  const pago=new Map(),unit=new Map(),ap=new Map();
  for(const r of [...outrasMedicoes,...pagos]) if(r.qtd) unit.set(r.op+"|"+r.cod,r.vlr/r.qtd);
  for(const r of pagos){const k=r.op+"|"+r.cod,a=pago.get(k)||{q:0,v:0};a.q+=r.qtd;a.v+=r.vlr;pago.set(k,a)}
  for(const r of apont){const k=r.op+"|"+r.cod,a=ap.get(k)||{q:0,v:0};a.q+=r.qtd;a.v+=r.qtd*(r.unit||0);ap.set(k,a)}
  const out={};
  for(const k of new Set([...pago.keys(),...ap.keys()])){
    const A=ap.get(k)||{q:0,v:0},P=pago.get(k)||{q:0,v:0};
    let sit;
    if(A.q>0&&P.q===0) sit="Pendente de pagamento";
    else if(A.q>P.q) sit="Pagamento parcial";
    else if(P.q>A.q&&A.q>0) sit="Pago a mais";
    else if(A.q===0&&P.q>0) sit="Pago sem apontamento";
    else sit="Conferido";
    const dif=Math.round((A.q-P.q)*100)/100;
    const u=unit.get(k)??(A.q?A.v/A.q:0);
    out[k]={dif,vdif:Math.round(dif*u*100)/100,sit};
  }
  const sum=f=>Object.values(out).filter(f).reduce((a,x)=>a+Math.abs(x.vdif),0);
  const pend=sum(x=>x.sit==="Pendente de pagamento"||x.sit==="Pagamento parcial");
  const maior=sum(x=>x.sit==="Pago a mais"), sem=sum(x=>x.sit==="Pago sem apontamento");
  return {linhas:out,pend,maior,sem,cobranca:Math.round((pend-maior-sem)*100)/100};
}

/* ---------- dados de teste (códigos e descrições do print da obra 020-26-01062) ---------- */
const OBRA="020-26-01062";
const APONT=[
  {op:"I",cod:"66011",desc:"CORTE ÁRVORE LM AMARRADA/TIFOR",qtd:4,unit:412.37},
  {op:"I",cod:"60134",desc:"PODA PESADA LM EM ÁREA AGRUPADA",qtd:42,unit:187.9},
  {op:"I",cod:"60135",desc:"PODA PESADA LM EM ÁREA ISOLADA",qtd:33,unit:245.12},
  {op:"I",cod:"60120",desc:"PODA LEVE",qtd:10.5,unit:12.345},     /* parcial, unitário com 3 casas */
  {op:"I",cod:"60121",desc:"PODA MÉDIA",qtd:5,unit:90},             /* conferido */
  {op:"R",cod:"63000",desc:"RETIRADA DE GALHADA",qtd:1,unit:80},    /* pago a mais */
];
const MEDICAO=[ /* o que a Energisa pagou nesta obra */
  {op:"I",cod:"60120",desc:"PODA LEVE",qtd:4,vlr:49.38},
  {op:"I",cod:"60121",desc:"PODA MÉDIA",qtd:5,vlr:450},
  {op:"R",cod:"63000",desc:"RETIRADA DE GALHADA",qtd:2,vlr:160},
  {op:"I",cod:"70001",desc:"ROÇADA",qtd:3,vlr:45},                  /* pago sem apontamento */
];
/* medição de outra obra, carregada antes: traz o preço pago dos serviços pendentes */
const OUTRA=[
  {op:"I",cod:"66011",desc:"CORTE ÁRVORE LM AMARRADA/TIFOR",qtd:1,vlr:412.37},
  {op:"I",cod:"60134",desc:"PODA PESADA LM EM ÁREA AGRUPADA",qtd:2,vlr:375.8},
  {op:"I",cod:"60135",desc:"PODA PESADA LM EM ÁREA ISOLADA",qtd:1,vlr:245.12},
];

let falhas=0, total=0;
const ok=(cond,msg)=>{total++; if(cond) console.log("  ✔ "+msg); else {falhas++; console.log("  ✘ "+msg)}};
const perto=(a,b)=>Math.abs(a-b)<0.006;

(async()=>{
  const tmp=fs.mkdtempSync(path.join(os.tmpdir(),"obra-paga-"));
  const b=await pw.chromium.launch();
  async function abrir(){
    const pg=await b.newPage({viewport:{width:1400,height:1400}});
    const erros=[]; pg.on("pageerror",e=>erros.push(e.message)); pg.on("dialog",d=>d.accept());
    await pg.goto("file://"+path.resolve(__dirname,"../index.html"));
    pg.erros=erros; return pg;
  }
  async function carregar(pg,xlsx,medicoes){
    const f=path.join(tmp,"fechamento.xlsx"); fs.writeFileSync(f,xlsx);
    await pg.setInputFiles("#arq",f); await pg.waitForFunction(()=>!!APONT);
    /* o PDF da medição entra pelo mesmo caminho do lerPDF: itens + aprenderPrecos */
    await pg.evaluate(ms=>{
      for(const it of ms.flat()) it.unit=it.vlr/it.qtd;
      PAGOS=ms.map((itens,i)=>({arquivo:"medicao"+i+".pdf",cab:{obra:"020-26-01062"},totais:{pagar:0},avisos:[],itens}));
      PAGOS.forEach(p=>aprenderPrecos(p.itens)); carregarAjustes(); render();
    },medicoes);
  }
  const app=pg=>pg.evaluate(()=>{
    const L=confronto().filter(x=>!x.inativo);
    return {linhas:Object.fromEntries(L.map(x=>[x.op+"|"+x.cod,x])),
      apontado:L.reduce((s,x)=>s+x.va,0),pend:aReceber(L),abate:abatimento(L),liq:liquido(L),
      cards:[...document.querySelectorAll(".card")].map(c=>c.innerText.replace(/\s+/g," ")),
      obra:obraApont()};
  });
  const compara=(a,ref,rot)=>{
    let iguais=true;
    for(const [k,r] of Object.entries(ref.linhas)){
      const x=a.linhas[k];
      if(!x||x.sit!==r.sit||!perto(x.vdif,r.vdif)){iguais=false; console.log("     ≠",k,x&&[x.sit,x.vdif],"ref",[r.sit,r.vdif])}
    }
    ok(iguais&&Object.keys(a.linhas).length===Object.keys(ref.linhas).length,rot+": situação e valor de cada serviço iguais ao app de obras");
    ok(perto(a.pend,ref.pend),rot+`: pendente ${a.pend.toFixed(2)} = ${ref.pend.toFixed(2)}`);
    ok(perto(a.abate,ref.maior+ref.sem),rot+`: pago a mais + sem apontamento ${a.abate.toFixed(2)} = ${(ref.maior+ref.sem).toFixed(2)}`);
    ok(perto(a.liq,ref.cobranca),rot+`: a receber líquido ${a.liq.toFixed(2)} = valor de cobrança ${ref.cobranca.toFixed(2)}`);
  };

  console.log("\n1) Planilha com o total em fórmula (=E*F), como o Excel grava");
  { const pg=await abrir(); await carregar(pg,fechamento(OBRA,APONT,"formula"),[MEDICAO]); const a=await app(pg);
    ok(a.obra===OBRA,`número da obra lido: ${a.obra}`);
    const esperado=APONT.reduce((s,x)=>s+x.qtd*x.unit,0);
    ok(perto(a.apontado,esperado),`apontado R$ ${a.apontado.toFixed(2)} (esperado ${esperado.toFixed(2)}) — antes saía R$ 0,00`);
    ok(perto(a.linhas["I|60120"].qa,10.5)&&perto(a.linhas["I|60120"].unit,12.345),"quantidade 10,5 e unitário 12,345 lidos sem virar milhar");
    compara(a,referencia(APONT,MEDICAO,[]),"fórmula");
    ok(!pg.erros.length,"sem erro na página "+(pg.erros.join(" | ")));
    await pg.close(); }

  console.log("\n2) Planilha só com quantidades + preço pago aprendido de outra medição");
  { const pg=await abrir();
    await pg.evaluate(o=>{for(const it of o) it.unit=it.vlr/it.qtd; aprenderPrecos(o)},OUTRA);
    await pg.reload(); /* o preço tem de sobreviver ao recarregar a página */
    await carregar(pg,fechamento(OBRA,APONT,"nada"),[MEDICAO]); const a=await app(pg);
    ok(perto(a.linhas["I|60134"].vdif,42*187.9),`PODA PESADA AGRUPADA pendente: 42 × 187,90 = ${a.linhas["I|60134"].vdif.toFixed(2)}`);
    ok(perto(a.linhas["I|66011"].vdif,4*412.37),`CORTE ÁRVORE pendente: 4 × 412,37 = ${a.linhas["I|66011"].vdif.toFixed(2)}`);
    compara(a,referencia(APONT.map(x=>({...x,unit:0})),MEDICAO,OUTRA),"só quantidade");
    ok(await pg.locator("text=sem preço").count()===0,"nenhum serviço ficou 'sem preço'");
    await pg.close(); }

  console.log("\n3) Sem preço em lugar nenhum: o app avisa em vez de mostrar R$ 0,00 calado");
  { const pg=await abrir(); await carregar(pg,fechamento(OBRA,APONT,"nada"),[MEDICAO]);
    ok(await pg.locator("td >> text=sem preço").count()===3,"os 3 pendentes sem preço aparecem marcados");
    ok(await pg.locator("text=3 serviço(s) sem preço").count()===1,"aviso no topo da ficha");
    await pg.close(); }

  console.log("\n4) Duas parcelas pagas somam");
  { const pg=await abrir();
    const p1=MEDICAO.map(x=>({...x})), p2=[{op:"I",cod:"60120",desc:"PODA LEVE",qtd:2,vlr:24.69}];
    await carregar(pg,fechamento(OBRA,APONT,"formula"),[p1,p2]); const a=await app(pg);
    ok(perto(a.linhas["I|60120"].qp,6)&&perto(a.linhas["I|60120"].vdif,4.5*12.345),`PODA LEVE: pago 4+2=6, falta 4,5 × 12,345 = ${a.linhas["I|60120"].vdif.toFixed(2)}`);
    compara(a,referencia(APONT,[...p1,...p2],[]),"parcelas");
    await pg.close(); }

  console.log("\n5) Quantidade editada na ficha refaz valor, situação e totais");
  { const pg=await abrir(); await carregar(pg,fechamento(OBRA,APONT,"formula"),[MEDICAO]);
    await pg.fill('input[data-ed="I|60120|qp"]',"10.5"); await pg.press('input[data-ed="I|60120|qp"]',"Enter");
    const a=await app(pg); const x=a.linhas["I|60120"];
    ok(x.sit==="Conferido"&&perto(x.vdif,0),"PODA LEVE com pago = apontado vira Conferido, valor 0");
    const ap2=APONT, pg2=MEDICAO.map(m=>m.cod==="60120"?{...m,qtd:10.5,vlr:10.5*12.345}:m);
    compara(a,referencia(ap2,pg2,[]),"editado");
    await pg.close(); }

  console.log("\n6) Serviço desativado sai dos totais; serviço acrescentado entra");
  { const pg=await abrir(); await carregar(pg,fechamento(OBRA,APONT,"formula"),[MEDICAO]);
    const antes=(await app(pg)).pend;
    await pg.click('button[data-inat="I|60134"]');
    ok(perto((await app(pg)).pend,antes-42*187.9),"desativar PODA PESADA AGRUPADA tira 42 × 187,90 do pendente");
    await pg.click("button[data-addserv]"); await pg.fill("#addCod","88888"); await pg.fill("#addServ","EXTRA");
    await pg.fill("#addQa","2"); await pg.fill("#addUnit","50"); await pg.click("button[data-addok]");
    ok(perto((await app(pg)).pend,antes-42*187.9+100),"acrescentar 2 × 50,00 soma 100,00 no pendente");
    await pg.close(); }

  console.log("\n7) Relatório, PNG e CSV com os mesmos números");
  { const pg=await abrir(); await carregar(pg,fechamento(OBRA,APONT,"formula"),[MEDICAO]);
    await pg.evaluate(()=>document.querySelectorAll("#relOpcoes input").forEach(i=>i.checked=true));
    const a=await app(pg);
    const r=await pg.evaluate(()=>({h:relHTML(),d:relDados()}));
    ok(perto(r.d.total,a.liq),`total do relatório ${r.d.total.toFixed(2)} = a receber ${a.liq.toFixed(2)}`);
    const brl=v=>"R$ "+v.toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2});
    ok(r.h.includes(brl(a.liq).replace(/ /g," "))||r.h.includes(brl(a.liq)),"valor líquido aparece no relatório");
    ok(r.h.includes("Divergência")&&!r.h.includes("Apontado R$"),"relatório com as colunas e rótulos do app de obras");
    await pg.click("#btnRel");
    const [dl]=await Promise.all([pg.waitForEvent("download"),pg.click("#relPng")]);
    ok(/\.png$/.test(dl.suggestedFilename()),"PNG gerado: "+dl.suggestedFilename());
    const [dc]=await Promise.all([pg.waitForEvent("download"),pg.click("#relCsv")]);
    const csv=fs.readFileSync(await dc.path(),"utf8");
    ok(csv.split("\r\n").length===1+Object.keys(a.linhas).length,"CSV com uma linha por serviço");
    ok(!pg.erros.length,"sem erro na página "+(pg.erros.join(" | ")));
    await pg.close(); }

  console.log("\n8) Total da planilha ao lado de \"R$\": o app confere o apontado com ele");
  { const pg=await abrir(); await carregar(pg,fechamento(OBRA,APONT,"formula",{total:true}),[MEDICAO]);
    const esperado=APONT.reduce((s,x)=>s+x.qtd*x.unit,0);
    const tot=await pg.evaluate(()=>num(APONT.cab.total));
    ok(perto(tot,esperado),`total lido da planilha: ${tot.toFixed(2)}`);
    ok(await pg.locator("#avisoConf >> text=Apontado confere com o total da planilha").count()===1,"aviso ✔ apontado confere");
    await pg.fill('input[data-ed="I|60134|qa"]',"40"); await pg.press('input[data-ed="I|60134|qa"]',"Enter");
    const txt=await pg.locator("#avisoConf").innerText();
    ok(/quantidades corrigidas à mão: − R\$ 375,80/.test(txt),"quantidade corrigida aparece como causa da diferença (− 2 × 187,90)");
    await pg.click("#btnRestaurar");
    ok(await pg.locator("#avisoConf >> text=Apontado confere com o total da planilha").count()===1,"Restaurar obra volta a bater com a planilha");
    await pg.close(); }

  console.log("\n9) Pasta com duas abas: lê a primeira do Excel, não o arquivo sheet1.xml");
  { const pg=await abrir();
    const antiga=APONT.map(x=>({...x,qtd:x.qtd+1}));
    await carregar(pg,fechamento(OBRA,APONT,"formula",{total:true,outraAba:antiga}),[MEDICAO]);
    const r=await pg.evaluate(()=>({aba:APONT.aba,soma:APONT.itens.reduce((s,x)=>s+x.vlr,0)}));
    ok(r.aba==="FECHAMENTO",`aba lida: ${r.aba}`);
    ok(perto(r.soma,APONT.reduce((s,x)=>s+x.qtd*x.unit,0)),`soma da aba certa: ${r.soma.toFixed(2)}`);
    await pg.close(); }

  await b.close();
  console.log(`\n${total-falhas} de ${total} verificações passaram.`);
  process.exit(falhas?1:0);
})().catch(e=>{console.error(e);process.exit(1)});

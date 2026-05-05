import { useState, useEffect, useCallback } from "react";

// ── Google Sheets via Vercel env ──
const SCRIPT_URL = import.meta.env.VITE_SHEETS_URL;

// ── Imágenes embebidas ──

// ── Storage: Google Sheets como fuente de verdad ──
// Cada "tabla" es una hoja en el Sheet. Guardamos 1 sola fila con columna A = key, columna B = JSON value
const ST = {
  get: async k => {
    if (SCRIPT_URL) {
      try {
        const r = await fetch(`${SCRIPT_URL}?hoja=${encodeURIComponent(k)}&t=${Date.now()}`);
        const d = await r.json();
        // El Apps Script devuelve {datos: [[key, jsonValue], ...]}
        // Buscamos la fila donde col[0] === k
        if (d.datos && d.datos.length > 0) {
          const row = d.datos.find(r => r[0] === k);
          if (row && row[1]) {
            try { return JSON.parse(row[1]); } catch(e) {}
          }
          // Si no hay fila con key, y hay datos, intentar parsear primera celda
          if (d.datos[0] && d.datos[0][0]) {
            try { return JSON.parse(d.datos[0][0]); } catch(e) {}
          }
        }
        return null;
      } catch(e) { console.error("ST.get error:", e); }
    }
    try { const r = await window.storage.get("wc26_"+k,true); return r?JSON.parse(r.value):null; } catch(e){return null;}
  },
  set: async (k,v) => {
    const json = JSON.stringify(v);
    if (SCRIPT_URL) {
      try {
        // Guardamos como una fila: [key, jsonValue]
        // Usamos accion="escribir" con valores=[[k, json]]
        // El Apps Script hace upsert por primera columna
        await fetch(SCRIPT_URL, {
          method:"POST",
          headers:{"Content-Type":"text/plain;charset=utf-8"},
          body: JSON.stringify({ hoja:k, accion:"escribir", fila:1, columna:1, valores:[[k, json]] })
        });
      } catch(e){ console.error("ST.set error:", e); }
    }
    try { await window.storage.set("wc26_"+k,json,true); } catch(e){}
  },
};

// ── Excel export ──
async function exportXLSX(users, preds, res, cl) {
  try {
    const { utils, writeFile } = await import("https://cdn.sheetjs.com/xlsx-0.20.2/package/xlsx.mjs");
    const wb = utils.book_new();
    const uRows = [["nombre","pin","createdAt"],...Object.values(users).map(u=>[u.name,u.pin,u.ts?new Date(u.ts).toLocaleString("es-AR"):""])];
    utils.book_append_sheet(wb, utils.aoa_to_sheet(uRows), "Usuarios");
    const pRows = [["nombre","partidoId","fase","local","visitante","gl","gv"]];
    Object.entries(preds).forEach(([name,ps])=>{
      Object.entries(ps||{}).forEach(([pid,v])=>{
        const p = ALL.find(x=>x.id===Number(pid));
        if(p&&(v.l!==""||v.v!==""))
          pRows.push([name,Number(pid),p.grupo||p.fase,p.local||p.lD,p.visitante||p.vD,v.l,v.v]);
      });
    });
    utils.book_append_sheet(wb, utils.aoa_to_sheet(pRows), "Pronósticos");
    const rRows = [["partidoId","fase","local","visitante","gl","gv"]];
    Object.entries(res||{}).forEach(([pid,v])=>{
      const p = ALL.find(x=>x.id===Number(pid));
      if(p&&(v.l!==""||v.v!==""))
        rRows.push([Number(pid),p.grupo||p.fase,p.local||p.lD,p.visitante||p.vD,v.l,v.v]);
    });
    utils.book_append_sheet(wb, utils.aoa_to_sheet(rRows), "Resultados");
    const cRows = [["partidoId","label","local","visitante"]];
    Object.entries(cl||{}).forEach(([pid,v])=>{
      const p = PO.find(x=>x.id===Number(pid));
      if(p) cRows.push([Number(pid),p.label,v.local||"",v.visitante||""]);
    });
    utils.book_append_sheet(wb, utils.aoa_to_sheet(cRows), "Clasificados");
    const tabla = Object.values(users).map(u=>{
      let pts=0,ex=0,pa=0;
      ALL.forEach(p=>{const pr=preds[u.name]?.[p.id],re=res[p.id],pt=cPts({l:pr?.l,v:pr?.v},re);pts+=pt;if(pt===3)ex++;else if(pt===1)pa++;});
      return [u.name,pts,ex,pa];
    }).sort((a,b)=>b[1]-a[1]);
    utils.book_append_sheet(wb, utils.aoa_to_sheet([["nombre","puntos","exactos","ganador_empate"],...tabla]), "Tabla");
    writeFile(wb, `BachiProde2026_${new Date().toISOString().slice(0,10)}.xlsx`);
    return true;
  } catch(e) { console.error(e); return false; }
}

const FL={"México":"🇲🇽","Sudáfrica":"🇿🇦","Corea del Sur":"🇰🇷","Chequia":"🇨🇿","Canadá":"🇨🇦","Bosnia y Herz.":"🇧🇦","Qatar":"🇶🇦","Suiza":"🇨🇭","Brasil":"🇧🇷","Marruecos":"🇲🇦","Haití":"🇭🇹","Escocia":"🏴󠁧󠁢󠁳󠁣󠁴󠁿","Estados Unidos":"🇺🇸","Paraguay":"🇵🇾","Australia":"🇦🇺","Türkiye":"🇹🇷","Alemania":"🇩🇪","Curazao":"🇨🇼","Costa de Marfil":"🇨🇮","Ecuador":"🇪🇨","Países Bajos":"🇳🇱","Japón":"🇯🇵","Suecia":"🇸🇪","Túnez":"🇹🇳","Bélgica":"🇧🇪","Egipto":"🇪🇬","Irán":"🇮🇷","Nueva Zelanda":"🇳🇿","España":"🇪🇸","Cabo Verde":"🇨🇻","Arabia Saudita":"🇸🇦","Uruguay":"🇺🇾","Francia":"🇫🇷","Senegal":"🇸🇳","Irak":"🇮🇶","Noruega":"🇳🇴","Argentina":"🇦🇷","Argelia":"🇩🇿","Austria":"🇦🇹","Jordania":"🇯🇴","Portugal":"🇵🇹","DR Congo":"🇨🇩","Uzbekistán":"🇺🇿","Colombia":"🇨🇴","Inglaterra":"🏴󠁧󠁢󠁥󠁮󠁧󠁿","Croacia":"🇭🇷","Ghana":"🇬🇭","Panamá":"🇵🇦"};
const TEAMS=Object.keys(FL);
const GL=["A","B","C","D","E","F","G","H","I","J","K","L"];

const PG=[
  [1,"A","México","Sudáfrica","2026-06-11T20:00Z","Estadio Azteca","Ciudad de México"],
  [2,"A","Corea del Sur","Chequia","2026-06-12T03:00Z","Estadio Akron","Guadalajara"],
  [3,"A","Chequia","Sudáfrica","2026-06-18T17:00Z","Mercedes-Benz Stadium","Atlanta"],
  [4,"A","México","Corea del Sur","2026-06-19T02:00Z","Estadio Akron","Guadalajara"],
  [5,"A","Chequia","México","2026-06-25T02:00Z","Estadio Azteca","Ciudad de México"],
  [6,"A","Sudáfrica","Corea del Sur","2026-06-25T02:00Z","Estadio BBVA","Monterrey"],
  [7,"B","Canadá","Bosnia y Herz.","2026-06-12T20:00Z","BMO Field","Toronto"],
  [8,"B","Qatar","Suiza","2026-06-13T20:00Z","Levi's Stadium","San Francisco"],
  [9,"B","Suiza","Bosnia y Herz.","2026-06-18T20:00Z","SoFi Stadium","Los Ángeles"],
  [10,"B","Canadá","Qatar","2026-06-18T23:00Z","BC Place","Vancouver"],
  [11,"B","Suiza","Canadá","2026-06-24T20:00Z","BC Place","Vancouver"],
  [12,"B","Bosnia y Herz.","Qatar","2026-06-24T20:00Z","Lumen Field","Seattle"],
  [13,"C","Brasil","Marruecos","2026-06-13T23:00Z","MetLife Stadium","Nueva York/NJ"],
  [14,"C","Haití","Escocia","2026-06-14T02:00Z","Gillette Stadium","Boston"],
  [15,"C","Escocia","Marruecos","2026-06-19T23:00Z","Gillette Stadium","Boston"],
  [16,"C","Brasil","Haití","2026-06-20T01:30Z","Lincoln Financial Field","Filadelfia"],
  [17,"C","Escocia","Brasil","2026-06-24T23:00Z","Hard Rock Stadium","Miami"],
  [18,"C","Marruecos","Haití","2026-06-24T23:00Z","Mercedes-Benz Stadium","Atlanta"],
  [19,"D","Estados Unidos","Paraguay","2026-06-13T02:00Z","SoFi Stadium","Los Ángeles"],
  [20,"D","Australia","Türkiye","2026-06-13T05:00Z","BC Place","Vancouver"],
  [21,"D","Estados Unidos","Australia","2026-06-19T20:00Z","Lumen Field","Seattle"],
  [22,"D","Türkiye","Paraguay","2026-06-20T04:00Z","Levi's Stadium","San Francisco"],
  [23,"D","Türkiye","Estados Unidos","2026-06-26T03:00Z","SoFi Stadium","Los Ángeles"],
  [24,"D","Paraguay","Australia","2026-06-26T03:00Z","Levi's Stadium","San Francisco"],
  [25,"E","Alemania","Curazao","2026-06-14T18:00Z","NRG Stadium","Houston"],
  [26,"E","Costa de Marfil","Ecuador","2026-06-15T00:00Z","Lincoln Financial Field","Filadelfia"],
  [27,"E","Alemania","Costa de Marfil","2026-06-20T21:00Z","BMO Field","Toronto"],
  [28,"E","Ecuador","Curazao","2026-06-21T01:00Z","Arrowhead Stadium","Kansas City"],
  [29,"E","Curazao","Costa de Marfil","2026-06-25T21:00Z","Lincoln Financial Field","Filadelfia"],
  [30,"E","Ecuador","Alemania","2026-06-25T21:00Z","MetLife Stadium","Nueva York/NJ"],
  [31,"F","Países Bajos","Japón","2026-06-14T21:00Z","AT&T Stadium","Dallas"],
  [32,"F","Suecia","Túnez","2026-06-15T03:00Z","Estadio BBVA","Monterrey"],
  [33,"F","Países Bajos","Suecia","2026-06-20T18:00Z","NRG Stadium","Houston"],
  [34,"F","Túnez","Japón","2026-06-20T05:00Z","Estadio BBVA","Monterrey"],
  [35,"F","Japón","Suecia","2026-06-26T00:00Z","AT&T Stadium","Dallas"],
  [36,"F","Túnez","Países Bajos","2026-06-26T00:00Z","Arrowhead Stadium","Kansas City"],
  [37,"G","Bélgica","Egipto","2026-06-15T20:00Z","Lumen Field","Seattle"],
  [38,"G","Irán","Nueva Zelanda","2026-06-16T02:00Z","SoFi Stadium","Los Ángeles"],
  [39,"G","Bélgica","Irán","2026-06-21T20:00Z","SoFi Stadium","Los Ángeles"],
  [40,"G","Nueva Zelanda","Egipto","2026-06-22T02:00Z","BC Place","Vancouver"],
  [41,"G","Egipto","Irán","2026-06-27T04:00Z","Lumen Field","Seattle"],
  [42,"G","Nueva Zelanda","Bélgica","2026-06-27T04:00Z","BC Place","Vancouver"],
  [43,"H","España","Cabo Verde","2026-06-15T17:00Z","Mercedes-Benz Stadium","Atlanta"],
  [44,"H","Arabia Saudita","Uruguay","2026-06-15T23:00Z","Hard Rock Stadium","Miami"],
  [45,"H","España","Arabia Saudita","2026-06-21T17:00Z","Mercedes-Benz Stadium","Atlanta"],
  [46,"H","Uruguay","Cabo Verde","2026-06-21T23:00Z","Hard Rock Stadium","Miami"],
  [47,"H","Cabo Verde","Arabia Saudita","2026-06-27T01:00Z","NRG Stadium","Houston"],
  [48,"H","Uruguay","España","2026-06-27T01:00Z","Estadio Akron","Guadalajara"],
  [49,"I","Francia","Senegal","2026-06-16T20:00Z","MetLife Stadium","Nueva York/NJ"],
  [50,"I","Irak","Noruega","2026-06-16T23:00Z","Gillette Stadium","Boston"],
  [51,"I","Francia","Irak","2026-06-22T22:00Z","Lincoln Financial Field","Filadelfia"],
  [52,"I","Noruega","Senegal","2026-06-23T01:00Z","MetLife Stadium","Nueva York/NJ"],
  [53,"I","Noruega","Francia","2026-06-26T20:00Z","Gillette Stadium","Boston"],
  [54,"I","Senegal","Irak","2026-06-26T20:00Z","BMO Field","Toronto"],
  [55,"J","Argentina","Argelia","2026-06-17T02:00Z","Arrowhead Stadium","Kansas City"],
  [56,"J","Austria","Jordania","2026-06-17T05:00Z","Levi's Stadium","San Francisco"],
  [57,"J","Argentina","Austria","2026-06-22T18:00Z","AT&T Stadium","Dallas"],
  [58,"J","Jordania","Argelia","2026-06-23T04:00Z","Levi's Stadium","San Francisco"],
  [59,"J","Jordania","Argentina","2026-06-28T03:00Z","AT&T Stadium","Dallas"],
  [60,"J","Argelia","Austria","2026-06-28T03:00Z","Arrowhead Stadium","Kansas City"],
  [61,"K","Portugal","DR Congo","2026-06-17T18:00Z","NRG Stadium","Houston"],
  [62,"K","Uzbekistán","Colombia","2026-06-18T03:00Z","Estadio Azteca","Ciudad de México"],
  [63,"K","Portugal","Uzbekistán","2026-06-23T18:00Z","NRG Stadium","Houston"],
  [64,"K","Colombia","DR Congo","2026-06-24T03:00Z","Estadio Akron","Guadalajara"],
  [65,"K","Colombia","Portugal","2026-06-28T00:30Z","Hard Rock Stadium","Miami"],
  [66,"K","DR Congo","Uzbekistán","2026-06-28T00:30Z","Mercedes-Benz Stadium","Atlanta"],
  [67,"L","Inglaterra","Croacia","2026-06-17T21:00Z","AT&T Stadium","Dallas"],
  [68,"L","Ghana","Panamá","2026-06-18T00:00Z","BMO Field","Toronto"],
  [69,"L","Inglaterra","Ghana","2026-06-23T21:00Z","Gillette Stadium","Boston"],
  [70,"L","Panamá","Croacia","2026-06-24T00:00Z","BMO Field","Toronto"],
  [71,"L","Panamá","Inglaterra","2026-06-27T22:00Z","MetLife Stadium","Nueva York/NJ"],
  [72,"L","Croacia","Ghana","2026-06-27T22:00Z","Lincoln Financial Field","Filadelfia"],
].map(r=>({id:r[0],grupo:r[1],local:r[2],visitante:r[3],dt:r[4],estadio:r[5],ciudad:r[6]}));

const PO=[
  [1001,"R32","M73","2°A","2°B","2026-06-28T20:00Z","SoFi Stadium","Los Ángeles"],
  [1002,"R32","M74","1°E","Mej.3°","2026-06-29T18:00Z","Gillette Stadium","Boston"],
  [1003,"R32","M75","1°F","2°C","2026-06-29T22:00Z","Estadio BBVA","Monterrey"],
  [1004,"R32","M76","1°C","2°F","2026-06-30T02:00Z","NRG Stadium","Houston"],
  [1005,"R32","M77","1°I","Mej.3°","2026-06-30T18:00Z","MetLife Stadium","Nueva York/NJ"],
  [1006,"R32","M78","2°E","2°I","2026-06-30T22:00Z","AT&T Stadium","Dallas"],
  [1007,"R32","M79","1°A","Mej.3°","2026-07-01T02:00Z","Estadio Azteca","Ciudad de México"],
  [1008,"R32","M80","1°L","Mej.3°","2026-07-01T17:00Z","Mercedes-Benz Stadium","Atlanta"],
  [1009,"R32","M81","1°D","Mej.3°","2026-07-01T21:00Z","Levi's Stadium","San Francisco"],
  [1010,"R32","M82","1°G","Mej.3°","2026-07-02T01:00Z","Lumen Field","Seattle"],
  [1011,"R32","M83","2°K","2°L","2026-07-02T20:00Z","BMO Field","Toronto"],
  [1012,"R32","M84","1°H","2°J","2026-07-03T00:00Z","SoFi Stadium","Los Ángeles"],
  [1013,"R32","M85","1°B","Mej.3°","2026-07-03T04:00Z","BC Place","Vancouver"],
  [1014,"R32","M86","1°J","2°H","2026-07-03T21:00Z","Hard Rock Stadium","Miami"],
  [1015,"R32","M87","1°K","Mej.3°","2026-07-04T02:30Z","Arrowhead Stadium","Kansas City"],
  [1016,"R32","M88","2°D","2°G","2026-07-04T20:00Z","AT&T Stadium","Dallas"],
  [2001,"R16","M89","G.M74","G.M77","2026-07-04T22:00Z","Lincoln Financial Field","Filadelfia"],
  [2002,"R16","M90","G.M73","G.M75","2026-07-04T18:00Z","NRG Stadium","Houston"],
  [2003,"R16","M91","G.M76","G.M78","2026-07-05T21:00Z","MetLife Stadium","Nueva York/NJ"],
  [2004,"R16","M92","G.M79","G.M80","2026-07-05T01:00Z","Estadio Azteca","Ciudad de México"],
  [2005,"R16","M93","G.M83","G.M84","2026-07-06T20:00Z","AT&T Stadium","Dallas"],
  [2006,"R16","M94","G.M81","G.M82","2026-07-06T01:00Z","Lumen Field","Seattle"],
  [2007,"R16","M95","G.M86","G.M88","2026-07-07T17:00Z","Mercedes-Benz Stadium","Atlanta"],
  [2008,"R16","M96","G.M85","G.M87","2026-07-07T21:00Z","BC Place","Vancouver"],
  [3001,"CF","CF1","G.M89","G.M90","2026-07-09T21:00Z","Gillette Stadium","Boston"],
  [3002,"CF","CF2","G.M91","G.M92","2026-07-10T20:00Z","SoFi Stadium","Los Ángeles"],
  [3003,"CF","CF3","G.M93","G.M94","2026-07-11T22:00Z","Hard Rock Stadium","Miami"],
  [3004,"CF","CF4","G.M95","G.M96","2026-07-12T02:00Z","Arrowhead Stadium","Kansas City"],
  [4001,"SF","SF1","G.CF1","G.CF2","2026-07-14T20:00Z","AT&T Stadium","Dallas"],
  [4002,"SF","SF2","G.CF3","G.CF4","2026-07-15T20:00Z","Mercedes-Benz Stadium","Atlanta"],
  [5001,"F","3°Pto","Perd.SF1","Perd.SF2","2026-07-18T22:00Z","Hard Rock Stadium","Miami"],
  [5002,"F","FINAL","G.SF1","G.SF2","2026-07-19T20:00Z","MetLife Stadium","Nueva York/NJ"],
].map(r=>({id:r[0],fase:r[1],label:r[2],lD:r[3],vD:r[4],dt:r[5],estadio:r[6],ciudad:r[7]}));

const ALL=[...PG,...PO];
const PO_TABS=["R32","R16","CF","SF","F"];
const SALUDOS=[
  n=>`¡Hola ${n}! ¿Otra vez por aquí? 🤔`,
  n=>`¿Te sentís de suerte, ${n}? 😉`,
  n=>`¡Hola ${n}! ¡Dale que sos vos! 🥹`,
  n=>`Nunca dije nada, pero si necesitás ayuda estoy... 🤫`,
];
const isLk=dt=>Date.now()>=new Date(dt).getTime()-300000;
const fmt=iso=>{try{return new Date(iso).toLocaleString("es-AR",{weekday:"short",day:"numeric",month:"short",hour:"2-digit",minute:"2-digit",timeZone:"America/Argentina/Buenos_Aires"});}catch(e){return "";}};
const cR=(a,b)=>{const x=parseInt(a),y=parseInt(b);if(isNaN(x)||isNaN(y))return null;return x>y?"L":x<y?"V":"E";};
const cPts=(p,r)=>{if(!p||!r||r.l===""||r.l===undefined)return 0;const a=cR(p.l,p.v),b=cR(r.l,r.v);if(!a||!b)return 0;if(String(p.l)===String(r.l)&&String(p.v)===String(r.v))return 3;if(a===b)return 1;return 0;};

const K={bg:"#06091a",gold:"#f0b429",gn:"#22c55e",rd:"#ef4444",mu:"#6b7280",tx:"#e5e7eb",di:"#9ca3af",bo:"rgba(255,255,255,0.09)",ca:"rgba(255,255,255,0.05)"};
const W={minHeight:"100vh",display:"flex",justifyContent:"center",background:K.bg,fontFamily:"'Trebuchet MS','Segoe UI',sans-serif"};
const BK={background:"none",border:"none",color:K.mu,cursor:"pointer",fontSize:13,padding:0,marginBottom:10,fontFamily:"inherit"};
const IN={background:"rgba(255,255,255,0.07)",border:`1px solid ${K.bo}`,borderRadius:8,color:"#fff",padding:"10px 13px",fontSize:15,outline:"none",width:"100%",boxSizing:"border-box",fontFamily:"inherit"};
const SC={width:36,textAlign:"center",background:"rgba(255,255,255,0.08)",border:"1px solid rgba(255,255,255,0.2)",borderRadius:7,color:"#fff",padding:"6px 2px",fontSize:14,fontWeight:700,outline:"none",fontFamily:"inherit"};
const CR=(
  <p style={{
    position:"fixed",bottom:"5mm",left:0,right:0,
    textAlign:"center",color:"#374151",fontSize:10,
    margin:0,padding:"0 16px",zIndex:999,
    pointerEvents:"none"
  }}>
    Copyright © 2026 Bachiprode Mundial LLC. Todos los derechos reservados.
  </p>
);

const Btn=({onClick,ch,v,disabled,sm})=>{
  const bg=v==="g"?`linear-gradient(135deg,${K.gn},#15803d)`:v==="r"?`linear-gradient(135deg,${K.rd},#b91c1c)`:v==="s"?"rgba(255,255,255,0.1)":`linear-gradient(135deg,${K.gold},#d97706)`;
  return <button onClick={onClick} disabled={!!disabled} style={{background:bg,border:"none",borderRadius:8,color:"#fff",padding:sm?"6px 11px":"11px 20px",fontSize:sm?12:14,fontWeight:700,cursor:disabled?"not-allowed":"pointer",opacity:disabled?0.45:1,fontFamily:"inherit"}}>{ch}</button>;
};
const Tabs=({items,sel,onSel})=>(
  <div style={{display:"flex",gap:4,flexWrap:"wrap",marginBottom:10}}>
    {items.map(it=><button key={it} onClick={()=>onSel(it)} style={{padding:"4px 10px",borderRadius:20,border:"1px solid",borderColor:sel===it?K.gold:K.bo,background:sel===it?K.gold+"22":"transparent",color:sel===it?K.gold:K.di,fontWeight:600,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{it}</button>)}
  </div>
);
const MRow=({p,lv,vv,onL,onV,real,locked,adm,others})=>{
  const pts=cPts({l:lv,v:vv},real);
  const hasReal=real&&real.l!==""&&real.l!==undefined;
  const bl=locked&&!adm;
  const isPO=p.lD!==undefined;
  const ln=isPO?(p.localTeam||p.lD):p.local;
  const vn=isPO?(p.visitanteTeam||p.vD):p.visitante;
  const [showO,setShowO]=useState(false);
  return(
    <div style={{background:"rgba(255,255,255,0.03)",borderRadius:10,padding:"9px 10px",border:`1px solid ${lv&&vv?K.gold+"25":K.bo}`}}>
      <div style={{display:"flex",alignItems:"center",gap:4}}>
        <span style={{fontSize:11,color:K.tx,flex:1,textAlign:"right",lineHeight:1.4,minWidth:0}}>{FL[ln]||"🏳"} {ln}</span>
        <div style={{display:"flex",gap:4,alignItems:"center",flexShrink:0}}>
          {bl?<span style={{color:K.mu,fontSize:14,fontWeight:700,minWidth:20,textAlign:"center"}}>{lv||"–"}</span>
             :<input inputMode="numeric" type="text" value={lv||""} onChange={e=>onL(e.target.value.replace(/\D/g,"").slice(0,2))} style={SC} placeholder="?"/>}
          <span style={{color:K.mu,fontWeight:700}}>-</span>
          {bl?<span style={{color:K.mu,fontSize:14,fontWeight:700,minWidth:20,textAlign:"center"}}>{vv||"–"}</span>
             :<input inputMode="numeric" type="text" value={vv||""} onChange={e=>onV(e.target.value.replace(/\D/g,"").slice(0,2))} style={SC} placeholder="?"/>}
        </div>
        <span style={{fontSize:11,color:K.tx,flex:1,lineHeight:1.4,minWidth:0}}>{FL[vn]||"🏳"} {vn}</span>
      </div>
      <div style={{marginTop:4,display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:3,alignItems:"center"}}>
        <span style={{fontSize:9,color:K.mu}}>{fmt(p.dt)} · {p.estadio}, {p.ciudad}</span>
        <div style={{display:"flex",gap:4,alignItems:"center"}}>
          {bl&&<span style={{fontSize:10,color:K.rd}}>🔒</span>}
          {hasReal&&<span style={{fontSize:10,fontWeight:700,color:pts===3?K.gn:pts===1?K.gold:K.rd}}>{pts===3?"✓+3":pts===1?"⚽+1":"✗+0"} ({real.l}-{real.v})</span>}
          {bl&&others&&Object.values(others).some(x=>x&&x.l!=="")&&(
            <button onClick={()=>setShowO(s=>!s)} style={{background:"none",border:"none",color:showO?K.gold:K.di,fontSize:9,cursor:"pointer",padding:"0 2px"}}>👥{showO?"▲":"▼"}</button>
          )}
        </div>
      </div>
      {showO&&bl&&others&&(
        <div style={{marginTop:5,paddingTop:5,borderTop:`1px solid ${K.bo}`}}>
          <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
            {Object.entries(others).filter(([,pr])=>pr&&pr.l!==undefined&&pr.l!=="").map(([name,pr])=>(
              <span key={name} style={{fontSize:9,color:K.di,background:"rgba(255,255,255,0.05)",borderRadius:6,padding:"2px 6px"}}>{name}: {pr.l}-{pr.v}</span>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};

export default function App(){
  const [sc,setSc]=useState("splash");
  const [tab,setTab]=useState("A");
  const [ptab,setPtab]=useState("R32");
  const [atab,setAtab]=useState("usuarios");
  const [users,setUsers]=useState({});
  const [user,setUser]=useState(null);
  const [isAdm,setIsAdm]=useState(false);
  const [fN,setFN]=useState("");
  const [fP,setFP]=useState("");
  const [fP2,setFP2]=useState("");
  const [fP3,setFP3]=useState("");
  const [fAdm,setFAdm]=useState("");
  const [fErr,setFErr]=useState("");
  const [preds,setPreds]=useState({});
  const [lp,setLp]=useState({});
  const [res,setRes]=useState({});
  const [lr,setLr]=useState({});
  const [cl,setCl]=useState({});
  const [lcl,setLcl]=useState({});
  const [busy,setBusy]=useState(false);
  const [msg,setMsg]=useState("");
  const [loading,setLoading]=useState(true);
  const [eU,setEU]=useState(null);
  const [eN,setEN]=useState("");
  const [eP,setEP]=useState("");
  const [addMode,setAddMode]=useState(false);
  const [addN,setAddN]=useState("");
  const [addP,setAddP]=useState("2026");
  const [mergeMode,setMergeMode]=useState(false);
  const [mergeA,setMergeA]=useState("");
  const [mergeB,setMergeB]=useState("");
  const [saludo]=useState(()=>SALUDOS[Math.floor(Math.random()*4)]);

  useEffect(()=>{
    const link=document.querySelector("link[rel~='icon']")||document.createElement("link");
    link.rel="icon";link.type="image/png";link.href="/favicon.png";
    document.head.appendChild(link);
    document.title="Bachi Prode Mundial 2026";
  },[]);

  const flash=t=>{setMsg(t);setTimeout(()=>setMsg(""),3000);};

  const loadAll=useCallback(async()=>{
    setLoading(true);
    const [u,p,r,c]=await Promise.all([ST.get("usuarios"),ST.get("predicciones"),ST.get("resultados"),ST.get("clasificados")]);
    setUsers(u||{});setPreds(p||{});setRes(r||{});setLr(r||{});setCl(c||{});setLcl(c||{});
    setLoading(false);
  },[]);

  useEffect(()=>{loadAll();},[loadAll]);
  useEffect(()=>{if(user)setLp(preds[user.name]||{});},[user?.name]);

  const saveU=async u2=>{await ST.set("usuarios",u2);setUsers(u2);};

  const doName=async()=>{
    setFErr("");const n=fN.trim();
    if(!n)return setFErr("Ingresá tu nombre");
    const latest=await ST.get("usuarios")||{};
    setUsers(latest);
    if(latest[n]){setSc("pin");}
    else{
      const nu={...latest,[n]:{name:n,pin:"2026",first:false,ts:Date.now()}};
      await ST.set("usuarios",nu);setUsers(nu);setUser(nu[n]);
      const p=await ST.get("predicciones")||{};setLp(p[n]||{});
      setSc("prode");
    }
  };
  const doPin=async()=>{
    setFErr("");
    const latest=await ST.get("usuarios")||{};
    const u=latest[fN.trim()];
    if(!u||fP!==u.pin)return setFErr("PIN incorrecto");
    setUser(u);const p=await ST.get("predicciones")||{};setLp(p[u.name]||{});
    setSc("prode");
  };
  const doChPin=async()=>{
    setFErr("");
    if(!/^\d{4}$/.test(fP2))return setFErr("Debe tener exactamente 4 dígitos");
    if(fP2!==fP3)return setFErr("Los PINs no coinciden");
    const nu={...users,[user.name]:{...user,pin:fP2}};
    await saveU(nu);setUser(nu[user.name]);setFP2("");setFP3("");flash("PIN actualizado 🔐");setSc("prode");
  };

  const chLp=useCallback((id,t,v)=>setLp(p=>({...p,[id]:{...(p[id]||{}),[t]:v}})),[]);
  const chLr=useCallback((id,t,v)=>setLr(p=>({...p,[id]:{...(p[id]||{}),[t]:v}})),[]);
  const chLcl=useCallback((id,t,v)=>setLcl(p=>({...p,[id]:{...(p[id]||{}),[t]:v}})),[]);

  const savePreds=async()=>{
    setBusy(true);
    const latest=await ST.get("predicciones")||{};
    const np={...latest,[user.name]:lp};
    await ST.set("predicciones",np);setPreds(np);flash("¡Guardado! 🎉");setBusy(false);
  };
  const saveRes=async()=>{setBusy(true);await ST.set("resultados",lr);setRes(lr);flash("Resultados guardados ✅");setBusy(false);};
  const saveCl=async()=>{setBusy(true);await ST.set("clasificados",lcl);setCl(lcl);flash("Clasificados guardados ✅");setBusy(false);};

  const fetchFifa=async()=>{
    setBusy(true);
    try{
      const r=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,tools:[{type:"web_search_20250305",name:"web_search"}],
          messages:[{role:"user",content:'Resultados FINALES del Mundial FIFA 2026. Solo JSON array sin markdown: [{"local":"nombre","visitante":"nombre","gl":N,"gv":N}]. Sin resultados: [].'}]})});
      const d=await r.json();
      const tx=(d.content||[]).filter(b=>b.type==="text").map(b=>b.text).join("").replace(/```json|```/g,"").trim();
      const m=tx.match(/\[[\s\S]*?\]/);
      if(m){const arr=JSON.parse(m[0]);if(arr.length){const nv={...lr};arr.forEach(x=>{const p=PG.find(p=>p.local===x.local&&p.visitante===x.visitante);if(p)nv[p.id]={l:String(x.gl),v:String(x.gv)};});setLr(nv);await ST.set("resultados",nv);setRes(nv);flash(`${arr.length} resultado(s) ✅`);}else flash("Sin nuevos resultados");}
    }catch(e){flash("Error de conexión");}
    setBusy(false);
  };

  const doDelete=async name=>{
    if(!window.confirm(`⚠️ ¿Eliminar a "${name}"?\nSe borran todos sus pronósticos.\nEsta acción no se puede deshacer.`))return;
    const nu={...users};delete nu[name];
    const np={...preds};delete np[name];
    await ST.set("usuarios",nu);await ST.set("predicciones",np);
    setUsers(nu);setPreds(np);flash(`"${name}" eliminado`);
  };

  const doMerge=async()=>{
    if(!mergeA||!mergeB||mergeA===mergeB)return flash("Elegí 2 usuarios distintos");
    if(!window.confirm(`¿Fusionar "${mergeA}" y "${mergeB}"?\nSe conserva el nombre de "${mergeA}".\nPrioridad a pronósticos más nuevos.`))return;
    const pA=preds[mergeA]||{},pB=preds[mergeB]||{};
    const merged={...pB,...pA};
    const nu={...users};delete nu[mergeB];
    const np={...preds};delete np[mergeB];np[mergeA]=merged;
    await ST.set("usuarios",nu);await ST.set("predicciones",np);
    setUsers(nu);setPreds(np);setMergeMode(false);setMergeA("");setMergeB("");
    flash(`Fusionados en "${mergeA}" ✅`);
  };

  const tabla=Object.values(users).map(u=>{
    let pts=0,ex=0,pa=0,co=0;
    ALL.forEach(p=>{const pr=preds[u.name]?.[p.id],re=res[p.id],pt=cPts({l:pr?.l,v:pr?.v},re);pts+=pt;if(pt===3)ex++;else if(pt===1)pa++;if(pr?.l!==undefined&&pr?.l!==""&&pr?.v!==undefined&&pr?.v!=="")co++;});
    return{...u,pts,ex,pa,co};
  }).sort((a,b)=>b.pts-a.pts);

  const jugados=ALL.filter(p=>{const r=res[p.id];return r&&r.l!==""&&r.l!==undefined;}).length;
  const myComp=user?ALL.filter(p=>{const pr=lp[p.id];return pr?.l!==undefined&&pr?.l!==""&&pr?.v!==undefined&&pr?.v!=="";}).length:0;
  const otherNames=user?Object.keys(preds).filter(n=>n!==user.name):[];

  const PW=({children,maxW=500,center=false})=>(
    <div style={{...W,alignItems:center?"center":"flex-start",flexDirection:"column",padding:0}}>
      <div style={{width:"100%",maxWidth:maxW,padding:"14px 12px 0",flex:1}}>{children}</div>
      {CR}
    </div>
  );

  if(sc==="splash")return(
    <div style={{...W,alignItems:"stretch",flexDirection:"column",padding:0}}>
      <div style={{position:"relative",overflow:"hidden",width:"100%",maxHeight:310,flexShrink:0}}>
        <img src={"/promo.jpg"} alt="" style={{width:"100%",objectFit:"cover",objectPosition:"center top",display:"block"}}/>
        <div style={{position:"absolute",inset:0,background:"linear-gradient(to bottom,rgba(6,9,26,0) 35%,rgba(6,9,26,1) 100%)"}}/>
      </div>
      <div style={{display:"flex",flexDirection:"column",alignItems:"center",padding:"0 20px",marginTop:-70,zIndex:2,width:"100%",maxWidth:480,margin:"-70px auto 0"}}>
        <img src={"/logo.webp"} alt="Bachi Prode 2026"
          style={{width:"min(320px,78vw)",filter:"drop-shadow(0 3px 18px rgba(0,0,0,0.75))"}}/>
        <p style={{color:"#9ca3af",fontSize:12,margin:"10px 0 2px",textAlign:"center"}}>🇺🇸 🇲🇽 🇨🇦 · 11 Jun – 19 Jul 2026</p>
        <p style={{color:"#6b7280",fontSize:11,margin:"0 0 18px",textAlign:"center"}}>48 equipos · 104 partidos · Base compartida 🌐</p>
        {loading?<p style={{color:"#6b7280"}}>Cargando…</p>:(
          <div style={{display:"flex",flexDirection:"column",gap:9,width:"100%",maxWidth:320}}>
            <p style={{color:"#22c55e",fontSize:11,margin:"0 0 4px",textAlign:"center"}}>✅ {Object.keys(users).length} jugadores registrados</p>
            <Btn onClick={()=>{setFErr("");setFN("");setFP("");setSc("auth");}} ch="▶ Ingresar"/>
            <Btn v="s" onClick={()=>{loadAll();setSc("tabla");}} ch="🏅 Tabla de Posiciones"/>
            <Btn v="s" onClick={()=>{setFAdm("");setSc("adminAuth");}} ch="⚙️ Administrador"/>
          </div>
        )}
      </div>
      {CR}
    </div>
  );

  if(sc==="auth")return(
    <PW maxW={360}>
      <button style={BK} onClick={()=>setSc("splash")}>← Volver</button>
      <h2 style={{color:"#fff",fontWeight:900,fontSize:20,margin:"0 0 14px"}}>¿Cómo te llamás?</h2>
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        <input autoFocus value={fN} onChange={e=>setFN(e.target.value)} onKeyDown={e=>e.key==="Enter"&&doName()} placeholder="Tu nombre o apodo" style={IN}/>
        {fErr&&<p style={{color:"#ef4444",fontSize:13,margin:0}}>{fErr}</p>}
        <Btn onClick={doName} ch="Continuar →" disabled={!fN.trim()}/>
      </div>
      {Object.keys(users).length>0&&(
        <div style={{marginTop:12}}>
          <p style={{color:"#6b7280",fontSize:11,marginBottom:5}}>Ya están jugando:</p>
          <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
            {Object.keys(users).map(n=><span key={n} onClick={()=>setFN(n)} style={{background:"rgba(255,255,255,0.07)",borderRadius:20,padding:"3px 10px",fontSize:12,color:"#9ca3af",cursor:"pointer"}}>{n}</span>)}
          </div>
        </div>
      )}
    </PW>
  );

  if(sc==="pin")return(
    <PW maxW={360}>
      <button style={BK} onClick={()=>setSc("auth")}>← Volver</button>
      <h2 style={{color:"#fff",fontWeight:900,fontSize:18,margin:"0 0 6px"}}>{saludo(fN)}</h2>
      <p style={{color:"#9ca3af",fontSize:13,marginBottom:14}}>Ingresá tu PIN para continuar</p>
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        <input type="password" inputMode="numeric" maxLength={4} autoFocus value={fP} onChange={e=>setFP(e.target.value.replace(/\D/g,"").slice(0,4))} onKeyDown={e=>e.key==="Enter"&&doPin()} placeholder="• • • •" style={{...IN,fontSize:22,letterSpacing:8,textAlign:"center"}}/>
        {fErr&&<p style={{color:"#ef4444",fontSize:13,margin:0}}>{fErr}</p>}
        <Btn onClick={doPin} ch="Ingresar →" disabled={fP.length<4}/>
      </div>
      <p style={{color:"#6b7280",fontSize:10,marginTop:12}}>💡 Son 4 dígitos · Guardalo en el llavero/Keychain</p>
    </PW>
  );

  if(sc==="changePin"&&user)return(
    <PW maxW={320} center>
      <div style={{textAlign:"center",fontSize:30,marginBottom:8}}>🔐</div>
      <h2 style={{color:"#fff",fontWeight:900,fontSize:17,textAlign:"center",marginBottom:14}}>Cambiar PIN</h2>
      <div style={{display:"flex",flexDirection:"column",gap:9}}>
        <div><p style={{color:"#6b7280",fontSize:11,margin:"0 0 3px"}}>Nuevo PIN</p>
          <input type="password" inputMode="numeric" maxLength={4} autoFocus value={fP2} onChange={e=>setFP2(e.target.value.replace(/\D/g,"").slice(0,4))} placeholder="• • • •" style={{...IN,textAlign:"center",fontSize:18,letterSpacing:6}}/>
        </div>
        <div><p style={{color:"#6b7280",fontSize:11,margin:"0 0 3px"}}>Repetí el PIN</p>
          <input type="password" inputMode="numeric" maxLength={4} value={fP3} onChange={e=>setFP3(e.target.value.replace(/\D/g,"").slice(0,4))} onKeyDown={e=>e.key==="Enter"&&doChPin()} placeholder="• • • •" style={{...IN,textAlign:"center",fontSize:18,letterSpacing:6}}/>
        </div>
        {fErr&&<p style={{color:"#ef4444",fontSize:13,margin:0}}>{fErr}</p>}
        <Btn v="g" onClick={doChPin} ch="Guardar PIN →" disabled={fP2.length<4||fP3.length<4}/>
        <button onClick={()=>setSc("prode")} style={{...BK,marginBottom:0,textAlign:"center"}}>Cancelar</button>
      </div>
    </PW>
  );

  if(sc==="prode"&&user){
    const isP=tab==="Playoff";
    return(
      <div style={{...W,alignItems:"flex-start",flexDirection:"column",padding:0}}>
        <div style={{width:"100%",maxWidth:500,padding:"14px 12px 0",margin:"0 auto",flex:1}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:2}}>
            <button style={BK} onClick={()=>setSc("splash")}>← Salir</button>
            <div style={{display:"flex",gap:7,alignItems:"center"}}>
              <button onClick={()=>setSc("changePin")} style={{background:"none",border:"none",cursor:"pointer",fontSize:15,padding:0}}>🔐</button>
              <span style={{color:"#f0b429",fontWeight:700,fontSize:12}}>👤 {user.name}</span>
            </div>
          </div>
          <div style={{height:3,background:"rgba(255,255,255,0.07)",borderRadius:4,margin:"6px 0 4px"}}>
            <div style={{width:`${(myComp/ALL.length)*100}%`,height:"100%",background:"linear-gradient(90deg,#f0b429,#d97706)",borderRadius:4}}/>
          </div>
          <p style={{color:"#6b7280",fontSize:10,marginBottom:10}}>{myComp}/{ALL.length} pronósticos · 🔒 5 min antes · 👥 ver ajenos post-cierre</p>
          <div style={{display:"flex",gap:4,flexWrap:"wrap",marginBottom:10}}>
            {GL.map(g=><button key={g} onClick={()=>setTab(g)} style={{padding:"4px 9px",borderRadius:20,border:"1px solid",borderColor:tab===g?"#f0b429":"rgba(255,255,255,0.09)",background:tab===g?"#f0b42922":"transparent",color:tab===g?"#f0b429":"#9ca3af",fontWeight:600,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>{g}</button>)}
            <button onClick={()=>setTab("Playoff")} style={{padding:"4px 9px",borderRadius:20,border:"1px solid",borderColor:isP?"#22c55e":"rgba(255,255,255,0.09)",background:isP?"#22c55e22":"transparent",color:isP?"#22c55e":"#9ca3af",fontWeight:600,fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>🏆</button>
          </div>
          {!isP&&(
            <div style={{display:"flex",flexDirection:"column",gap:7}}>
              {PG.filter(p=>p.grupo===tab).map(p=>{
                const others=isLk(p.dt)?Object.fromEntries(otherNames.map(n=>[n,preds[n]?.[p.id]]).filter(([,v])=>v&&v.l!==undefined)):null;
                return <MRow key={p.id} p={p} lv={lp[p.id]?.l||""} vv={lp[p.id]?.v||""} onL={v=>chLp(p.id,"l",v)} onV={v=>chLp(p.id,"v",v)} real={res[p.id]} locked={isLk(p.dt)} adm={false} others={others}/>;
              })}
            </div>
          )}
          {isP&&(
            <div>
              <Tabs items={PO_TABS} sel={ptab} onSel={setPtab}/>
              <div style={{display:"flex",flexDirection:"column",gap:7}}>
                {PO.filter(p=>p.fase===ptab).map(p=>{
                  const c=cl[p.id]||{};
                  const pp={...p,localTeam:c.local||"",visitanteTeam:c.visitante||""};
                  const others=isLk(p.dt)?Object.fromEntries(otherNames.map(n=>[n,preds[n]?.[p.id]]).filter(([,v])=>v&&v.l!==undefined)):null;
                  return <MRow key={p.id} p={pp} lv={lp[p.id]?.l||""} vv={lp[p.id]?.v||""} onL={v=>chLp(p.id,"l",v)} onV={v=>chLp(p.id,"v",v)} real={res[p.id]} locked={isLk(p.dt)} adm={false} others={others}/>;
                })}
              </div>
              <p style={{color:"#6b7280",fontSize:10,marginTop:8}}>Los equipos aparecen cuando el admin carga los clasificados.</p>
            </div>
          )}
          {msg&&<div style={{marginTop:10,color:"#22c55e",fontWeight:700,textAlign:"center",fontSize:13}}>{msg}</div>}
          <div style={{display:"flex",gap:8,marginTop:12}}>
            <Btn v="g" onClick={savePreds} disabled={busy} ch={busy?"Guardando…":"💾 Guardar"}/>
            <Btn v="s" onClick={()=>{loadAll();setSc("tabla");}} ch="🏅 Tabla"/>
          </div>
        </div>
        {CR}
      </div>
    );
  }

  if(sc==="tabla")return(
    <div style={{...W,alignItems:"flex-start",flexDirection:"column",padding:0}}>
      <div style={{width:"100%",maxWidth:500,padding:"14px 12px 0",margin:"0 auto",flex:1}}>
        <button style={BK} onClick={()=>setSc("splash")}>← Inicio</button>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
          <h2 style={{color:"#fff",fontWeight:900,fontSize:20,margin:0}}>🏆 Tabla</h2>
          <span style={{color:"#6b7280",fontSize:10}}>{jugados}/{ALL.length} jugados</span>
        </div>
        <button onClick={loadAll} style={{...BK,color:"#f0b429",marginBottom:10}}>🔄 Actualizar</button>
        {tabla.length===0?<p style={{color:"#6b7280",textAlign:"center",marginTop:20}}>Sin jugadores aún</p>:(
          <div style={{display:"flex",flexDirection:"column",gap:6}}>
            {tabla.map((u,i)=>(
              <div key={u.name} style={{background:i===0?"#f0b42912":"rgba(255,255,255,0.03)",border:`1px solid ${i===0?"#f0b42944":"rgba(255,255,255,0.09)"}`,borderRadius:10,padding:"10px 13px",display:"flex",alignItems:"center",gap:10}}>
                <span style={{fontSize:16,width:22}}>{i===0?"🥇":i===1?"🥈":i===2?"🥉":`${i+1}.`}</span>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,color:"#fff",fontSize:13}}>{u.name}</div>
                  <div style={{fontSize:10,color:"#6b7280"}}>{u.co}/{ALL.length} pronósticos</div>
                </div>
                <div style={{display:"flex",gap:9,alignItems:"center"}}>
                  <span style={{color:"#22c55e",fontSize:10}}>✓{u.ex}</span>
                  <span style={{color:"#f0b429",fontSize:10}}>⚽{u.pa}</span>
                  <span style={{background:i===0?"linear-gradient(135deg,#f0b429,#d97706)":"rgba(255,255,255,0.08)",borderRadius:8,padding:"3px 9px",fontWeight:900,color:"#fff",fontSize:15}}>{u.pts}</span>
                </div>
              </div>
            ))}
          </div>
        )}
        <p style={{marginTop:12,color:"#6b7280",fontSize:10,padding:"8px",background:"rgba(255,255,255,0.02)",borderRadius:8}}>Exacto = 3pts · Ganador/empate correcto = 1pt</p>
      </div>
      {CR}
    </div>
  );

  if(sc==="adminAuth")return(
    <div style={{...W,alignItems:"center",flexDirection:"column",padding:"0 20px"}}>
      <div style={{width:"100%",maxWidth:290}}>
        <button style={BK} onClick={()=>setSc("splash")}>← Volver</button>
        <div style={{textAlign:"center"}}>
          <div style={{fontSize:32,marginBottom:6}}>🕵️</div>
          <h2 style={{color:"#fff",fontWeight:900,fontSize:17,marginBottom:4}}>¿Sos vos, Bachi? 🤨</h2>
          <p style={{color:"#6b7280",fontSize:12,marginBottom:14}}>Demostralo con el PIN secreto…</p>
          <div style={{display:"flex",flexDirection:"column",gap:9}}>
            <input type="password" inputMode="numeric" maxLength={4} autoFocus value={fAdm} onChange={e=>setFAdm(e.target.value.replace(/\D/g,"").slice(0,4))} onKeyDown={e=>{if(e.key==="Enter"){if(fAdm==="2026"){setIsAdm(true);setSc("admin");}else flash("¡Impostor! 😤");}}} placeholder="• • • •" style={{...IN,textAlign:"center",fontSize:22,letterSpacing:8}}/>
            <Btn onClick={()=>{if(fAdm==="2026"){setIsAdm(true);setSc("admin");}else flash("¡Impostor! 😤");}} ch="Soy yo, dejame entrar 🚪" disabled={fAdm.length<4}/>
          </div>
          {msg&&<p style={{color:"#ef4444",marginTop:10,fontSize:13}}>{msg}</p>}
        </div>
      </div>
      {CR}
    </div>
  );

  if(sc==="admin"&&isAdm)return(
    <div style={{...W,alignItems:"flex-start",flexDirection:"column",padding:0}}>
      <div style={{width:"100%",maxWidth:520,padding:"14px 12px 0",margin:"0 auto",flex:1}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
          <button style={BK} onClick={()=>{setIsAdm(false);setSc("splash");}}>← Salir</button>
          <span style={{fontSize:10,fontWeight:700,color:"#ef4444",background:"#ef444422",padding:"2px 8px",borderRadius:10}}>⚙️ ADMIN</span>
        </div>
        <h2 style={{color:"#fff",fontWeight:900,fontSize:18,marginBottom:10}}>Administrador</h2>
        <Tabs items={["usuarios","fase de grupos","playoffs"]} sel={atab} onSel={t=>{setAtab(t);setAddMode(false);setMergeMode(false);}}/>

        {atab==="usuarios"&&(
          <div>
            <div style={{display:"flex",gap:7,marginBottom:10,flexWrap:"wrap"}}>
              <Btn sm ch="+ Agregar" onClick={()=>{setAddMode(m=>!m);setMergeMode(false);}}/>
              <Btn sm v="s" ch="⇄ Fusionar" onClick={()=>{setMergeMode(m=>!m);setAddMode(false);}}/>

            </div>
            {addMode&&(
              <div style={{background:"rgba(255,255,255,0.04)",borderRadius:10,padding:"12px",border:"1px solid #f0b42955",marginBottom:10}}>
                <p style={{color:"#f0b429",fontSize:12,fontWeight:700,margin:"0 0 8px"}}>Nuevo usuario</p>
                <div style={{display:"flex",flexDirection:"column",gap:7}}>
                  <input value={addN} onChange={e=>setAddN(e.target.value)} placeholder="Nombre" style={{...IN,fontSize:13,padding:"7px 10px"}}/>
                  <input value={addP} onChange={e=>setAddP(e.target.value.replace(/\D/g,"").slice(0,4))} placeholder="PIN (4 dígitos)" inputMode="numeric" maxLength={4} style={{...IN,fontSize:13,padding:"7px 10px"}}/>
                  <div style={{display:"flex",gap:7}}>
                    <Btn sm v="g" ch="Crear" onClick={async()=>{
                      const n=addN.trim();
                      if(!n)return flash("Ingresá un nombre");
                      if(users[n])return flash("Ya existe ese usuario");
                      if(addP.length!==4)return flash("PIN de 4 dígitos requerido");
                      const nu={...users,[n]:{name:n,pin:addP,first:false,ts:Date.now()}};
                      await saveU(nu);setAddMode(false);setAddN("");setAddP("2026");flash(`"${n}" creado ✅`);
                    }}/>
                    <Btn sm v="s" ch="Cancelar" onClick={()=>setAddMode(false)}/>
                  </div>
                </div>
              </div>
            )}
            {mergeMode&&(
              <div style={{background:"rgba(255,255,255,0.04)",borderRadius:10,padding:"12px",border:"1px solid #f0b42955",marginBottom:10}}>
                <p style={{color:"#f0b429",fontSize:12,fontWeight:700,margin:"0 0 4px"}}>Fusionar usuarios</p>
                <p style={{color:"#6b7280",fontSize:10,margin:"0 0 8px"}}>Se conserva el nombre del Usuario A. Prioridad a pronósticos más nuevos.</p>
                <div style={{display:"flex",gap:7,marginBottom:7,flexWrap:"wrap"}}>
                  {[["A",mergeA,setMergeA],["B",mergeB,setMergeB]].map(([lbl,val,set])=>(
                    <div key={lbl} style={{flex:1,minWidth:120}}>
                      <p style={{color:"#6b7280",fontSize:10,margin:"0 0 3px"}}>Usuario {lbl}</p>
                      <select value={val} onChange={e=>set(e.target.value)} style={{...IN,fontSize:12,padding:"6px 8px"}}>
                        <option value="">— Elegir —</option>
                        {Object.keys(users).map(n=><option key={n} value={n}>{n}</option>)}
                      </select>
                    </div>
                  ))}
                </div>
                <div style={{display:"flex",gap:7}}>
                  <Btn sm v="g" ch="Fusionar" onClick={doMerge}/>
                  <Btn sm v="s" ch="Cancelar" onClick={()=>setMergeMode(false)}/>
                </div>
              </div>
            )}
            <p style={{color:"#6b7280",fontSize:11,marginBottom:8}}>{Object.keys(users).length} jugador(es)</p>
            <div style={{display:"flex",flexDirection:"column",gap:7}}>
              {Object.values(users).map(u=>(
                <div key={u.name} style={{background:"rgba(255,255,255,0.03)",borderRadius:9,padding:"10px 12px",border:"1px solid rgba(255,255,255,0.09)"}}>
                  {eU===u.name?(
                    <div style={{display:"flex",flexDirection:"column",gap:7}}>
                      <input value={eN} onChange={e=>setEN(e.target.value)} placeholder="Nombre" style={{...IN,fontSize:13,padding:"7px 10px"}}/>
                      <input value={eP} onChange={e=>setEP(e.target.value.replace(/\D/g,"").slice(0,4))} placeholder="Nuevo PIN (vacío=no cambia)" inputMode="numeric" maxLength={4} style={{...IN,fontSize:13,padding:"7px 10px"}}/>
                      <div style={{display:"flex",gap:7}}>
                        <Btn sm v="g" ch="Guardar" onClick={async()=>{
                          const old=u.name,nn=eN.trim()||old,np=eP.length===4?eP:u.pin;
                          const nu={...users};delete nu[old];nu[nn]={...u,name:nn,pin:np};
                          if(old!==nn){const np2={...preds};if(np2[old]){np2[nn]=np2[old];delete np2[old];}await ST.set("predicciones",np2);setPreds(np2);}
                          await saveU(nu);setEU(null);flash("✅ Guardado");
                        }}/>
                        <Btn sm v="s" ch="Cancelar" onClick={()=>setEU(null)}/>
                      </div>
                    </div>
                  ):(
                    <div style={{display:"flex",alignItems:"center",gap:9}}>
                      <div style={{flex:1}}>
                        <div style={{fontWeight:700,color:"#fff",fontSize:13}}>{u.name}</div>
                        <div style={{fontSize:10,color:"#6b7280"}}>PIN: {u.pin}</div>
                      </div>
                      <Btn sm v="s" ch="✏️" onClick={()=>{setEU(u.name);setEN(u.name);setEP("");}}/>
                      <Btn sm v="r" ch="🗑️" onClick={()=>doDelete(u.name)}/>
                    </div>
                  )}
                </div>
              ))}
            </div>
            {msg&&<div style={{marginTop:8,color:"#22c55e",fontWeight:700,textAlign:"center",fontSize:13}}>{msg}</div>}
          </div>
        )}

        {atab==="fase de grupos"&&(
          <div>
            <div style={{marginBottom:10,display:"flex",gap:7,flexWrap:"wrap"}}>
              <Btn v="g" onClick={fetchFifa} disabled={busy} ch={busy?"⏳ Buscando…":"🌐 Auto-actualizar FIFA"}/>

            </div>
            <Tabs items={GL} sel={tab} onSel={setTab}/>
            <div style={{display:"flex",flexDirection:"column",gap:7}}>
              {PG.filter(p=>p.grupo===tab).map(p=>(
                <MRow key={p.id} p={p} lv={lr[p.id]?.l||""} vv={lr[p.id]?.v||""} onL={v=>chLr(p.id,"l",v)} onV={v=>chLr(p.id,"v",v)} real={null} locked={false} adm={true}/>
              ))}
            </div>
            {msg&&<div style={{marginTop:8,color:"#22c55e",fontWeight:700,textAlign:"center",fontSize:13}}>{msg}</div>}
            <div style={{marginTop:10}}><Btn v="g" ch={busy?"Guardando…":"💾 Guardar Resultados"} disabled={busy} onClick={saveRes}/></div>
          </div>
        )}

        {atab==="playoffs"&&(
          <div>
            <div style={{marginBottom:10}}><Btn v="g" onClick={fetchFifa} disabled={busy} ch={busy?"⏳ Buscando…":"🌐 Auto-actualizar FIFA"}/></div>
            <Tabs items={PO_TABS} sel={ptab} onSel={setPtab}/>
            <div style={{display:"flex",flexDirection:"column",gap:9}}>
              {PO.filter(p=>p.fase===ptab).map(p=>(
                <div key={p.id} style={{background:"rgba(255,255,255,0.03)",borderRadius:9,padding:"10px 12px",border:"1px solid rgba(255,255,255,0.09)"}}>
                  <p style={{color:"#f0b429",fontSize:11,fontWeight:700,margin:"0 0 4px"}}>{p.label} <span style={{color:"#6b7280",fontWeight:400,fontSize:10}}>· {p.estadio}, {p.ciudad}</span></p>
                  <div style={{display:"flex",gap:7,marginBottom:7,flexWrap:"wrap"}}>
                    {["local","visitante"].map(t=>(
                      <div key={t} style={{flex:1,minWidth:120}}>
                        <p style={{color:"#6b7280",fontSize:10,margin:"0 0 3px"}}>{t==="local"?`Local (${p.lD})`:`Visitante (${p.vD})`}</p>
                        <select value={lcl[p.id]?.[t]||""} onChange={e=>chLcl(p.id,t,e.target.value)} style={{...IN,fontSize:12,padding:"6px 8px"}}>
                          <option value="">— Por definir —</option>
                          {TEAMS.map(eq=><option key={eq} value={eq}>{FL[eq]||""} {eq}</option>)}
                        </select>
                      </div>
                    ))}
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:6}}>
                    <span style={{fontSize:11,color:"#9ca3af",flex:1,textAlign:"right"}}>{lcl[p.id]?.local||p.lD}</span>
                    <input inputMode="numeric" type="text" value={lr[p.id]?.l||""} onChange={e=>chLr(p.id,"l",e.target.value.replace(/\D/g,"").slice(0,2))} style={SC} placeholder="?"/>
                    <span style={{color:"#6b7280",fontWeight:700}}>-</span>
                    <input inputMode="numeric" type="text" value={lr[p.id]?.v||""} onChange={e=>chLr(p.id,"v",e.target.value.replace(/\D/g,"").slice(0,2))} style={SC} placeholder="?"/>
                    <span style={{fontSize:11,color:"#9ca3af",flex:1}}>{lcl[p.id]?.visitante||p.vD}</span>
                  </div>
                </div>
              ))}
            </div>
            {msg&&<div style={{marginTop:8,color:"#22c55e",fontWeight:700,textAlign:"center",fontSize:13}}>{msg}</div>}
            <div style={{display:"flex",gap:8,marginTop:10,flexWrap:"wrap"}}>
              <Btn v="g" ch={busy?"Guardando…":"💾 Guardar Clasificados"} disabled={busy} onClick={saveCl}/>
              <Btn v="s" ch="💾 Guardar Resultados" disabled={busy} onClick={saveRes}/>
            </div>
          </div>
        )}
      </div>
      {/* Botón Excel global - siempre visible en admin */}
      <div style={{width:"100%",maxWidth:520,margin:"16px auto 0",padding:"0 12px 32px"}}>
        <div style={{borderTop:"1px solid rgba(255,255,255,0.09)",paddingTop:14,display:"flex",justifyContent:"center"}}>
          <Btn v="s" ch="⬇ Exportar toda la base a Excel" onClick={async()=>{const ok=await exportXLSX(users,preds,res,cl);if(!ok)flash("Error exportando");}}/>
        </div>
      </div>
      {CR}
    </div>
  );

  return <div style={{...W,alignItems:"center"}}><p style={{color:"#6b7280"}}>Cargando…</p></div>;
}

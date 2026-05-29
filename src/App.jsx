import { useState, useEffect } from "react";
import * as XLSX from 'xlsx';

const MESI_NOMI = ['GENNAIO','FEBBRAIO','MARZO','APRILE','MAGGIO','GIUGNO','LUGLIO','AGOSTO','SETTEMBRE','OTTOBRE','NOVEMBRE','DICEMBRE'];
const MESI_BREVI = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];
const DEFAULT_SETTINGS = { batteria: 64, prezzo: 0.50, targa: 'MGS5', p1a: 20, p1b: 30, p2a: 80, p2b: 100 };
const DEFAULT_SHEETS_URL = 'https://script.google.com/macros/s/AKfycbyb70PcCzlVoACCdodeyJ2vmhg9v1_NaEAIlkj_jpBvE4RuFK0kO5d24zsKkAC2Ccwc/exec';
function getSheetsUrl() { return storage.get('mgs5_sheetsUrl') || DEFAULT_SHEETS_URL; }

const memStore = {};
const storage = {
  get(key) { try { return localStorage.getItem(key); } catch { return memStore[key] ?? null; } },
  set(key, val) { try { localStorage.setItem(key, val); } catch { memStore[key] = val; } }
};

async function syncToSheets(ricariche) {
  try {
    // Ordina per km decrescente: km più alto → riga 2 del foglio
    const ordinate = [...ricariche].sort((a, b) => (b.km || 0) - (a.km || 0));
    const ricaricheSheets = ordinate.map(r => ({
      ...r,
      pctPrima: r.pctPrima > 1 ? r.pctPrima / 100 : r.pctPrima,
      pctDopo:  r.pctDopo  > 1 ? r.pctDopo  / 100 : r.pctDopo,
    }));
    await fetch(getSheetsUrl() + '?action=write', {
      method: 'POST',
      body: JSON.stringify(ricaricheSheets)
    });
  } catch(e) { console.error('Sync to Sheets failed:', e); }
}

async function syncFromSheets() {
  try {
    const res = await fetch(getSheetsUrl() + '?action=read');
    const json = await res.json();
    return json.data || [];
  } catch(e) { console.error('Sync from Sheets failed:', e); return null; }
}

function today() { return new Date().toISOString().split('T')[0]; }
function toNum(x, def = 0) { const n = Number(x); return Number.isFinite(n) ? n : def; }

function normalizzaData(d) {
  if (!d) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(String(d))) return String(d).slice(0,10);
  const dt = new Date(d);
  if (!isNaN(dt)) return dt.getFullYear()+'-'+String(dt.getMonth()+1).padStart(2,'0')+'-'+String(dt.getDate()).padStart(2,'0');
  return String(d).slice(0,10);
}

function sanitizzaRicarica(r) {
  const pctPrima  = toNum(r?.pctPrima ?? r?.percPrima, 0);
  const pctDopo   = toNum(r?.pctDopo  ?? r?.percDopo,  0);
  const kwhEff    = toNum(r?.kwhEff,  0);
  const prezzoKwh = toNum(r?.prezzoKwh ?? r?.prezzoKWh ?? r?.prezzo, 0);
  const costoRaw  = toNum(r?.costo ?? r?.costoEuro ?? r?.euro, 0);
  const costoCalc = kwhEff * prezzoKwh;
  const costo     = costoCalc > 0 ? costoCalc : (costoRaw > 0 ? costoRaw : 0);
  const luogo = (() => {
    const l = (r?.luogo || r?.dove || r?.location || '').toString().trim().toUpperCase();
    if (l) return l;
    if (Math.abs(prezzoKwh - 0.20) < 0.001) return 'CASA';
    return '';
  })();
  return {
    ...r, luogo, data: normalizzaData(r?.data), kwhEff, kwhTeor: toNum(r?.kwhTeor, 0),
    costo, prezzoKwh, pctPrima, pctDopo,
    kwh100:     r?.kwh100     == null ? null : (Number.isFinite(Number(r.kwh100))     ? Number(r.kwh100)     : null),
    km:         r?.km         == null ? null : (Number.isFinite(Number(r.km))         ? Number(r.km)         : null),
    kmParziali: r?.kmParziali == null ? null : (Number.isFinite(Number(r.kmParziali)) ? Number(r.kmParziali) : null),
  };
}

function useToast() {
  const [toast, setToast] = useState(null);
  const show = (msg, color = '#10b981') => {
    setToast({ msg, color });
    setTimeout(() => setToast(null), 2500);
  };
  return { toast, show };
}

function ricalcolaKmParziali(lista) {
  return lista.map((r, i) => {
    if (r.km && r.km > 0) {
      const prev = [...lista].slice(0, i).filter(x => x.km && x.km > 0).pop();
      const kmParziali = prev ? r.km - prev.km : r.kmParziali;
      const kwh100 = kmParziali && kmParziali > 0 ? (r.kwhEff / kmParziali) * 100 : r.kwh100;
      return { ...r, kmParziali: kmParziali || r.kmParziali, kwh100 };
    }
    return r;
  });
}

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('/sw.js').catch(() => {});
  });
}

function BarChart({ data, color = '#00e5ff' }) {
  if (!data.length) return <div style={{ height:160, display:'flex', alignItems:'center', justifyContent:'center', color:'#94a3b8', fontSize:'0.75rem' }}>Nessun dato</div>;
  const max = Math.max(...data.map(d => d.y));
  return (
    <div style={{ display:'flex', alignItems:'flex-end', gap:6, height:160, padding:'8px 0' }}>
      {data.map((d,i) => (
        <div key={i} style={{ flex:1, display:'flex', flexDirection:'column', alignItems:'center', gap:4, height:'100%', justifyContent:'flex-end' }}>
          <div style={{ width:'100%', background:color+'99', borderRadius:4, height:max?`${(d.y/max)*120}px`:0, minHeight:d.y?4:0, transition:'height 0.4s' }}/>
          <div style={{ fontSize:'0.5rem', color:'#94a3b8', textAlign:'center', whiteSpace:'nowrap', overflow:'hidden', maxWidth:'100%' }}>{d.x}</div>
        </div>
      ))}
    </div>
  );
}

function LineChart({ data, color = '#7c3aed' }) {
  if (data.length < 2) return <div style={{ height:160, display:'flex', alignItems:'center', justifyContent:'center', color:'#94a3b8', fontSize:'0.75rem' }}>Servono almeno 2 punti</div>;
  const w=300, h=120;
  const ys=data.map(d=>d.y), minY=Math.min(...ys), maxY=Math.max(...ys);
  const px=i=>(i/(data.length-1))*w;
  const py=y=>h-((y-minY)/(maxY-minY||1))*(h-10)-5;
  const pts=data.map((d,i)=>`${px(i)},${py(d.y)}`).join(' ');
  return (
    <svg viewBox={`0 0 ${w} ${h}`} style={{ width:'100%', height:160 }}>
      <defs>
        <linearGradient id={`g${color.replace('#','')}`} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor={color} stopOpacity="0.3"/>
          <stop offset="100%" stopColor={color} stopOpacity="0"/>
        </linearGradient>
      </defs>
      <polygon points={`0,${h} ${pts} ${w},${h}`} fill={`url(#g${color.replace('#','')})`}/>
      <polyline points={pts} fill="none" stroke={color} strokeWidth="2" strokeLinejoin="round"/>
      {data.map((d,i)=><circle key={i} cx={px(i)} cy={py(d.y)} r="3" fill={color}/>)}
    </svg>
  );
}

const S = {
  bg:'#0a0f1e', bg2:'rgba(255,255,255,0.04)',
  border:'rgba(0,229,255,0.15)', accent:'#00e5ff',
  accent2:'#7c3aed', green:'#10b981', text:'#e2e8f0', text2:'#94a3b8'
};
const inputSt = {
  background:'rgba(255,255,255,0.06)', border:`1px solid ${S.border}`, borderRadius:10,
  color:S.text, fontFamily:'monospace', fontSize:'1rem', padding:'10px 12px',
  width:'100%', outline:'none', boxSizing:'border-box'
};
const chipSt = {
  padding:'4px 10px', background:'rgba(0,229,255,0.08)',
  border:'1px solid rgba(0,229,255,0.2)', borderRadius:20,
  color:S.accent, fontFamily:'monospace', fontSize:'0.75rem', cursor:'pointer'
};

function ConfirmDialog({ messaggio, onConfirm, onCancel }) {
  return (
    <div style={{ position:'fixed', inset:0, zIndex:300, background:'rgba(0,0,0,0.7)', backdropFilter:'blur(4px)', display:'flex', alignItems:'center', justifyContent:'center', padding:24 }}>
      <div style={{ background:'#111827', border:'1px solid rgba(239,68,68,0.4)', borderRadius:20, padding:28, maxWidth:320, width:'100%', textAlign:'center' }}>
        <div style={{ fontSize:'2rem', marginBottom:12 }}>🗑</div>
        <div style={{ fontSize:'1rem', fontWeight:700, marginBottom:8 }}>{messaggio}</div>
        <div style={{ fontSize:'0.8rem', color:S.text2, marginBottom:24 }}>Questa azione non può essere annullata.</div>
        <div style={{ display:'flex', gap:10 }}>
          <button onClick={onCancel} style={{ flex:1, padding:'12px 0', background:'transparent', border:`1px solid ${S.border}`, borderRadius:12, color:S.text2, fontSize:'0.9rem', cursor:'pointer' }}>Annulla</button>
          <button onClick={onConfirm} style={{ flex:1, padding:'12px 0', background:'rgba(239,68,68,0.15)', border:'1px solid rgba(239,68,68,0.5)', borderRadius:12, color:'#ef4444', fontSize:'0.9rem', fontWeight:700, cursor:'pointer' }}>Conferma</button>
        </div>
      </div>
    </div>
  );
}

function ImportModal({ onImportJSON, onImportFoglio, onClose }) {
  return (
    <div style={{ position:'fixed', inset:0, zIndex:300, background:'rgba(0,0,0,0.75)', backdropFilter:'blur(6px)', display:'flex', alignItems:'flex-end', justifyContent:'center', padding:0 }}>
      <div style={{ background:'#111827', border:`1px solid ${S.border}`, borderTopLeftRadius:24, borderTopRightRadius:24, padding:28, width:'100%', maxWidth:480 }}>
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:20 }}>
          <div>
            <div style={{ fontSize:'1.1rem', fontWeight:800, letterSpacing:'0.04em' }}>📥 Importa dati</div>
            <div style={{ fontSize:'0.7rem', color:S.text2, marginTop:2 }}>Scegli il formato del file da caricare</div>
          </div>
          <button onClick={onClose} style={{ background:'rgba(255,255,255,0.08)', border:'none', borderRadius:20, color:S.text2, width:32, height:32, fontSize:'1rem', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center' }}>✕</button>
        </div>
        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:12, marginBottom:16 }}>
          <button onClick={() => { const inp = document.createElement('input'); inp.type='file'; inp.accept='.json,application/json'; inp.onchange=onImportJSON; inp.click(); onClose(); }}
            style={{ background:'rgba(124,58,237,0.1)', border:'1px solid rgba(124,58,237,0.35)', borderRadius:16, padding:'20px 12px', cursor:'pointer', textAlign:'center', color:S.text }}>
            <div style={{ fontSize:'2.2rem', marginBottom:8 }}>💾</div>
            <div style={{ fontSize:'0.8rem', fontWeight:800, textTransform:'uppercase', color:'#a78bfa', letterSpacing:'0.06em' }}>Backup JSON</div>
            <div style={{ fontSize:'0.62rem', color:S.text2, marginTop:6, lineHeight:1.5 }}>Ripristina un backup<br/>completo precedente</div>
          </button>
          <button onClick={() => { const inp = document.createElement('input'); inp.type='file'; inp.accept='.xlsx,.xls,.csv,.ods'; inp.onchange=onImportFoglio; inp.click(); onClose(); }}
            style={{ background:'rgba(0,229,255,0.08)', border:'1px solid rgba(0,229,255,0.3)', borderRadius:16, padding:'20px 12px', cursor:'pointer', textAlign:'center', color:S.text }}>
            <div style={{ fontSize:'2.2rem', marginBottom:8 }}>📊</div>
            <div style={{ fontSize:'0.8rem', fontWeight:800, textTransform:'uppercase', color:S.accent, letterSpacing:'0.06em' }}>Excel / Sheets</div>
            <div style={{ fontSize:'0.62rem', color:S.text2, marginTop:6, lineHeight:1.5 }}>xlsx · xls · csv · ods<br/>Google Sheets export</div>
          </button>
        </div>
        <button onClick={onClose} style={{ width:'100%', padding:'13px 0', background:'transparent', border:`1px solid rgba(255,255,255,0.1)`, borderRadius:12, color:S.text2, fontSize:'0.85rem', cursor:'pointer' }}>Annulla</button>
      </div>
    </div>
  );
}

function FormRicarica({ settings, ricariche, editIdx, onSave, onCancel, showToast }) {
  const ex = editIdx !== null ? ricariche[editIdx] : null;
  const [fData,   setFData]   = useState(ex ? ex.data : today());
  const [fLuogo,  setFLuogo]  = useState(ex ? (ex.luogo||'') : '');
  const [fKm,     setFKm]     = useState(ex && ex.km !== null ? String(ex.km) : '');
  const [fKmParz, setFKmParz] = useState(ex && ex.kmParziali !== null ? String(ex.kmParziali) : '');
  const [fPrima,  setFPrima]  = useState(ex ? String(ex.pctPrima) : '');
  const [fDopo,   setFDopo]   = useState(ex ? String(ex.pctDopo) : '');
  const [fKwh,    setFKwh]    = useState(ex ? String(ex.kwhEff) : '');
  const [fPrezzo, setFPrezzo] = useState(ex ? String(ex.prezzoKwh) : String(settings.prezzo));

  const prima=parseFloat(fPrima)||0, dopo=parseFloat(fDopo)||0;
  const diff=dopo-prima, kwhTeor=settings.batteria*(diff/100);
  const kwhEff=parseFloat(fKwh)||kwhTeor;
  const prezzo=parseFloat(fPrezzo)||settings.prezzo;
  const costo=kwhEff*prezzo;
  const kmParzN=parseFloat(fKmParz);
  const kmN=parseFloat(fKm);

  let kmParzAuto=null;
  if (!isNaN(kmN)&&kmN>0&&!(kmParzN>0)) {
    const altri=editIdx!==null?ricariche.filter((_,i)=>i!==editIdx):ricariche;
    const prev=[...altri].filter(r=>r.km&&r.km<kmN).sort((a,b)=>b.km-a.km)[0];
    if (prev) kmParzAuto=kmN-prev.km;
  }
  const kmParzEff=(kmParzN>0)?kmParzN:kmParzAuto;
  const kwh100=kmParzEff?(kwhEff/kmParzEff)*100:null;

  function salva() {
    if (!fData||prima<=0||dopo<=0){showToast('⚠️ Compila data e %','#f59e0b');return;}
    if (dopo<=prima){showToast('⚠️ % dopo > % prima','#f59e0b');return;}
    onSave({ data:fData, luogo:fLuogo||null, km:parseFloat(fKm)||null,
      pctPrima:prima, pctDopo:dopo, kwhEff, kwhTeor, prezzoKwh:prezzo, costo,
      kmParziali:kmParzEff||null, kwh100:kmParzEff?(kwhEff/kmParzEff)*100:null });
  }

  return (
    <div style={{ background:S.bg2, border:`1px solid ${S.border}`, borderRadius:16, padding:16, marginBottom:12 }}>
      <div style={{ fontSize:'0.65rem', textTransform:'uppercase', letterSpacing:'0.12em', color:S.text2, marginBottom:12 }}>
        {editIdx!==null?'✏️ Modifica Ricarica':'Nuova Ricarica'}
      </div>
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:10 }}>
        <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
          <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>Data</label>
          <input type="date" value={fData} onChange={e=>setFData(e.target.value)} style={inputSt}/>
        </div>
        <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
          <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>Dove si ricarica</label>
          <input type="text" value={fLuogo} onChange={e=>setFLuogo(e.target.value.toUpperCase())} placeholder="CASA..." style={inputSt}/>
        </div>
        <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
          <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>KM Totali</label>
          <input type="number" value={fKm} onChange={e=>setFKm(e.target.value)} placeholder="0" style={inputSt}/>
        </div>
        <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
          <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>KM Parziali</label>
          <input type="number" value={fKmParz} onChange={e=>setFKmParz(e.target.value)} placeholder={kmParzAuto?String(kmParzAuto):''} style={inputSt}/>
          {kmParzAuto&&!(kmParzN>0)&&<div style={{ fontSize:'0.6rem', color:S.accent, marginTop:2 }}>⚡ calcolato automaticamente</div>}
        </div>
        <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
          <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>% Prima</label>
          <input type="number" value={fPrima} onChange={e=>setFPrima(e.target.value)} placeholder="0" style={inputSt}/>
          <div style={{ display:'flex', gap:6 }}>
            {[settings.p1a,settings.p1b].map(v=><button key={v} onClick={()=>setFPrima(String(v))} style={chipSt}>{v}%</button>)}
          </div>
        </div>
        <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
          <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>% Dopo</label>
          <input type="number" value={fDopo} onChange={e=>setFDopo(e.target.value)} placeholder="100" style={inputSt}/>
          <div style={{ display:'flex', gap:6 }}>
            {[settings.p2a,settings.p2b].map(v=><button key={v} onClick={()=>setFDopo(String(v))} style={chipSt}>{v}%</button>)}
          </div>
        </div>
        <div style={{ gridColumn:'1/-1', background:'rgba(0,229,255,0.04)', border:`1px solid ${S.border}`, borderRadius:10, padding:10, display:'flex', justifyContent:'space-between' }}>
          <div>
            <div style={{ fontSize:'0.65rem', color:S.text2, textTransform:'uppercase' }}>Differenza %</div>
            <div style={{ fontFamily:'monospace', fontSize:'1rem', color:S.accent }}>{diff>0?diff.toFixed(0)+' %':'—'}</div>
          </div>
          <div style={{ textAlign:'right' }}>
            <div style={{ fontSize:'0.65rem', color:S.text2, textTransform:'uppercase' }}>kWh teorici</div>
            <div style={{ fontFamily:'monospace', fontSize:'1rem', color:S.accent }}>{kwhTeor>0?kwhTeor.toFixed(2)+' kWh':'—'}</div>
          </div>
        </div>
        <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
          <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>Prezzo €/kWh</label>
          <input type="number" value={fPrezzo} onChange={e=>setFPrezzo(e.target.value)} step="0.001" style={inputSt}/>
          <div style={{ display:'flex', gap:6 }}>
            <button onClick={()=>{setFPrezzo('0.2');setFLuogo('CASA');}} style={chipSt}>€0.20</button>
            <button onClick={()=>setFPrezzo('0.5')} style={chipSt}>€0.50</button>
          </div>
        </div>
        <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
          <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>Costo €</label>
          <div style={{ fontFamily:'monospace', fontSize:'1.2rem', color:S.green, padding:'10px 12px', background:'rgba(255,255,255,0.06)', border:`1px solid ${S.border}`, borderRadius:10 }}>
            {costo>0?'€ '+costo.toFixed(2):'—'}
          </div>
        </div>
        <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
          <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>kWh effettivi</label>
          <input type="number" value={fKwh} onChange={e=>setFKwh(e.target.value)} placeholder="dal display" step="0.01" style={inputSt}/>
        </div>
        <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
          <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>kWh / 100km</label>
          <div style={{ fontFamily:'monospace', fontSize:'1.2rem', color:S.accent2, padding:'10px 12px', background:'rgba(255,255,255,0.06)', border:`1px solid ${S.border}`, borderRadius:10 }}>
            {kwh100?kwh100.toFixed(2):'—'}
          </div>
        </div>
      </div>
      <div style={{ display:'flex', gap:8, marginTop:12 }}>
        {onCancel&&<button onClick={onCancel} style={{ flex:1, padding:16, background:'transparent', border:`1px solid ${S.border}`, borderRadius:12, color:S.text2, fontSize:'0.85rem', cursor:'pointer' }}>Annulla</button>}
        <button onClick={salva} style={{ flex:2, padding:16, background:`linear-gradient(135deg,${S.accent},${S.accent2})`, border:'none', borderRadius:12, color:'#fff', fontSize:'0.9rem', fontWeight:800, textTransform:'uppercase', cursor:'pointer' }}>
          {editIdx!==null?'✏️ Aggiorna':'⚡ Salva Ricarica'}
        </button>
      </div>
    </div>
  );
}

export default function App() {
  const [view,       setView]       = useState('home');
  const [ricariche,  setRicariche]  = useState(() => {
    try { const s = storage.get('mgs5_ricariche'); return s ? JSON.parse(s).map(sanitizzaRicarica) : []; } catch { return []; }
  });
  const [settings,   setSettings]   = useState(() => {
    try { const s = storage.get('mgs5_settings'); return s ? {...DEFAULT_SETTINGS,...JSON.parse(s)} : {...DEFAULT_SETTINGS}; } catch { return {...DEFAULT_SETTINGS}; }
  });
  const [editIdx,    setEditIdx]    = useState(null);
  const [confirmIdx, setConfirmIdx] = useState(null);
  const [showImport, setShowImport] = useState(false);
  const [showExport, setShowExport] = useState(false);
  const [mesiAperti, setMesiAperti] = useState({});
  const { toast, show: showToast }  = useToast();
  const [syncing,    setSyncing]    = useState(false);

  useEffect(() => { try { storage.set('mgs5_ricariche', JSON.stringify(ricariche)); } catch {} }, [ricariche]);
  useEffect(() => { try { storage.set('mgs5_settings', JSON.stringify(settings)); } catch {} }, [settings]);

  useEffect(() => {
    async function autoSync() {
      setSyncing(true);
      const data = await syncFromSheets();
      setSyncing(false);
      if (data && data.length > 0) {
        setRicariche(ricalcolaKmParziali(data.sort((a,b)=>(a.km||0)-(b.km||0)).map(sanitizzaRicarica)));
      }
    }
    autoSync();
  }, []);

  async function sincronizzaVersoSheet() {
    setSyncing(true); await syncToSheets(ricariche); setSyncing(false);
    showToast('☁️ Dati inviati a Google Sheets!');
  }

  async function sincronizzaDaSheet() {
    setSyncing(true); const data = await syncFromSheets(); setSyncing(false);
    if (data && data.length > 0) {
      setRicariche(ricalcolaKmParziali(data.sort((a,b)=>(a.km||0)-(b.km||0)).map(sanitizzaRicarica)));
      showToast('☁️ Dati importati da Google Sheets!'); setView('home');
    } else { showToast('❌ Nessun dato trovato', '#ef4444'); }
  }

  function salvaRicarica(record) {
    const safe = sanitizzaRicarica(record);
    let nuove = editIdx!==null ? ricariche.map((r,i)=>i===editIdx?safe:r) : [...ricariche,safe];
    nuove = ricalcolaKmParziali(nuove.sort((a,b)=>(a.km||0)-(b.km||0)));
    setRicariche(nuove); syncToSheets(nuove);
    showToast(editIdx!==null?'✅ Aggiornata e salvata!':'✅ Salvata!');
    setEditIdx(null); setView('home');
  }

  function cancellaRicarica(idx) {
    const nuove = ricalcolaKmParziali(ricariche.filter((_,i)=>i!==idx));
    setRicariche(nuove); syncToSheets(nuove); setConfirmIdx(null); showToast('🗑 Eliminata');
  }

  function apriModifica(idx) { setEditIdx(idx); setView('add'); }
  function annullaModifica() { setEditIdx(null); setView('home'); }

  function esportaCSV() {
    const header='Data,KM Totali,KM Parziali,% Prima,% Dopo,kWh Effettivi,Prezzo Euro/kWh,Costo Euro,kWh/100km,Luogo\n';
    const rows=ricariche.map(r=>[r.data,r.km||'',r.kmParziali||'',r.pctPrima,r.pctDopo,(r.kwhEff||0).toFixed(3),r.prezzoKwh,(r.costo||0).toFixed(3),r.kwh100?(r.kwh100||0).toFixed(2):'',r.luogo||''].join(',')).join('\n');
    scarica('ricariche.csv','text/csv',header+rows); showToast('📊 CSV esportato');
  }

  function esportaJSON() {
    scarica('ricariche_backup.json','application/json',JSON.stringify({ricariche,settings,esportato:new Date().toISOString()},null,2));
    showToast('💾 Backup salvato');
  }

  function importaJSON(e) {
    const file=e.target.files[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=(ev)=>{
      try {
        const data=JSON.parse(ev.target.result);
        if (data.ricariche) {
          setRicariche(ricalcolaKmParziali(data.ricariche.sort((a,b)=>(a.km||0)-(b.km||0)).map(sanitizzaRicarica)));
          if (data.settings) setSettings({...DEFAULT_SETTINGS,...data.settings});
          showToast('✅ Dati importati!'); setView('home');
        }
      } catch { showToast('❌ File non valido','#ef4444'); }
    };
    reader.readAsText(file); e.target.value='';
  }

  function importaFoglio(e) {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type:'array', cellDates:true });
        const sheetName = wb.SheetNames.find(n=>n.toLowerCase().trim()==='dati')||wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
        if (rows.length < 2) { showToast('❌ File vuoto o non valido','#ef4444'); return; }
        let headerIdx = 0;
        for (let i=0;i<Math.min(5,rows.length);i++) {
          const r=rows[i].map(c=>String(c).toLowerCase());
          if (r.some(c=>c.includes('kwh')||c.includes('data')||c.includes('prima')||c.includes('dopo'))) { headerIdx=i; break; }
        }
        const headers = rows[headerIdx].map(c=>String(c).toLowerCase().trim());
        const col = (kws) => { const i=headers.findIndex(h=>kws.some(k=>h.includes(k))); return i>=0?i:-1; };
        const iData=col(['data','date','giorno']), iKm=col(['km tot','km_tot','totali','odometro']),
              iKmParz=col(['km parz','km_parz','parziali','percorsi']), iPrima=col(['prima','start','inizio','% prima','%prima']),
              iDopo=col(['dopo','end','fine','% dopo','%dopo']), iKwh=col(['kwh eff','kwh_eff','effettivi','kwh ricaricati','kwh']),
              iPrezzo=col(['prezzo','euro/kwh','€/kwh','price','tariffa']),
              iCosto=col(['costo €','costo euro','cost€','€ tot','costo','cost','spesa']),
              iKwh100=col(['kwh/100','kwh100','consumo','efficienza']), iLuogo=col(['luogo','dove','location','posto','stazione']);

        function parseData(val) {
          if (!val) return null;
          if (val instanceof Date) return val.getFullYear()+'-'+String(val.getMonth()+1).padStart(2,'0')+'-'+String(val.getDate()).padStart(2,'0');
          const s=String(val).trim();
          if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) { const [gg,mm,aaaa]=s.split('/'); return `${aaaa}-${mm.padStart(2,'0')}-${gg.padStart(2,'0')}`; }
          if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
          const n=parseFloat(s);
          if (!isNaN(n)&&n>40000) return new Date(Math.round((n-25569)*86400*1000)).toISOString().split('T')[0];
          return null;
        }

        const nuove=[];
        for (let i=headerIdx+1;i<rows.length;i++) {
          const row=rows[i];
          if (row.every(c=>c===''||c===null||c===undefined)) continue;
          const data=parseData(iData>=0?row[iData]:null);
          const kwhEff=parseFloat(iKwh>=0?row[iKwh]:0)||0;
          if (!data||!kwhEff) continue;
          let pctPrima=iPrima>=0?parseFloat(row[iPrima])||0:0;
          let pctDopo=iDopo>=0?parseFloat(row[iDopo])||0:0;
          if (pctPrima>0&&pctPrima<=1) pctPrima*=100;
          if (pctDopo>0&&pctDopo<=1) pctDopo*=100;
          nuove.push({ data, kwhEff, kwhTeor:0,
            luogo: iLuogo>=0?String(row[iLuogo]||'').toUpperCase()||null:null,
            km: iKm>=0?parseFloat(row[iKm])||null:null,
            kmParziali: iKmParz>=0?parseFloat(row[iKmParz])||null:null,
            pctPrima, pctDopo,
            prezzoKwh: iPrezzo>=0?parseFloat(row[iPrezzo])||0:0,
            costo: iCosto>=0?parseFloat(row[iCosto])||0:0,
            kwh100: iKwh100>=0?parseFloat(row[iKwh100])||null:null,
          });
        }
        if (!nuove.length) { showToast('❌ Nessun dato riconosciuto','#ef4444'); return; }
        const ordinate=ricalcolaKmParziali(nuove.sort((a,b)=>(a.km||0)-(b.km||0)).map(sanitizzaRicarica));
        setRicariche(ordinate); showToast(`✅ Importate ${ordinate.length} ricariche!`); setView('home');
      } catch(err) { console.error(err); showToast('❌ Errore nel file','#ef4444'); }
    };
    reader.readAsArrayBuffer(file); e.target.value='';
  }

  function scarica(nome, tipo, contenuto) {
    try {
      const a=document.createElement('a');
      a.href='data:'+tipo+';charset=utf-8,'+encodeURIComponent(contenuto);
      a.download=nome; document.body.appendChild(a); a.click(); document.body.removeChild(a);
    } catch { showToast('❌ Errore esportazione','#ef4444'); }
  }

  const costoTot=ricariche.reduce((s,r)=>s+r.costo,0);
  const kmMax=ricariche.length?Math.max(...ricariche.map(r=>r.km||0)):null;

  const perMese={};
  ricariche.forEach((r,i)=>{
    const d=new Date(r.data), key=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');
    if(!perMese[key]) perMese[key]=[];
    perMese[key].push({...r,idx:i});
  });
  Object.keys(perMese).forEach(k=>perMese[k].sort((a,b)=>(b.km||0)-(a.km||0)));

  const perMeseChart={};
  ricariche.forEach(r=>{
    const d=new Date(r.data), k=MESI_BREVI[d.getMonth()]+'\''+String(d.getFullYear()).slice(2);
    if(!perMeseChart[k]) perMeseChart[k]={costo:0};
    perMeseChart[k].costo+=r.costo;
  });
  const chartMensile=Object.entries(perMeseChart).map(([x,v])=>({x,y:v.costo}));
  const chartEff=ricariche.filter(r=>r.kwh100).map(r=>{const d=new Date(r.data);return{x:String(d.getDate()).padStart(2,'0')+'/'+(d.getMonth()+1),y:r.kwh100};});
  const chartPrezzo=ricariche.filter(r=>r.prezzoKwh>0).map(r=>{const d=new Date(r.data);return{x:String(d.getDate()).padStart(2,'0')+'/'+(d.getMonth()+1),y:r.prezzoKwh};});

  const Card=({children,style={}})=><div style={{background:S.bg2,border:`1px solid ${S.border}`,borderRadius:16,padding:16,marginBottom:12,...style}}>{children}</div>;
  const CardTitle=({children,style={}})=><div style={{fontSize:'0.65rem',textTransform:'uppercase',letterSpacing:'0.12em',color:S.text2,marginBottom:12,...style}}>{children}</div>;

  return (
    <div style={{ fontFamily:'system-ui,sans-serif', background:S.bg, color:S.text, minHeight:'100vh', maxWidth:480, margin:'0 auto', position:'relative' }}>
      {confirmIdx!==null&&(
        <ConfirmDialog
          messaggio={confirmIdx==='all'?'Cancellare TUTTI i dati?':'Cancellare questa ricarica?'}
          onConfirm={()=>confirmIdx==='all'?(setRicariche([]),setConfirmIdx(null),showToast('🗑 Dati cancellati','#ef4444')):cancellaRicarica(confirmIdx)}
          onCancel={()=>setConfirmIdx(null)}/>
      )}
      {showImport&&<ImportModal onImportJSON={importaJSON} onImportFoglio={importaFoglio} onClose={()=>setShowImport(false)}/>}
      {showExport&&(
        <div style={{ position:'fixed', inset:0, zIndex:300, background:'rgba(0,0,0,0.75)', backdropFilter:'blur(6px)', display:'flex', alignItems:'flex-end', justifyContent:'center' }}>
          <div style={{ background:'#111827', border:`1px solid ${S.border}`, borderTopLeftRadius:24, borderTopRightRadius:24, padding:28, width:'100%', maxWidth:480 }}>
            <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:20 }}>
              <div><div style={{ fontSize:'1.1rem', fontWeight:800 }}>📤 Esporta dati</div><div style={{ fontSize:'0.7rem', color:S.text2, marginTop:2 }}>Scegli il formato</div></div>
              <button onClick={()=>setShowExport(false)} style={{ background:'rgba(255,255,255,0.08)', border:'none', borderRadius:20, color:S.text2, width:32, height:32, fontSize:'1rem', cursor:'pointer' }}>✕</button>
            </div>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:12, marginBottom:16 }}>
              <button onClick={()=>{esportaCSV();setShowExport(false);}} style={{ background:'rgba(0,229,255,0.08)', border:'1px solid rgba(0,229,255,0.3)', borderRadius:16, padding:'20px 12px', cursor:'pointer', textAlign:'center', color:S.text }}>
                <div style={{ fontSize:'2.2rem', marginBottom:8 }}>📊</div>
                <div style={{ fontSize:'0.8rem', fontWeight:800, textTransform:'uppercase', color:S.accent }}>CSV</div>
                <div style={{ fontSize:'0.62rem', color:S.text2, marginTop:6 }}>Apri in Excel<br/>o Google Sheets</div>
              </button>
              <button onClick={()=>{esportaJSON();setShowExport(false);}} style={{ background:'rgba(124,58,237,0.1)', border:'1px solid rgba(124,58,237,0.35)', borderRadius:16, padding:'20px 12px', cursor:'pointer', textAlign:'center', color:S.text }}>
                <div style={{ fontSize:'2.2rem', marginBottom:8 }}>💾</div>
                <div style={{ fontSize:'0.8rem', fontWeight:800, textTransform:'uppercase', color:'#a78bfa' }}>Backup JSON</div>
                <div style={{ fontSize:'0.62rem', color:S.text2, marginTop:6 }}>Ripristino<br/>completo</div>
              </button>
            </div>
            <a href="https://docs.google.com/spreadsheets/d/1egavj34-1EM3lY91kSikV48G7zxrnM7tOKdJ5GqPhY4/edit?gid=433900317#gid=433900317" target="_blank" rel="noreferrer" style={{ display:'block', width:'100%', padding:'13px 0', background:'rgba(0,229,255,0.08)', border:'1px solid rgba(0,229,255,0.3)', borderRadius:12, color:S.accent, fontSize:'0.85rem', fontWeight:700, textAlign:'center', textDecoration:'none', marginBottom:8, boxSizing:'border-box' }}>📊 Apri Google Sheet</a>
            <button onClick={()=>setShowExport(false)} style={{ width:'100%', padding:'13px 0', background:'transparent', border:`1px solid rgba(255,255,255,0.1)`, borderRadius:12, color:S.text2, fontSize:'0.85rem', cursor:'pointer' }}>Annulla</button>
          </div>
        </div>
      )}

      <div style={{ position:'sticky', top:0, zIndex:100, background:'rgba(10,15,30,0.95)', backdropFilter:'blur(20px)', borderBottom:`1px solid ${S.border}`, padding:'12px 16px', display:'flex', alignItems:'center', justifyContent:'space-between' }}>
        <div style={{ display:'flex', alignItems:'center', gap:10 }}>
          <div style={{ fontSize:'1.1rem', fontWeight:800, letterSpacing:'0.05em' }}>MG<span style={{color:'#ef4444'}}>S5</span></div>
          <a href="https://docs.google.com/spreadsheets/d/1egavj34-1EM3lY91kSikV48G7zxrnM7tOKdJ5GqPhY4/edit?gid=433900317#gid=433900317" target="_blank" rel="noreferrer" style={{ display:'flex', alignItems:'center', gap:4, padding:'3px 9px', background:'rgba(0,229,255,0.08)', border:'1px solid rgba(0,229,255,0.25)', borderRadius:20, color:S.accent, fontSize:'0.65rem', fontWeight:700, textDecoration:'none' }}>📊</a>
        </div>
        <div style={{ display:'flex', alignItems:'center', gap:16 }}>
          <div style={{ textAlign:'right' }}>
            <div style={{ fontFamily:'monospace', fontSize:'1rem', fontWeight:700, color:S.accent }}>{kmMax?kmMax.toLocaleString('it'):'—'}</div>
            <div style={{ fontSize:'0.62rem', color:S.text2, textTransform:'uppercase', letterSpacing:'0.06em' }}>km totali</div>
          </div>
          <div style={{ textAlign:'right' }}>
            <div style={{ fontFamily:'monospace', fontSize:'1rem', fontWeight:700, color:S.accent }}>{costoTot?'€'+(costoTot||0).toFixed(2):'—'}</div>
            <div style={{ fontSize:'0.62rem', color:S.text2, textTransform:'uppercase', letterSpacing:'0.06em' }}>costo tot.</div>
          </div>
        </div>
      </div>

      <div style={{ padding:'16px 16px 120px' }}>
        {view==='home'&&<>
          {!ricariche.length?(
            <div style={{ textAlign:'center', padding:'60px 20px', color:S.text2 }}>
              <div style={{ fontSize:'4rem', marginBottom:16, opacity:0.4 }}>🔋</div>
              <div style={{ fontSize:'1rem', marginBottom:24 }}>Nessuna ricarica.<br/>Premi + per aggiungere.</div>
              <button onClick={()=>setShowImport(true)} style={{ display:'inline-flex', alignItems:'center', gap:8, padding:'14px 28px', background:'rgba(0,229,255,0.1)', border:`1px solid rgba(0,229,255,0.3)`, borderRadius:40, color:S.accent, fontSize:'0.85rem', fontWeight:700, textTransform:'uppercase', letterSpacing:'0.06em', cursor:'pointer' }}>
                <span style={{ fontSize:'1.1rem' }}>📥</span> Importa dati
              </button>
            </div>
          ):(() => {
            const chiavi=Object.keys(perMese).sort();
            return chiavi.slice().reverse().map((key,keyIdx)=>{
              const [anno,mese]=key.split('-'), lista=perMese[key];
              const totCosto=lista.reduce((s,r)=>s+r.costo,0), totKwh=lista.reduce((s,r)=>s+r.kwhEff,0);
              let acc=0; const progMap={};
              [...lista].reverse().forEach(r=>{acc+=r.costo;progMap[r.idx]=acc;});
              const aperto=key in mesiAperti?mesiAperti[key]:keyIdx===0;
              const toggleMese=()=>setMesiAperti(prev=>({...prev,[key]:!aperto}));
              const listaCasa=lista.filter(r=>(r.luogo||'').toUpperCase()==='CASA');
              const totCostoCasa=listaCasa.reduce((s,r)=>s+r.costo,0), totKwhCasa=listaCasa.reduce((s,r)=>s+r.kwhEff,0);
              return (
                <div key={key} style={{ marginBottom:10 }}>
                  <div onClick={toggleMese} style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'12px 16px', background:'linear-gradient(135deg,rgba(0,229,255,0.1),rgba(124,58,237,0.1))', border:`1px solid ${S.border}`, borderRadius:aperto?'14px 14px 0 0':14, cursor:'pointer', userSelect:'none' }}>
                    <div style={{ display:'flex', alignItems:'center', flexShrink:0, width:140 }}>
                      <div style={{ width:10, height:10, borderRadius:'50%', background:S.accent, flexShrink:0, marginRight:8 }}/>
                      <div>
                        <div style={{ fontWeight:900, fontSize:'1.2rem', textTransform:'uppercase', letterSpacing:'0.08em', color:'#ffffff', lineHeight:1.15 }}>{MESI_NOMI[parseInt(mese)-1]}</div>
                        <div style={{ fontWeight:900, fontSize:'1.2rem', letterSpacing:'0.08em', color:'#ffffff', lineHeight:1.15 }}>{anno}</div>
                      </div>
                    </div>
                    <div style={{ display:'flex', alignItems:'center', flex:1, justifyContent:'center' }}>
                      <div style={{ textAlign:'right' }}>
                        <div style={{ fontFamily:'monospace', fontSize:'0.83rem', color:S.text2, lineHeight:1.6 }}>{lista.length} ric</div>
                        <div style={{ fontFamily:'monospace', fontSize:'0.83rem', color:S.text2, lineHeight:1.6 }}>tot</div>
                        <div style={{ fontFamily:'monospace', fontSize:'0.83rem', color:S.text2, lineHeight:1.6 }}>{listaCasa.length} ric</div>
                        <div style={{ fontFamily:'monospace', fontSize:'0.83rem', color:S.text2, lineHeight:1.6 }}>casa</div>
                      </div>
                    </div>
                    <div style={{ textAlign:'right' }}>
                      <div style={{ fontFamily:'monospace', fontSize:'1rem', fontWeight:700, color:S.accent, lineHeight:1.6 }}>€{(totCosto||0).toFixed(2)}</div>
                      <div style={{ fontFamily:'monospace', fontSize:'1rem', fontWeight:700, color:S.accent, lineHeight:1.6 }}>{(totKwh||0).toFixed(1)} kWh</div>
                      <div style={{ fontFamily:'monospace', fontSize:'1rem', fontWeight:700, color:'#34d399', lineHeight:1.6 }}>€{(totCostoCasa||0).toFixed(2)}</div>
                      <div style={{ fontFamily:'monospace', fontSize:'1rem', fontWeight:700, color:'#34d399', lineHeight:1.6 }}>{(totKwhCasa||0).toFixed(1)} kWh</div>
                    </div>
                  </div>
                  {aperto&&(
                    <div style={{ border:`1px solid ${S.border}`, borderTop:'none', borderRadius:'0 0 14px 14px', overflow:'hidden' }}>
                      {lista.map(r=>{
                        const d=new Date(r.data);
                        const dataStr=String(d.getDate()).padStart(2,'0')+'/'+String(d.getMonth()+1).padStart(2,'0')+'/'+d.getFullYear();
                        return (
                          <div key={r.idx} style={{ padding:'14px 16px', background:'rgba(255,255,255,0.02)', borderTop:'1px solid rgba(255,255,255,0.05)', display:'flex', flexDirection:'column', gap:8 }}>
                            <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between' }}>
                              <div style={{ fontFamily:'monospace', fontSize:'0.85rem', color:S.text2, fontWeight:600 }}>{dataStr}{r.luogo?' · '+r.luogo:''}</div>
                              <div style={{ display:'flex', alignItems:'center', gap:6 }}>
                                <div style={{ fontFamily:'monospace', fontSize:'0.8rem', color:S.text2 }}>{Math.round(r.pctPrima*(r.pctPrima<=1?100:1))}→{Math.round(r.pctDopo*(r.pctDopo<=1?100:1))}%</div>
                                <button onClick={e=>{e.stopPropagation();apriModifica(r.idx);}} style={{ background:'none', border:'none', color:S.accent, cursor:'pointer', padding:'4px 6px', fontSize:'1rem' }}>✏️</button>
                                <button onClick={e=>{e.stopPropagation();setConfirmIdx(r.idx);}} style={{ background:'none', border:'none', color:'#ef4444', cursor:'pointer', padding:'4px 8px', fontSize:'1.1rem', fontWeight:700 }}>✕</button>
                              </div>
                            </div>
                            <div style={{ display:'flex', alignItems:'baseline', justifyContent:'space-between' }}>
                              <div style={{ fontSize:'1.3rem', fontWeight:700, fontFamily:'monospace', color:S.text }}>{(r.kwhEff||0).toFixed(2)} kWh</div>
                              <div style={{ fontFamily:'monospace', fontSize:'0.9rem', color:'#f59e0b' }}>{r.kwh100?r.kwh100.toFixed(2)+' kWh/100km':'—'}</div>
                            </div>
                            <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between' }}>
                              <div style={{ fontSize:'0.8rem', color:S.text2, lineHeight:1.6 }}>
                                {r.prezzoKwh?'€'+r.prezzoKwh+'/kWh':'—'}
                                {r.kmParziali?'  ·  +'+r.kmParziali+' km':''}
                                {r.km?'  ·  '+r.km.toLocaleString('it')+' km tot':''}
                              </div>
                              <div style={{ textAlign:'right' }}>
                                <div style={{ fontFamily:'monospace', fontSize:'1.05rem', fontWeight:700, color:S.green }}>€{(r.costo||0).toFixed(2)}</div>
                                <div style={{ fontFamily:'monospace', fontSize:'0.7rem', color:S.accent }}>↑€{(progMap[r.idx]||0).toFixed(2)}</div>
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              );
            });
          })()}
        </>}

        {view==='add'&&<FormRicarica settings={settings} ricariche={ricariche} editIdx={editIdx} onSave={salvaRicarica} onCancel={editIdx!==null?annullaModifica:null} showToast={showToast}/>}

        {view==='charts'&&<>
          <Card><CardTitle>Costo mensile (€)</CardTitle><BarChart data={chartMensile} color={S.accent}/></Card>
          <Card><CardTitle>kWh / 100km per ricarica</CardTitle><LineChart data={chartEff} color={S.accent2}/></Card>
          <Card><CardTitle>Prezzo €/kWh nel tempo</CardTitle><LineChart data={chartPrezzo} color="#f59e0b"/></Card>
        </>}

        {view==='export'&&<>
          <Card>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:10 }}>
              <button onClick={()=>setShowImport(true)} style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', gap:8, padding:'18px 0', background:'rgba(0,229,255,0.1)', border:`1px solid rgba(0,229,255,0.3)`, borderRadius:14, color:S.accent, fontSize:'0.85rem', fontWeight:800, textTransform:'uppercase', letterSpacing:'0.06em', cursor:'pointer' }}><span style={{ fontSize:'1.6rem' }}>📥</span> Importa</button>
              <button onClick={()=>setShowExport(true)} style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', gap:8, padding:'18px 0', background:'rgba(124,58,237,0.1)', border:'1px solid rgba(124,58,237,0.35)', borderRadius:14, color:'#a78bfa', fontSize:'0.85rem', fontWeight:800, textTransform:'uppercase', letterSpacing:'0.06em', cursor:'pointer' }}><span style={{ fontSize:'1.6rem' }}>📤</span> Esporta</button>
              <button onClick={sincronizzaDaSheet} disabled={syncing} style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', gap:8, padding:'18px 0', background:'rgba(16,185,129,0.1)', border:'1px solid rgba(16,185,129,0.35)', borderRadius:14, color:'#10b981', fontSize:'0.85rem', fontWeight:800, textTransform:'uppercase', letterSpacing:'0.06em', cursor:'pointer', opacity:syncing?0.6:1 }}><span style={{ fontSize:'1.6rem' }}>☁️</span> {syncing?'...':'Da Sheet'}</button>
              <button onClick={sincronizzaVersoSheet} disabled={syncing} style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', gap:8, padding:'18px 0', background:'rgba(245,158,11,0.1)', border:'1px solid rgba(245,158,11,0.35)', borderRadius:14, color:'#f59e0b', fontSize:'0.85rem', fontWeight:800, textTransform:'uppercase', letterSpacing:'0.06em', cursor:'pointer', opacity:syncing?0.6:1 }}><span style={{ fontSize:'1.6rem' }}>📡</span> {syncing?'...':'A Sheet'}</button>
            </div>
          </Card>
          <Card>
            <CardTitle>Riepilogo</CardTitle>
            {Object.entries((() => {
              const pm={};
              ricariche.forEach(r=>{const d=new Date(r.data);const k=MESI_NOMI[d.getMonth()]+' '+d.getFullYear();if(!pm[k])pm[k]={costo:0,kwh:0,count:0};pm[k].costo+=r.costo;pm[k].kwh+=r.kwhEff;pm[k].count++;});
              return pm;
            })()).map(([mese,v])=>(
              <div key={mese} style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'12px 0', borderBottom:`1px solid rgba(255,255,255,0.05)` }}>
                <div><div style={{ fontSize:'0.85rem' }}>{mese}</div><div style={{ fontSize:'0.65rem', color:S.text2 }}>{v.count} ricariche · {(v.kwh||0).toFixed(1)} kWh</div></div>
                <div style={{ fontFamily:'monospace', color:S.green }}>€{(v.costo||0).toFixed(2)}</div>
              </div>
            ))}
            {!ricariche.length&&<div style={{ textAlign:'center', padding:20, color:S.text2, fontSize:'0.85rem' }}>Nessun dato</div>}
          </Card>
        </>}

        {view==='settings'&&<>
          <Card>
            <CardTitle>Impostazioni</CardTitle>
            {[{label:'Capacità batteria (kWh)',sub:'kWh totali della batteria',key:'batteria',type:'number',step:1},{label:'Prezzo default €/kWh',key:'prezzo',type:'number',step:0.001},{label:'Targa / Nome veicolo',key:'targa',type:'text'}].map(f=>(
              <div key={f.key} style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'12px 0', borderBottom:`1px solid rgba(255,255,255,0.05)` }}>
                <div><div style={{ fontSize:'0.85rem' }}>{f.label}</div>{f.sub&&<div style={{ fontSize:'0.65rem', color:S.text2 }}>{f.sub}</div>}</div>
                <input type={f.type} step={f.step} value={settings[f.key]} onChange={e=>setSettings({...settings,[f.key]:f.type==='number'?parseFloat(e.target.value)||0:e.target.value})} style={{...inputSt,width:100,textAlign:'right'}}/>
              </div>
            ))}
          </Card>
          <Card>
            <CardTitle>Preset % batteria</CardTitle>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:10 }}>
              {[{label:'% Prima — op.1',key:'p1a'},{label:'% Prima — op.2',key:'p1b'},{label:'% Dopo — op.1',key:'p2a'},{label:'% Dopo — op.2',key:'p2b'}].map(f=>(
                <div key={f.key}>
                  <div style={{ fontSize:'0.65rem', color:S.text2, textTransform:'uppercase', marginBottom:6 }}>{f.label}</div>
                  <input type="number" value={settings[f.key]} onChange={e=>setSettings({...settings,[f.key]:parseFloat(e.target.value)||0})} style={{...inputSt,textAlign:'right'}}/>
                </div>
              ))}
            </div>
          </Card>
          <Card style={{ borderColor:'rgba(239,68,68,0.3)' }}>
            <CardTitle style={{ color:'#ef4444' }}>⚠️ Zona pericolosa</CardTitle>
            <button onClick={()=>setConfirmIdx('all')} style={{ width:'100%', padding:12, background:'transparent', border:'1px solid rgba(239,68,68,0.4)', borderRadius:10, color:'#ef4444', fontSize:'0.85rem', cursor:'pointer' }}>🗑 Cancella tutti i dati</button>
          </Card>
          <Card>
            <CardTitle>☁️ Google Sheets (opzionale)</CardTitle>
            <div style={{ fontSize:'0.75rem', color:S.text2, marginBottom:12, lineHeight:1.6 }}>
              Collega un tuo foglio Google per sincronizzare i dati su più dispositivi.
            </div>
            <div style={{ display:'flex', flexDirection:'column', gap:6, marginBottom:12 }}>
              <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>URL Apps Script</label>
              <input type="text" defaultValue={getSheetsUrl()} onChange={e=>storage.set('mgs5_sheetsUrl', e.target.value.trim())} placeholder="https://script.google.com/macros/s/..." style={{...inputSt, fontSize:'0.7rem'}}/>
              <div style={{ fontSize:'0.6rem', color:S.text2, marginTop:2 }}>Incolla qui l'URL del tuo Google Apps Script dopo averlo pubblicato</div>
            </div>
            <div style={{ background:'rgba(0,229,255,0.04)', border:`1px solid ${S.border}`, borderRadius:10, padding:12 }}>
              <div style={{ fontSize:'0.7rem', fontWeight:700, color:S.accent, marginBottom:8 }}>Come configurare:</div>
              <div style={{ fontSize:'0.65rem', color:S.text2, lineHeight:1.8 }}>
                1. Vai su <span style={{ color:S.accent }}>script.google.com</span> → Nuovo progetto<br/>
                2. Incolla il codice da <span style={{ color:S.accent }}>google-apps-script.js</span><br/>
                3. Clicca Distribuisci → Nuova distribuzione → App web<br/>
                4. Accesso: <strong style={{ color:S.text }}>Chiunque</strong> → Distribuisci<br/>
                5. Copia l'URL e incollalo qui sopra
              </div>
            </div>
          </Card>
          
        </>}
      </div>

      <div style={{ position:'fixed', bottom:0, left:'50%', transform:'translateX(-50%)', width:'100%', maxWidth:480, background:'rgba(10,15,30,0.95)', backdropFilter:'blur(20px)', borderTop:`1px solid ${S.border}`, display:'flex', zIndex:100, paddingBottom:'env(safe-area-inset-bottom)' }}>
        {[{id:'home',icon:'⚡',label:'Home'},{id:'add',icon:'＋',label:'Aggiungi'},{id:'charts',icon:'📈',label:'Grafici'},{id:'export',icon:'📤',label:'Esporta'},{id:'settings',icon:'⚙️',label:'Config'}].map(btn=>(
          <button key={btn.id} onClick={()=>{if(btn.id==='add')setEditIdx(null);setView(btn.id);}} style={{ flex:1, display:'flex', flexDirection:'column', alignItems:'center', gap:4, padding:'8px 4px', background:'none', border:'none', color:view===btn.id?S.accent:S.text2, fontFamily:'system-ui', fontSize:'0.6rem', textTransform:'uppercase', letterSpacing:'0.08em', cursor:'pointer' }}>
            <span style={{ fontSize:'1.3rem', lineHeight:1 }}>{btn.icon}</span>
            {btn.label}
          </button>
        ))}
      </div>

      {toast&&(
        <div style={{ position:'fixed', bottom:90, left:'50%', transform:'translateX(-50%)', background:toast.color, color:'#fff', padding:'10px 20px', borderRadius:20, fontSize:'0.8rem', fontWeight:600, zIndex:200, whiteSpace:'nowrap', boxShadow:'0 4px 20px rgba(0,0,0,0.4)' }}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}

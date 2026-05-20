 import { useState, useRef, useEffect } from "react"; import * as XLSX from 'xlsx';

const MESI_NOMI = ['GENNAIO','FEBBRAIO','MARZO','APRILE','MAGGIO','GIUGNO','LUGLIO','AGOSTO','SETTEMBRE','OTTOBRE','NOVEMBRE','DICEMBRE']; const MESI_BREVI = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic']; const DEFAULT_SETTINGS = { batteria: 64, prezzo: 0.50, targa: 'MGS5', p1a: 20, p1b: 30, p2a: 80, p2b: 100 }; const SHEETS_URL = '
script.google.com
';

async function syncToSheets(ricariche) { try { const encoded = encodeURIComponent(JSON.stringify(ricariche)); await fetch(SHEETS_URL + '?action=write&data=' + encoded); } catch(e) { console.error('Sync to Sheets failed:', e); } }

async function syncFromSheets() { try { const res = await fetch(SHEETS_URL + '?action=read'); const json = await res.json(); return json.data || []; } catch(e) { console.error('Sync from Sheets failed:', e); return null; } }

function today() { return new Date().toISOString().split('T')[0]; }

function toNum(x, def = 0) { const n = Number(x); return Number.isFinite(n) ? n : def; } function sanitizzaRicarica(r) { return { ...r, kwhEff: toNum(r?.kwhEff, 0), kwhTeor: toNum(r?.kwhTeor, 0), costo: toNum(r?.costo, 0), prezzoKwh: toNum(r?.prezzoKwh, 0), pctPrima: toNum(r?.pctPrima, 0), pctDopo: toNum(r?.pctDopo, 0), kwh100: r?.kwh100 == null ? null : (Number.isFinite(Number(r.kwh100)) ? Number(r.kwh100) : null), km: r?.km == null ? null : (Number.isFinite(Number(r.km)) ? Number(r.km) : null), kmParziali:r?.kmParziali == null ? null : (Number.isFinite(Number(r.kmParziali)) ? Number(r.kmParziali) : null), }; }

function useToast() { const [toast, setToast] = useState(null); const show = (msg, color = '#10b981') => { setToast({ msg, color }); setTimeout(() => setToast(null), 2500); }; return { toast, show }; }

function ricalcolaKmParziali(lista) { return lista.map((r, i) => { if (r.km && r.km > 0) { const prev = [...lista].slice(0, i).filter(x => x.km && x.km > 0).pop(); const kmParziali = prev ? r.km - prev.km : r.kmParziali; const kwh100 = kmParziali && kmParziali > 0 ? (r.kwhEff / kmParziali) * 100 : r.kwh100; return { ...r, kmParziali: kmParziali || r.kmParziali, kwh100 }; } return r; }); }

function BarChart({ data, color = '#00e5ff' }) { if (!data.length) return <div style={{ height:160, display:'flex', alignItems:'center', justifyContent:'center', color:'#94a3b8', fontSize:'0.75rem' }}>Nessun dato; const max = Math.max(...data.map(d => d.y)); return ( <div style={{ display:'flex', alignItems:'flex-end', gap:6, height:160, padding:'8px 0' }}> {data.map((d,i) => ( <div key={i} style={{ flex:1, display:'flex', flexDirection:'column', alignItems:'center', gap:4, height:'100%', justifyContent:'flex-end' }}> <div style={{ width:'100%', background:color+'99', borderRadius:4, height:max?${(d.y/max)*120}px:0, minHeight:d.y?4:0, transition:'height 0.4s' }}/> <div style={{ fontSize:'0.5rem', color:'#94a3b8', textAlign:'center', whiteSpace:'nowrap', overflow:'hidden', maxWidth:'100%' }}>{d.x} ))} ); }

function LineChart({ data, color = '#7c3aed' }) { if (data.length < 2) return <div style={{ height:160, display:'flex', alignItems:'center', justifyContent:'center', color:'#94a3b8', fontSize:'0.75rem' }}>Servono almeno 2 punti; const w=300, h=120; const ys=data.map(d=>d.y), minY=Math.min(...ys), maxY=Math.max(...ys); const px=i=>(i/(data.length-1))w; const py=y=>h-((y-minY)/(maxY-minY||1))(h-10)-5; const pts=data.map((d,i)=>${px(i)},${py(d.y)}).join(' '); return ( <svg viewBox={0 0 ${w} ${h}} style={{ width:'100%', height:160 }}> <linearGradient id={g${color.replace('#','')}} x1="0" y1="0" x2="0" y2="1"> <polygon points={0,${h} ${pts} ${w},${h}} fill={url(#g${color.replace('#','')})}/> {data.map((d,i)=>)} ); }

const S = { bg:'#0a0f1e', bg2:'rgba(255,255,255,0.04)', border:'rgba(0,229,255,0.15)', accent:'#00e5ff', accent2:'#7c3aed', green:'#10b981', text:'#e2e8f0', text2:'#94a3b8' }; const inputSt = { background:'rgba(255,255,255,0.06)', border:1px solid ${S.border}, borderRadius:10, color:S.text, fontFamily:'monospace', fontSize:'1rem', padding:'10px 12px', width:'100%', outline:'none', boxSizing:'border-box' }; const chipSt = { padding:'4px 10px', background:'rgba(0,229,255,0.08)', border:'1px solid rgba(0,229,255,0.2)', borderRadius:20, color:S.accent, fontFamily:'monospace', fontSize:'0.75rem', cursor:'pointer' };

function ConfirmDialog({ messaggio, onConfirm, onCancel }) { return ( <div style={{ position:'fixed', inset:0, zIndex:300, background:'rgba(0,0,0,0.7)', backdropFilter:'blur(4px)', display:'flex', alignItems:'center', justifyContent:'center', padding:24 }}> <div style={{ background:'#111827', border:'1px solid rgba(239,68,68,0.4)', borderRadius:20, padding:28, maxWidth:320, width:'100%', textAlign:'center' }}> <div style={{ fontSize:'2rem', marginBottom:12 }}>🗑 <div style={{ fontSize:'1rem', fontWeight:700, marginBottom:8 }}>{messaggio} <div style={{ fontSize:'0.8rem', color:S.text2, marginBottom:24 }}>Questa azione non può essere annullata. <div style={{ display:'flex', gap:10 }}> <button onClick={onCancel} style={{ flex:1, padding:'12px 0', background:'transparent', border:1px solid ${S.border}, borderRadius:12, color:S.text2, fontSize:'0.9rem', cursor:'pointer' }}>Annulla <button onClick={onConfirm} style={{ flex:1, padding:'12px 0', background:'rgba(239,68,68,0.15)', border:'1px solid rgba(239,68,68,0.5)', borderRadius:12, color:'#ef4444', fontSize:'0.9rem', fontWeight:700, cursor:'pointer' }}>Conferma ); }

function ImportModal({ onImportJSON, onImportFoglio, onClose }) { return ( <div style={{ position:'fixed', inset:0, zIndex:300, background:'rgba(0,0,0,0.75)', backdropFilter:'blur(6px)', display:'flex', alignItems:'flex-end', justifyContent:'center', padding:0 }}> <div style={{ background:'#111827', border:1px solid ${S.border}, borderTopLeftRadius:24, borderTopRightRadius:24, padding:28, width:'100%', maxWidth:480 }}> <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:20 }}>

<div style={{ fontSize:'1.1rem', fontWeight:800, letterSpacing:'0.04em' }}>📥 Importa dati
<div style={{ fontSize:'0.7rem', color:S.text2, marginTop:2 }}>Scegli il formato del file da caricare <button onClick={onClose} style={{ background:'rgba(255,255,255,0.08)', border:'none', borderRadius:20, color:S.text2, width:32, height:32, fontSize:'1rem', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center' }}>✕


    <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:12, marginBottom:16 }}>
      <button
        onClick={() => { const inp = document.createElement('input'); inp.type='file'; inp.accept='.json,application/json'; inp.onchange=onImportJSON; inp.click(); onClose(); }}
        style={{ background:'rgba(124,58,237,0.1)', border:'1px solid rgba(124,58,237,0.35)', borderRadius:16, padding:'20px 12px', cursor:'pointer', textAlign:'center', color:S.text, transition:'all 0.2s' }}
      >
        <div style={{ fontSize:'2.2rem', marginBottom:8 }}>💾</div>
        <div style={{ fontSize:'0.8rem', fontWeight:800, textTransform:'uppercase', color:'#a78bfa', letterSpacing:'0.06em' }}>Backup JSON</div>
        <div style={{ fontSize:'0.62rem', color:S.text2, marginTop:6, lineHeight:1.5 }}>Ripristina un backup<br/>completo precedente</div>
      </button>
      <button
        onClick={() => { const inp = document.createElement('input'); inp.type='file'; inp.accept='.xlsx,.xls,.csv,.ods,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv'; inp.onchange=onImportFoglio; inp.click(); onClose(); }}
        style={{ background:'rgba(0,229,255,0.08)', border:'1px solid rgba(0,229,255,0.3)', borderRadius:16, padding:'20px 12px', cursor:'pointer', textAlign:'center', color:S.text, transition:'all 0.2s' }}
      >
        <div style={{ fontSize:'2.2rem', marginBottom:8 }}>📊</div>
        <div style={{ fontSize:'0.8rem', fontWeight:800, textTransform:'uppercase', color:S.accent, letterSpacing:'0.06em' }}>Excel / Sheets</div>
        <div style={{ fontSize:'0.62rem', color:S.text2, marginTop:6, lineHeight:1.5 }}>xlsx · xls · csv · ods<br/>Google Sheets export</div>
      </button>
    </div>
    <button onClick={onClose} style={{ width:'100%', padding:'13px 0', background:'transparent', border:`1px solid rgba(255,255,255,0.1)`, borderRadius:12, color:S.text2, fontSize:'0.85rem', cursor:'pointer' }}>
      Annulla
    </button>
  </div>
</div>
); }

function FormRicarica({ settings, ricariche, editIdx, onSave, onCancel, showToast }) { const ex = editIdx !== null ? ricariche[editIdx] : null; const [fData, setFData] = useState(ex ? ex.data : today()); const [fLuogo, setFLuogo] = useState(ex ? (ex.luogo||'') : ''); const [fKm, setFKm] = useState(ex && ex.km !== null ? String(ex.km) : ''); const [fKmParz, setFKmParz] = useState(ex && ex.kmParziali !== null ? String(ex.kmParziali) : ''); const [fPrima, setFPrima] = useState(ex ? String(ex.pctPrima) : ''); const [fDopo, setFDopo] = useState(ex ? String(ex.pctDopo) : ''); const [fKwh, setFKwh] = useState(ex ? String(ex.kwhEff) : ''); const [fPrezzo, setFPrezzo] = useState(ex ? String(ex.prezzoKwh) : String(settings.prezzo));

const prima=parseFloat(fPrima)||0, dopo=parseFloat(fDopo)||0; const diff=dopo-prima, kwhTeor=settings.batteria*(diff/100); const kwhEff=parseFloat(fKwh)||kwhTeor; const prezzo=parseFloat(fPrezzo)||settings.prezzo; const costo=kwhEff*prezzo; const kmParzN=parseFloat(fKmParz); const kmN=parseFloat(fKm);

let kmParzAuto=null; if (!isNaN(kmN)&&kmN>0&&!(kmParzN>0)) { const altri=editIdx!==null?ricariche.filter((_,i)=>i!==editIdx):ricariche; const prev=[...altri].filter(r=>r.km&&r.km<kmN).sort((a,b)=>b.km-a.km)[0]; if (prev) kmParzAuto=kmN-prev.km; } const kmParzEff=(kmParzN>0)?kmParzN:kmParzAuto; const kwh100=kmParzEff?(kwhEff/kmParzEff)*100:null;

function salva() { if (!fData||prima<=0||dopo<=0){showToast('⚠️ Compila data e %','#f59e0b');return;} if (dopo<=prima){showToast('⚠️ % dopo > % prima','#f59e0b');return;} onSave({ data:fData, luogo:fLuogo||null, km:parseFloat(fKm)||null, pctPrima:prima, pctDopo:dopo, kwhEff, kwhTeor, prezzoKwh:prezzo, costo, kmParziali:kmParzEff||null, kwh100:kmParzEff?(kwhEff/kmParzEff)*100:null }); }

return ( <div style={{ background:S.bg2, border:1px solid ${S.border}, borderRadius:16, padding:16, marginBottom:12 }}> <div style={{ fontSize:'0.65rem', textTransform:'uppercase', letterSpacing:'0.12em', color:S.text2, marginBottom:12 }}> {editIdx!==null?'✏️ Modifica Ricarica':'Nuova Ricarica'} <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:10 }}> <div style={{ display:'flex', flexDirection:'column', gap:6 }}> <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>Data <input type="date" value={fData} onChange={e=>setFData(e.target.value)} style={inputSt}/> <div style={{ display:'flex', flexDirection:'column', gap:6 }}> <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>Dove si ricarica <input type="text" value={fLuogo} onChange={e=>setFLuogo(e.target.value.toUpperCase())} placeholder="CASA..." style={inputSt}/> <div style={{ display:'flex', flexDirection:'column', gap:6 }}> <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>KM Totali <input type="number" value={fKm} onChange={e=>setFKm(e.target.value)} placeholder="0" style={inputSt}/> <div style={{ display:'flex', flexDirection:'column', gap:6 }}> <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>KM Parziali <input type="number" value={fKmParz} onChange={e=>setFKmParz(e.target.value)} placeholder={kmParzAuto?String(kmParzAuto):''} style={inputSt}/> {kmParzAuto&&!(kmParzN>0)&&<div style={{ fontSize:'0.6rem', color:S.accent, marginTop:2 }}>⚡ calcolato automaticamente} <div style={{ display:'flex', flexDirection:'column', gap:6 }}> <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>% Prima <input type="number" value={fPrima} onChange={e=>setFPrima(e.target.value)} placeholder="0" style={inputSt}/> <div style={{ display:'flex', gap:6 }}> {[settings.p1a,settings.p1b].map(v=><button key={v} onClick={()=>setFPrima(String(v))} style={chipSt}>{v}%)} <div style={{ display:'flex', flexDirection:'column', gap:6 }}> <label style={{ fontSize:'0.62rem', textTransform:'uppercase', color:S.text2 }}>% Dopo <input type="number" value={fDopo} onChange={e=>setFDopo(e.target.value)} placeholder="100" style={inputSt}/> <div style={{ display:'flex', gap:6 }}> {[settings.p2a,settings.p2b].map(v=><button key={v} onClick={()=>setFDopo(String(v))} style={chipSt}>{v}%)} <div style={{ gridColumn:'1/-1', background:'rgba(0,229,255,0.04)', border:1px solid ${S.border}, borderRadius:10, padding:10, display:'flex', justifyContent:'space-between' }}>

<div style={{ fontSize:'0.65rem', color:S.text2, textTransform:'uppercase' }}>Differenza %
<div style={{ fontFamily:'monospace', fontSize:'1rem', color:S.accent }}>{diff>0?diff.toFixed(0)+' %':'—'} <div style={{ textAlign:'right' }}> <div style={{ fontSize:'0.65rem', color:S.text2, textTransform:'uppercase' }}>kWh teorici




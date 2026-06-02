import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const C = {
  bg:"#0F1117", panel:"#161B24", card:"#1E2530", border:"#2A3140",
  accent:"#3B82F6", accentL:"#60A5FA", green:"#22C55E",
  red:"#EF4444", amber:"#F59E0B", text:"#F0F4FF", muted:"#6B7A99", dim:"#3A4560",
};
const FONT = "'IBM Plex Mono','Courier New',monospace";
const uid = () => Math.random().toString(36).slice(2,8);

function cropCanvas(canvas, px, py, cw, ch) {
  const x = Math.max(0, Math.round(px - cw/2));
  const y = Math.max(0, Math.round(py - ch/2));
  const w = Math.min(cw, canvas.width  - x);
  const h = Math.min(ch, canvas.height - y);
  if (w <= 0 || h <= 0) return null;
  const tmp = document.createElement("canvas");
  tmp.width = w; tmp.height = h;
  tmp.getContext("2d").drawImage(canvas, x, y, w, h, 0, 0, w, h);
  return tmp.toDataURL("image/png").split(",")[1];
}

// Call Claude API — works inside Claude artifact sandbox
async function callClaude(imageB64, tapX, tapY, canvasW, canvasH) {
  const resp = await fetch("/api/claude", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 150,
      system: 'You read text from PDF page images. Respond ONLY with raw JSON, no markdown: {"fieldName":"...","value":"..."}',
      messages: [{
        role: "user",
        content: [
          { type: "image", source: { type: "base64", media_type: "image/png", data: imageB64 } },
          { type: "text", text: `The user tapped at approximately (${tapX}, ${tapY}) on this ${canvasW}x${canvasH} image. What field label and value is closest to that point? Reply only with JSON: {"fieldName":"...","value":"..."}` }
        ]
      }]
    })
  });

  // If /api/claude proxy not available, try direct
  if (resp.status === 404) {
    const resp2 = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-sonnet-4-20250514",
        max_tokens: 150,
        system: 'You read text from PDF page images. Respond ONLY with raw JSON, no markdown: {"fieldName":"...","value":"..."}',
        messages: [{
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: "image/png", data: imageB64 } },
            { type: "text", text: `The user tapped at approximately (${tapX}, ${tapY}) on this ${canvasW}x${canvasH} image. What field label and value is closest to that point? Reply only with JSON: {"fieldName":"...","value":"..."}` }
          ]
        }]
      })
    });
    return resp2;
  }
  return resp;
}

export default function App() {
  const [pdfDoc,     setPdfDoc]     = useState(null);
  const [pageNum,    setPageNum]    = useState(1);
  const [totalPages, setTotalPages] = useState(0);
  const [scale,      setScale]      = useState(1.4);
  const [rotation,   setRotation]   = useState(0);
  const [fields,     setFields]     = useState([]);
  const [dots,       setDots]       = useState([]);
  const [downloaded, setDownloaded] = useState(false);
  const [pdfReady,   setPdfReady]   = useState(false);
  const [renderErr,  setRenderErr]  = useState(null);
  const [editingId,  setEditingId]  = useState(null);
  const [exportMsg,  setExportMsg]  = useState("");
  const [pageReady,  setPageReady]  = useState(false);
  const [isScanned,  setIsScanned]  = useState(false); // true = image-only PDF

  const canvasRef = useRef(null);
  const pdfJsRef  = useRef(null);
  const renderRef = useRef(null);
  const fileRef   = useRef(null);
  const textItems = useRef([]);

  useEffect(() => {
    if (window.pdfjsLib) { pdfJsRef.current = window.pdfjsLib; setPdfReady(true); return; }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
    s.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      pdfJsRef.current = window.pdfjsLib;
      setPdfReady(true);
    };
    document.head.appendChild(s);
  }, []);

  useEffect(() => {
    if (!pdfDoc || !canvasRef.current) return;
    setPageReady(false);
    textItems.current = [];
    let cancelled = false;
    (async () => {
      try {
        if (renderRef.current) renderRef.current.cancel();
        const page = await pdfDoc.getPage(pageNum);
        if (cancelled) return;
        const vp = page.getViewport({ scale, rotation });
        const canvas = canvasRef.current;
        canvas.width  = vp.width;
        canvas.height = vp.height;
        const task = page.render({ canvasContext: canvas.getContext("2d"), viewport: vp });
        renderRef.current = task;
        await task.promise;
        if (cancelled) return;

        // Try text extraction
        const content = await page.getTextContent();
        if (cancelled) return;
        const items = [];
        for (const item of content.items) {
          if (!item.str?.trim()) continue;
          const pt = vp.convertToViewportPoint(item.transform[4], item.transform[5]);
          items.push({ text: item.str.trim(), x: pt[0], y: pt[1] });
        }
        textItems.current = items;
        const scanned = items.length === 0;
        setIsScanned(scanned);
        setPageReady(true);
        setRenderErr(null);
      } catch(e) {
        if (e?.name !== "RenderingCancelledException") setRenderErr(String(e.message));
      }
    })();
    return () => { cancelled = true; };
  }, [pdfDoc, pageNum, scale, rotation]);

  const loadPDF = useCallback(async (file) => {
    if (!pdfReady) return;
    setFields([]); setDots([]); setDownloaded(false);
    setRenderErr(null); setExportMsg(""); setPageReady(false);
    textItems.current = [];
    try {
      const buf = await file.arrayBuffer();
      const doc = await pdfJsRef.current.getDocument({ data: buf }).promise;
      setPdfDoc(doc);
      setTotalPages(doc.numPages);
      setPageNum(1);
    } catch(e) {
      setRenderErr("Could not load PDF: " + e.message);
    }
  }, [pdfReady]);

  const handleTap = useCallback(async (e) => {
    if (!pdfDoc || !canvasRef.current || !pageReady) return;

    let clientX, clientY;
    if (e.type === "touchend") {
      if (!e.changedTouches?.length) return;
      clientX = e.changedTouches[0].clientX;
      clientY = e.changedTouches[0].clientY;
    } else {
      clientX = e.clientX;
      clientY = e.clientY;
    }
    e.preventDefault();
    e.stopPropagation();

    const canvas = canvasRef.current;
    const rect   = canvas.getBoundingClientRect();
    const relX   = clientX - rect.left;
    const relY   = clientY - rect.top;
    const px     = relX * (canvas.width  / rect.width);
    const py     = relY * (canvas.height / rect.height);

    const dotId = uid();
    setDots(d => [...d, { id:dotId, x:relX, y:relY }]);
    setTimeout(() => setDots(d => d.filter(x => x.id !== dotId)), 900);

    // ── SCANNED PDF: use AI vision ──────────────────────────────
    if (isScanned) {
      const fid = uid();
      setFields(prev => [...prev, { id:fid, label:"Reading…", value:"…", x:relX, y:relY, loading:true, error:false }]);

      try {
        // Crop generously around tap — bigger helps with scanned docs
        const cropB64 = cropCanvas(canvas, px, py, 520, 160);
        if (!cropB64) throw new Error("Crop failed");

        const res = await callClaude(cropB64, Math.round(px), Math.round(py), canvas.width, canvas.height);
        const rawText = await res.text();

        if (!res.ok) {
          let msg = `HTTP ${res.status}`;
          try { msg = JSON.parse(rawText)?.error?.message || msg; } catch{}
          throw new Error(msg);
        }

        const data = JSON.parse(rawText);
        if (data.error) throw new Error(data.error.message);

        const txt = (data.content||[]).filter(b=>b.type==="text").map(b=>b.text).join("").trim();
        const match = txt.match(/\{[^{}]+\}/);
        if (!match) throw new Error("No JSON in: " + txt.slice(0,60));

        const parsed = JSON.parse(match[0]);
        const label  = String(parsed.fieldName || parsed.label || parsed.field || "Field").trim();
        const value  = String(parsed.value || parsed.val || "—").trim();

        setFields(prev => prev.map(f =>
          f.id === fid ? { ...f, label, value, loading:false, error:false } : f
        ));
      } catch(err) {
        setFields(prev => prev.map(f =>
          f.id === fid ? { ...f, label:"Error", value:String(err.message).slice(0,100), loading:false, error:true } : f
        ));
      }
      return;
    }

    // ── TEXT PDF: local matching ────────────────────────────────
    const items = textItems.current;
    if (!items.length) return;

    let best = null, bestDist = Infinity;
    for (const item of items) {
      const d = Math.hypot(item.x - px, item.y - py);
      if (d < bestDist) { bestDist = d; best = item; }
    }
    if (!best) return;

    const rowTol = 18;
    const candidates = items.filter(t =>
      t !== best && (
        (Math.abs(t.y - best.y) < rowTol && t.x < best.x) ||
        (t.y < best.y - rowTol && t.y > best.y - 55 && Math.abs(t.x - best.x) < 100)
      )
    );
    let label = "Field";
    if (candidates.length) {
      let lb = candidates[0], lbD = Math.hypot(lb.x-best.x, lb.y-best.y);
      for (const c of candidates) {
        const d = Math.hypot(c.x-best.x, c.y-best.y);
        if (d < lbD) { lb = c; lbD = d; }
      }
      label = lb.text;
    }
    setFields(prev => [...prev, { id:uid(), label, value:best.text, x:relX, y:relY, loading:false, error:false }]);
  }, [pdfDoc, pageReady, isScanned]);

  const handleExport = useCallback(() => {
    const ready = fields.filter(f => !f.error && !f.loading);
    if (!ready.length) { setExportMsg("No fields to export."); return; }
    try {
      const row = {};
      ready.forEach(f => { row[f.label] = f.value; });
      const ws = XLSX.utils.json_to_sheet([row]);
      ws["!cols"] = Object.keys(row).map(k => ({ wch: Math.max(k.length, String(row[k]).length)+2 }));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Extracted");
      const arr  = XLSX.write(wb, { bookType:"xlsx", type:"array" });
      const blob = new Blob([arr], { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement("a");
      a.href=url; a.download="extracted_data.xlsx"; a.style.display="none";
      document.body.appendChild(a); a.click();
      setTimeout(()=>{ document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
      setDownloaded(true); setExportMsg("✓ Downloaded!");
    } catch(err) { setExportMsg("Error: "+err.message); }
  }, [fields]);

  const canExport = fields.some(f => !f.error && !f.loading);

  const statusColor = !pageReady ? C.amber : isScanned ? C.accent : C.green;
  const statusMsg   = !pdfDoc ? "" :
    !pageReady ? "Rendering…" :
    isScanned  ? `📷 Scanned PDF — tap any area, AI will read it` :
                 `✓ ${textItems.current.length} text items — tap to select`;

  return (
    <div style={{ minHeight:"100vh", background:C.bg, color:C.text, fontFamily:FONT, display:"flex", flexDirection:"column" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:5px} ::-webkit-scrollbar-thumb{background:${C.dim};border-radius:3px}
        @keyframes pulse{0%{transform:scale(.3);opacity:1}100%{transform:scale(3.5);opacity:0}}
        @keyframes spin{to{transform:rotate(360deg)}}
        @keyframes fadeUp{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:none}}
        .frow{animation:fadeUp .18s ease forwards}
        .tdot{position:absolute;width:32px;height:32px;border-radius:50%;
              border:2.5px solid ${C.accent};margin:-16px 0 0 -16px;
              pointer-events:none;animation:pulse .85s ease-out forwards}
      `}</style>

      {/* HEADER */}
      <div style={{ background:C.panel, borderBottom:`1px solid ${C.border}`, padding:"11px 14px", display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
        <span style={{ color:C.accent, fontSize:15 }}>◈</span>
        <span style={{ fontSize:13, fontWeight:600, letterSpacing:".07em" }}>PDF → EXCEL</span>
        {statusMsg && <span style={{ fontSize:10, color:statusColor, marginLeft:2 }}>{statusMsg}</span>}
        {pdfDoc && (
          <div style={{ marginLeft:"auto", display:"flex", gap:5, alignItems:"center" }}>
            <SmBtn disabled={pageNum<=1}          onClick={()=>setPageNum(p=>p-1)}>◂</SmBtn>
            <span style={{ fontSize:11, color:C.muted }}>{pageNum}/{totalPages}</span>
            <SmBtn disabled={pageNum>=totalPages} onClick={()=>setPageNum(p=>p+1)}>▸</SmBtn>
            <Sep/>
            <SmBtn onClick={()=>setScale(s=>Math.min(2.5,+(s+.2).toFixed(1)))}>＋</SmBtn>
            <SmBtn onClick={()=>setScale(s=>Math.max(.5, +(s-.2).toFixed(1)))}>－</SmBtn>
            <Sep/>
            <SmBtn onClick={()=>{ setRotation(r=>(r+90)%360); setFields([]); }}>↻</SmBtn>
          </div>
        )}
      </div>

      {/* BODY */}
      <div style={{ flex:1, display:"flex", overflow:"hidden" }}>

        {/* PDF VIEWER */}
        <div style={{ flex:1, overflow:"auto", background:C.bg, display:"flex", justifyContent:"center", alignItems:pdfDoc?"flex-start":"center", padding:pdfDoc?"20px":0 }}
          onDragOver={e=>e.preventDefault()}
          onDrop={e=>{ e.preventDefault(); const f=e.dataTransfer.files[0]; if(f?.type==="application/pdf") loadPDF(f); }}>

          {!pdfDoc ? (
            <div onClick={()=>fileRef.current.click()}
              style={{ border:`2px dashed ${C.border}`, borderRadius:10, padding:"48px 32px", textAlign:"center", cursor:"pointer", maxWidth:340 }}>
              <div style={{ fontSize:36, marginBottom:12 }}>⬆</div>
              <div style={{ fontSize:14, fontWeight:600, marginBottom:6 }}>Upload a PDF</div>
              <div style={{ fontSize:12, color:C.muted, lineHeight:1.7 }}>
                Works with both text PDFs<br/>and scanned image PDFs.
              </div>
              <input ref={fileRef} type="file" accept=".pdf,application/pdf" style={{ display:"none" }}
                onChange={e=>{ if(e.target.files[0]) loadPDF(e.target.files[0]); e.target.value=""; }}/>
            </div>
          ) : (
            <div style={{ position:"relative", display:"inline-block", boxShadow:"0 6px 32px #0008" }}>
              <canvas ref={canvasRef} style={{ display:"block", userSelect:"none" }}/>
              <div
                style={{ position:"absolute", inset:0, zIndex:5, cursor:pageReady?"crosshair":"wait", WebkitTapHighlightColor:"transparent", touchAction:"none" }}
                onClick={handleTap}
                onTouchEnd={handleTap}>
                {dots.map(d=>(
                  <div key={d.id} className="tdot" style={{ left:d.x, top:d.y }}/>
                ))}
                {fields.map(f=>(
                  <div key={f.id} style={{
                    position:"absolute", left:f.x, top:f.y, transform:"translate(-50%,-50%)",
                    width:11, height:11, borderRadius:"50%", pointerEvents:"none", zIndex:6,
                    background: f.loading ? C.amber : f.error ? C.red : C.green,
                    border:`2px solid ${C.bg}`, boxShadow:`0 0 0 3px ${f.loading?C.amber:f.error?C.red:C.green}55`,
                    transition:"background .3s"
                  }}/>
                ))}
              </div>
              {!pageReady && (
                <div style={{ position:"absolute", inset:0, background:"#0007", display:"flex", alignItems:"center", justifyContent:"center", zIndex:8, pointerEvents:"none" }}>
                  <span style={{ fontSize:12, color:C.amber }}>Rendering…</span>
                </div>
              )}
              {renderErr && (
                <div style={{ position:"absolute", inset:0, background:"#000b", display:"flex", alignItems:"center", justifyContent:"center", zIndex:10 }}>
                  <span style={{ color:C.red, fontSize:12, padding:16, textAlign:"center" }}>{renderErr}</span>
                </div>
              )}
            </div>
          )}
        </div>

        {/* FIELDS PANEL */}
        <div style={{ width:255, background:C.panel, borderLeft:`1px solid ${C.border}`, display:"flex", flexDirection:"column", flexShrink:0 }}>
          <div style={{ padding:"10px 13px", borderBottom:`1px solid ${C.border}`, display:"flex", alignItems:"center", justifyContent:"space-between" }}>
            <span style={{ fontSize:10, fontWeight:600, letterSpacing:".1em", color:C.muted, textTransform:"uppercase" }}>Fields</span>
            {fields.length>0 && <span style={{ fontSize:10, background:C.accent+"22", color:C.accentL, padding:"1px 6px", borderRadius:3, fontWeight:700 }}>{fields.length}</span>}
          </div>

          <div style={{ flex:1, overflowY:"auto", padding:fields.length?"7px 8px":0 }}>
            {fields.length===0 ? (
              <div style={{ padding:"24px 14px", textAlign:"center" }}>
                <div style={{ fontSize:20, opacity:.2, marginBottom:10 }}>⊕</div>
                <div style={{ fontSize:11, color:C.muted, lineHeight:1.8 }}>
                  {!pdfDoc ? "Upload a PDF to begin." :
                   !pageReady ? "Rendering page…" :
                   isScanned ? <>Scanned PDF detected.<br/>Tap any area — AI<br/>will read the text.<br/><span style={{color:C.amber,fontSize:10}}>Takes ~3 sec per tap.</span></> :
                   <>Tap any text on the<br/>PDF to capture it.</>}
                </div>
              </div>
            ) : fields.map((f,i)=>(
              <div key={f.id} className="frow" style={{ background:C.card, border:`1px solid ${f.error?C.red+"44":C.border}`, borderRadius:6, padding:"8px 10px", marginBottom:6 }}>
                <div style={{ display:"flex", alignItems:"center", gap:5, marginBottom:4 }}>
                  <span style={{ fontSize:9, color:C.muted, flexShrink:0 }}>#{i+1}</span>
                  {editingId===f.id ? (
                    <input type="text" defaultValue={f.label} autoFocus
                      style={{ flex:1, background:"transparent", border:`1px solid ${C.accent}`, borderRadius:3, padding:"1px 5px", color:C.text, fontSize:11, fontFamily:FONT, fontWeight:600 }}
                      onBlur={e=>{ const v=e.target.value.trim()||f.label; setFields(fs=>fs.map(ff=>ff.id===f.id?{...ff,label:v}:ff)); setEditingId(null); }}
                      onKeyDown={e=>{ if(e.key==="Enter") e.target.blur(); }}/>
                  ) : (
                    <span style={{ flex:1, fontSize:11, fontWeight:600, color:f.error?C.red:f.loading?C.amber:C.accentL, cursor:"pointer", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}
                      onClick={()=>{ if(!f.loading&&!f.error) setEditingId(f.id); }}>
                      {f.loading ? <span style={{display:"inline-flex",alignItems:"center",gap:4}}><span style={{animation:"spin .8s linear infinite",display:"inline-block"}}>◌</span> reading…</span> : f.label}
                    </span>
                  )}
                  <button onClick={()=>setFields(fs=>fs.filter(ff=>ff.id!==f.id))}
                    style={{ background:"none", border:"none", cursor:"pointer", color:C.muted, fontSize:13, lineHeight:1, padding:"0 2px", flexShrink:0 }}>×</button>
                </div>
                <div style={{ fontSize:12, color:f.error?C.red+"bb":C.text, lineHeight:1.4, wordBreak:"break-word" }}>{f.value}</div>
              </div>
            ))}
          </div>

          <div style={{ padding:"10px 8px", borderTop:`1px solid ${C.border}`, display:"flex", flexDirection:"column", gap:7 }}>
            {exportMsg && <div style={{ fontSize:10, color:exportMsg.startsWith("✓")?C.green:C.red, textAlign:"center" }}>{exportMsg}</div>}
            {fields.length>0 && (
              <button onClick={()=>{ setFields([]); setExportMsg(""); setDownloaded(false); }}
                style={{ ...base(), background:"transparent", color:C.muted, border:`1px solid ${C.border}`, fontSize:11 }}>
                Clear All
              </button>
            )}
            <button disabled={!canExport} onClick={handleExport}
              style={{ ...base(!canExport), background:downloaded?C.dim:C.green, color:"#fff", fontSize:12 }}>
              {downloaded?"✓ Downloaded!":"↓ Export to Excel"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function SmBtn({ children, onClick, disabled }) {
  return (
    <button onClick={onClick} disabled={disabled} style={{
      background:"transparent", border:`1px solid ${C.border}`,
      color:disabled?C.dim:C.muted, cursor:disabled?"not-allowed":"pointer",
      padding:"3px 8px", borderRadius:4, fontSize:12, fontFamily:FONT,
      opacity:disabled?.5:1, WebkitTapHighlightColor:"transparent",
    }}>{children}</button>
  );
}
function Sep() {
  return <span style={{ width:1, height:14, background:C.dim, display:"inline-block", margin:"0 1px" }}/>;
}
function base(disabled=false) {
  return {
    width:"100%", padding:"9px 13px", borderRadius:6, border:"none",
    cursor:disabled?"not-allowed":"pointer", fontFamily:FONT, fontWeight:600,
    letterSpacing:".04em", opacity:disabled?.4:1, transition:"opacity .15s",
    WebkitTapHighlightColor:"transparent",
  };
}

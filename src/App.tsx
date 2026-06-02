// @ts-nocheck
import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

const STORAGE_KEY = "dave-dashboard-v6";
const TIME_BLOCKS = ["5:00 AM","6:00 AM","7:00 AM","8:00 AM","9:00 AM","10:00 AM","11:00 AM","12:00 PM","1:00 PM","2:00 PM","3:00 PM","4:00 PM","5:00 PM","6:00 PM","7:00 PM","8:00 PM","9:00 PM","10:00 PM","11:00 PM"];
const DAYS        = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"];
const CLASS_TYPES = ["Lecture","Seminar","Lab","Tutorial","Workshop"];
const PRIORITIES  = ["High","Medium","Low"];
const TENDER_STATUSES = ["Under Estimation","Takeoff","Pricing","1st Submission","Technical Submission","PTC1","PTA-02(VE)","EOI","Submitted","Won","Lost","Regretted"];
const TENDER_STAGES   = ["","1st Submission","Technical Submission","PTC1","PTA-02(VE)","EOI","PTA-01","2nd Submission"];

const CAT        = { work:{label:"Work",color:"#3B82F6"}, study:{label:"Study",color:"#F59E0B"}, personal:{label:"Personal",color:"#10B981"} };
const PRIO_CLR   = { High:"#EF4444", Medium:"#F59E0B", Low:"#10B981" };
const TYPE_CLR   = { Lecture:"#3B82F6", Seminar:"#A78BFA", Lab:"#F59E0B", Tutorial:"#10B981", Workshop:"#F472B6" };
const STATUS_CLR = { "Under Estimation":"#64748b","Takeoff":"#F59E0B","Pricing":"#3B82F6","1st Submission":"#A78BFA","Technical Submission":"#8B5CF6","PTC1":"#F472B6","Submitted":"#10B981","Won":"#22c55e","Lost":"#EF4444","Regretted":"#94a3b8","EOI":"#38BDF8" };

const DK = { bg:"#0f172a",sb:"#0a101e",panel:"#131f33",b1:"#1e2e45",b2:"#1e293b",text:"#e2e8f0",muted:"#94a3b8",dim:"#64748b",inp:"#0f172a",chip:"#1e293b",stat:"#1e293b" };
const LT = { bg:"#f1f5f9",sb:"#ffffff",panel:"#ffffff",b1:"#e2e8f0",b2:"#e2e8f0",text:"#0f172a",muted:"#475569",dim:"#94a3b8",inp:"#f8fafc",chip:"#e2e8f0",stat:"#f1f5f9" };

const defaultTender = { id:"", f10ref:"", stage:"", projectTitle:"", client:"", location:"", duration:"", clarDeadline:"", siteVisit:"", initialSub:"", latestSub:"", totalBUA:"", amount:"", status:"Under Estimation", remarks:"", tenderBond:"" };
const defaultData   = { tasks:[], goals:[], schedule:[], classes:[], seminars:[], studySessions:[], assignments:[], tenders:[], notes:"" };

const uid      = () => Math.random().toString(36).slice(2,9);
const todayDay = () => ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"][new Date().getDay()];
const fmtDate  = () => new Date().toLocaleDateString("en-US",{weekday:"long",month:"long",day:"numeric",year:"numeric"});
const daysLeft = (d) => { if(!d) return null; return Math.ceil((new Date(d).getTime()-new Date(new Date().toDateString()).getTime())/86400000); };
const countdown = (n) => {
  if(n===null) return null;
  if(n<0)  return {label:`${Math.abs(n)}d overdue`,color:"#EF4444"};
  if(n===0) return {label:"Today",color:"#EF4444"};
  if(n<=3) return {label:`${n}d`,color:"#F59E0B"};
  return {label:`${n}d`,color:"#64748b"};
};
const fmtAED = (v) => { if(!v) return ""; const n=parseFloat(v); if(isNaN(n)) return v; return "AED "+n.toLocaleString(); };

// ── Shared sub-components (OUTSIDE App) ──────────────────────────────────────
function Panel({ title, accent, children, style, T }) {
  return (
    <div style={{background:T.panel,border:`1px solid ${T.b1}`,borderRadius:12,padding:"13px 14px 10px",display:"flex",flexDirection:"column",gap:9,overflow:"hidden",...style}}>
      <div style={{fontSize:"0.7rem",fontWeight:700,letterSpacing:"0.08em",textTransform:"uppercase",color:T.muted,borderLeft:`3px solid ${accent}`,paddingLeft:10}}>{title}</div>
      {children}
    </div>
  );
}
function AddRow({ children }) { return <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>{children}</div>; }
function Btn({ onClick, children, color }) { return <button style={{background:color||"#1d4ed8",color:"#fff",border:"none",borderRadius:6,padding:"5px 13px",fontSize:"0.81rem",cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap",fontWeight:600}} onClick={onClick}>{children}</button>; }
function DelBtn({ onClick, T }) { return <button style={{background:"none",border:"none",color:T.dim,fontSize:"1.1rem",cursor:"pointer",lineHeight:1,padding:"0 2px",flexShrink:0}} onClick={onClick}>×</button>; }
function EditBtn({ onClick, T }) { return <button style={{background:"none",border:`1px solid ${T.b2}`,color:T.dim,fontSize:"0.65rem",cursor:"pointer",borderRadius:4,padding:"1px 5px",flexShrink:0}} onClick={onClick}>edit</button>; }
function Empty({ T }) { return <li style={{color:T.dim,fontSize:"0.8rem",padding:"8px 0",listStyle:"none"}}>Nothing here yet.</li>; }
function CatTag({ cat, T }) { const c=CAT[cat]; return c?<span style={{fontSize:"0.65rem",fontWeight:700,padding:"2px 7px",borderRadius:20,background:c.color+"22",color:c.color,whiteSpace:"nowrap"}}>{c.label}</span>:null; }
function CatSelect({ value, onChange, T }) {
  const SE = {background:T.inp,border:`1px solid ${T.b2}`,borderRadius:6,padding:"5px 7px",fontSize:"0.81rem",color:T.muted,cursor:"pointer",fontFamily:"inherit"};
  return <select style={SE} value={value} onChange={e=>onChange(e.target.value)}>{Object.entries(CAT).map(([k,v])=><option key={k} value={k}>{v.label}</option>)}</select>;
}
function Countdown({ n, done }) { const c=countdown(n); return c&&!done?<span style={{fontSize:"0.66rem",fontWeight:600,color:c.color,whiteSpace:"nowrap"}}>{c.label}</span>:null; }

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [data,setData]         = useState(defaultData);
  const [flash,setFlash]       = useState(false);
  const [dark,setDark]         = useState(true);
  const [search,setSearch]     = useState("");
  const [studyTab,setStudyTab] = useState("schedule");
  const [editingId,setEditingId] = useState(null);
  const [editForm,setEditForm]   = useState({});
  const fileRef = useRef(null);

  const [task,setTask]     = useState({text:"",cat:"work",due:"",priority:"Medium"});
  const [goal,setGoal]     = useState({text:"",cat:"personal"});
  const [block,setBlock]   = useState({time:"9:00 AM",label:"",cat:"work"});
  const [cls,setCls]       = useState({subject:"",day:"Mon",time:"9:00 AM",room:"",type:"Lecture"});
  const [sem,setSem]       = useState({title:"",date:"",time:"9:00 AM",location:"",note:""});
  const [sess,setSess]     = useState({subject:"",duration:"",note:""});
  const [asgn,setAsgn]     = useState({subject:"",title:"",due:""});
  const [tender,setTender] = useState({...defaultTender});

  const T  = dark ? DK : LT;
  const I  = {background:T.inp,border:`1px solid ${T.b2}`,borderRadius:6,padding:"5px 8px",fontSize:"0.81rem",color:T.text,outline:"none",flex:1,minWidth:70,fontFamily:"inherit"};
  const SE = {background:T.inp,border:`1px solid ${T.b2}`,borderRadius:6,padding:"5px 7px",fontSize:"0.81rem",color:T.muted,cursor:"pointer",fontFamily:"inherit"};
  const LI = {display:"flex",alignItems:"center",gap:7,padding:"7px 6px",borderBottom:`1px solid ${T.b2}`,listStyle:"none",fontSize:"0.83rem"};

  useEffect(()=>{ try{ const s=localStorage.getItem(STORAGE_KEY); if(s) setData(JSON.parse(s)); }catch{} },[]);

  function save(next) {
    setData(next);
    try{ localStorage.setItem(STORAGE_KEY,JSON.stringify(next)); setFlash(true); setTimeout(()=>setFlash(false),1000); }catch{}
  }

  // ── Editing ────────────────────────────────────────────────────────────────
  function startEdit(item) { setEditingId(item.id); setEditForm({...item}); }
  function cancelEdit()    { setEditingId(null); setEditForm({}); }
  function saveEdit(type) {
    const next = {...data};
    next[type] = data[type].map(x => x.id===editingId ? {...editForm} : x);
    save(next); cancelEdit();
  }

  // ── CRUD ──────────────────────────────────────────────────────────────────
  const addTask  = () => { if(!task.text.trim()) return; save({...data,tasks:[{id:uid(),...task,done:false},...data.tasks]}); setTask({text:"",cat:"work",due:"",priority:"Medium"}); };
  const togTask  = (id) => save({...data,tasks:data.tasks.map(t=>t.id===id?{...t,done:!t.done}:t)});
  const delTask  = (id) => save({...data,tasks:data.tasks.filter(t=>t.id!==id)});
  const addGoal  = () => { if(!goal.text.trim()) return; save({...data,goals:[{id:uid(),...goal,done:false},...data.goals]}); setGoal({text:"",cat:"personal"}); };
  const togGoal  = (id) => save({...data,goals:data.goals.map(g=>g.id===id?{...g,done:!g.done}:g)});
  const delGoal  = (id) => save({...data,goals:data.goals.filter(g=>g.id!==id)});
  const addBlock = () => { if(!block.label.trim()) return; save({...data,schedule:[...data.schedule,{id:uid(),...block}].sort((a,b)=>TIME_BLOCKS.indexOf(a.time)-TIME_BLOCKS.indexOf(b.time))}); setBlock({time:"9:00 AM",label:"",cat:"work"}); };
  const delBlock = (id) => save({...data,schedule:data.schedule.filter(b=>b.id!==id)});
  const addClass = () => { if(!cls.subject.trim()) return; save({...data,classes:[...data.classes,{id:uid(),...cls}].sort((a,b)=>DAYS.indexOf(a.day)-DAYS.indexOf(b.day)||TIME_BLOCKS.indexOf(a.time)-TIME_BLOCKS.indexOf(b.time))}); setCls({subject:"",day:"Mon",time:"9:00 AM",room:"",type:"Lecture"}); };
  const delClass = (id) => save({...data,classes:data.classes.filter(c=>c.id!==id)});
  const addSem   = () => { if(!sem.title.trim()) return; save({...data,seminars:[...data.seminars,{id:uid(),...sem}].sort((a,b)=>new Date(a.date).getTime()-new Date(b.date).getTime())}); setSem({title:"",date:"",time:"9:00 AM",location:"",note:""}); };
  const delSem   = (id) => save({...data,seminars:data.seminars.filter(s=>s.id!==id)});
  const addSess  = () => { if(!sess.subject.trim()) return; save({...data,studySessions:[{id:uid(),...sess,date:new Date().toLocaleDateString()},...data.studySessions]}); setSess({subject:"",duration:"",note:""}); };
  const delSess  = (id) => save({...data,studySessions:data.studySessions.filter(s=>s.id!==id)});
  const addAsgn  = () => { if(!asgn.title.trim()) return; save({...data,assignments:[{id:uid(),...asgn,done:false},...data.assignments]}); setAsgn({subject:"",title:"",due:""}); };
  const togAsgn  = (id) => save({...data,assignments:data.assignments.map(a=>a.id===id?{...a,done:!a.done}:a)});
  const delAsgn  = (id) => save({...data,assignments:data.assignments.filter(a=>a.id!==id)});

  const addTender = () => {
    if(!tender.projectTitle.trim()&&!tender.f10ref.trim()) return;
    save({...data,tenders:[{...tender,id:uid()},...data.tenders]});
    setTender({...defaultTender});
  };
  const delTender = (id) => save({...data,tenders:data.tenders.filter(t=>t.id!==id)});

  // ── Excel Import ──────────────────────────────────────────────────────────
  function handleImport(e) {
    const file = e.target.files[0]; if(!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, {type:"binary",cellDates:true});
        const ws = wb.Sheets["Data"];
        const rows = XLSX.utils.sheet_to_json(ws, {header:1});
        const imported = [];
        for(let i=1; i<rows.length; i++) {
          const r = rows[i];
          if(!r[0]) continue;
          const fmtD = (v) => { if(!v) return ""; if(v instanceof Date) return v.toISOString().split("T")[0]; return String(v); };
          imported.push({
            id: uid(),
            f10ref:        String(r[0]||""),
            stage:         String(r[1]||""),
            projectTitle:  String(r[2]||""),
            client:        String(r[3]||""),
            location:      String(r[4]||""),
            duration:      String(r[5]||""),
            clarDeadline:  fmtD(r[6]),
            siteVisit:     fmtD(r[7]),
            initialSub:    fmtD(r[8]),
            latestSub:     fmtD(r[9]),
            totalBUA:      String(r[10]||""),
            amount:        String(r[11]||""),
            status:        String(r[12]||""),
            remarks:       String(r[13]||""),
            tenderBond:    String(r[15]||""),
          });
        }
        save({...data,tenders:imported});
        alert(`✓ Imported ${imported.length} tenders`);
      } catch(err) { alert("Import failed: "+err.message); }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  }

  // ── Excel Export ──────────────────────────────────────────────────────────
  function handleExport() {
    const headers = ["F10 Ref","Stage","Project Title","Client","Location","Duration (in months)","Clarification Deadline","Site Visit Date","Initial Submission Date","Latest Submission Date","Total BUA(m2)","Tender Amount Submitted","Status","Remarks","Remarks1","Tender Bond"];
    const rows = data.tenders.map(t => [
      t.f10ref, t.stage, t.projectTitle, t.client, t.location,
      t.duration, t.clarDeadline, t.siteVisit, t.initialSub, t.latestSub,
      t.totalBUA, t.amount, t.status, t.remarks, "", t.tenderBond
    ]);
    const dataSheet = XLSX.utils.aoa_to_sheet([headers,...rows]);
    // Column widths matching original
    dataSheet["!cols"] = [8,15,60,25,15,10,14,12,16,16,10,18,14,30,10,10].map(w=>({wch:w}));

    const reportHeaders = ["F10 Ref","Stage","Project Title","Client","Location","Duration","Clarification","Site Visit","Initial Submission Date","Latest Submission Date","Total BUA(m2)","Tender Amount Submitted","Status"];
    const today = new Date().toLocaleDateString("en-GB",{day:"2-digit",month:"long",year:"numeric"});
    const reportData = [
      [`Tenders Report Under Estimation/Planning as of ${today}`,...Array(12).fill("")],
      reportHeaders,
      ...data.tenders.filter(t=>t.status==="Under Estimation"||t.status==="Takeoff"||t.status==="Pricing").map(t=>[
        t.f10ref,t.stage,t.projectTitle,t.client,t.location,t.duration,t.clarDeadline,t.siteVisit,t.initialSub,t.latestSub,t.totalBUA,t.amount,t.status
      ])
    ];
    const reportSheet = XLSX.utils.aoa_to_sheet(reportData);

    const priceHeaders = ["F10 Ref","Latest Offer Date","Latest Offer Amount"];
    const priceRows = data.tenders.map(t=>[t.f10ref, t.latestSub||"", t.amount||""]);
    const priceSheet = XLSX.utils.aoa_to_sheet([priceHeaders,...priceRows]);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, dataSheet, "Data");
    XLSX.utils.book_append_sheet(wb, reportSheet, "Report");
    XLSX.utils.book_append_sheet(wb, priceSheet, "Prices");
    XLSX.writeFile(wb, `Tender_Dashboard_v2.xlsx`);
  }

  const studyMins     = data.studySessions.reduce((s,x)=>s+(parseInt(x.duration)||0),0);
  const todayClasses  = data.classes.filter(c=>c.day===todayDay());
  const sq            = search.toLowerCase();
  const filteredTasks = data.tasks.filter(t=>!sq||t.text.toLowerCase().includes(sq));
  const filteredGoals = data.goals.filter(g=>!sq||g.text.toLowerCase().includes(sq));

  const subjectProgress = useMemo(()=>{
    const map={};
    data.studySessions.forEach(s=>{ map[s.subject]=(map[s.subject]||0)+(parseInt(s.duration)||0); });
    const max=Math.max(...Object.values(map),1);
    return Object.entries(map).map(([subject,mins])=>({subject,mins,pct:Math.round(mins/max*100)})).sort((a,b)=>b.mins-a.mins);
  },[data.studySessions]);

  // ── Inline edit renderer ──────────────────────────────────────────────────
  function EditRow({ type, fields, onSave }) {
    return (
      <li style={{...LI,flexWrap:"wrap",background:T.stat,borderRadius:6,padding:"8px",gap:6}}>
        {fields.map(([key,placeholder,width,isSelect,opts])=>(
          isSelect
            ? <select key={key} style={{...SE,width:width||"auto"}} value={editForm[key]||""} onChange={e=>setEditForm({...editForm,[key]:e.target.value})}>
                {opts.map(o=><option key={o}>{o}</option>)}
              </select>
            : <input key={key} style={{...I,flex:"none",width:width||100}} placeholder={placeholder} value={editForm[key]||""} onChange={e=>setEditForm({...editForm,[key]:e.target.value})}/>
        ))}
        <div style={{display:"flex",gap:5}}>
          <Btn onClick={onSave} color="#10B981">Save</Btn>
          <Btn onClick={cancelEdit} color="#64748b">Cancel</Btn>
        </div>
      </li>
    );
  }

  return (
    <div style={{display:"flex",width:"100vw",height:"100vh",overflow:"hidden",background:T.bg,color:T.text,fontFamily:"'Outfit','Segoe UI',sans-serif",boxSizing:"border-box"}}>

      {/* SIDEBAR */}
      <aside style={{width:224,minWidth:224,flexShrink:0,background:T.sb,borderRight:`1px solid ${T.b2}`,padding:"18px 14px",display:"flex",flexDirection:"column",gap:14,overflowY:"auto"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:9,height:9,borderRadius:"50%",background:"#3B82F6",boxShadow:"0 0 8px #3B82F6",flexShrink:0}}/>
          <div>
            <div style={{fontWeight:700,fontSize:"0.95rem",letterSpacing:"-0.02em",color:T.text}}>Dashboard</div>
            <div style={{fontSize:"0.6rem",color:T.dim,marginTop:1,lineHeight:1.4}}>{fmtDate()}</div>
          </div>
        </div>
        <button style={{background:T.stat,border:`1px solid ${T.b2}`,borderRadius:6,padding:"6px 10px",fontSize:"0.74rem",cursor:"pointer",fontFamily:"inherit",color:T.muted,textAlign:"left"}} onClick={()=>setDark(!dark)}>
          {dark?"☀️  Light mode":"🌙  Dark mode"}
        </button>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:7}}>
          {[
            [data.tasks.filter(t=>!t.done).length,"Tasks","#3B82F6"],
            [studyMins,"Study min","#F59E0B"],
            [data.goals.filter(g=>!g.done).length,"Goals","#10B981"],
            [data.schedule.length,"Blocks","#A78BFA"],
            [data.tenders.filter(t=>t.status!=="Won"&&t.status!=="Lost"&&t.status!=="Regretted").length,"Active Tenders","#F472B6"],
            [data.assignments.filter(a=>!a.done).length,"Due","#EF4444"],
          ].map(([n,label,color])=>(
            <div key={label} style={{background:T.stat,borderRadius:8,padding:"9px 10px",borderTop:`2px solid ${color}`}}>
              <div style={{fontSize:"1.3rem",fontWeight:700,color}}>{n}</div>
              <div style={{fontSize:"0.6rem",color:T.dim,letterSpacing:"0.06em",textTransform:"uppercase",marginTop:2}}>{label}</div>
            </div>
          ))}
        </div>
        {todayClasses.length>0&&(
          <div style={{background:T.stat,borderRadius:8,border:`1px solid ${T.b2}`,padding:"10px"}}>
            <div style={{fontSize:"0.6rem",fontWeight:700,letterSpacing:"0.08em",textTransform:"uppercase",color:T.dim,marginBottom:7}}>Today's Classes</div>
            {todayClasses.map(c=>(
              <div key={c.id} style={{display:"flex",alignItems:"center",gap:5,marginBottom:5}}>
                <span style={{fontSize:"0.64rem",fontWeight:700,padding:"1px 5px",borderRadius:10,background:(TYPE_CLR[c.type]||"#888")+"22",color:TYPE_CLR[c.type]||"#888"}}>{c.type}</span>
                <span style={{flex:1,fontSize:"0.75rem",color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.subject}</span>
                <span style={{fontSize:"0.62rem",color:T.dim,whiteSpace:"nowrap"}}>{c.time}</span>
              </div>
            ))}
          </div>
        )}
        <div style={{display:"flex",flexDirection:"column",gap:5,flex:1}}>
          <div style={{fontSize:"0.6rem",fontWeight:700,letterSpacing:"0.08em",textTransform:"uppercase",color:T.dim}}>Quick Notes</div>
          <textarea style={{background:T.inp,border:`1px solid ${T.b2}`,borderRadius:6,padding:"7px 9px",fontSize:"0.79rem",resize:"none",outline:"none",fontFamily:"inherit",color:T.text,flex:1,minHeight:90,lineHeight:1.5}} placeholder="Jot something down…" value={data.notes||""} onChange={e=>save({...data,notes:e.target.value})}/>
        </div>
        <div style={{fontSize:"0.72rem",color:"#10B981",fontWeight:600,transition:"opacity 0.3s",opacity:flash?1:0}}>✓ Saved</div>
      </aside>

      {/* MAIN */}
      <main style={{flex:1,minWidth:0,overflow:"auto",padding:14,display:"flex",flexDirection:"column",gap:11}}>

        {/* Search */}
        <div style={{display:"flex",alignItems:"center",gap:8,background:T.panel,border:`1px solid ${T.b2}`,borderRadius:8,padding:"5px 12px"}}>
          <span style={{color:T.dim}}>🔍</span>
          <input style={{flex:1,fontSize:"0.83rem",outline:"none",fontFamily:"inherit",background:"transparent",color:T.text,border:"none"}} placeholder="Search tasks, goals, tenders…" value={search} onChange={e=>setSearch(e.target.value)}/>
          {search&&<button style={{background:"none",border:"none",color:T.dim,cursor:"pointer",fontSize:"1rem"}} onClick={()=>setSearch("")}>×</button>}
        </div>

        {/* ── TENDER TRACKER ── */}
        <div style={{background:T.panel,border:`1px solid ${T.b1}`,borderRadius:12,padding:"13px 14px 10px",display:"flex",flexDirection:"column",gap:10}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
            <div style={{fontSize:"0.7rem",fontWeight:700,letterSpacing:"0.08em",textTransform:"uppercase",color:T.muted,borderLeft:"3px solid #F472B6",paddingLeft:10}}>Tender Tracker</div>
            <div style={{display:"flex",gap:6}}>
              <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={handleImport}/>
              <Btn onClick={()=>fileRef.current?.click()} color="#1e3a5f">⬆ Import Excel</Btn>
              <Btn onClick={handleExport} color="#064e3b">⬇ Export Excel</Btn>
            </div>
          </div>

          {/* Add form */}
          <div style={{background:T.stat,borderRadius:8,padding:"10px",display:"flex",flexDirection:"column",gap:7}}>
            <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
              <input style={{...I,flex:"none",width:80}} placeholder="F10 Ref" value={tender.f10ref} onChange={e=>setTender({...tender,f10ref:e.target.value})}/>
              <select style={SE} value={tender.stage} onChange={e=>setTender({...tender,stage:e.target.value})}>{TENDER_STAGES.map(s=><option key={s} value={s}>{s||"-- Stage --"}</option>)}</select>
              <input style={{...I,flex:2,minWidth:200}} placeholder="Project Title" value={tender.projectTitle} onChange={e=>setTender({...tender,projectTitle:e.target.value})}/>
              <input style={{...I,flex:"none",width:120}} placeholder="Client" value={tender.client} onChange={e=>setTender({...tender,client:e.target.value})}/>
              <input style={{...I,flex:"none",width:100}} placeholder="Location" value={tender.location} onChange={e=>setTender({...tender,location:e.target.value})}/>
            </div>
            <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
              <input style={{...I,flex:"none",width:70}} placeholder="Duration" value={tender.duration} onChange={e=>setTender({...tender,duration:e.target.value})}/>
              <input type="date" style={{...I,flex:"none",width:130}} title="Clarification Deadline" value={tender.clarDeadline} onChange={e=>setTender({...tender,clarDeadline:e.target.value})}/>
              <input type="date" style={{...I,flex:"none",width:130}} title="Site Visit Date" value={tender.siteVisit} onChange={e=>setTender({...tender,siteVisit:e.target.value})}/>
              <input type="date" style={{...I,flex:"none",width:130}} title="Latest Submission Date" value={tender.latestSub} onChange={e=>setTender({...tender,latestSub:e.target.value})}/>
              <input style={{...I,flex:"none",width:100}} placeholder="Total BUA m²" value={tender.totalBUA} onChange={e=>setTender({...tender,totalBUA:e.target.value})}/>
              <input style={{...I,flex:"none",width:130}} placeholder="Tender Amount" value={tender.amount} onChange={e=>setTender({...tender,amount:e.target.value})}/>
              <select style={SE} value={tender.status} onChange={e=>setTender({...tender,status:e.target.value})}>{TENDER_STATUSES.map(s=><option key={s}>{s}</option>)}</select>
              <input style={I} placeholder="Remarks" value={tender.remarks} onChange={e=>setTender({...tender,remarks:e.target.value})}/>
              <Btn onClick={addTender}>Add</Btn>
            </div>
          </div>

          {/* Tender list */}
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:"0.78rem"}}>
              <thead>
                <tr style={{borderBottom:`2px solid ${T.b2}`}}>
                  {["F10 Ref","Stage","Project Title","Client","Location","Submission Date","Amount","Status","Remarks",""].map(h=>(
                    <th key={h} style={{textAlign:"left",padding:"5px 8px",color:T.muted,fontWeight:600,fontSize:"0.65rem",letterSpacing:"0.04em",textTransform:"uppercase",whiteSpace:"nowrap"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {data.tenders.filter(t=>!sq||t.projectTitle.toLowerCase().includes(sq)||t.f10ref.toLowerCase().includes(sq)||t.client.toLowerCase().includes(sq)).length===0
                  ? <tr><td colSpan={10} style={{color:T.dim,padding:"12px 8px",fontSize:"0.8rem"}}>No tenders yet. Add one above or import an Excel file.</td></tr>
                  : data.tenders.filter(t=>!sq||t.projectTitle.toLowerCase().includes(sq)||t.f10ref.toLowerCase().includes(sq)||t.client.toLowerCase().includes(sq)).map(t=>{
                    const isEditing = editingId===t.id;
                    const cd = countdown(daysLeft(t.latestSub));
                    return isEditing ? (
                      <tr key={t.id} style={{background:T.stat}}>
                        <td style={{padding:"4px"}}><input style={{...I,width:70}} value={editForm.f10ref||""} onChange={e=>setEditForm({...editForm,f10ref:e.target.value})}/></td>
                        <td style={{padding:"4px"}}><select style={SE} value={editForm.stage||""} onChange={e=>setEditForm({...editForm,stage:e.target.value})}>{TENDER_STAGES.map(s=><option key={s}>{s}</option>)}</select></td>
                        <td style={{padding:"4px"}}><input style={{...I,width:220}} value={editForm.projectTitle||""} onChange={e=>setEditForm({...editForm,projectTitle:e.target.value})}/></td>
                        <td style={{padding:"4px"}}><input style={{...I,width:100}} value={editForm.client||""} onChange={e=>setEditForm({...editForm,client:e.target.value})}/></td>
                        <td style={{padding:"4px"}}><input style={{...I,width:90}} value={editForm.location||""} onChange={e=>setEditForm({...editForm,location:e.target.value})}/></td>
                        <td style={{padding:"4px"}}><input type="date" style={{...I,width:130}} value={editForm.latestSub||""} onChange={e=>setEditForm({...editForm,latestSub:e.target.value})}/></td>
                        <td style={{padding:"4px"}}><input style={{...I,width:110}} value={editForm.amount||""} onChange={e=>setEditForm({...editForm,amount:e.target.value})}/></td>
                        <td style={{padding:"4px"}}><select style={SE} value={editForm.status||""} onChange={e=>setEditForm({...editForm,status:e.target.value})}>{TENDER_STATUSES.map(s=><option key={s}>{s}</option>)}</select></td>
                        <td style={{padding:"4px"}}><input style={{...I,width:140}} value={editForm.remarks||""} onChange={e=>setEditForm({...editForm,remarks:e.target.value})}/></td>
                        <td style={{padding:"4px"}}><div style={{display:"flex",gap:4}}><Btn onClick={()=>saveEdit("tenders")} color="#10B981">✓</Btn><Btn onClick={cancelEdit} color="#64748b">✕</Btn></div></td>
                      </tr>
                    ) : (
                      <tr key={t.id} style={{borderBottom:`1px solid ${T.b2}`,transition:"background 0.15s"}} onMouseEnter={e=>e.currentTarget.style.background=T.stat} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                        <td style={{padding:"6px 8px",color:T.muted,fontWeight:600,whiteSpace:"nowrap"}}>{t.f10ref}</td>
                        <td style={{padding:"6px 8px",color:T.dim,whiteSpace:"nowrap",fontSize:"0.72rem"}}>{t.stage}</td>
                        <td style={{padding:"6px 8px",color:T.text,maxWidth:260,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t.projectTitle}</td>
                        <td style={{padding:"6px 8px",color:T.muted,whiteSpace:"nowrap"}}>{t.client}</td>
                        <td style={{padding:"6px 8px",color:T.dim,whiteSpace:"nowrap",fontSize:"0.72rem"}}>{t.location}</td>
                        <td style={{padding:"6px 8px",whiteSpace:"nowrap"}}>
                          <div style={{display:"flex",flexDirection:"column",gap:2}}>
                            <span style={{color:T.text,fontSize:"0.72rem"}}>{t.latestSub}</span>
                            {cd&&<span style={{fontSize:"0.65rem",fontWeight:600,color:cd.color}}>{cd.label}</span>}
                          </div>
                        </td>
                        <td style={{padding:"6px 8px",color:t.amount?"#10B981":T.dim,whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"}}>{t.amount?fmtAED(t.amount):"—"}</td>
                        <td style={{padding:"6px 8px"}}>
                          <span style={{fontSize:"0.65rem",fontWeight:700,padding:"2px 8px",borderRadius:20,background:(STATUS_CLR[t.status]||"#888")+"22",color:STATUS_CLR[t.status]||"#888",whiteSpace:"nowrap"}}>{t.status}</span>
                        </td>
                        <td style={{padding:"6px 8px",color:T.dim,fontSize:"0.72rem",maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t.remarks}</td>
                        <td style={{padding:"6px 8px",whiteSpace:"nowrap"}}>
                          <div style={{display:"flex",gap:4}}>
                            <EditBtn onClick={()=>startEdit(t)} T={T}/>
                            <DelBtn onClick={()=>delTender(t.id)} T={T}/>
                          </div>
                        </td>
                      </tr>
                    );
                  })
                }
              </tbody>
            </table>
          </div>
        </div>

        {/* ROW 2: Tasks + Schedule */}
        <div style={{display:"flex",gap:11,minWidth:0}}>
          <Panel title="Work Reminders" accent={CAT.work.color} style={{flex:1.3,minWidth:0}} T={T}>
            <AddRow>
              <input style={I} placeholder="New task…" value={task.text} onChange={e=>setTask({...task,text:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addTask()}/>
              <CatSelect value={task.cat} onChange={v=>setTask({...task,cat:v})} T={T}/>
              <select style={SE} value={task.priority} onChange={e=>setTask({...task,priority:e.target.value})}>{PRIORITIES.map(p=><option key={p}>{p}</option>)}</select>
              <input type="date" style={{...I,flex:"none",width:128}} value={task.due} onChange={e=>setTask({...task,due:e.target.value})}/>
              <Btn onClick={addTask}>Add</Btn>
            </AddRow>
            <ul style={{margin:0,padding:0,overflowY:"auto",maxHeight:160,display:"flex",flexDirection:"column"}}>
              {filteredTasks.length===0&&<Empty T={T}/>}
              {filteredTasks.map(t=>
                editingId===t.id ? (
                  <li key={t.id} style={{...LI,flexWrap:"wrap",background:T.stat,borderRadius:6,gap:5}}>
                    <input style={{...I,flex:2}} value={editForm.text||""} onChange={e=>setEditForm({...editForm,text:e.target.value})}/>
                    <select style={SE} value={editForm.cat||""} onChange={e=>setEditForm({...editForm,cat:e.target.value})}>{Object.keys(CAT).map(k=><option key={k}>{k}</option>)}</select>
                    <select style={SE} value={editForm.priority||""} onChange={e=>setEditForm({...editForm,priority:e.target.value})}>{PRIORITIES.map(p=><option key={p}>{p}</option>)}</select>
                    <input type="date" style={{...I,flex:"none",width:120}} value={editForm.due||""} onChange={e=>setEditForm({...editForm,due:e.target.value})}/>
                    <Btn onClick={()=>saveEdit("tasks")} color="#10B981">✓</Btn>
                    <Btn onClick={cancelEdit} color="#64748b">✕</Btn>
                  </li>
                ) : (
                  <li key={t.id} style={{...LI,opacity:t.done?0.4:1}}>
                    <input type="checkbox" checked={t.done} onChange={()=>togTask(t.id)} style={{cursor:"pointer",accentColor:CAT.work.color,flexShrink:0}}/>
                    <span style={{width:7,height:7,borderRadius:"50%",background:PRIO_CLR[t.priority]||"#888",flexShrink:0}}/>
                    <span style={{flex:1,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",textDecoration:t.done?"line-through":"none"}}>{t.text}</span>
                    <CatTag cat={t.cat} T={T}/><Countdown n={daysLeft(t.due)} done={t.done}/>
                    <EditBtn onClick={()=>startEdit(t)} T={T}/><DelBtn onClick={()=>delTask(t.id)} T={T}/>
                  </li>
                )
              )}
            </ul>
          </Panel>

          <Panel title="Daily Schedule" accent="#A78BFA" style={{flex:0.75,minWidth:0}} T={T}>
            <AddRow>
              <select style={SE} value={block.time} onChange={e=>setBlock({...block,time:e.target.value})}>{TIME_BLOCKS.map(t=><option key={t}>{t}</option>)}</select>
              <input style={I} placeholder="Label…" value={block.label} onChange={e=>setBlock({...block,label:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addBlock()}/>
              <CatSelect value={block.cat} onChange={v=>setBlock({...block,cat:v})} T={T}/>
              <Btn onClick={addBlock}>Add</Btn>
            </AddRow>
            <ul style={{margin:0,padding:0,overflowY:"auto",maxHeight:160,display:"flex",flexDirection:"column"}}>
              {data.schedule.length===0&&<Empty T={T}/>}
              {data.schedule.map(b=>
                editingId===b.id ? (
                  <li key={b.id} style={{...LI,flexWrap:"wrap",background:T.stat,borderRadius:6,gap:5}}>
                    <select style={SE} value={editForm.time||""} onChange={e=>setEditForm({...editForm,time:e.target.value})}>{TIME_BLOCKS.map(t=><option key={t}>{t}</option>)}</select>
                    <input style={I} value={editForm.label||""} onChange={e=>setEditForm({...editForm,label:e.target.value})}/>
                    <select style={SE} value={editForm.cat||""} onChange={e=>setEditForm({...editForm,cat:e.target.value})}>{Object.keys(CAT).map(k=><option key={k}>{k}</option>)}</select>
                    <Btn onClick={()=>saveEdit("schedule")} color="#10B981">✓</Btn>
                    <Btn onClick={cancelEdit} color="#64748b">✕</Btn>
                  </li>
                ) : (
                  <li key={b.id} style={{...LI,borderLeft:`2px solid ${CAT[b.cat]?.color}`,paddingLeft:10}}>
                    <span style={{fontSize:"0.69rem",color:T.dim,width:60,flexShrink:0}}>{b.time}</span>
                    <span style={{flex:1,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{b.label}</span>
                    <CatTag cat={b.cat} T={T}/><EditBtn onClick={()=>startEdit(b)} T={T}/><DelBtn onClick={()=>delBlock(b.id)} T={T}/>
                  </li>
                )
              )}
            </ul>
          </Panel>
        </div>

        {/* ROW 3: Goals + Study */}
        <div style={{display:"flex",gap:11,minWidth:0}}>
          <Panel title="Personal Goals" accent={CAT.personal.color} style={{flex:0.75,minWidth:0}} T={T}>
            <AddRow>
              <input style={I} placeholder="New goal…" value={goal.text} onChange={e=>setGoal({...goal,text:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addGoal()}/>
              <CatSelect value={goal.cat} onChange={v=>setGoal({...goal,cat:v})} T={T}/>
              <Btn onClick={addGoal}>Add</Btn>
            </AddRow>
            <ul style={{margin:0,padding:0,overflowY:"auto",maxHeight:160,display:"flex",flexDirection:"column"}}>
              {filteredGoals.length===0&&<Empty T={T}/>}
              {filteredGoals.map(g=>
                editingId===g.id ? (
                  <li key={g.id} style={{...LI,flexWrap:"wrap",background:T.stat,borderRadius:6,gap:5}}>
                    <input style={I} value={editForm.text||""} onChange={e=>setEditForm({...editForm,text:e.target.value})}/>
                    <select style={SE} value={editForm.cat||""} onChange={e=>setEditForm({...editForm,cat:e.target.value})}>{Object.keys(CAT).map(k=><option key={k}>{k}</option>)}</select>
                    <Btn onClick={()=>saveEdit("goals")} color="#10B981">✓</Btn>
                    <Btn onClick={cancelEdit} color="#64748b">✕</Btn>
                  </li>
                ) : (
                  <li key={g.id} style={{...LI,opacity:g.done?0.4:1}}>
                    <input type="checkbox" checked={g.done} onChange={()=>togGoal(g.id)} style={{cursor:"pointer",accentColor:CAT.personal.color,flexShrink:0}}/>
                    <span style={{flex:1,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",textDecoration:g.done?"line-through":"none"}}>{g.text}</span>
                    <CatTag cat={g.cat} T={T}/><EditBtn onClick={()=>startEdit(g)} T={T}/><DelBtn onClick={()=>delGoal(g.id)} T={T}/>
                  </li>
                )
              )}
            </ul>
          </Panel>

          {/* Study */}
          <div style={{background:T.panel,border:`1px solid ${T.b1}`,borderRadius:12,padding:"13px 14px 10px",display:"flex",flexDirection:"column",gap:9,overflow:"hidden",flex:1.5,minWidth:0}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:6}}>
              <div style={{fontSize:"0.7rem",fontWeight:700,letterSpacing:"0.08em",textTransform:"uppercase",color:T.muted,borderLeft:`3px solid ${CAT.study.color}`,paddingLeft:10}}>Study</div>
              <div style={{display:"flex",gap:3,flexWrap:"wrap"}}>
                {[["schedule","Schedule"],["seminars","Seminars"],["assignments","Assignments"],["log","Session Log"],["progress","Progress"]].map(([k,label])=>(
                  <button key={k} style={{background:studyTab===k?T.stat:"none",border:`1px solid ${studyTab===k?T.b2:"transparent"}`,borderRadius:6,padding:"3px 9px",fontSize:"0.69rem",cursor:"pointer",color:studyTab===k?T.text:T.dim,fontFamily:"inherit"}} onClick={()=>setStudyTab(k)}>{label}</button>
                ))}
              </div>
            </div>

            {studyTab==="schedule"&&<>
              <AddRow>
                <input style={I} placeholder="Subject / Course…" value={cls.subject} onChange={e=>setCls({...cls,subject:e.target.value})}/>
                <select style={SE} value={cls.day} onChange={e=>setCls({...cls,day:e.target.value})}>{DAYS.map(d=><option key={d}>{d}</option>)}</select>
                <select style={SE} value={cls.time} onChange={e=>setCls({...cls,time:e.target.value})}>{TIME_BLOCKS.map(t=><option key={t}>{t}</option>)}</select>
                <select style={SE} value={cls.type} onChange={e=>setCls({...cls,type:e.target.value})}>{CLASS_TYPES.map(t=><option key={t}>{t}</option>)}</select>
                <input style={{...I,flex:"none",width:75}} placeholder="Room" value={cls.room} onChange={e=>setCls({...cls,room:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addClass()}/>
                <Btn onClick={addClass}>Add</Btn>
              </AddRow>
              <ul style={{margin:0,padding:0,overflowY:"auto",maxHeight:160,display:"flex",flexDirection:"column"}}>
                {data.classes.length===0&&<Empty T={T}/>}
                {data.classes.map(c=>(
                  <li key={c.id} style={{...LI,borderLeft:`2px solid ${TYPE_CLR[c.type]||"#888"}`,paddingLeft:10}}>
                    <span style={{fontSize:"0.65rem",fontWeight:700,color:"#38BDF8",background:"#0c2a3d",borderRadius:4,padding:"2px 5px",flexShrink:0}}>{c.day}</span>
                    <span style={{fontSize:"0.69rem",color:T.dim,width:60,flexShrink:0}}>{c.time}</span>
                    <span style={{flex:1,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.subject}</span>
                    <span style={{fontSize:"0.64rem",fontWeight:700,padding:"2px 6px",borderRadius:10,background:(TYPE_CLR[c.type]||"#888")+"22",color:TYPE_CLR[c.type]||"#888",whiteSpace:"nowrap"}}>{c.type}</span>
                    {c.room&&<span style={{fontSize:"0.65rem",color:T.dim,background:T.chip,borderRadius:4,padding:"2px 5px"}}>{c.room}</span>}
                    <DelBtn onClick={()=>delClass(c.id)} T={T}/>
                  </li>
                ))}
              </ul>
            </>}
            {studyTab==="seminars"&&<>
              <AddRow>
                <input style={I} placeholder="Seminar title…" value={sem.title} onChange={e=>setSem({...sem,title:e.target.value})}/>
                <input type="date" style={{...I,flex:"none",width:128}} value={sem.date} onChange={e=>setSem({...sem,date:e.target.value})}/>
                <select style={SE} value={sem.time} onChange={e=>setSem({...sem,time:e.target.value})}>{TIME_BLOCKS.map(t=><option key={t}>{t}</option>)}</select>
                <input style={{...I,flex:"none",width:95}} placeholder="Location" value={sem.location} onChange={e=>setSem({...sem,location:e.target.value})}/>
                <input style={I} placeholder="Note" value={sem.note} onChange={e=>setSem({...sem,note:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addSem()}/>
                <Btn onClick={addSem}>Add</Btn>
              </AddRow>
              <ul style={{margin:0,padding:0,overflowY:"auto",maxHeight:160,display:"flex",flexDirection:"column"}}>
                {data.seminars.length===0&&<Empty T={T}/>}
                {data.seminars.map(s=>{
                  const cd=countdown(daysLeft(s.date));
                  const past=s.date&&new Date(s.date)<new Date(new Date().toDateString());
                  return (
                    <li key={s.id} style={{...LI,opacity:past?0.4:1}}>
                      <span style={{fontSize:"0.64rem",fontWeight:700,padding:"2px 6px",borderRadius:10,background:"#A78BFA22",color:"#A78BFA",whiteSpace:"nowrap"}}>Seminar</span>
                      <span style={{flex:1,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}><strong>{s.title}</strong>{s.note?<span style={{color:T.dim}}> — {s.note}</span>:""}</span>
                      {s.date&&<span style={{fontSize:"0.65rem",color:T.dim,background:T.chip,borderRadius:4,padding:"2px 5px"}}>{s.date}</span>}
                      {cd&&!past&&<span style={{fontSize:"0.66rem",fontWeight:600,color:cd.color}}>{cd.label}</span>}
                      <DelBtn onClick={()=>delSem(s.id)} T={T}/>
                    </li>
                  );
                })}
              </ul>
            </>}
            {studyTab==="assignments"&&<>
              <AddRow>
                <input style={{...I,flex:"none",width:110}} placeholder="Subject…" value={asgn.subject} onChange={e=>setAsgn({...asgn,subject:e.target.value})}/>
                <input style={I} placeholder="Assignment title…" value={asgn.title} onChange={e=>setAsgn({...asgn,title:e.target.value})}/>
                <input type="date" style={{...I,flex:"none",width:128}} value={asgn.due} onChange={e=>setAsgn({...asgn,due:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addAsgn()}/>
                <Btn onClick={addAsgn}>Add</Btn>
              </AddRow>
              <ul style={{margin:0,padding:0,overflowY:"auto",maxHeight:160,display:"flex",flexDirection:"column"}}>
                {data.assignments.length===0&&<Empty T={T}/>}
                {data.assignments.map(a=>(
                  <li key={a.id} style={{...LI,opacity:a.done?0.4:1}}>
                    <input type="checkbox" checked={a.done} onChange={()=>togAsgn(a.id)} style={{cursor:"pointer",accentColor:CAT.study.color,flexShrink:0}}/>
                    <span style={{fontSize:"0.7rem",fontWeight:700,color:CAT.study.color,background:CAT.study.color+"22",borderRadius:4,padding:"2px 6px",whiteSpace:"nowrap"}}>{a.subject}</span>
                    <span style={{flex:1,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",textDecoration:a.done?"line-through":"none"}}>{a.title}</span>
                    <Countdown n={daysLeft(a.due)} done={a.done}/><DelBtn onClick={()=>delAsgn(a.id)} T={T}/>
                  </li>
                ))}
              </ul>
            </>}
            {studyTab==="log"&&<>
              <AddRow>
                <input style={I} placeholder="Subject / topic…" value={sess.subject} onChange={e=>setSess({...sess,subject:e.target.value})}/>
                <input style={{...I,flex:"none",width:75}} placeholder="Mins" type="number" value={sess.duration} onChange={e=>setSess({...sess,duration:e.target.value})}/>
                <input style={I} placeholder="Note (opt.)" value={sess.note} onChange={e=>setSess({...sess,note:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addSess()}/>
                <Btn onClick={addSess}>Log</Btn>
              </AddRow>
              <div style={{fontSize:"0.7rem",color:CAT.study.color,fontWeight:600}}>{studyMins} min total</div>
              <ul style={{margin:0,padding:0,overflowY:"auto",maxHeight:140,display:"flex",flexDirection:"column"}}>
                {data.studySessions.length===0&&<Empty T={T}/>}
                {data.studySessions.map(s=>(
                  <li key={s.id} style={{...LI}}>
                    <span style={{width:7,height:7,borderRadius:"50%",background:CAT.study.color,flexShrink:0}}/>
                    <span style={{flex:1,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}><strong>{s.subject}</strong>{s.note?` — ${s.note}`:""}</span>
                    {s.duration&&<span style={{fontSize:"0.65rem",color:T.dim,background:T.chip,borderRadius:4,padding:"2px 5px"}}>{s.duration}m</span>}
                    <span style={{fontSize:"0.65rem",color:T.dim,background:T.chip,borderRadius:4,padding:"2px 5px"}}>{s.date}</span>
                    <DelBtn onClick={()=>delSess(s.id)} T={T}/>
                  </li>
                ))}
              </ul>
            </>}
            {studyTab==="progress"&&(
              subjectProgress.length===0?<Empty T={T}/>:
              <div style={{display:"flex",flexDirection:"column",gap:10,overflowY:"auto",maxHeight:180}}>
                {subjectProgress.map(({subject,mins,pct})=>(
                  <div key={subject}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                      <span style={{fontSize:"0.82rem",color:T.text,fontWeight:600}}>{subject}</span>
                      <span style={{fontSize:"0.72rem",color:T.dim}}>{mins} min</span>
                    </div>
                    <div style={{height:6,borderRadius:3,background:T.chip}}>
                      <div style={{height:"100%",width:`${pct}%`,background:CAT.study.color,borderRadius:3}}/>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* ── FULL WEEKLY CALENDAR ── */}
        <div style={{background:T.panel,border:`1px solid ${T.b1}`,borderRadius:12,padding:"13px 14px 14px",flex:1}}>
          <div style={{fontSize:"0.7rem",fontWeight:700,letterSpacing:"0.08em",textTransform:"uppercase",color:T.muted,borderLeft:"3px solid #38BDF8",paddingLeft:10,marginBottom:12}}>Weekly Class Schedule</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:8,height:"calc(100% - 30px)"}}>
            {DAYS.map(day=>{
              const items = data.classes.filter(c=>c.day===day).sort((a,b)=>TIME_BLOCKS.indexOf(a.time)-TIME_BLOCKS.indexOf(b.time));
              const isToday = day===todayDay();
              return (
                <div key={day} style={{background:isToday?T.stat:T.inp,borderRadius:10,border:`1.5px solid ${isToday?"#38BDF8":T.b2}`,padding:"10px 8px",display:"flex",flexDirection:"column",gap:6,minHeight:180}}>
                  <div style={{fontSize:"0.7rem",fontWeight:700,letterSpacing:"0.06em",textTransform:"uppercase",color:isToday?"#38BDF8":T.dim,marginBottom:4,textAlign:"center"}}>{day}{isToday&&<span style={{marginLeft:5,fontSize:"0.55rem",background:"#38BDF8",color:"#0f172a",borderRadius:10,padding:"1px 5px"}}>TODAY</span>}</div>
                  {items.length===0
                    ? <div style={{fontSize:"0.68rem",color:T.b2,textAlign:"center",marginTop:8}}>Free</div>
                    : items.map(c=>(
                        <div key={c.id} style={{fontSize:"0.7rem",padding:"6px 8px",borderRadius:7,background:(TYPE_CLR[c.type]||"#888")+"22",color:TYPE_CLR[c.type]||"#888",borderLeft:`3px solid ${TYPE_CLR[c.type]||"#888"}`,lineHeight:1.4}}>
                          <div style={{fontWeight:700,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.subject}</div>
                          <div style={{fontSize:"0.65rem",opacity:0.85,marginTop:2}}>{c.time}{c.room?` · ${c.room}`:""}</div>
                          <div style={{fontSize:"0.6rem",opacity:0.7,marginTop:1}}>{c.type}</div>
                        </div>
                      ))
                  }
                </div>
              );
            })}
          </div>
        </div>

      </main>
    </div>
  );
}

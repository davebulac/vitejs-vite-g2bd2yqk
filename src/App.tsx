// @ts-nocheck
import { useState, useEffect, useMemo } from "react";

// ── Types ────────────────────────────────────────────────────────────────────
const STORAGE_KEY = "dave-dashboard-v5";
const TIME_BLOCKS = ["5:00 AM","6:00 AM","7:00 AM","8:00 AM","9:00 AM","10:00 AM","11:00 AM","12:00 PM","1:00 PM","2:00 PM","3:00 PM","4:00 PM","5:00 PM","6:00 PM","7:00 PM","8:00 PM","9:00 PM","10:00 PM","11:00 PM"];
const DAYS        = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"];
const CLASS_TYPES = ["Lecture","Seminar","Lab","Tutorial","Workshop"];
const PRIORITIES  = ["High","Medium","Low"];
const TENDER_ST   = ["Takeoff","Pricing","Submitted","Won","Lost"];
const CAT         = { work:{label:"Work",color:"#3B82F6"}, study:{label:"Study",color:"#F59E0B"}, personal:{label:"Personal",color:"#10B981"} };
const PRIO_CLR    = { High:"#EF4444", Medium:"#F59E0B", Low:"#10B981" };
const TYPE_CLR    = { Lecture:"#3B82F6", Seminar:"#A78BFA", Lab:"#F59E0B", Tutorial:"#10B981", Workshop:"#F472B6" };
const STATUS_CLR  = { Takeoff:"#F59E0B", Pricing:"#3B82F6", Submitted:"#A78BFA", Won:"#10B981", Lost:"#64748b" };
const DK = { bg:"#0f172a",sb:"#0a101e",panel:"#131f33",b1:"#1e2e45",b2:"#1e293b",text:"#e2e8f0",muted:"#94a3b8",dim:"#64748b",inp:"#0f172a",chip:"#1e293b",stat:"#1e293b" };
const LT = { bg:"#f1f5f9",sb:"#ffffff",panel:"#ffffff",b1:"#e2e8f0",b2:"#e2e8f0",text:"#0f172a",muted:"#475569",dim:"#94a3b8",inp:"#f8fafc",chip:"#e2e8f0",stat:"#f1f5f9" };
const defaultData = { tasks:[], goals:[], schedule:[], classes:[], seminars:[], studySessions:[], assignments:[], tenders:[], notes:"" };

// ── Helpers ──────────────────────────────────────────────────────────────────
const uid      = () => Math.random().toString(36).slice(2,8);
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

// ── Sub-components (defined OUTSIDE App to prevent re-mount on keystroke) ───
function Panel({ title, accent, children, style, T }) {
  return (
    <div style={{background:T.panel,border:`1px solid ${T.b1}`,borderRadius:12,padding:"13px 14px 10px",display:"flex",flexDirection:"column",gap:9,overflow:"hidden",...style}}>
      <div style={{fontSize:"0.7rem",fontWeight:700,letterSpacing:"0.08em",textTransform:"uppercase",color:T.muted,borderLeft:`3px solid ${accent}`,paddingLeft:10}}>{title}</div>
      {children}
    </div>
  );
}
function AddRow({ children }) { return <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{children}</div>; }
function List({ children }) { return <ul style={{margin:0,padding:0,overflowY:"auto",maxHeight:185,display:"flex",flexDirection:"column"}}>{children}</ul>; }
function Btn({ onClick, children }) {
  return <button style={{background:"#1d4ed8",color:"#fff",border:"none",borderRadius:6,padding:"5px 13px",fontSize:"0.81rem",cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap",fontWeight:600}} onClick={onClick}>{children}</button>;
}
function DelBtn({ onClick, T }) {
  return <button style={{background:"none",border:"none",color:T.dim,fontSize:"1.1rem",cursor:"pointer",lineHeight:1,padding:"0 2px",flexShrink:0}} onClick={onClick}>×</button>;
}
function Empty({ T }) { return <li style={{color:T.dim,fontSize:"0.8rem",padding:"8px 0",listStyle:"none"}}>Nothing here yet.</li>; }
function CatTag({ cat, T }) {
  const c = CAT[cat]; if(!c) return null;
  return <span style={{fontSize:"0.65rem",fontWeight:700,padding:"2px 7px",borderRadius:20,background:c.color+"22",color:c.color,whiteSpace:"nowrap"}}>{c.label}</span>;
}
function CatSelect({ value, onChange, T }) {
  const SE = {background:T.inp,border:`1px solid ${T.b2}`,borderRadius:6,padding:"5px 7px",fontSize:"0.81rem",color:T.muted,cursor:"pointer",fontFamily:"inherit"};
  return <select style={SE} value={value} onChange={e=>onChange(e.target.value)}>{Object.entries(CAT).map(([k,v])=><option key={k} value={k}>{v.label}</option>)}</select>;
}
function Countdown({ n, done }) {
  const c = countdown(n);
  return c&&!done ? <span style={{fontSize:"0.66rem",fontWeight:600,color:c.color,whiteSpace:"nowrap"}}>{c.label}</span> : null;
}

// ── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [data,setData]         = useState(defaultData);
  const [flash,setFlash]       = useState(false);
  const [dark,setDark]         = useState(true);
  const [search,setSearch]     = useState("");
  const [studyTab,setStudyTab] = useState("schedule");
  const [cmd,setCmd]           = useState("");
  const [cmdBusy,setCmdBusy]   = useState(false);
  const [cmdMsg,setCmdMsg]     = useState({text:"",ok:true});

  const [task,setTask]     = useState({text:"",cat:"work",due:"",priority:"Medium"});
  const [goal,setGoal]     = useState({text:"",cat:"personal"});
  const [block,setBlock]   = useState({time:"9:00 AM",label:"",cat:"work"});
  const [cls,setCls]       = useState({subject:"",day:"Mon",time:"9:00 AM",room:"",type:"Lecture"});
  const [sem,setSem]       = useState({title:"",date:"",time:"9:00 AM",location:"",note:""});
  const [sess,setSess]     = useState({subject:"",duration:"",note:""});
  const [asgn,setAsgn]     = useState({subject:"",title:"",due:""});
  const [tender,setTender] = useState({name:"",status:"Takeoff",deadline:"",trades:""});

  const T  = dark ? DK : LT;
  const I  = {background:T.inp,border:`1px solid ${T.b2}`,borderRadius:6,padding:"5px 8px",fontSize:"0.81rem",color:T.text,outline:"none",flex:1,minWidth:70,fontFamily:"inherit"};
  const SE = {background:T.inp,border:`1px solid ${T.b2}`,borderRadius:6,padding:"5px 7px",fontSize:"0.81rem",color:T.muted,cursor:"pointer",fontFamily:"inherit"};
  const LI = {display:"flex",alignItems:"center",gap:7,padding:"7px 6px",borderBottom:`1px solid ${T.b2}`,listStyle:"none",fontSize:"0.83rem"};

  useEffect(()=>{
    try { const s=localStorage.getItem(STORAGE_KEY); if(s) setData(JSON.parse(s)); } catch{}
  },[]);

  function save(next) {
    setData(next);
    try { localStorage.setItem(STORAGE_KEY,JSON.stringify(next)); setFlash(true); setTimeout(()=>setFlash(false),1000); } catch{}
  }

  // ── AI Command ──────────────────────────────────────────────────────────────
  async function executeCommand() {
    if(!cmd.trim()||cmdBusy) return;
    setCmdBusy(true); setCmdMsg({text:"Thinking…",ok:true});
    const today = new Date().toISOString().split("T")[0];
    try {
      const res = await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST", headers:{"Content-Type":"application/json"},
        body:JSON.stringify({
          model:"claude-sonnet-4-20250514", max_tokens:600,
          system:`You are a dashboard assistant. Parse natural language and return ONE JSON object.
Today is ${today}. Convert relative dates to YYYY-MM-DD.
Actions: addTask→{text,cat(work/study/personal),due,priority(High/Medium/Low)}, addGoal→{text,cat}, addBlock→{time("9:00 AM"),label,cat}, addClass→{subject,day(Mon-Sun),time,room,type(Lecture/Seminar/Lab/Tutorial/Workshop)}, addSeminar→{title,date,time,location,note}, addSession→{subject,duration,note}, addAssignment→{subject,title,due}, addTender→{name,status(Takeoff/Pricing/Submitted/Won/Lost),deadline,trades}
Return ONLY: {"action":"...","data":{...},"message":"short confirmation"}`,
          messages:[{role:"user",content:cmd}]
        })
      });
      const json = await res.json();
      const raw = json.content?.[0]?.text||"";
      const parsed = JSON.parse(raw.replace(/```json|```/g,"").trim());
      const d = parsed.data||{};
      const next = {...data};
      switch(parsed.action){
        case "addTask": next.tasks=[{id:uid(),text:d.text||"",cat:d.cat||"work",due:d.due||"",priority:d.priority||"Medium",done:false},...data.tasks]; break;
        case "addGoal": next.goals=[{id:uid(),text:d.text||"",cat:d.cat||"personal",done:false},...data.goals]; break;
        case "addBlock": next.schedule=[...data.schedule,{id:uid(),time:d.time||"9:00 AM",label:d.label||"",cat:d.cat||"work"}].sort((a,b)=>TIME_BLOCKS.indexOf(a.time)-TIME_BLOCKS.indexOf(b.time)); break;
        case "addClass": next.classes=[...data.classes,{id:uid(),subject:d.subject||"",day:d.day||"Mon",time:d.time||"9:00 AM",room:d.room||"",type:d.type||"Lecture"}].sort((a,b)=>DAYS.indexOf(a.day)-DAYS.indexOf(b.day)||TIME_BLOCKS.indexOf(a.time)-TIME_BLOCKS.indexOf(b.time)); break;
        case "addSeminar": next.seminars=[...data.seminars,{id:uid(),title:d.title||"",date:d.date||"",time:d.time||"9:00 AM",location:d.location||"",note:d.note||""}].sort((a,b)=>new Date(a.date).getTime()-new Date(b.date).getTime()); break;
        case "addSession": next.studySessions=[{id:uid(),subject:d.subject||"",duration:d.duration||"",note:d.note||"",date:new Date().toLocaleDateString()},...data.studySessions]; break;
        case "addAssignment": next.assignments=[{id:uid(),subject:d.subject||"",title:d.title||"",due:d.due||"",done:false},...data.assignments]; break;
        case "addTender": next.tenders=[{id:uid(),name:d.name||"",status:d.status||"Takeoff",deadline:d.deadline||"",trades:d.trades||""},...data.tenders]; break;
      }
      save(next); setCmdMsg({text:parsed.message||"Done!",ok:true}); setCmd("");
    } catch { setCmdMsg({text:"Couldn't understand — try rephrasing.",ok:false}); }
    setCmdBusy(false);
  }

  // ── CRUD ─────────────────────────────────────────────────────────────────────
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
  const addTender  = () => { if(!tender.name.trim()) return; save({...data,tenders:[{id:uid(),...tender},...data.tenders]}); setTender({name:"",status:"Takeoff",deadline:"",trades:""}); };
  const togTSt     = (id,status) => save({...data,tenders:data.tenders.map(t=>t.id===id?{...t,status}:t)});
  const delTender  = (id) => save({...data,tenders:data.tenders.filter(t=>t.id!==id)});

  const studyMins     = data.studySessions.reduce((s,x)=>s+(parseInt(x.duration)||0),0);
  const todayClasses  = data.classes.filter(c=>c.day===todayDay());
  const sq            = search.toLowerCase();
  const filteredTasks = data.tasks.filter(t=>!sq||t.text.toLowerCase().includes(sq));
  const filteredGoals = data.goals.filter(g=>!sq||g.text.toLowerCase().includes(sq));
  const subjectProgress = useMemo(()=>{
    const map={};
    data.studySessions.forEach(s=>{ map[s.subject]=(map[s.subject]||0)+(parseInt(s.duration)||0); });
    const max=Math.max(...Object.values(map),1);
    return Object.entries(map).map(([subject,mins])=>({subject,mins:mins,pct:Math.round((mins/max)*100)})).sort((a,b)=>b.mins-a.mins);
  },[data.studySessions]);

  return (
    <div style={{display:"flex",width:"100vw",height:"100vh",overflow:"hidden",background:T.bg,color:T.text,fontFamily:"'Outfit','Segoe UI',sans-serif",boxSizing:"border-box"}}>

      {/* ── SIDEBAR ── */}
      <aside style={{width:230,minWidth:230,flexShrink:0,background:T.sb,borderRight:`1px solid ${T.b2}`,padding:"18px 14px",display:"flex",flexDirection:"column",gap:14,overflowY:"auto"}}>
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
          {[[data.tasks.filter(t=>!t.done).length,"Tasks","#3B82F6"],[studyMins,"Study min","#F59E0B"],[data.goals.filter(g=>!g.done).length,"Goals","#10B981"],[data.schedule.length,"Blocks","#A78BFA"],[data.tenders.filter(t=>t.status!=="Won"&&t.status!=="Lost").length,"Tenders","#F472B6"],[data.assignments.filter(a=>!a.done).length,"Due","#EF4444"]].map(([n,label,color])=>(
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

      {/* ── MAIN ── */}
      <main style={{flex:1,minWidth:0,overflow:"auto",padding:14,display:"flex",flexDirection:"column",gap:11}}>

        {/* AI Bar */}
        <div style={{background:dark?"#0d1f35":"#e8f0fe",border:`1.5px solid ${dark?"#1e3a5f":"#93c5fd"}`,borderRadius:10,padding:"10px 14px",display:"flex",flexDirection:"column",gap:7}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span>✦</span>
            <span style={{fontSize:"0.72rem",fontWeight:700,letterSpacing:"0.06em",textTransform:"uppercase",color:dark?"#60a5fa":"#1d4ed8"}}>AI Assistant</span>
            <span style={{fontSize:"0.68rem",color:T.dim}}>— tell me what to add</span>
          </div>
          <div style={{display:"flex",gap:8}}>
            <input style={{...I,fontSize:"0.87rem",padding:"7px 11px"}} placeholder='e.g. "Add high priority task: Review facade drawings for 10637, due Friday"' value={cmd} onChange={e=>setCmd(e.target.value)} onKeyDown={e=>e.key==="Enter"&&executeCommand()} disabled={cmdBusy}/>
            <button style={{background:cmdBusy?"#334155":"#1d4ed8",color:"#fff",border:"none",borderRadius:7,padding:"7px 18px",fontSize:"0.84rem",cursor:cmdBusy?"not-allowed":"pointer",fontFamily:"inherit",fontWeight:700,whiteSpace:"nowrap",opacity:cmdBusy?0.7:1}} onClick={executeCommand} disabled={cmdBusy}>{cmdBusy?"…":"Send"}</button>
          </div>
          {cmdMsg.text&&<div style={{fontSize:"0.76rem",color:cmdMsg.ok?(dark?"#86efac":"#15803d"):"#f87171"}}>{cmdMsg.ok?"✓ ":"✕ "}{cmdMsg.text}</div>}
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {["Add task: Review facade drawings for 10637, high priority, due Friday","Log 45 min study on Structural Analysis","Add seminar: Fire Safety Workshop, June 10, 2pm, Room B3","Add assignment: Thesis draft due next Monday"].map(ex=>(
              <button key={ex} style={{background:T.stat,border:`1px solid ${T.b2}`,borderRadius:5,padding:"3px 9px",fontSize:"0.68rem",cursor:"pointer",color:T.dim,fontFamily:"inherit",whiteSpace:"nowrap"}} onClick={()=>setCmd(ex)}>{ex.length>45?ex.slice(0,45)+"…":ex}</button>
            ))}
          </div>
        </div>

        {/* Search */}
        <div style={{display:"flex",alignItems:"center",gap:8,background:T.panel,border:`1px solid ${T.b2}`,borderRadius:8,padding:"5px 12px"}}>
          <span style={{color:T.dim}}>🔍</span>
          <input style={{flex:1,fontSize:"0.83rem",outline:"none",fontFamily:"inherit",background:"transparent",color:T.text,border:"none"}} placeholder="Search tasks & goals…" value={search} onChange={e=>setSearch(e.target.value)}/>
          {search&&<button style={{background:"none",border:"none",color:T.dim,cursor:"pointer",fontSize:"1rem"}} onClick={()=>setSearch("")}>×</button>}
        </div>

        {/* ROW 1 */}
        <div style={{display:"flex",gap:11,minWidth:0}}>

          {/* Tenders */}
          <Panel title="Tender Tracker" accent="#F472B6" style={{flex:1,minWidth:0}} T={T}>
            <AddRow>
              <input style={I} placeholder="Tender name / ref…" value={tender.name} onChange={e=>setTender({...tender,name:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addTender()}/>
              <select style={SE} value={tender.status} onChange={e=>setTender({...tender,status:e.target.value})}>{TENDER_ST.map(s=><option key={s}>{s}</option>)}</select>
              <input type="date" style={{...I,flex:"none",width:128}} value={tender.deadline} onChange={e=>setTender({...tender,deadline:e.target.value})}/>
              <input style={I} placeholder="Trades…" value={tender.trades} onChange={e=>setTender({...tender,trades:e.target.value})}/>
              <Btn onClick={addTender}>Add</Btn>
            </AddRow>
            <List>
              {data.tenders.length===0&&<Empty T={T}/>}
              {data.tenders.map(t=>{
                const cd=countdown(daysLeft(t.deadline));
                return (
                  <li key={t.id} style={{...LI,borderLeft:`2px solid ${STATUS_CLR[t.status]||"#888"}`,paddingLeft:10}}>
                    <div style={{flex:1,overflow:"hidden"}}>
                      <div style={{display:"flex",alignItems:"center",gap:5,flexWrap:"wrap"}}>
                        <span style={{color:T.text,fontWeight:600}}>{t.name}</span>
                        <select style={{...SE,fontSize:"0.67rem",padding:"1px 4px"}} value={t.status} onChange={e=>togTSt(t.id,e.target.value)}>{TENDER_ST.map(s=><option key={s}>{s}</option>)}</select>
                        {cd&&<span style={{fontSize:"0.66rem",fontWeight:600,color:cd.color}}>{cd.label}</span>}
                      </div>
                      {t.trades&&<div style={{fontSize:"0.69rem",color:T.dim,marginTop:2}}>{t.trades}</div>}
                    </div>
                    <DelBtn onClick={()=>delTender(t.id)} T={T}/>
                  </li>
                );
              })}
            </List>
          </Panel>

          {/* Tasks */}
          <Panel title="Work Reminders" accent={CAT.work.color} style={{flex:1.2,minWidth:0}} T={T}>
            <AddRow>
              <input style={I} placeholder="New task…" value={task.text} onChange={e=>setTask({...task,text:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addTask()}/>
              <CatSelect value={task.cat} onChange={v=>setTask({...task,cat:v})} T={T}/>
              <select style={SE} value={task.priority} onChange={e=>setTask({...task,priority:e.target.value})}>{PRIORITIES.map(p=><option key={p}>{p}</option>)}</select>
              <input type="date" style={{...I,flex:"none",width:128}} value={task.due} onChange={e=>setTask({...task,due:e.target.value})}/>
              <Btn onClick={addTask}>Add</Btn>
            </AddRow>
            <List>
              {filteredTasks.length===0&&<Empty T={T}/>}
              {filteredTasks.map(t=>(
                <li key={t.id} style={{...LI,opacity:t.done?0.4:1}}>
                  <input type="checkbox" checked={t.done} onChange={()=>togTask(t.id)} style={{cursor:"pointer",accentColor:CAT.work.color,flexShrink:0}}/>
                  <span style={{width:7,height:7,borderRadius:"50%",background:PRIO_CLR[t.priority]||"#888",flexShrink:0}}/>
                  <span style={{flex:1,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",textDecoration:t.done?"line-through":"none"}}>{t.text}</span>
                  <CatTag cat={t.cat} T={T}/><Countdown n={daysLeft(t.due)} done={t.done}/><DelBtn onClick={()=>delTask(t.id)} T={T}/>
                </li>
              ))}
            </List>
          </Panel>

          {/* Schedule */}
          <Panel title="Daily Schedule" accent="#A78BFA" style={{flex:0.85,minWidth:0}} T={T}>
            <AddRow>
              <select style={SE} value={block.time} onChange={e=>setBlock({...block,time:e.target.value})}>{TIME_BLOCKS.map(t=><option key={t}>{t}</option>)}</select>
              <input style={I} placeholder="Label…" value={block.label} onChange={e=>setBlock({...block,label:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addBlock()}/>
              <CatSelect value={block.cat} onChange={v=>setBlock({...block,cat:v})} T={T}/>
              <Btn onClick={addBlock}>Add</Btn>
            </AddRow>
            <List>
              {data.schedule.length===0&&<Empty T={T}/>}
              {data.schedule.map(b=>(
                <li key={b.id} style={{...LI,borderLeft:`2px solid ${CAT[b.cat]?.color}`,paddingLeft:10}}>
                  <span style={{fontSize:"0.69rem",color:T.dim,width:60,flexShrink:0}}>{b.time}</span>
                  <span style={{flex:1,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{b.label}</span>
                  <CatTag cat={b.cat} T={T}/><DelBtn onClick={()=>delBlock(b.id)} T={T}/>
                </li>
              ))}
            </List>
          </Panel>
        </div>

        {/* ROW 2 */}
        <div style={{display:"flex",gap:11,minWidth:0}}>

          {/* Goals */}
          <Panel title="Personal Goals" accent={CAT.personal.color} style={{flex:0.8,minWidth:0}} T={T}>
            <AddRow>
              <input style={I} placeholder="New goal…" value={goal.text} onChange={e=>setGoal({...goal,text:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addGoal()}/>
              <CatSelect value={goal.cat} onChange={v=>setGoal({...goal,cat:v})} T={T}/>
              <Btn onClick={addGoal}>Add</Btn>
            </AddRow>
            <List>
              {filteredGoals.length===0&&<Empty T={T}/>}
              {filteredGoals.map(g=>(
                <li key={g.id} style={{...LI,opacity:g.done?0.4:1}}>
                  <input type="checkbox" checked={g.done} onChange={()=>togGoal(g.id)} style={{cursor:"pointer",accentColor:CAT.personal.color,flexShrink:0}}/>
                  <span style={{flex:1,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",textDecoration:g.done?"line-through":"none"}}>{g.text}</span>
                  <CatTag cat={g.cat} T={T}/><DelBtn onClick={()=>delGoal(g.id)} T={T}/>
                </li>
              ))}
            </List>
          </Panel>

          {/* Study */}
          <div style={{background:T.panel,border:`1px solid ${T.b1}`,borderRadius:12,padding:"13px 14px 10px",display:"flex",flexDirection:"column",gap:9,overflow:"hidden",flex:1.6,minWidth:0}}>
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
              <List>
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
              </List>
            </>}

            {studyTab==="seminars"&&<>
              <AddRow>
                <input style={I} placeholder="Seminar title…" value={sem.title} onChange={e=>setSem({...sem,title:e.target.value})}/>
                <input type="date" style={{...I,flex:"none",width:128}} value={sem.date} onChange={e=>setSem({...sem,date:e.target.value})}/>
                <select style={SE} value={sem.time} onChange={e=>setSem({...sem,time:e.target.value})}>{TIME_BLOCKS.map(t=><option key={t}>{t}</option>)}</select>
                <input style={{...I,flex:"none",width:95}} placeholder="Location" value={sem.location} onChange={e=>setSem({...sem,location:e.target.value})}/>
                <input style={I} placeholder="Note (opt.)" value={sem.note} onChange={e=>setSem({...sem,note:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addSem()}/>
                <Btn onClick={addSem}>Add</Btn>
              </AddRow>
              <List>
                {data.seminars.length===0&&<Empty T={T}/>}
                {data.seminars.map(s=>{
                  const cd=countdown(daysLeft(s.date));
                  const past=s.date&&new Date(s.date)<new Date(new Date().toDateString());
                  return (
                    <li key={s.id} style={{...LI,opacity:past?0.4:1}}>
                      <span style={{fontSize:"0.64rem",fontWeight:700,padding:"2px 6px",borderRadius:10,background:"#A78BFA22",color:"#A78BFA",whiteSpace:"nowrap"}}>Seminar</span>
                      <span style={{flex:1,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}><strong>{s.title}</strong>{s.note?<span style={{color:T.dim}}> — {s.note}</span>:""}</span>
                      {s.date&&<span style={{fontSize:"0.65rem",color:T.dim,background:T.chip,borderRadius:4,padding:"2px 5px"}}>{s.date}</span>}
                      <span style={{fontSize:"0.65rem",color:T.dim,background:T.chip,borderRadius:4,padding:"2px 5px"}}>{s.time}</span>
                      {s.location&&<span style={{fontSize:"0.65rem",color:T.dim,background:T.chip,borderRadius:4,padding:"2px 5px"}}>{s.location}</span>}
                      {cd&&!past&&<span style={{fontSize:"0.66rem",fontWeight:600,color:cd.color}}>{cd.label}</span>}
                      <DelBtn onClick={()=>delSem(s.id)} T={T}/>
                    </li>
                  );
                })}
              </List>
            </>}

            {studyTab==="assignments"&&<>
              <AddRow>
                <input style={{...I,flex:"none",width:110}} placeholder="Subject…" value={asgn.subject} onChange={e=>setAsgn({...asgn,subject:e.target.value})}/>
                <input style={I} placeholder="Assignment title…" value={asgn.title} onChange={e=>setAsgn({...asgn,title:e.target.value})}/>
                <input type="date" style={{...I,flex:"none",width:128}} value={asgn.due} onChange={e=>setAsgn({...asgn,due:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addAsgn()}/>
                <Btn onClick={addAsgn}>Add</Btn>
              </AddRow>
              <List>
                {data.assignments.length===0&&<Empty T={T}/>}
                {data.assignments.map(a=>(
                  <li key={a.id} style={{...LI,opacity:a.done?0.4:1}}>
                    <input type="checkbox" checked={a.done} onChange={()=>togAsgn(a.id)} style={{cursor:"pointer",accentColor:CAT.study.color,flexShrink:0}}/>
                    <span style={{fontSize:"0.7rem",fontWeight:700,color:CAT.study.color,background:CAT.study.color+"22",borderRadius:4,padding:"2px 6px",whiteSpace:"nowrap"}}>{a.subject}</span>
                    <span style={{flex:1,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",textDecoration:a.done?"line-through":"none"}}>{a.title}</span>
                    <Countdown n={daysLeft(a.due)} done={a.done}/><DelBtn onClick={()=>delAsgn(a.id)} T={T}/>
                  </li>
                ))}
              </List>
            </>}

            {studyTab==="log"&&<>
              <AddRow>
                <input style={I} placeholder="Subject / topic…" value={sess.subject} onChange={e=>setSess({...sess,subject:e.target.value})}/>
                <input style={{...I,flex:"none",width:75}} placeholder="Mins" type="number" value={sess.duration} onChange={e=>setSess({...sess,duration:e.target.value})}/>
                <input style={I} placeholder="Note (opt.)" value={sess.note} onChange={e=>setSess({...sess,note:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addSess()}/>
                <Btn onClick={addSess}>Log</Btn>
              </AddRow>
              <div style={{fontSize:"0.7rem",color:CAT.study.color,fontWeight:600}}>{studyMins} min total</div>
              <List>
                {data.studySessions.length===0&&<Empty T={T}/>}
                {data.studySessions.map(s=>(
                  <li key={s.id} style={LI}>
                    <span style={{width:7,height:7,borderRadius:"50%",background:CAT.study.color,flexShrink:0}}/>
                    <span style={{flex:1,color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}><strong>{s.subject}</strong>{s.note?` — ${s.note}`:""}</span>
                    {s.duration&&<span style={{fontSize:"0.65rem",color:T.dim,background:T.chip,borderRadius:4,padding:"2px 5px"}}>{s.duration}m</span>}
                    <span style={{fontSize:"0.65rem",color:T.dim,background:T.chip,borderRadius:4,padding:"2px 5px"}}>{s.date}</span>
                    <DelBtn onClick={()=>delSess(s.id)} T={T}/>
                  </li>
                ))}
              </List>
            </>}

            {studyTab==="progress"&&(
              subjectProgress.length===0?<Empty T={T}/>:
              <div style={{display:"flex",flexDirection:"column",gap:10,overflowY:"auto",maxHeight:185}}>
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

        {/* ROW 3: Weekly View */}
        <div style={{background:T.panel,border:`1px solid ${T.b1}`,borderRadius:12,padding:"13px 14px 12px"}}>
          <div style={{fontSize:"0.7rem",fontWeight:700,letterSpacing:"0.08em",textTransform:"uppercase",color:T.muted,borderLeft:"3px solid #38BDF8",paddingLeft:10,marginBottom:10}}>Weekly Class View</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:7}}>
            {DAYS.map(day=>{
              const items=data.classes.filter(c=>c.day===day);
              const isToday=day===todayDay();
              return (
                <div key={day} style={{background:isToday?T.stat:T.inp,borderRadius:8,border:`1px solid ${isToday?"#38BDF8":T.b2}`,padding:"8px 7px",minHeight:56}}>
                  <div style={{fontSize:"0.63rem",fontWeight:700,letterSpacing:"0.06em",textTransform:"uppercase",color:isToday?"#38BDF8":T.dim,marginBottom:5}}>{day}</div>
                  {items.length===0?<div style={{fontSize:"0.63rem",color:T.b2}}>—</div>
                    :items.map(c=>(
                      <div key={c.id} style={{fontSize:"0.67rem",padding:"3px 5px",borderRadius:4,marginBottom:3,background:(TYPE_CLR[c.type]||"#888")+"22",color:TYPE_CLR[c.type]||"#888",lineHeight:1.35}}>
                        <div style={{fontWeight:600}}>{c.subject}</div>
                        <div style={{opacity:0.8}}>{c.time}</div>
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

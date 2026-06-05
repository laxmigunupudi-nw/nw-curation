import { useState, useEffect, useRef, useCallback } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";

const SUPABASE_URL = "https://hyoiwwmhrhkhjhrueuon.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imh5b2l3d21ocmhraGpocnVldW9uIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzkwMzgzNDksImV4cCI6MjA5NDYxNDM0OX0.DUvfR7q0qPpIe5gkDFVQf40JUhrNPEg98GVmb3HoMpg";
const sb = createClient(SUPABASE_URL, SUPABASE_ANON);

const css = `
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500;9..40,600&family=DM+Mono:wght@400;500&display=swap');
  *{box-sizing:border-box;margin:0;padding:0}
  :root{
    --bg:#f4f6fb;
    --bg2:#ffffff;
    --bg3:#f0f2f8;
    --bg4:#e4e7f2;
    --border:#dde1ef;
    --border2:#b8bfd4;
    --text:#1a1d2e;
    --text2:#52566b;
    --text3:#9296a8;
    --accent:#3b6ff5;
    --accent2:#2d5ee0;
    --ag:rgba(59,111,245,.1);
    --green:#0a7c55;
    --gbg:rgba(10,124,85,.09);
    --red:#c0392b;
    --rbg:rgba(192,57,43,.09);
    --amber:#b45309;
    --abg:rgba(180,83,9,.09);
    --purple:#6d4fc2;
    --pbg:rgba(109,79,194,.09);
    --r:8px;--r2:12px;--r3:16px;
    --f:'DM Sans',sans-serif;--mono:'DM Mono',monospace;
    --shadow:0 1px 4px rgba(0,0,0,.07);
    --shadow2:0 4px 24px rgba(0,0,0,.12);
  }
  html,body,#root{background:var(--bg);color:var(--text);font-family:var(--f);font-size:14px;line-height:1.6;min-height:100vh}
  input,textarea,select{background:#fff;border:1.5px solid var(--border);border-radius:var(--r);color:var(--text);font-family:var(--f);font-size:14px;padding:8px 12px;width:100%;outline:none;transition:border-color .18s,box-shadow .18s}
  input:focus,textarea:focus,select:focus{border-color:var(--accent);box-shadow:0 0 0 3px var(--ag)}
  input::placeholder,textarea::placeholder{color:var(--text3)}
  button{cursor:pointer;font-family:var(--f);font-size:14px;border:none;border-radius:var(--r);padding:8px 16px;transition:all .15s;font-weight:500;line-height:1.4}
  .bp{background:var(--accent);color:#fff;box-shadow:0 1px 3px rgba(59,111,245,.3)}
  .bp:hover{background:var(--accent2);transform:translateY(-1px)}
  .bp:active{transform:translateY(0)}
  .bp:disabled{opacity:.55;cursor:not-allowed;transform:none;box-shadow:none}
  .bg{background:#fff;color:var(--text2);border:1.5px solid var(--border)}
  .bg:hover{border-color:var(--border2);color:var(--text);background:var(--bg3)}
  .bg:disabled{opacity:.5;cursor:not-allowed}
  .bd{background:var(--rbg);color:var(--red);border:1.5px solid rgba(192,57,43,.3)}
  .bd:hover{background:rgba(192,57,43,.15)}
  .bsm{padding:5px 11px;font-size:13px}
  .bxs{padding:3px 9px;font-size:12px}
  .card{background:#fff;border:1.5px solid var(--border);border-radius:var(--r2);padding:18px;box-shadow:var(--shadow)}
  .badge{display:inline-flex;align-items:center;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:600;letter-spacing:.01em}
  .b-green{background:var(--gbg);color:var(--green)}
  .b-red{background:var(--rbg);color:var(--red)}
  .b-amber{background:var(--abg);color:var(--amber)}
  .b-blue{background:var(--ag);color:var(--accent)}
  .b-purple{background:var(--pbg);color:var(--purple)}
  .b-gray{background:var(--bg4);color:var(--text2)}
  table{width:100%;border-collapse:collapse}
  th{text-align:left;padding:10px 14px;font-size:11px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;border-bottom:1.5px solid var(--border);background:var(--bg3)}
  td{padding:11px 14px;border-bottom:1px solid var(--border);color:var(--text2);font-size:13px}
  tr:last-child td{border-bottom:none}
  tr:hover td{background:var(--bg3);color:var(--text)}
  .tag{display:inline-block;padding:2px 7px;border-radius:5px;font-size:11px;font-weight:500;background:var(--bg4);color:var(--text2);border:1px solid var(--border)}
  .div{height:1.5px;background:var(--border);margin:16px 0}
  .mo{position:fixed;inset:0;background:rgba(20,24,50,.4);display:flex;align-items:center;justify-content:center;z-index:1000;padding:16px;backdrop-filter:blur(2px)}
  .md{background:#fff;border:1.5px solid var(--border);border-radius:var(--r3);padding:26px;width:100%;max-width:520px;max-height:92vh;overflow-y:auto;box-shadow:var(--shadow2)}
  .md-lg{max-width:760px}
  .fg{margin-bottom:14px}
  .fl{display:block;font-size:11px;font-weight:600;color:var(--text2);margin-bottom:5px;text-transform:uppercase;letter-spacing:.05em}
  .row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
  .row3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}
  .fx{display:flex}.g2{gap:8px}.g3{gap:12px}.g4{gap:16px}
  .ac{align-items:center}.jb{justify-content:space-between}.wrap{flex-wrap:wrap}
  .g2c{display:grid;grid-template-columns:1fr 1fr;gap:14px}
  .g3c{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px}
  .g4c{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
  .sm{font-size:13px}.xs{font-size:12px}.xl{font-size:22px}
  .m2{color:var(--text2)}.m3{color:var(--text3)}
  .mono{font-family:var(--mono)}
  .fw5{font-weight:500}.fw6{font-weight:600}
  .wf{width:100%}
  .mt2{margin-top:8px}.mt3{margin-top:12px}.mt4{margin-top:16px}
  .mb2{margin-bottom:8px}.mb3{margin-bottom:12px}.mb4{margin-bottom:16px}
  .sp{width:16px;height:16px;border:2px solid var(--border2);border-top-color:var(--accent);border-radius:50%;animation:spin .7s linear infinite;display:inline-block;vertical-align:middle;flex-shrink:0}
  @keyframes spin{to{transform:rotate(360deg)}}
  .fi{animation:fi .2s ease}
  @keyframes fi{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:translateY(0)}}
  .sidebar{width:214px;min-width:214px;background:#fff;border-right:1.5px solid var(--border);height:100vh;display:flex;flex-direction:column;position:fixed;left:0;top:0;z-index:10;box-shadow:2px 0 8px rgba(0,0,0,.05)}
  .mc{margin-left:214px;min-height:100vh;padding:28px;background:var(--bg)}
  .ni{display:flex;align-items:center;gap:9px;padding:9px 14px;border-radius:var(--r);color:var(--text2);cursor:pointer;transition:all .15s;font-size:13px;margin:1px 8px}
  .ni:hover{background:var(--bg3);color:var(--text)}
  .ni.act{background:var(--ag);color:var(--accent);font-weight:600}
  .ns{font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.1em;padding:14px 16px 5px}
  .pt{font-size:22px;font-weight:700;color:var(--text);margin-bottom:3px}
  .ps{font-size:13px;color:var(--text2);margin-bottom:22px}
  .sc{background:#fff;border:1.5px solid var(--border);border-radius:var(--r2);padding:18px 20px;box-shadow:var(--shadow)}
  .sv{font-size:28px;font-weight:700;color:var(--text);line-height:1.2}
  .sl{font-size:11px;font-weight:600;color:var(--text3);margin-top:4px;text-transform:uppercase;letter-spacing:.05em}
  .pb{height:6px;background:var(--bg4);border-radius:3px;overflow:hidden}
  .pf{height:100%;border-radius:3px;background:var(--accent);transition:width .4s ease}
  .tabs{display:flex;gap:3px;background:var(--bg3);padding:3px;border-radius:var(--r2);margin-bottom:18px;border:1.5px solid var(--border)}
  .tab{padding:7px 16px;border-radius:var(--r);cursor:pointer;font-size:13px;color:var(--text2);transition:all .15s;font-weight:500}
  .tab.act{background:#fff;color:var(--text);font-weight:600;box-shadow:0 1px 4px rgba(0,0,0,.1)}
  .tsb{width:194px;min-width:194px;background:#fff;border-right:1.5px solid var(--border);height:100vh;overflow-y:auto;padding:14px 8px;display:flex;flex-direction:column;box-shadow:2px 0 6px rgba(0,0,0,.04)}
  .td{width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;cursor:pointer;transition:all .15s;flex-shrink:0;border:1.5px solid transparent}
  .td-ns{background:var(--bg4);color:var(--text3);border-color:var(--border)}
  .td-ip{background:var(--abg);color:var(--amber);border-color:var(--amber)}
  .td-sb{background:var(--gbg);color:var(--green);border-color:var(--green)}
  .td-act{box-shadow:0 0 0 3px var(--accent);transform:scale(1.12)}
  .fsc{background:var(--gbg);border:1.5px solid rgba(10,124,85,.25);border-radius:var(--r);padding:10px 12px;margin-bottom:8px}
  .fsw{background:var(--rbg);border:1.5px solid rgba(192,57,43,.25);border-radius:var(--r);padding:10px 12px;margin-bottom:8px}
  .fsp{background:var(--abg);border:1.5px solid rgba(180,83,9,.25);border-radius:var(--r);padding:10px 12px;margin-bottom:8px}
  .cert-l3{background:var(--pbg);border:1.5px solid rgba(109,79,194,.35);color:var(--purple)}
  .cert-l2{background:var(--gbg);border:1.5px solid rgba(10,124,85,.35);color:var(--green)}
  .cert-l1{background:var(--abg);border:1.5px solid rgba(180,83,9,.35);color:var(--amber)}
  .cert-l0{background:var(--rbg);border:1.5px solid rgba(192,57,43,.35);color:var(--red)}
  .imt{width:130px;height:130px;object-fit:cover;border-radius:var(--r);border:1.5px solid var(--border);background:var(--bg3);transition:opacity .15s}
  .cb2{background:var(--bg3);border:1.5px solid var(--border);border-radius:var(--r);padding:10px 12px;font-family:var(--mono);font-size:12px;color:var(--accent);word-break:break-all;cursor:pointer;transition:background .15s}
  .cb2:hover{background:var(--bg4)}
  .al{padding:11px 14px;border-radius:var(--r);margin-bottom:14px;font-size:13px;line-height:1.5}
  .al-e{background:#fff5f5;color:var(--red);border:1.5px solid rgba(192,57,43,.3)}
  .al-s{background:#f0faf6;color:var(--green);border:1.5px solid rgba(10,124,85,.3)}
  .al-i{background:#f0f5ff;color:var(--accent);border:1.5px solid rgba(59,111,245,.3)}
  .hc{padding:8px 12px;border-radius:8px;font-size:12px;text-align:center;font-weight:600;min-width:108px}
  ::-webkit-scrollbar{width:5px;height:5px}
  ::-webkit-scrollbar-thumb{background:var(--border2);border-radius:3px}
  select option{background:#fff;color:var(--text)}
  input[type=file]{padding:6px 11px;font-size:12px;background:var(--bg3)}
  input[type=number]::-webkit-inner-spin-button,input[type=number]::-webkit-outer-spin-button{-webkit-appearance:none;margin:0}
  input[type=number]{-moz-appearance:textfield}
  input:disabled,select:disabled,textarea:disabled{opacity:.55;cursor:not-allowed;background:var(--bg3)}
  .toast-wrap{position:fixed;bottom:22px;right:22px;z-index:3000;animation:fi .2s ease}
  .topbar{background:#fff;border-bottom:1.5px solid var(--border);padding:13px 28px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 1px 3px rgba(0,0,0,.04)}
`;

// ── Scoring ──────────────────────────────────────────────────
const SW = new Set(["a","an","the","and","or","of","in","on","with","for","to","is","are","was","were","be","been","by","at","from"]);
function tok(s) {
  if (!s) return new Set();
  return new Set(s.toLowerCase().replace(/[^a-z0-9\s]/g," ").split(/\s+/).filter(t=>t.length>1&&!SW.has(t)));
}
const SYN = {
  "round neck":["crew neck","round collar"],"crew neck":["round neck","round collar"],
  "hand wash":["handwash","hand washing"],"machine wash":["machine washable","machine washing"],
  "pure cotton":["100% cotton","cotton"],"viscose rayon":["rayon","viscose"],
  "waterproof":["water resistant","water proof"],"antibacterial":["anti bacterial","antimicrobial"],
};
function exSyn(s) {
  let r = s.toLowerCase();
  for (const [k,vs] of Object.entries(SYN)) {
    if (r.includes(k)) vs.forEach(v=>{r+=" "+v;});
    vs.forEach(v=>{if(r.includes(v))r+=" "+k;});
  }
  return r;
}
function scoreF(uv, gv, ct, st=0.7) {
  if (!uv&&!gv) return 1;
  if (!uv||!gv) return 0;
  const u=String(uv).trim(), g=String(gv).trim();
  if (ct==="as_is") return u.toLowerCase()===g.toLowerCase()?1:0;
  if (ct==="multiselect"||ct==="list") {
    const us=new Set(u.toLowerCase().split(",").map(x=>x.trim()).filter(Boolean));
    const gs=new Set(g.toLowerCase().split(",").map(x=>x.trim()).filter(Boolean));
    if(us.size===0&&gs.size===0)return 1;
    const inter=[...us].filter(x=>gs.has(x)).length;
    const union=new Set([...us,...gs]).size;
    return union===0?0:inter/union;
  }
  if (ct==="numeric") {
    const un=parseFloat(u),gn=parseFloat(g);
    if(isNaN(un)||isNaN(gn))return 0;
    return Math.abs(un-gn)<0.001?1:0;
  }
  if (ct==="url") {
    const n=s=>s.toLowerCase().replace(/\/$/,"").replace(/^https?:\/\//,"");
    return n(u)===n(g)?1:0;
  }
  if (ct==="semantic") {
    const ue=exSyn(u),ge=exSyn(g);
    const ut=tok(ue),gt=tok(ge);
    if(ut.size===0&&gt.size===0)return 1;
    const inter=[...ut].filter(x=>gt.has(x)).length;
    const union=new Set([...ut,...gt]).size;
    return union===0?0:inter/union;
  }
  return 0;
}

function fmt(d) {
  if (!d) return "—";
  return new Date(d).toLocaleString("en-IN",{dateStyle:"medium",timeStyle:"short"});
}

function useToast() {
  const [t, st] = useState(null);
  const show = (msg, type="success") => { st({msg,type}); setTimeout(()=>st(null),3500); };
  return [t, show];
}
function Toast({ t }) {
  if (!t) return null;
  return (
    <div className="toast-wrap">
      <div className={`al al-${t.type==="error"?"e":"s"}`} style={{minWidth:260,boxShadow:"var(--shadow2)",marginBottom:0}}>
        {t.msg}
      </div>
    </div>
  );
}

// ── LOGIN ────────────────────────────────────────────────────
function LoginPage({ onLogin }) {
  const [email,setEmail] = useState("");
  const [pw,setPw] = useState("");
  const [load,setLoad] = useState(false);
  const [err,setErr] = useState("");

  async function doLogin(e) {
    e.preventDefault(); setLoad(true); setErr("");
    const {data,error} = await sb.auth.signInWithPassword({email:email.trim(),password:pw});
    if (error) { setErr(error.message); setLoad(false); return; }
    const {data:prof,error:pe} = await sb.rpc("get_my_profile");
    if (pe||!prof) { setErr("Could not load profile. Please try again."); setLoad(false); return; }
    if (prof.status==="disabled") { await sb.auth.signOut(); setErr("Account disabled. Contact admin."); setLoad(false); return; }
    onLogin(prof); setLoad(false);
  }

  return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"var(--bg)"}}>
      <div style={{width:"100%",maxWidth:400,padding:"0 20px"}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <div style={{fontSize:30,fontWeight:700,letterSpacing:"-0.5px",color:"var(--text)"}}>NW Curation</div>
          <div style={{color:"var(--text3)",marginTop:6,fontSize:14}}>Retail Item Quality Assessment</div>
        </div>
        <div className="card" style={{borderRadius:"var(--r3)",padding:28,boxShadow:"var(--shadow2)"}}>
          {err && <div className="al al-e">{err}</div>}
          <form onSubmit={doLogin}>
            <div className="fg"><label className="fl">Email</label><input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="you@example.com" required/></div>
            <div className="fg"><label className="fl">Password</label><input type="password" value={pw} onChange={e=>setPw(e.target.value)} placeholder="Enter your password" required/></div>
            <button type="submit" className="bp wf" style={{marginTop:8}} disabled={load}>
              {load?<><span className="sp"/> &nbsp;Signing in...</>:"Sign in"}
            </button>
          </form>
          <div style={{marginTop:16,textAlign:"center",fontSize:13,color:"var(--text3)"}}>
            Forgot your password? Contact your admin to reset it.
          </div>
        </div>
      </div>
    </div>
  );
}

// ── ADMIN SHELL ──────────────────────────────────────────────
function AdminShell({ user, onLogout }) {
  const [page, setPage] = useState("dash");
  const [toast, showToast] = useToast();
  const nav = [
    {id:"dash",icon:"▦",label:"Dashboard"},
    {id:"users",icon:"◈",label:"Users"},
    {id:"domains",icon:"◉",label:"Domains"},
    {id:"contests",icon:"◎",label:"Contests"},
    {id:"progress",icon:"◐",label:"Live Progress"},
  ];
  return (
    <div style={{display:"flex"}}>
      <Toast t={toast}/>
      <div className="sidebar">
        <div style={{padding:"20px 16px 12px"}}>
          <div style={{fontSize:15,fontWeight:700,letterSpacing:"-0.4px",color:"var(--text)"}}>NW Curation</div>
          <div style={{fontSize:10,color:"var(--text3)",marginTop:2,fontWeight:600,letterSpacing:".04em",textTransform:"uppercase"}}>Admin</div>
        </div>
        <div className="div" style={{margin:"0 12px"}}/>
        <div style={{flex:1,paddingTop:6}}>
          <div className="ns">Navigation</div>
          {nav.map(n=>(
            <div key={n.id} className={`ni ${page===n.id?"act":""}`} onClick={()=>setPage(n.id)}>
              <span style={{fontSize:14,lineHeight:1}}>{n.icon}</span>
              <span>{n.label}</span>
            </div>
          ))}
        </div>
        <div style={{padding:"12px 16px 20px",borderTop:"1.5px solid var(--border)"}}>
          <div style={{fontSize:12,color:"var(--text3)",marginBottom:8,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{user.email}</div>
          <button className="bg bsm wf" onClick={onLogout}>Sign out</button>
        </div>
      </div>
      <div className="mc fi">
        {page==="dash"&&<AdminDash/>}
        {page==="users"&&<AdminUsers showToast={showToast}/>}
        {page==="domains"&&<AdminDomains showToast={showToast}/>}
        {page==="contests"&&<AdminContests showToast={showToast}/>}
        {page==="progress"&&<AdminProgress/>}
      </div>
    </div>
  );
}

function AdminDash() {
  const [s,setS] = useState({u:0,d:0,c:0,a:0});
  useEffect(()=>{
    Promise.all([
      sb.from("users").select("*",{count:"exact",head:true}).eq("role","participant"),
      sb.from("domains").select("*",{count:"exact",head:true}),
      sb.from("contests").select("*",{count:"exact",head:true}),
      sb.from("contests").select("*",{count:"exact",head:true}).eq("status","active"),
    ]).then(([{count:u},{count:d},{count:c},{count:a}])=>setS({u:u||0,d:d||0,c:c||0,a:a||0}));
  },[]);
  return (
    <div>
      <div className="pt">Dashboard</div>
      <div className="ps">Overview of your curation platform</div>
      <div className="g4c" style={{marginBottom:28}}>
        {[
          {label:"Participants",v:s.u,col:"var(--accent)"},
          {label:"Domains",v:s.d,col:"var(--purple)"},
          {label:"Contests",v:s.c,col:"var(--green)"},
          {label:"Active now",v:s.a,col:"var(--amber)"},
        ].map(x=>(
          <div className="sc" key={x.label}>
            <div className="sv" style={{color:x.col}}>{x.v}</div>
            <div className="sl">{x.label}</div>
          </div>
        ))}
      </div>
      <div className="card">
        <div className="fw6" style={{marginBottom:12}}>Quick start guide</div>
        {[
          ["1","Add users","Users → add emails individually or upload a CSV"],
          ["2","Upload domains","Domains → upload golden dataset + config file (.xlsx)"],
          ["3","Create a contest","Contests → set mode, tasks, domains, assign users"],
          ["4","Monitor live","Live Progress → watch scores update in real time"],
        ].map(([n,t,d])=>(
          <div key={n} className="fx g3 ac" style={{padding:"10px 0",borderBottom:"1px solid var(--border)"}}>
            <div style={{width:26,height:26,borderRadius:"50%",background:"var(--ag)",color:"var(--accent)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,flexShrink:0,border:"1.5px solid rgba(59,111,245,.2)"}}>{n}</div>
            <div>
              <div className="fw5 sm">{t}</div>
              <div className="xs m3">{d}</div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── ADMIN USERS ──────────────────────────────────────────────
function AdminUsers({ showToast }) {
  const [users,setUsers] = useState([]);
  const [load,setLoad] = useState(true);
  const [showAdd,setShowAdd] = useState(false);
  const [showReset,setShowReset] = useState(null);
  const [newEmail,setNewEmail] = useState("");
  const [newName,setNewName] = useState("");
  const [newPass,setNewPass] = useState("");
  const [resetPass,setResetPass] = useState("");
  const [csvTxt,setCsvTxt] = useState("");
  const [addMode,setAddMode] = useState("single");
  const [saving,setSaving] = useState(false);
  const [bulkResult,setBulkResult] = useState(null);
  const fRef = useRef();

  async function load2() {
    setLoad(true);
    const {data} = await sb.from("users").select("*").eq("role","participant").order("created_at",{ascending:false});
    setUsers(data||[]); setLoad(false);
  }
  useEffect(()=>{load2();},[]);

  async function addUser(email, name, password) {
    const cleanEmail = email.toLowerCase().trim();
    if (!cleanEmail||!cleanEmail.includes("@")) return {ok:false,msg:`Invalid email: ${email}`};
    if (!password||password.length<6) return {ok:false,msg:"Password must be at least 6 characters"};
    const {data:existing} = await sb.from("users").select("id").eq("email",cleanEmail).maybeSingle();
    if (existing) return {ok:false,msg:`${cleanEmail} already exists`};
    // Create auth account — trigger auto-creates users row
    const {data:su,error:suErr} = await sb.auth.signUp({
      email: cleanEmail, password,
      options: {data:{full_name:name||"",role:"participant"}}
    });
    if (suErr) return {ok:false,msg:suErr.message};
    // Update full_name
    if (su?.user && name) await sb.from("users").update({full_name:name}).eq("email",cleanEmail);
    return {ok:true};
  }

  async function handleSingle(e) {
    e.preventDefault(); setSaving(true);
    const result = await addUser(newEmail.trim(), newName.trim(), newPass);
    if (result.ok) {
      showToast(`✓ User ${newEmail} added`);
      setNewEmail(""); setNewName(""); setNewPass(""); setShowAdd(false); load2();
    } else showToast(result.msg,"error");
    setSaving(false);
  }

  async function handleCSV(e) {
    e.preventDefault(); setSaving(true);
    setBulkResult(null);
    if (!newPass||newPass.length<6) { showToast("Enter a temporary password (min 6 chars)","error"); setSaving(false); return; }

    // Parse emails and names from CSV text
    const lines = csvTxt.trim().split("\n").filter(l=>l.trim());
    const emails = [];
    const names = {}; // email -> name map
    lines.forEach(line=>{
      const clean = line.trim().replace(/^["'"]|["'"]$/g,"");
      const parts = clean.split(",").map(s=>s.trim().replace(/^["'"]|["'"]$/g,""));
      const email = parts[0];
      const name = parts.slice(1).join(" ").trim();
      if (email && email.includes("@")) {
        emails.push(email);
        if (name) names[email] = name;
      }
    });

    if (emails.length===0) { showToast("No valid emails found","error"); setSaving(false); return; }

    try {
      // Call Edge Function — bypasses rate limits
      const {data:{session}} = await sb.auth.getSession();
      const resp = await fetch(`${SUPABASE_URL}/functions/v1/create-users`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${session?.access_token}`,
          "apikey": SUPABASE_ANON,
        },
        body: JSON.stringify({ emails, names, password: newPass }),
      });

      const result = await resp.json();
      if (!resp.ok) throw new Error(result.error||"Edge Function error");

      setBulkResult(result);
      showToast(`✓ ${result.summary.created} created, ${result.summary.skipped} skipped, ${result.summary.failed} failed`);
      setCsvTxt(""); setNewPass("");
      load2();
    } catch(err) {
      showToast("Bulk create failed: "+err.message, "error");
    }
    setSaving(false);
  }

  async function deleteUser(u) {
    if (!confirm(`Delete ${u.email}? Removes them and all task data.`)) return;
    if (u.id) {
      const {data:userTasks} = await sb.from("tasks").select("id").eq("user_id",u.id);
      if (userTasks?.length>0) await sb.from("responses").delete().in("task_id",userTasks.map(t=>t.id));
      await sb.from("tasks").delete().eq("user_id",u.id);
      await sb.from("contest_users").delete().eq("user_id",u.id);
      // Delete from auth.users via security definer function
      await sb.rpc("delete_auth_user", {user_id: u.id});
    }
    const {error} = await sb.from("users").delete().eq("email",u.email);
    if (error) { showToast("Delete failed: "+error.message,"error"); return; }
    showToast(`User ${u.email} deleted`); load2();
  }

  async function handleReset(e) {
    e.preventDefault(); setSaving(true);
    if (!resetPass||resetPass.length<6) { showToast("Min 6 characters","error"); setSaving(false); return; }
    try {
      const {data:{session}} = await sb.auth.getSession();
      const resp = await fetch(`${SUPABASE_URL}/functions/v1/reset-password`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${session?.access_token}`,
          "apikey": SUPABASE_ANON,
        },
        body: JSON.stringify({ userId: showReset.id, password: resetPass }),
      });
      const result = await resp.json();
      if (!resp.ok) throw new Error(result.error||"Reset failed");
      showToast(`✓ Password reset for ${showReset.email}. Share the new password with them.`);
      setShowReset(null); setResetPass("");
    } catch(err) {
      showToast("Reset failed: "+err.message, "error");
    }
    setSaving(false);
  }


  return (
    <div>
      <div className="fx ac jb" style={{marginBottom:20}}>
        <div><div className="pt">Users</div><div className="ps">{users.length} participant(s)</div></div>
        <button className="bp" onClick={()=>setShowAdd(true)}>+ Add users</button>
      </div>
      <div className="card" style={{padding:0,overflow:"hidden"}}>
        {load ? (
          <div style={{padding:36,textAlign:"center"}}><span className="sp"/></div>
        ) : (
          <table>
            <thead><tr>{["Name","Email","Status","Added","Actions"].map(c=><th key={c}>{c}</th>)}</tr></thead>
            <tbody>
              {users.length===0&&<tr><td colSpan={5} style={{textAlign:"center",color:"var(--text3)",padding:32}}>No users yet. Add your first participant.</td></tr>}
              {users.map(u=>(
                <tr key={u.id}>
                  <td className="fw5">{u.full_name||"—"}</td>
                  <td className="mono xs">{u.email}</td>
                  <td><span className={`badge b-${u.status==="active"?"green":"red"}`}>{u.status}</span></td>
                  <td>{fmt(u.created_at)}</td>
                  <td>
                    <div className="fx g2">
                      <button className="bg bxs" onClick={()=>setShowReset(u)}>Reset pwd</button>
                      <button className="bd bxs" onClick={()=>deleteUser(u)}>Delete</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {showAdd&&(
        <div className="mo" onClick={()=>setShowAdd(false)}>
          <div className="md" onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:18,fontWeight:700,marginBottom:18}}>Add users</div>
            <div className="tabs">
              <div className={`tab ${addMode==="single"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setAddMode("single")}>Single</div>
              <div className={`tab ${addMode==="csv"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setAddMode("csv")}>Bulk CSV</div>
            </div>
            {addMode==="single" ? (
              <form onSubmit={handleSingle}>
                <div className="fg"><label className="fl">Full name</label><input value={newName} onChange={e=>setNewName(e.target.value)} placeholder="Jane Doe"/></div>
                <div className="fg"><label className="fl">Email *</label><input type="email" value={newEmail} onChange={e=>setNewEmail(e.target.value)} placeholder="jane@company.com" required/></div>
                <div className="fg"><label className="fl">Temporary password *</label><input type="text" value={newPass} onChange={e=>setNewPass(e.target.value)} placeholder="Min 6 chars — share with user" required/></div>
                <div className="fx g3 jb" style={{marginTop:18}}>
                  <button type="button" className="bg" onClick={()=>setShowAdd(false)}>Cancel</button>
                  <button type="submit" className="bp" disabled={saving}>{saving?<><span className="sp"/> &nbsp;Adding...</>:"Add user"}</button>
                </div>
              </form>
            ) : (
              <>{bulkResult ? (
                <div>
                  <div style={{marginBottom:12}}>
                    <div className="fw6" style={{marginBottom:8}}>Results</div>
                    <div className="fx g3" style={{marginBottom:12}}>
                      <div style={{padding:"8px 16px",borderRadius:8,background:"#dcfce7",color:"#16a34a",fontWeight:600,fontSize:14}}>✓ {bulkResult.summary.created} created</div>
                      {bulkResult.summary.skipped>0&&<div style={{padding:"8px 16px",borderRadius:8,background:"#fef9c3",color:"#ca8a04",fontWeight:600,fontSize:14}}>⟳ {bulkResult.summary.skipped} skipped</div>}
                      {bulkResult.summary.failed>0&&<div style={{padding:"8px 16px",borderRadius:8,background:"#fee2e2",color:"#dc2626",fontWeight:600,fontSize:14}}>✗ {bulkResult.summary.failed} failed</div>}
                    </div>
                    {bulkResult.results.filter(r=>r.status!=='created').length>0&&(
                      <div style={{maxHeight:200,overflowY:"auto",border:"1px solid var(--border)",borderRadius:8,padding:8}}>
                        <div className="xs fw6" style={{marginBottom:6,color:"var(--text3)"}}>Skipped / Failed:</div>
                        {bulkResult.results.filter(r=>r.status!=='created').map((r,i)=>(
                          <div key={i} className="fx ac g2" style={{padding:"3px 0",fontSize:12,borderBottom:"1px solid var(--border)"}}>
                            <span style={{color:r.status==='skipped'?"#ca8a04":"#dc2626",fontWeight:600,minWidth:52}}>{r.status}</span>
                            <span className="mono">{r.email}</span>
                            <span style={{color:"var(--text3)",fontSize:11}}>— {r.reason}</span>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                  <button className="bp" onClick={()=>{setBulkResult(null);setShowAdd(false);}}>Done</button>
                </div>
              ) : (
              <form onSubmit={handleCSV}>
                <div className="al al-i" style={{fontSize:12}}>One per line: <code>email, Full Name</code> — name is optional.</div>
                <div className="fg">
                  <label className="fl">Paste emails</label>
                  <textarea value={csvTxt} onChange={e=>setCsvTxt(e.target.value)} placeholder={"jane@co.com, Jane Doe\nbob@co.com, Bob Smith"} rows={6} style={{fontFamily:"var(--mono)",fontSize:12}}/>
                </div>
                <div className="fx g3 ac" style={{marginBottom:16}}>
                  <button type="button" className="bg bsm" onClick={()=>fRef.current.click()}>Upload .csv</button>
                  <input ref={fRef} type="file" accept=".csv,.txt" style={{display:"none"}} onChange={ev=>{const r=new FileReader();r.onload=e=>setCsvTxt(e.target.result);r.readAsText(ev.target.files[0]);}}/>
                  <span className="xs m3">or paste above</span>
                </div>
                <div className="fg"><label className="fl">Temporary password for all *</label><input type="text" value={newPass} onChange={e=>setNewPass(e.target.value)} placeholder="Min 6 chars — all users get this" required/></div>
                <div className="al al-i" style={{fontSize:12}}>All users in this batch get this password. Share it with them.</div>
                <div className="fx g3 jb">
                  <button type="button" className="bg" onClick={()=>setShowAdd(false)}>Cancel</button>
                  <button type="submit" className="bp" disabled={saving||!csvTxt.trim()}>{saving?<><span className="sp"/> &nbsp;Adding... (this may take a minute)</>:"Add all"}</button>
                </div>
              </form>
              )}</>
            )}
          </div>
        </div>
      )}

      {showReset&&(
        <div className="mo" onClick={()=>{setShowReset(null);setResetPass("");}}>
          <div className="md" onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:18,fontWeight:700,marginBottom:14}}>Reset password — {showReset.email}</div>
            <form onSubmit={handleReset}>
              <div className="fg"><label className="fl">New temporary password *</label><input type="text" value={resetPass} onChange={e=>setResetPass(e.target.value)} placeholder="Min 6 characters" required/></div>
              <div className="al al-i" style={{fontSize:12}}>Share this with the user. They can change it in Profile after signing in.</div>
              <div className="fx g3 jb" style={{marginTop:18}}>
                <button type="button" className="bg" onClick={()=>{setShowReset(null);setResetPass("");}}>Cancel</button>
                <button type="submit" className="bp" disabled={saving}>{saving?<span className="sp"/>:"Set password"}</button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

// ── ADMIN DOMAINS ────────────────────────────────────────────
function AdminDomains({ showToast }) {
  const [domains,setDomains] = useState([]);
  const [load,setLoad] = useState(true);
  const [show,setShow] = useState(false);
  const [uploading,setUploading] = useState(false);
  const [name,setName] = useState("");
  const [gFile,setGFile] = useState(null);
  const [cFile,setCFile] = useState(null);
  const [prog,setProg] = useState("");
  const [curUser,setCurUser] = useState(null);

  useEffect(()=>{
    sb.auth.getUser().then(({data})=>setCurUser(data?.user));
    load2();
  },[]);

  async function load2() {
    setLoad(true);
    const {data} = await sb.from("domains").select("*, domain_items(count), domain_fields(count)").order("created_at",{ascending:false});
    setDomains(data||[]);
    setLoad(false);
  }

  async function doUpload(e) {
    e.preventDefault();
    if (!gFile||!cFile||!name.trim()) return;
    setUploading(true);
    try {
      setProg("Parsing config file...");
      const cBuf = await new Promise(r=>{const fr=new FileReader();fr.onload=ev=>r(ev.target.result);fr.readAsArrayBuffer(cFile);});
      const cWb = XLSX.read(cBuf,{type:"array"});
      const cRows = XLSX.utils.sheet_to_json(cWb.Sheets[cWb.SheetNames[0]]);

      setProg("Parsing golden dataset...");
      const gBuf = await new Promise(r=>{const fr=new FileReader();fr.onload=ev=>r(ev.target.result);fr.readAsArrayBuffer(gFile);});
      const gWb = XLSX.read(gBuf,{type:"array"});
      const gRows = XLSX.utils.sheet_to_json(gWb.Sheets[gWb.SheetNames[0]]);

      if (!gRows.length) throw new Error("Golden dataset is empty");
      if (!cRows.length) throw new Error("Config file is empty");

      setProg("Creating domain...");
      const {data:dom,error:de} = await sb.from("domains").insert({name:name.trim(),created_by:curUser.id}).select().single();
      if (de) throw new Error(de.message);

      setProg("Saving field config...");
      const fieldRows = cRows.filter(r=>r["Field Name"]&&r["Field Role"]).map(r=>({
        domain_id: dom.id,
        field_name: String(r["Field Name"]).trim(),
        field_role: String(r["Field Role"]).trim(),
        input_type: String(r["Input Type"]||"as_is").trim(),
        comparison_type: String(r["Comparison Type"]||"as_is").trim().toLowerCase(),
        dropdown_values: r["Dropdown Values"]?String(r["Dropdown Values"]).trim():null,
        display_order: parseInt(r["Display Order"]||0),
      }));
      if (fieldRows.length>0) {
        const {error:fe} = await sb.from("domain_fields").insert(fieldRows);
        if (fe) throw new Error("Field config error: "+fe.message);
      }

      // Auto-detect key columns
      const keys = Object.keys(gRows[0]);
      // For item_key, always use row index to guarantee uniqueness.
      // Optionally also look for a dedicated ID column.
      const ikc = keys.find(k=>k==="Item ID")||keys.find(k=>k.toLowerCase().includes("item id"))||null;
      const catc = keys.find(k=>k==="Category")||keys.find(k=>k==="Product Type")||keys.find(k=>k.toLowerCase().includes("categor"))||keys[0];
      const atrc = keys.find(k=>k==="All Attributes for Category")||keys.find(k=>k.toLowerCase().trim()==="all attributes for category");

      const items = gRows.map((row,i)=>{
        // Exclude bulky/redundant columns from json_value storage
        const excludeCols = new Set(["JSONValue","jsonvalue","json_value","All Attributes for Category"]);
        const cleanRow = Object.fromEntries(
          Object.entries(row).filter(([k])=>!excludeCols.has(k))
        );
        return {
          domain_id: dom.id,
          // Always use row index as the key to guarantee uniqueness across any file format
          item_key: ikc ? `${String(row[ikc]||"").trim()}-${i+1}` : `item-${i+1}`,
          category: String(row[catc]||"Unknown").trim(),
          json_value: cleanRow,
          attributes_for_category: atrc?String(row[atrc]||"").trim():"",
        };
      });

      for (let i=0;i<items.length;i+=50) {
        const {error:ie} = await sb.from("domain_items").insert(items.slice(i,i+50));
        if (ie) throw new Error("Item insert error: "+ie.message);
        setProg(`Saved ${Math.min(i+50,items.length)} / ${items.length} items...`);
      }

      showToast(`✓ Domain "${name}" created — ${items.length} items, ${fieldRows.length} fields`);
      setShow(false); setName(""); setGFile(null); setCFile(null); setProg(""); load2();
    } catch(err) {
      showToast(err.message,"error");
      setProg("");
    }
    setUploading(false);
  }

  async function del(d) {
    if (!confirm(`Delete domain "${d.name}"? This removes all items and field configs.`)) return;
    const {error} = await sb.from("domains").delete().eq("id",d.id);
    if (error) { showToast("Delete failed: "+error.message,"error"); return; }
    showToast("Domain deleted"); load2();
  }

  return (
    <div>
      <div className="fx ac jb" style={{marginBottom:20}}>
        <div><div className="pt">Domains</div><div className="ps">Upload golden datasets and field config files</div></div>
        <button className="bp" onClick={()=>setShow(true)}>+ Add domain</button>
      </div>
      <div className="g2c">
        {load ? <div style={{padding:36}}><span className="sp"/></div> :
         domains.length===0 ? (
           <div className="card" style={{gridColumn:"1/-1",textAlign:"center",color:"var(--text3)",padding:40}}>
             No domains yet. Upload your first domain to get started.
           </div>
         ) :
         domains.map(d=>(
          <div className="card" key={d.id}>
            <div className="fx ac jb" style={{marginBottom:10}}>
              <div className="fw6" style={{fontSize:16}}>{d.name}</div>
              <button className="bd bxs" onClick={()=>del(d)}>Delete</button>
            </div>
            <div className="fx g4 sm m2">
              <span>{d.domain_items?.[0]?.count||0} items</span>
              <span>{d.domain_fields?.[0]?.count||0} fields</span>
            </div>
            <div className="xs m3" style={{marginTop:8}}>{fmt(d.created_at)}</div>
          </div>
         ))
        }
      </div>

      {show&&(
        <div className="mo" onClick={()=>!uploading&&setShow(false)}>
          <div className="md" onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:18,fontWeight:700,marginBottom:20}}>Add new domain</div>
            <form onSubmit={doUpload}>
              <div className="fg"><label className="fl">Domain name *</label><input value={name} onChange={e=>setName(e.target.value)} placeholder="e.g. Apparel" required/></div>
              <div className="fg">
                <label className="fl">Golden dataset (.xlsx) *</label>
                <input type="file" accept=".xlsx,.xls" onChange={e=>setGFile(e.target.files[0])} required/>
                <div className="xs m3" style={{marginTop:4}}>e.g. Apparel_80_golden_fixed.xlsx</div>
              </div>
              <div className="fg">
                <label className="fl">Config file (.xlsx) *</label>
                <input type="file" accept=".xlsx,.xls" onChange={e=>setCFile(e.target.files[0])} required/>
                <div className="xs m3" style={{marginTop:4}}>e.g. config_apparel_v2.xlsx</div>
              </div>
              {prog&&<div className="al al-i" style={{fontSize:12}}>{prog}</div>}
              <div className="fx g3 jb" style={{marginTop:20}}>
                <button type="button" className="bg" onClick={()=>setShow(false)} disabled={uploading}>Cancel</button>
                <button type="submit" className="bp" disabled={uploading}>{uploading?<><span className="sp"/> &nbsp;Uploading...</>:"Upload & create"}</button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

// ── ADMIN CONTESTS ───────────────────────────────────────────
function AdminContests({ showToast }) {
  const [contests,setContests] = useState([]);
  const [domains,setDomains] = useState([]);
  const [users,setUsers] = useState([]);
  const [load,setLoad] = useState(true);
  const [showCreate,setShowCreate] = useState(false);
  const [showAssign,setShowAssign] = useState(null);
  const [curUser,setCurUser] = useState(null);

  useEffect(()=>{ sb.auth.getUser().then(({data})=>setCurUser(data?.user)); load2(); },[]);

  async function load2() {
    setLoad(true);
    const [{data:c},{data:d},{data:u}] = await Promise.all([
      sb.from("contests").select("*, contest_domains(domain_id,task_count,domains(name)), contest_users(count)").order("created_at",{ascending:false}),
      sb.from("domains").select("id,name"),
      sb.from("users").select("id,email,full_name").eq("role","participant").eq("status","active"),
    ]);
    setContests(c||[]); setDomains(d||[]); setUsers(u||[]); setLoad(false);
  }

  async function toggleStatus(c) {
    const next = c.status==="active"?"closed":c.status==="draft"?"active":"active";
    const {error} = await sb.from("contests").update({status:next}).eq("id",c.id);
    if (error) { showToast("Update failed: "+error.message,"error"); return; }

    // On activate: pre-create all task rows for all assigned users
    if (next==="active") {
      const [{data:ciItems},{data:cuUsers}] = await Promise.all([
        sb.from("contest_items").select("id").eq("contest_id",c.id),
        sb.from("contest_users").select("user_id").eq("contest_id",c.id),
      ]);
      const taskRows = [];
      for (const u of (cuUsers||[])) {
        for (const item of (ciItems||[])) {
          taskRows.push({
            contest_id: c.id,
            user_id: u.user_id,
            contest_item_id: item.id,
            status: "not_started",
            started_at: new Date().toISOString(),
          });
        }
      }
      if (taskRows.length > 0) {
        // Insert in batches of 100 to avoid payload limits
        for (let i=0; i<taskRows.length; i+=100) {
          const {error:tce} = await sb.from("tasks").upsert(taskRows.slice(i,i+100), {onConflict:"contest_id,user_id,contest_item_id",ignoreDuplicates:true});
          if (tce) { showToast("⚠️ Task pre-creation error — retry activation"); throw tce; }
        }
      }
      showToast(`Contest activated — ${taskRows.length} tasks pre-created`);
    } else {
      showToast(`Contest ${next}`);
    }
    load2();
  }

  async function del(c) {
    if (!confirm(`Delete contest "${c.name}"? This cannot be undone.`)) return;
    await sb.from("contests").delete().eq("id",c.id);
    showToast("Contest deleted"); load2();
  }

  const sc = {draft:"gray",active:"green",closed:"red"};

  return (
    <div>
      <div className="fx ac jb" style={{marginBottom:20}}>
        <div><div className="pt">Contests</div><div className="ps">{contests.length} total</div></div>
        <button className="bp" onClick={()=>setShowCreate(true)}>+ Create contest</button>
      </div>
      <div className="card" style={{padding:0,overflow:"hidden"}}>
        {load ? <div style={{padding:36,textAlign:"center"}}><span className="sp"/></div> : (
          <table>
            <thead><tr>{["Name","Mode","Domains","Tasks","Users","Status","Actions"].map(c=><th key={c}>{c}</th>)}</tr></thead>
            <tbody>
              {contests.length===0&&<tr><td colSpan={7} style={{textAlign:"center",color:"var(--text3)",padding:32}}>No contests yet.</td></tr>}
              {contests.map(c=>(
                <tr key={c.id}>
                  <td className="fw5">{c.name}</td>
                  <td><span className={`badge b-${c.mode==="practice"?"purple":"blue"}`}>{c.mode}</span></td>
                  <td className="xs">{c.contest_domains?.map(cd=>cd.domains?.name).filter(Boolean).join(", ")||"—"}</td>
                  <td>{c.task_count}</td>
                  <td>{c.contest_users?.[0]?.count||0}</td>
                  <td><span className={`badge b-${sc[c.status]||"gray"}`}>{c.status}</span></td>
                  <td>
                    <div className="fx g2">
                      <button className="bg bxs" onClick={()=>setShowAssign(c)}>Users</button>
                      <button className="bg bxs" onClick={()=>toggleStatus(c)}>
                        {c.status==="active"?"Close":c.status==="draft"?"Activate":"Reopen"}
                      </button>
                      <button className="bd bxs" onClick={()=>del(c)}>Del</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
      {showCreate&&<CreateContestModal domains={domains} users={users} uid={curUser?.id} onClose={()=>setShowCreate(false)} onDone={()=>{setShowCreate(false);load2();showToast("Contest created!");}}/>}
      {showAssign&&<AssignModal contest={showAssign} users={users} onClose={()=>setShowAssign(null)} onSaved={()=>{setShowAssign(null);load2();showToast("Users updated");}}/>}
    </div>
  );
}

function CreateContestModal({ domains, users, uid, onClose, onDone }) {
  const [name,setName]=useState("");
  const [mode,setMode]=useState("assessment");
  const [tc,setTc]=useState(10);
  const [sa,setSa]=useState(""); const [ea,setEa]=useState("");
  const [l0,setL0]=useState(70); const [l1,setL1]=useState(80); const [l2,setL2]=useState(90); const [st,setSt]=useState(0.7);
  const [allocs,setAllocs]=useState([{did:"",tc:10}]);
  const [selU,setSelU]=useState([]); const [csvU,setCsvU]=useState(""); const [amode,setAmode]=useState("sel");
  const [saving,setSaving]=useState(false); const [err,setErr]=useState("");
  const fRef=useRef();

  async function doCreate(e) {
    e.preventDefault(); setErr("");
    const tot = allocs.reduce((s,a)=>s+parseInt(a.tc||0),0);
    if (tot!==parseInt(tc)) { setErr(`Domain task counts must sum to ${tc}. Currently: ${tot}`); return; }
    if (allocs.some(a=>!a.did)) { setErr("Please select a domain for each row"); return; }
    if (!name.trim()) { setErr("Contest name is required"); return; }
    setSaving(true);
    try {
      const {data:con,error:ce} = await sb.from("contests").insert({
        name:name.trim(),mode,task_count:parseInt(tc),
        start_at:sa||null,end_at:ea||null,
        l0_threshold:l0,l1_threshold:l1,l2_threshold:l2,
        semantic_correct_threshold:st,
        status:"draft",created_by:uid,
      }).select().single();
      if (ce) throw new Error(ce.message);

      let off=0;
      for (const a of allocs) {
        await sb.from("contest_domains").insert({contest_id:con.id,domain_id:a.did,task_count:parseInt(a.tc)});
        const {data:ditems} = await sb.from("domain_items").select("id").eq("domain_id",a.did);
        if (!ditems||!ditems.length) continue;
        const n = parseInt(a.tc);
        const shuffled = [...ditems].sort(()=>Math.random()-.5).slice(0,n);
        await sb.from("contest_items").insert(shuffled.map((it,i)=>({contest_id:con.id,domain_item_id:it.id,item_order:off+i+1})));
        off+=n;
      }

      let uids=[...selU];
      if (amode==="csv"&&csvU.trim()) {
        const em=csvU.trim().split("\n").map(l=>l.split(",")[0].trim().toLowerCase()).filter(Boolean);
        const {data:fu}=await sb.from("users").select("id,email").in("email",em);
        uids=[...new Set([...uids,...(fu||[]).map(u=>u.id||u.email).filter(Boolean)])];
      }
      // Separate signed-up users (have auth uuid) from pending (no auth yet)
      const signedUpIds = uids.filter(u=>u&&u.includes("-")&&u.length===36);
      const pendingEmails = uids.filter(u=>u&&u.includes("@"));
      if (signedUpIds.length>0) await sb.from("contest_users").insert(signedUpIds.map(uid2=>({contest_id:con.id,user_id:uid2})));
      // Store pending users by email — contest_users.pending_email column
      if (pendingEmails.length>0) await sb.from("contest_users").insert(pendingEmails.map(email=>({contest_id:con.id,pending_email:email})));
      onDone();
    } catch(er) { setErr(er.message); }
    setSaving(false);
  }

  return (
    <div className="mo" onClick={onClose}>
      <div className="md md-lg" onClick={e=>e.stopPropagation()} style={{maxHeight:"92vh",overflowY:"auto"}}>
        <div style={{fontSize:18,fontWeight:700,marginBottom:20}}>Create contest</div>
        <form onSubmit={doCreate}>
          {err&&<div className="al al-e">{err}</div>}
          <div className="row2">
            <div className="fg"><label className="fl">Contest name *</label><input value={name} onChange={e=>setName(e.target.value)} placeholder="e.g. Apparel Assessment June 2025" required/></div>
            <div className="fg"><label className="fl">Mode *</label>
              <select value={mode} onChange={e=>setMode(e.target.value)}>
                <option value="assessment">Assessment (scores stored, no answer reveal)</option>
                <option value="practice">Practice (see answers after each task)</option>
              </select>
            </div>
          </div>
          <div className="row3">
            <div className="fg"><label className="fl">Total tasks *</label><input type="number" min={1} max={200} value={tc} onChange={e=>setTc(e.target.value)} required/></div>
            <div className="fg"><label className="fl">Start time (optional)</label><input type="datetime-local" value={sa} onChange={e=>setSa(e.target.value)}/></div>
            <div className="fg"><label className="fl">End time (optional)</label><input type="datetime-local" value={ea} onChange={e=>setEa(e.target.value)}/></div>
          </div>

          <div style={{marginBottom:16}}>
            <div className="fx ac jb" style={{marginBottom:8}}>
              <label className="fl" style={{margin:0}}>Domain & task allocation *</label>
              <button type="button" className="bg bxs" onClick={()=>setAllocs(p=>[...p,{did:"",tc:5}])}>+ Add domain</button>
            </div>
            {allocs.map((a,i)=>(
              <div key={i} className="fx g3 ac" style={{marginBottom:8}}>
                <select style={{flex:2}} value={a.did} onChange={e=>setAllocs(p=>p.map((x,idx)=>idx===i?{...x,did:e.target.value}:x))}>
                  <option value="">Select domain...</option>
                  {domains.map(d=><option key={d.id} value={d.id}>{d.name}</option>)}
                </select>
                <input type="number" min={1} style={{flex:1}} value={a.tc} onChange={e=>setAllocs(p=>p.map((x,idx)=>idx===i?{...x,tc:e.target.value}:x))} placeholder="Tasks"/>
                {allocs.length>1&&<button type="button" className="bd bxs" onClick={()=>setAllocs(p=>p.filter((_,idx)=>idx!==i))}>✕</button>}
              </div>
            ))}
            <div className="xs m3">Task counts must sum to {tc}</div>
          </div>

          {mode==="assessment"&&(
            <div style={{background:"var(--bg3)",borderRadius:"var(--r)",padding:14,marginBottom:16,border:"1px solid var(--border)"}}>
              <div className="sm fw6" style={{marginBottom:10}}>Certification thresholds</div>
              <div className="row3" style={{marginBottom:10}}>
                {[["L1 min %",l0,setL0],["L2 min %",l1,setL1],["L3 min %",l2,setL2]].map(([lbl,val,setter])=>(
                  <div key={lbl}><label className="fl">{lbl}</label><input type="number" min={0} max={100} value={val} onChange={e=>setter(Number(e.target.value))}/></div>
                ))}
              </div>
              <div>
                <label className="fl">Semantic match threshold (0–1)</label>
                <input type="number" min={0} max={1} step={0.05} value={st} onChange={e=>setSt(Number(e.target.value))}/>
                <div className="xs m3" style={{marginTop:4}}>Semantic score ≥ this value counts as correct. Default: 0.7</div>
              </div>
            </div>
          )}

          <div style={{marginBottom:16}}>
            <label className="fl" style={{marginBottom:10}}>Assign users</label>
            <div className="tabs" style={{marginBottom:10}}>
              <div className={`tab ${amode==="sel"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setAmode("sel")}>Select individually</div>
              <div className={`tab ${amode==="csv"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setAmode("csv")}>Paste emails</div>
            </div>
            {amode==="sel" ? (
              <div style={{maxHeight:180,overflowY:"auto",background:"var(--bg3)",borderRadius:"var(--r)",padding:8,border:"1px solid var(--border)"}}>
                {users.map(u=>(
                  <label key={u.id} className="fx g3 ac" style={{padding:"6px 8px",cursor:"pointer",borderRadius:6}}>
                    <input type="checkbox" checked={selU.includes(u.id||u.email)} onChange={ev=>setSelU(p=>ev.target.checked?[...p,u.id||u.email]:p.filter(id=>id!==(u.id||u.email)))}/>
                    <span className="sm">{u.full_name||u.email}</span>
                    <span className="xs m3" style={{marginLeft:"auto"}}>{u.full_name?u.email:""}{!u.id&&<span style={{color:"var(--amber)",marginLeft:4}}> (not yet signed in)</span>}</span>
                  </label>
                ))}
                {users.length===0&&<div className="xs m3" style={{padding:8}}>No participants yet. Add users first.</div>}
              </div>
            ) : (
              <div>
                <textarea value={csvU} onChange={e=>setCsvU(e.target.value)} placeholder={"jane@co.com\nbob@co.com"} rows={4} style={{fontFamily:"var(--mono)",fontSize:12}}/>
                <input ref={fRef} type="file" accept=".csv,.txt" style={{display:"none"}} onChange={ev=>{const r=new FileReader();r.onload=e=>setCsvU(e.target.result);r.readAsText(ev.target.files[0]);}}/>
                <button type="button" className="bg bxs" style={{marginTop:6}} onClick={()=>fRef.current.click()}>Upload CSV</button>
              </div>
            )}
          </div>

          <div className="fx g3 jb" style={{marginTop:20}}>
            <button type="button" className="bg" onClick={onClose}>Cancel</button>
            <button type="submit" className="bp" disabled={saving}>{saving?<><span className="sp"/> &nbsp;Creating...</>:"Create contest"}</button>
          </div>
        </form>
      </div>
    </div>
  );
}

function AssignModal({ contest, users, onClose, onSaved }) {
  const [assigned,setAssigned]=useState([]); const [sel,setSel]=useState([]);
  const [csv,setCsv]=useState(""); const [mode,setMode]=useState("sel"); const [saving,setSaving]=useState(false);
  const fRef=useRef();
  useEffect(()=>{
    sb.from("contest_users").select("user_id").eq("contest_id",contest.id).then(({data})=>{
      const ids=(data||[]).map(r=>r.user_id);
      setAssigned(ids); setSel(ids);
    });
  },[]);
  async function save() {
    setSaving(true);
    let toAdd=[...sel];
    if (mode==="csv"&&csv.trim()) {
      const em=csv.trim().split("\n").map(l=>l.split(",")[0].trim().toLowerCase()).filter(Boolean);
      const {data}=await sb.from("users").select("id").in("email",em);
      toAdd=[...new Set([...toAdd,...(data||[]).map(u=>u.id).filter(Boolean)])];
    }
    const signedUp=toAdd.filter(u=>u&&u.includes("-")&&u.length===36);
    const pending=toAdd.filter(u=>u&&u.includes("@"));
    const newSU=signedUp.filter(id=>!assigned.includes(id));
    const newP=pending.filter(e=>!assigned.includes(e));
    if(newSU.length>0) await sb.from("contest_users").insert(newSU.map(uid=>({contest_id:contest.id,user_id:uid})));
    if(newP.length>0) await sb.from("contest_users").insert(newP.map(email=>({contest_id:contest.id,pending_email:email})));
    const rem=assigned.filter(id=>!toAdd.includes(id));
    if(rem.length>0) await sb.from("contest_users").delete().eq("contest_id",contest.id).in("user_id",rem);
    onSaved(); setSaving(false);
  }
  return (
    <div className="mo" onClick={onClose}>
      <div className="md" onClick={e=>e.stopPropagation()}>
        <div style={{fontSize:18,fontWeight:700,marginBottom:16}}>Assign users — {contest.name}</div>
        <div className="tabs">
          <div className={`tab ${mode==="sel"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setMode("sel")}>Select</div>
          <div className={`tab ${mode==="csv"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setMode("csv")}>By email</div>
        </div>
        {mode==="sel" ? (
          <div style={{maxHeight:280,overflowY:"auto",background:"var(--bg3)",borderRadius:"var(--r)",padding:8,marginBottom:14,border:"1px solid var(--border)"}}>
            {users.map(u=>(
              <label key={u.id||u.email} className="fx g3 ac" style={{padding:"7px 8px",cursor:"pointer",borderRadius:6}}>
                <input type="checkbox" checked={sel.includes(u.id||u.email)} onChange={ev=>setSel(p=>ev.target.checked?[...p,u.id||u.email]:p.filter(id=>id!==(u.id||u.email)))}/>
                <span className="sm">{u.full_name||u.email}</span>
                <span className="xs m3" style={{marginLeft:"auto"}}>
                  {u.full_name?u.email:""}
                  {!u.id&&<span style={{color:"var(--amber)",fontSize:11,marginLeft:4}}> (not yet signed in)</span>}
                </span>
              </label>
            ))}
          </div>
        ) : (
          <div style={{marginBottom:14}}>
            <textarea value={csv} onChange={e=>setCsv(e.target.value)} placeholder={"jane@co.com\nbob@co.com"} rows={5} style={{fontFamily:"var(--mono)",fontSize:12}}/>
            <input ref={fRef} type="file" accept=".csv,.txt" style={{display:"none"}} onChange={ev=>{const r=new FileReader();r.onload=e=>setCsv(e.target.result);r.readAsText(ev.target.files[0]);}}/>
            <button type="button" className="bg bxs" style={{marginTop:6}} onClick={()=>fRef.current.click()}>Upload CSV</button>
          </div>
        )}
        <div className="sm m2" style={{marginBottom:14}}>{sel.length} user(s) selected</div>
        <div className="fx g3 jb">
          <button className="bg" onClick={onClose}>Cancel</button>
          <button className="bp" onClick={save} disabled={saving}>{saving?<><span className="sp"/> &nbsp;Saving...</>:"Save"}</button>
        </div>
      </div>
    </div>
  );
}

// ── ADMIN PROGRESS ───────────────────────────────────────────
function AdminProgress() {
  const [contests,setContests]=useState([]); const [sel,setSel]=useState("");
  const [progress,setProgress]=useState([]); const [fa,setFa]=useState([]);
  const [load,setLoad]=useState(false); const [tab,setTab]=useState("users");

  useEffect(()=>{
    sb.from("contests").select("id,name,mode,status").order("created_at",{ascending:false}).then(({data})=>setContests(data||[]));
  },[]);

  useEffect(()=>{
    if (!sel) return;
    loadP();
    // Fix 4: No realtime — manual refresh only to reduce IO
  },[sel]);

  async function loadP() {
    if (!sel) return; setLoad(true);
    const con=contests.find(c=>c.id===sel);
    try {
      // Fetch users separately — views don't support join syntax
      const {data:allUsers} = await sb.from("users").select("id,email,full_name");
      const userMap = {};
      (allUsers||[]).forEach(u=>{ if(u.id) userMap[u.id]=u; });

      if (con?.mode==="assessment") {
        const [{data:p,error:pe},{data:f,error:fe},{data:allTasks},{data:assigned}] = await Promise.all([
          sb.from("v_user_contest_accuracy").select("*").eq("contest_id",sel),
          sb.from("v_field_accuracy").select("*").eq("contest_id",sel).order("field_accuracy_pct"),
          sb.from("tasks").select("user_id,status").eq("contest_id",sel),
          sb.from("contest_users").select("user_id").eq("contest_id",sel),
        ]);
        if(pe) console.error("Progress error:",pe.message);
        if(fe) console.error("Field accuracy error:",fe.message);

        // Build score map from view
        const scoreMap = {};
        (p||[]).forEach(row=>{ scoreMap[row.user_id]=row; });

        // Build task status map — count submitted tasks per user
        const taskMap = {};
        (allTasks||[]).forEach(t=>{
          if(!taskMap[t.user_id]) taskMap[t.user_id]={submitted:0,in_progress:0,not_started:0};
          taskMap[t.user_id][t.status]=(taskMap[t.user_id][t.status]||0)+1;
        });

        // Build full list from all assigned users
        const enriched = (assigned||[]).map(row=>{
          const uid = row.user_id;
          const score = scoreMap[uid]||null;
          const tstat = taskMap[uid]||{submitted:0,in_progress:0,not_started:0};
          const tasksDone = tstat.submitted + tstat.in_progress;
          const status = tstat.submitted>0&&tstat.in_progress===0&&tstat.not_started===0 ? 'submitted'
                       : tstat.in_progress>0 ? 'in_progress' : 'not_started';
          return {
            user_id: uid,
            users: userMap[uid]||null,
            tasks_submitted: score?.tasks_submitted||tstat.submitted||0,
            tasks_done: tasksDone,
            total_attributes: score?.total_attributes||0,
            correct_attributes: score?.correct_attributes||0,
            accuracy_pct: score?.accuracy_pct||null,
            cert_level: score?.cert_level||null,
            user_status: status,
          };
        });

        // Sort: submitted first (by accuracy desc), then in_progress, then not_started
        enriched.sort((a,b)=>{
          const order = {submitted:0, in_progress:1, not_started:2};
          if(order[a.user_status]!==order[b.user_status]) return order[a.user_status]-order[b.user_status];
          return (b.accuracy_pct||0)-(a.accuracy_pct||0);
        });

        setProgress(enriched); setFa(f||[]);
      } else {
        const {data:p,error:pe}=await sb.from("v_practice_progress").select("*").eq("contest_id",sel);
        if(pe) console.error("Practice progress error:",pe.message);
        const enriched = (p||[]).map(row=>({...row, users: userMap[row.user_id]||null}));
        setProgress(enriched); setFa([]);
      }
    } catch(e) { console.error("loadP exception:",e); }
    setLoad(false);
  }

  function downloadCSV() {
    const con = contests.find(c=>c.id===sel);
    if (!progress.length) return;
    let csv, filename;
    if (con?.mode==="assessment") {
      const headers = ["Name","Email","Status","Tasks Done","Total Attributes","Correct Attributes","Accuracy %","Certification"];
      const rows = progress.map(p=>[
        p.users?.full_name||"",
        p.users?.email||"",
        p.user_status||"submitted",
        p.tasks_submitted||0,
        p.total_attributes||0,
        p.correct_attributes||0,
        p.accuracy_pct||"",
        p.cert_level||"",
      ]);
      csv = [headers, ...rows].map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(",")).join("\n");
      filename = `${con.name}_results.csv`;
    } else {
      const headers = ["Name","Email","Tasks Completed","Total Tasks"];
      const rows = progress.map(p=>[
        p.users?.full_name||"",
        p.users?.email||"",
        p.tasks_completed||0,
        p.total_tasks||0,
      ]);
      csv = [headers, ...rows].map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(",")).join("\n");
      filename = `${con.name}_practice_progress.csv`;
    }
    const blob = new Blob([csv], {type:"text/csv"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);
  }


  const con=contests.find(c=>c.id===sel);

  return (
    <div>
      <div className="pt">Live Progress</div>
      <div className="ps">Real-time participant performance</div>
      <div style={{maxWidth:500,marginBottom:24,display:"flex",gap:8,alignItems:"flex-end"}}>
        <div style={{flex:1}}>
          <label className="fl">Select contest</label>
          <select value={sel} onChange={e=>{setSel(e.target.value);setProgress([]);setFa([]);}}>
            <option value="">Choose a contest...</option>
            {contests.map(c=><option key={c.id} value={c.id}>[{c.mode}] {c.name} — {c.status}</option>)}
          </select>
        </div>
        {sel&&<button className="bg bsm" onClick={loadP} disabled={load} style={{whiteSpace:"nowrap",paddingBottom:8}}>
          {load?<span className="sp"/>:"↻ Refresh"}
        </button>}
      </div>
      {!sel&&<div className="card" style={{textAlign:"center",color:"var(--text3)",padding:40}}>Select a contest above to see progress</div>}
      {sel&&(
        <>
          <div className="fx ac jb" style={{marginBottom:4}}>
            {con?.mode==="assessment"?(
              <div className="tabs" style={{maxWidth:300,marginBottom:0}}>
                <div className={`tab ${tab==="users"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setTab("users")}>Per user</div>
                <div className={`tab ${tab==="fields"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setTab("fields")}>Field heatmap</div>
              </div>
            ):<div/>}
            {progress.length>0&&(
              <button className="bg bsm" onClick={downloadCSV} style={{whiteSpace:"nowrap"}}>
                ↓ Download CSV
              </button>
            )}
          </div>
          {load ? <div style={{textAlign:"center",padding:40}}><span className="sp"/></div> : (
            <>
              {(tab==="users"||con?.mode==="practice")&&(
                <div className="card" style={{padding:0,overflow:"hidden"}}>
                  <table>
                    <thead><tr>
                      <th>User</th>
                      <th>Tasks done</th>
                      {con?.mode==="assessment"&&<><th>Total attrs</th><th>Correct</th><th>Accuracy</th><th>Cert</th></>}
                    </tr></thead>
                    <tbody>
                      {progress.length===0&&<tr><td colSpan={7} style={{textAlign:"center",color:"var(--text3)",padding:24}}>No data yet</td></tr>}
                      {progress.map((p,i)=>(
                        <tr key={i} style={{opacity:p.user_status==='not_started'?0.5:1}}>
                          <td className="fw5">
                            <div>{p.users?.full_name||p.users?.email||"—"}</div>
                            {con?.mode==="assessment"&&(
                              <div style={{fontSize:11,marginTop:2}}>
                                {p.user_status==='submitted'&&<span style={{color:'var(--green)'}}>✓ Submitted</span>}
                                {p.user_status==='in_progress'&&<span style={{color:'var(--amber)'}}>⟳ In progress ({p.tasks_done}/24)</span>}
                                {p.user_status==='not_started'&&<span style={{color:'var(--text3)'}}>— Not started</span>}
                              </div>
                            )}
                          </td>
                          <td>{p.tasks_submitted||p.tasks_completed||0}</td>
                          {con?.mode==="assessment"&&<>
                            <td>{p.total_attributes||"—"}</td>
                            <td>{p.correct_attributes||"—"}</td>
                            <td>
                              {p.accuracy_pct!=null?(
                                <div className="fx ac g2">
                                  <div className="pb" style={{width:80}}>
                                    <div className="pf" style={{width:`${p.accuracy_pct||0}%`,background:p.accuracy_pct>=90?"var(--green)":p.accuracy_pct>=70?"var(--amber)":"var(--red)"}}/>
                                  </div>
                                  <span className="mono xs">{p.accuracy_pct}%</span>
                                </div>
                              ):<span style={{color:'var(--text3)',fontSize:12}}>—</span>}
                            </td>
                            <td>
                              {p.cert_level
                                ? <span className={`badge cert-${p.cert_level?.toLowerCase()}`}>{p.cert_level}</span>
                                : <span style={{color:'var(--text3)',fontSize:12}}>—</span>}
                            </td>
                          </>}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
              {tab==="fields"&&con?.mode==="assessment"&&(
                <div className="card">
                  <div className="fw6" style={{marginBottom:16}}>Field accuracy — lowest first</div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
                    {fa.map(f=>{
                      const pct=parseFloat(f.field_accuracy_pct)||0;
                      const bg=pct>=80?"var(--gbg)":pct>=60?"var(--abg)":"var(--rbg)";
                      const col=pct>=80?"var(--green)":pct>=60?"var(--amber)":"var(--red)";
                      return (
                        <div key={f.field_name} className="hc" style={{background:bg,color:col,border:`1.5px solid ${col}`}}>
                          <div style={{fontSize:10,marginBottom:2,opacity:.8}}>{f.field_name}</div>
                          <div style={{fontSize:16,fontWeight:700}}>{pct}%</div>
                          <div style={{fontSize:10,opacity:.7}}>{f.correct_responses}/{f.total_responses}</div>
                        </div>
                      );
                    })}
                    {fa.length===0&&<div className="sm m2">No field data yet</div>}
                  </div>
                </div>
              )}
            </>
          )}
        </>
      )}
    </div>
  );
}

// ── USER SHELL ───────────────────────────────────────────────
function UserShell({ user, onLogout }) {
  const [page,setPage]=useState("contests");
  const [ac,setAc]=useState(null);
  const [toast,showToast]=useToast();

  function openContest(c) { setAc(c); setPage("task"); }
  function closeContest() { setAc(null); setPage("contests"); }

  return (
    <div style={{minHeight:"100vh",background:"var(--bg)"}}>
      <Toast t={toast}/>
      {page!=="task"&&(
        <div className="topbar">
          <div style={{fontSize:15,fontWeight:700,letterSpacing:"-0.3px"}}>NW Curation</div>
          <div className="fx g3 ac">
            <span className="sm m2">{user.full_name||user.email}</span>
            <button className="bg bsm" onClick={()=>setPage(page==="profile"?"contests":"profile")}>Profile</button>
            <button className="bg bsm" onClick={onLogout}>Sign out</button>
          </div>
        </div>
      )}
      <div style={{padding:page==="task"?0:28}} className="fi">
        {page==="contests"&&<UserContests user={user} onOpen={openContest}/>}
        {page==="profile"&&<UserProfile user={user} showToast={showToast}/>}
        {page==="task"&&ac&&<ContestTaskView contest={ac} user={user} onClose={closeContest} showToast={showToast}/>}
      </div>
    </div>
  );
}

function UserContests({ user, onOpen }) {
  const [live,setLive]=useState([]); const [past,setPast]=useState([]); const [load,setLoad]=useState(true); const [tab,setTab]=useState("live");

  useEffect(()=>{load2();},[]);

  async function load2() {
    setLoad(true);
    const {data:cu}=await sb.from("contest_users").select("contest_id").eq("user_id",user.id);
    if (!cu||!cu.length) { setLoad(false); return; }
    const ids=cu.map(r=>r.contest_id);
    const [{data:all},{data:tasks},{data:scores}]=await Promise.all([
      sb.from("contests").select("*, contest_domains(domain_id,task_count,domains(name))").in("id",ids).order("created_at",{ascending:false}),
      sb.from("tasks").select("contest_id,status").eq("user_id",user.id),
      sb.from("v_user_contest_accuracy").select("*").eq("user_id",user.id),
    ]);
    const tm={};
    (tasks||[]).forEach(t=>{
      if(!tm[t.contest_id])tm[t.contest_id]={total:0,submitted:0};
      tm[t.contest_id].total++;
      if(t.status==="submitted")tm[t.contest_id].submitted++;
    });
    const sm={};(scores||[]).forEach(s=>{sm[s.contest_id]=s;});
    setLive((all||[]).filter(c=>c.status!=="closed").map(c=>({...c,progress:tm[c.id]||null})));
    setPast((all||[]).filter(c=>c.status==="closed").map(c=>({...c,progress:tm[c.id]||null,score:sm[c.id]||null})));
    setLoad(false);
  }

  if (load) return <div style={{textAlign:"center",padding:60}}><span className="sp" style={{width:28,height:28}}/></div>;

  return (
    <div style={{maxWidth:900,margin:"0 auto"}}>
      <div style={{marginBottom:24}}>
        <div className="xl fw6">My contests</div>
        <div className="sm m2" style={{marginTop:4}}>Welcome back, {user.full_name||user.email.split("@")[0]}</div>
      </div>
      <div className="tabs" style={{maxWidth:300}}>
        <div className={`tab ${tab==="live"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setTab("live")}>Live ({live.length})</div>
        <div className={`tab ${tab==="past"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setTab("past")}>Past ({past.length})</div>
      </div>
      {tab==="live"&&(
        <div>
          {live.length===0&&<div className="card" style={{textAlign:"center",color:"var(--text3)",padding:40}}>No live contests assigned to you right now.</div>}
          {["practice","assessment"].map(m=>{
            const fl=live.filter(c=>c.mode===m);
            if(!fl.length)return null;
            return (
              <div key={m}>
                <div className="xs m3" style={{marginBottom:8,marginTop:16,textTransform:"uppercase",letterSpacing:".07em",fontWeight:700}}>{m}</div>
                <div className="g2c">{fl.map(c=><ContestCard key={c.id} contest={c} onOpen={()=>onOpen(c)}/>)}</div>
              </div>
            );
          })}
        </div>
      )}
      {tab==="past"&&(
        <div>
          {past.length===0&&<div className="card" style={{textAlign:"center",color:"var(--text3)",padding:40}}>No past contests yet.</div>}
          {past.map(c=>(
            <div className="card" key={c.id} style={{marginBottom:12}}>
              <div className="fx ac jb">
                <div>
                  <div className="fw5">{c.name}</div>
                  <div className="xs m3" style={{marginTop:2}}>{c.contest_domains?.map(cd=>cd.domains?.name).filter(Boolean).join(", ")} · {c.task_count} tasks · {c.mode}</div>
                </div>
                <div className="fx g3 ac">
                  {c.mode==="assessment"&&c.score ? (
                    <>
                      <div style={{textAlign:"right"}}>
                        <div className="mono fw6" style={{fontSize:18}}>{c.score.accuracy_pct}%</div>
                        <div className="xs m3">accuracy</div>
                      </div>
                      <span className={`badge cert-${c.score.cert_level?.toLowerCase()}`} style={{fontSize:13,padding:"5px 12px"}}>{c.score.cert_level}</span>
                    </>
                  ) : (
                    <div className="sm m2">{c.progress?.submitted||0}/{c.task_count} completed</div>
                  )}
                </div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ContestCard({ contest, onOpen }) {
  const p=contest.progress; const pct=p&&contest.task_count?Math.round((p.submitted/contest.task_count)*100):0; const done=p?.submitted===contest.task_count;
  return (
    <div className="card" style={{cursor:"pointer",transition:"border-color .15s,box-shadow .15s"}} onClick={onOpen}
      onMouseEnter={e=>{e.currentTarget.style.borderColor="var(--accent)";e.currentTarget.style.boxShadow="0 2px 12px rgba(59,111,245,.12)";}}
      onMouseLeave={e=>{e.currentTarget.style.borderColor="var(--border)";e.currentTarget.style.boxShadow="var(--shadow)";}}>
      <div className="fx ac jb" style={{marginBottom:10}}>
        <span className={`badge b-${contest.mode==="practice"?"purple":"blue"}`}>{contest.mode}</span>
        {done&&<span className="badge b-green">✓ Complete</span>}
      </div>
      <div className="fw6" style={{marginBottom:4,fontSize:15}}>{contest.name}</div>
      <div className="xs m3" style={{marginBottom:12}}>{contest.contest_domains?.map(cd=>cd.domains?.name).filter(Boolean).join(", ")} · {contest.task_count} tasks</div>
      {contest.end_at&&<div className="xs m2" style={{marginBottom:10}}>Ends: {fmt(contest.end_at)}</div>}
      <div className="pb" style={{marginBottom:6}}><div className="pf" style={{width:`${pct}%`}}/></div>
      <div className="xs m3">{p?.submitted||0} / {contest.task_count} submitted</div>
    </div>
  );
}

function UserProfile({ user, showToast }) {
  const [np,setNp]=useState(""); const [cf,setCf]=useState(""); const [fn,setFn]=useState(user.full_name||""); const [saving,setSaving]=useState(false);
  async function save(e) {
    e.preventDefault(); setSaving(true);
    if (fn!==user.full_name) await sb.from("users").update({full_name:fn}).eq("id",user.id);
    if (np) {
      if(np!==cf){showToast("Passwords do not match","error");setSaving(false);return;}
      const{error}=await sb.auth.updateUser({password:np});
      if(error){showToast(error.message,"error");setSaving(false);return;}
    }
    showToast("Profile updated"); setSaving(false);
  }
  return (
    <div style={{maxWidth:480,margin:"0 auto"}}>
      <div className="xl fw6" style={{marginBottom:24}}>Profile</div>
      <div className="card">
        <form onSubmit={save}>
          <div className="fg"><label className="fl">Full name</label><input value={fn} onChange={e=>setFn(e.target.value)} placeholder="Your name"/></div>
          <div className="fg"><label className="fl">Email</label><input value={user.email} disabled/></div>
          <div className="div"/>
          <div className="sm fw6" style={{marginBottom:12}}>Change password</div>
          <div className="fg"><label className="fl">New password</label><input type="password" value={np} onChange={e=>setNp(e.target.value)} placeholder="Leave blank to keep current"/></div>
          <div className="fg"><label className="fl">Confirm new password</label><input type="password" value={cf} onChange={e=>setCf(e.target.value)} placeholder="Repeat password"/></div>
          <button type="submit" className="bp" disabled={saving}>{saving?<><span className="sp"/> &nbsp;Saving...</>:"Save changes"}</button>
        </form>
      </div>
    </div>
  );
}


// ── CURATE FIELD — local state only, DB saves happen on navigate ─
function CurateField({ fieldDef, initialValue, disabled, onSave }) {
  const [val, setVal] = useState(initialValue||"");

  useEffect(()=>{ setVal(initialValue||""); }, [initialValue]);

  function handleChange(newVal) {
    setVal(newVal);
    onSave(fieldDef.field_name, newVal); // just updates parent state, no DB
  }

  const opts = fieldDef.dropdown_values ? fieldDef.dropdown_values.split(";").map(s=>s.trim()).filter(Boolean) : [];
  // For multiselect: val is stored as comma-separated string, parse to array for UI
  const selectedVals = fieldDef.input_type==="multiselect"
    ? (val ? val.split(",").map(s=>s.trim()).filter(Boolean) : [])
    : [];

  function handleMultiChange(opt) {
    const curr = val ? val.split(",").map(s=>s.trim()).filter(Boolean) : [];
    const next = curr.includes(opt) ? curr.filter(x=>x!==opt) : [...curr, opt];
    const newVal = next.join(", ");
    setVal(newVal);
    onSave(fieldDef.field_name, newVal);
  }

  return (
    <div key={fieldDef.field_name}>
      <label className="fl">{fieldDef.field_name}</label>
      {fieldDef.input_type==="multiselect" ? (
        <div style={{display:"flex",flexWrap:"wrap",gap:8,marginTop:4}}>
          {opts.map(o=>{
            const sel = selectedVals.includes(o);
            return (
              <button key={o} type="button" disabled={disabled}
                onClick={()=>!disabled&&handleMultiChange(o)}
                style={{
                  display:"flex",alignItems:"center",gap:5,
                  cursor:disabled?"default":"pointer",
                  padding:"6px 14px",borderRadius:20,
                  border: sel ? "2px solid #16a34a" : "2px solid #d1d5db",
                  background: sel ? "#22c55e" : "#f9fafb",
                  color: sel ? "#ffffff" : "#374151",
                  fontWeight: sel ? 600 : 400,
                  fontSize:13,transition:"all 0.15s",
                  opacity:disabled?0.6:1,
                }}>
                {sel && <span style={{fontSize:12,fontWeight:700}}>✓</span>}
                <span>{o}</span>
              </button>
            );
          })}
          {opts.length===0&&<span style={{fontSize:12,color:"#9ca3af"}}>No options configured</span>}
        </div>
      ) : fieldDef.input_type==="dropdown" ? (
        <select value={val} disabled={disabled} onChange={e=>handleChange(e.target.value)}>
          <option value="">Select...</option>
          {opts.map(o=><option key={o} value={o}>{o}</option>)}
        </select>
      ) : fieldDef.input_type==="numeric" ? (
        <input type="number" value={val} disabled={disabled} onChange={e=>handleChange(e.target.value)}/>
      ) : (
        <input type="text" value={val} disabled={disabled}
          placeholder={`Enter ${fieldDef.field_name.toLowerCase()}...`}
          onChange={e=>handleChange(e.target.value)}/>
      )}
    </div>
  );
}


// ── CONTEST TASK VIEW ────────────────────────────────────────
function ContestTaskView({ contest, user, onClose, showToast }) {
  const [items,setItems]=useState([]);
  const [tasks,setTasks]=useState({});
  const [fields,setFields]=useState({});
  const [idx,setIdx]=useState(0);
  const [saveStatus,setSaveStatus]=useState('idle'); // idle|saving|saved|failed
  const isSaving = useRef(false);
  const pendingNav = useRef(null); // queued navigation target idx
  const [answers,setAnswers]=useState({});
  const [loading,setLoading]=useState(true);
  const [submitting,setSubmitting]=useState(false);
  const [showVal,setShowVal]=useState(false);
  const [submitted,setSubmitted]=useState(false);
  const [score,setScore]=useState(null);

  const [contestClosed,setContestClosed] = useState(contest.status==="closed");
  const closingRef = useRef(false);
  // Ref to prevent concurrent ensureTask calls
  const creatingTask = useRef({});

  useEffect(()=>{ loadC(); },[]);

  // Realtime + polling: watch for contest close
  useEffect(()=>{
    // Realtime subscription
    const sub = sb.channel("contest-status-"+contest.id)
      .on("postgres_changes",{
        event:"UPDATE", schema:"public", table:"contests",
        filter:`id=eq.${contest.id}`
      }, async(payload)=>{
        if (payload.new.status==="closed" && !closingRef.current) {
          closingRef.current = true;
          setContestClosed(true);
          await autoSubmitAll();
        }
      })
      .subscribe();

    return ()=>sb.removeChannel(sub);
  },[]);

  async function loadC() {
    try { await sb.rpc("close_expired_contests"); } catch(e) {}
    // Fix 6: Simpler query — no deep nesting
    const {data:ci} = await sb.from("contest_items")
      .select("*, domain_items(*)")
      .eq("contest_id",contest.id)
      .order("item_order");
    // Attach domain name separately
    const domainIds = [...new Set((ci||[]).map(i=>i.domain_items?.domain_id).filter(Boolean))];
    const {data:doms} = await sb.from("domains").select("id,name").in("id",domainIds);
    const domMap = {};
    (doms||[]).forEach(d=>{ domMap[d.id]=d; });
    const items = (ci||[]).map(i=>({
      ...i,
      domain_items: i.domain_items ? {
        ...i.domain_items,
        domains: domMap[i.domain_items.domain_id] || null
      } : null
    }));
    setItems(items);

    const dids=[...new Set((ci||[]).map(i=>i.domain_items?.domain_id).filter(Boolean))];
    // Fix 3: Single query for all domain fields instead of one per domain
    const {data:allFields} = await sb.from("domain_fields").select("*").in("domain_id",dids).order("display_order");
    const fm={};
    (allFields||[]).forEach(f=>{ if(!fm[f.domain_id]) fm[f.domain_id]=[]; fm[f.domain_id].push(f); });
    setFields(fm);

    const {data:et}=await sb.from("tasks").select("*, responses(*)").eq("contest_id",contest.id).eq("user_id",user.id);
    const tm={}; const am={};
    (et||[]).forEach(t=>{
      tm[t.contest_item_id]=t;
      am[t.id]={};
      (t.responses||[]).forEach(r=>{am[t.id][r.field_name]=r.user_value;});
    });
    setTasks(tm); setAnswers(am);

    const allSub=(ci||[]).length>0&&(ci||[]).every(item=>tm[item.id]?.status==="submitted");
    if (allSub&&contest.mode==="assessment") {
      setSubmitted(true);
      // Stagger score queries — random delay prevents thundering herd
      await new Promise(r=>setTimeout(r, Math.random()*3000));
      try { const {data:s}=await sb.from("v_user_contest_accuracy").select("*").eq("contest_id",contest.id).eq("user_id",user.id).single(); if(s) setScore(s); } catch(e){}
    }
    setLoading(false);
  }

  // Tasks are pre-created on activate — just update status to in_progress
  async function ensureTask(item) {
    const existing = tasks[item.id];
    if (existing) {
      // Update to in_progress if not_started
      if (existing.status==="not_started") {
        await sb.from("tasks").update({status:"in_progress"}).eq("id",existing.id);
        setTasks(prev=>({...prev,[item.id]:{...prev[item.id],status:"in_progress"}}));
      }
      return existing;
    }
    // Fallback: create if somehow missing
    if (creatingTask.current[item.id]) {
      await new Promise(r=>setTimeout(r,500));
      return tasks[item.id];
    }
    creatingTask.current[item.id]=true;
    const {data:t}=await sb.from("tasks").insert({
      contest_id:contest.id,user_id:user.id,contest_item_id:item.id,
      status:"in_progress",started_at:new Date().toISOString(),
    }).select().single();
    creatingTask.current[item.id]=false;
    if (t) setTasks(prev=>({...prev,[item.id]:t}));
    return t;
  }

  // Update local state only — no DB write
  function saveAns(item, fieldName, val) {
    // Get existing task id if available
    const task = tasks[item.id];
    const taskId = task?.id || item.id; // fallback key
    setAnswers(prev=>({...prev,[taskId]:{...(prev[taskId]||{}),[fieldName]:val}}));
  }

  // Batch save all answers for a task to DB — called on navigate/submit
  // Includes: save lock, silent null detection, timeout race, 4-attempt retry
  async function flushTask(item, {silent=false}={}) {
    if (contest.mode==="practice") return true;
    const task = tasks[item.id];
    if (!task || task.status==="submitted") return true;

    const taskAnswers = answers[task.id] || {};
    if (Object.keys(taskAnswers).length === 0) return true;

    const did = item.domain_items?.domain_id;
    const rows = [];
    for (const [fieldName, val] of Object.entries(taskAnswers)) {
      if (!val && val !== 0) continue;
      const golden = String(item.domain_items?.json_value?.[fieldName]||"");
      const fd = (fields[did]||[]).find(f=>f.field_name===fieldName);
      const ct = fd?.comparison_type||"as_is";
      const sc = scoreF(val, golden, ct, contest.semantic_correct_threshold||0.7);
      rows.push({task_id:task.id, field_name:fieldName, user_value:String(val), golden_value:golden, score:sc, comparison_type:ct, is_draft:true});
    }

    if (rows.length === 0) return true;

    // Lock navigation
    isSaving.current = true;
    if (!silent) setSaveStatus('saving');

    // Helper: single save attempt with timeout race and silent null detection
    async function attemptSave() {
      const timeout = new Promise((_,reject)=>setTimeout(()=>reject(new Error('timeout')),8000));
      const save = sb.from("responses").upsert(rows,{onConflict:"task_id,field_name"}).select('id');
      const {error, data} = await Promise.race([save, timeout.then(()=>{throw new Error('timeout')})]).catch(e=>({error:e,data:null}));
      if (error || !data || data.length === 0) throw new Error(error?.message||'silent-null');
    }

    // 4 attempts with increasing delays
    const delays = [0, 2000, 5000, 10000];
    let saved = false;
    for (let i=0; i<delays.length; i++) {
      if (delays[i]>0) await new Promise(r=>setTimeout(r,delays[i]));
      try {
        await attemptSave();
        saved = true;
        break;
      } catch(e) {
        console.warn(`Save attempt ${i+1} failed:`, e.message);
      }
    }

    if (saved) {
      if (!silent) { setSaveStatus('saved'); setTimeout(()=>setSaveStatus('idle'),2000); }
      // Update task status to in_progress if not_started
      if (task.status==="not_started") {
        await sb.from("tasks").update({status:"in_progress"}).eq("id",task.id);
        setTasks(prev=>({...prev,[item.id]:{...prev[item.id],status:"in_progress"}}));
      }
    } else {
      if (!silent) setSaveStatus('failed');
    }

    // Release lock and execute any queued navigation
    isSaving.current = false;
    if (pendingNav.current !== null && saved) {
      const target = pendingNav.current;
      pendingNav.current = null;
      // Brief pause so user sees "Saved ✓" before navigating
      await new Promise(r=>setTimeout(r,400));
      setIdx(target);
      setShowVal(false);
    }

    return saved;
  }

  // Check contest is still active before any user action
  async function checkContestOpen() {
    if (closingRef.current) return false;
    const {data:c} = await sb.from("contests").select("status").eq("id",contest.id).single();
    if (c?.status==="closed") {
      closingRef.current = true;
      setContestClosed(true);
      await autoSubmitAll();
      return false;
    }
    return true;
  }

  async function submitPractice(item) {
    // Save all answers to DB before validating (practice mode saves on validate)
    let task = tasks[item.id];

    // Get answers — may be keyed by item.id if task not yet created
    const existingAnswers = answers[task?.id] || answers[item.id] || {};

    // Create task if it doesn't exist yet
    if (!task) {
      const {data:t}=await sb.from("tasks").insert({
        contest_id:contest.id,user_id:user.id,contest_item_id:item.id,
        status:"in_progress",started_at:new Date().toISOString(),
      }).select().single();
      task=t;
      if(task) {
        setTasks(prev=>({...prev,[item.id]:task}));
        // Move answers from item.id key to task.id key
        setAnswers(prev=>({...prev,[task.id]:existingAnswers}));
      }
    }

    // Save responses for practice mode
    // answers may be keyed by item.id (fallback) if task didn't exist when user typed
    if (task) {
      const taskAnswers = answers[task.id] || answers[item.id] || {};
      const did = item.domain_items?.domain_id;
      const rows = [];
      for (const [fieldName, val] of Object.entries(taskAnswers)) {
        if (!val && val !== 0) continue;
        const golden = String(item.domain_items?.json_value?.[fieldName]||"");
        const fd = (fields[did]||[]).find(f=>f.field_name===fieldName);
        const ct = fd?.comparison_type||"as_is";
        const sc = scoreF(val, golden, ct, contest.semantic_correct_threshold||0.7);
        rows.push({task_id:task.id, field_name:fieldName, user_value:String(val), golden_value:golden, score:sc, comparison_type:ct, is_draft:false});
      }
      if (rows.length > 0) {
        const {error:pre} = await sb.from("responses").upsert(rows, {onConflict:"task_id,field_name"});
        if (pre) {
          await new Promise(r=>setTimeout(r,2000));
          const {error:pre2} = await sb.from("responses").upsert(rows, {onConflict:"task_id,field_name"});
          if (pre2) { showToast("⚠️ Save failed — please try validating again"); return; }
        }
      }
      // Mark task submitted
      const {error:pte} = await sb.from("tasks").update({status:"submitted",submitted_at:new Date().toISOString()}).eq("id",task.id);
      if (pte) {
        await new Promise(r=>setTimeout(r,2000));
        await sb.from("tasks").update({status:"submitted",submitted_at:new Date().toISOString()}).eq("id",task.id);
      }
      setTasks(prev=>({...prev,[item.id]:{...prev[item.id],status:"submitted"}}));
    }
    setShowVal(true);
  }

  // Silent auto-submit — called when admin closes contest
  async function autoSubmitAll() {
    try {
      setSubmitting(true);

      // Build current task rows before any async operations
      const currentTaskRows = buildCurrentTaskRows(items[idx]);

      // ── PRIMARY PATH: Edge Function ───────────────────────────────────────
      // Don't fetch score on auto-submit — DB is under load from contest close
      const efResult = await callSubmitEdgeFunction(currentTaskRows, false);

      if (efResult?.success) {
        // Edge Function succeeded — score will load when user checks Past tab
        if (contest.mode==="assessment" && efResult.score) setScore(efResult.score);
        setSubmitting(false);
        return;
      }

      // ── FALLBACK PATH: Browser submit ─────────────────────────────────────
      console.warn("Edge Function failed on auto-submit — falling back to browser submit");

      // Save current task first
      if (items[idx]) {
        await flushTask(items[idx], {silent:true});
      }

      const {data:latestItems} = await sb.from("contest_items")
        .select("*, domain_items(*)")
        .eq("contest_id",contest.id).order("item_order");
      const {data:latestTasks} = await sb.from("tasks").select("*, responses(*)")
        .eq("contest_id",contest.id).eq("user_id",user.id);
      const tm={};
      (latestTasks||[]).forEach(t=>{ tm[t.contest_item_id]=t; });

      const inProgressIds = (latestItems||[])
        .filter(item=>tm[item.id]&&(tm[item.id].status==="in_progress"||tm[item.id].status==="not_started"))
        .map(item=>tm[item.id].id);

      if (inProgressIds.length>0) {
        const {error:ate} = await sb.from("tasks").update({status:"submitted",submitted_at:new Date().toISOString()}).in("id",inProgressIds);
        if (ate) {
          await new Promise(r=>setTimeout(r,2000));
          await sb.from("tasks").update({status:"submitted",submitted_at:new Date().toISOString()}).in("id",inProgressIds);
        }
      }

      const missingItems = (latestItems||[]).filter(item=>!tm[item.id]);
      for (const item of missingItems) {
        await sb.from("tasks").insert({
          contest_id:contest.id, user_id:user.id, contest_item_id:item.id,
          status:"submitted", started_at:new Date().toISOString(), submitted_at:new Date().toISOString(),
        });
      }

      if (contest.mode==="assessment") {
        // Stagger score queries — random delay prevents thundering herd
        await new Promise(r=>setTimeout(r, Math.random()*5000));
        try { const {data:s} = await sb.from("v_user_contest_accuracy").select("*").eq("contest_id",contest.id).eq("user_id",user.id).single(); if(s) setScore(s); } catch(e){}
      }
      setSubmitting(false);
    } catch(e) { console.error("Auto-submit error:",e); setSubmitting(false); }
  }

  // Build pre-scored rows for current task — used by Edge Function
  function buildCurrentTaskRows(item) {
    if (!item) return [];
    const task = tasks[item.id];
    if (!task || task.status==="submitted") return [];
    const taskAnswers = answers[task.id] || {};
    const did = item.domain_items?.domain_id;
    const rows = [];
    for (const [fieldName, val] of Object.entries(taskAnswers)) {
      if (!val && val !== 0) continue;
      const golden = String(item.domain_items?.json_value?.[fieldName]||"");
      const fd = (fields[did]||[]).find(f=>f.field_name===fieldName);
      const ct = fd?.comparison_type||"as_is";
      const sc = scoreF(val, golden, ct, contest.semantic_correct_threshold||0.7);
      rows.push({task_id:task.id, field_name:fieldName, user_value:String(val), golden_value:golden, score:sc, comparison_type:ct, is_draft:true});
    }
    return rows;
  }

  // Call submit-contest Edge Function — primary submit path
  async function callSubmitEdgeFunction(currentTaskRows=[], fetchScore=true) {
    const delays = [0, 3000, 6000];
    for (let i=0; i<delays.length; i++) {
      if (delays[i]>0) await new Promise(r=>setTimeout(r,delays[i]));
      try {
        const {data:{session}} = await sb.auth.getSession();
        const timeout = new Promise((_,reject)=>setTimeout(()=>reject(new Error('timeout')),15000));
        const call = fetch(`${SUPABASE_URL}/functions/v1/submit-contest`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${session?.access_token}`,
            "apikey": SUPABASE_ANON,
          },
          body: JSON.stringify({ contestId: contest.id, currentTaskRows, fetchScore }),
        }).then(r=>r.json());
        const result = await Promise.race([call, timeout]);
        if (result?.success) return result;
        console.warn(`Edge Function attempt ${i+1} failed:`, result?.error);
      } catch(e) {
        console.warn(`Edge Function attempt ${i+1} error:`, e.message);
      }
    }
    return null; // All attempts failed — caller falls back to browser submit
  }

  async function submitAll() {
    const open = await checkContestOpen();
    if (!open) return;
    // Wait if save is in progress
    if (isSaving.current) {
      showToast("Please wait — saving current task...");
      await new Promise(r=>{ const t=setInterval(()=>{ if(!isSaving.current){clearInterval(t);r();} },200); });
    }
    const incomplete=items.filter(i=>!tasks[i.id]||tasks[i.id].status==="not_started");
    if (incomplete.length>0&&!confirm(`${incomplete.length} task(s) have no answers. Submit all anyway?`)) return;
    setSubmitting(true);

    // Build current task rows to send to Edge Function
    const currentTaskRows = buildCurrentTaskRows(items[idx]);

    // ── PRIMARY PATH: Edge Function (server-side atomic submit) ──────────────
    const efResult = await callSubmitEdgeFunction(currentTaskRows);

    if (efResult?.success) {
      // Edge Function succeeded — update local state and show score
      const updatedTasks = {...tasks};
      items.forEach(item=>{
        if(tasks[item.id]) updatedTasks[item.id]={...tasks[item.id],status:"submitted"};
      });
      setTasks(updatedTasks);
      if (efResult.score) setScore(efResult.score);
      setSubmitted(true);
      setSubmitting(false);
      return;
    }

    // ── FALLBACK PATH: Browser submit (if Edge Function failed) ──────────────
    console.warn("Edge Function failed — falling back to browser submit");

    // Flush current task first
    const saved = await flushTask(items[idx]);
    if (!saved) {
      const proceed = window.confirm("⚠️ Current task could not be saved. Submit anyway? (answers for this task will be lost)");
      if (!proceed) { setSubmitting(false); return; }
    }

    const updatedTasks={...tasks};
    const inProgressTaskIds = items
      .filter(i=>tasks[i.id]&&(tasks[i.id].status==="in_progress"||tasks[i.id].status==="not_started"))
      .map(i=>tasks[i.id].id);
    const missingItems = items.filter(i=>!tasks[i.id]);

    if (inProgressTaskIds.length>0) {
      const {error:te} = await sb.from("tasks").update({status:"submitted",submitted_at:new Date().toISOString()}).in("id",inProgressTaskIds);
      if (te) {
        await new Promise(r=>setTimeout(r,2000));
        await sb.from("tasks").update({status:"submitted",submitted_at:new Date().toISOString()}).in("id",inProgressTaskIds);
      }
      inProgressTaskIds.forEach(tid=>{
        const item = items.find(i=>tasks[i.id]?.id===tid);
        if(item) updatedTasks[item.id]={...tasks[item.id],status:"submitted"};
      });
    }
    for (const item of missingItems) {
      const {data:t}=await sb.from("tasks").insert({
        contest_id:contest.id,user_id:user.id,contest_item_id:item.id,
        status:"submitted",started_at:new Date().toISOString(),submitted_at:new Date().toISOString(),
      }).select().single();
      if(t)updatedTasks[item.id]=t;
    }
    setTasks(updatedTasks);
    // Stagger score queries — random delay prevents thundering herd on contest close
    await new Promise(r=>setTimeout(r, Math.random()*3000));
    try { const {data:s}=await sb.from("v_user_contest_accuracy").select("*").eq("contest_id",contest.id).eq("user_id",user.id).single(); if(s) setScore(s); } catch(e){}
    setSubmitted(true);
    setSubmitting(false);
  }

  if (loading) return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:"var(--bg)"}}>
      <span className="sp" style={{width:32,height:32}}/>
    </div>
  );

  // Block draft contests
  if (contest.status==="draft") return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:"var(--bg)"}}>
      <div className="card" style={{maxWidth:440,textAlign:"center",padding:40}}>
        <div style={{fontSize:36,marginBottom:16}}>⏳</div>
        <div className="fw6" style={{fontSize:20,marginBottom:8}}>Contest not yet started</div>
        <div className="sm m2" style={{marginBottom:24}}>Your admin hasn't activated this contest yet. Please check back later.</div>
        <button className="bg" onClick={onClose}>← Back to contests</button>
      </div>
    </div>
  );

  // Show closing screen — auto-submitted, show scores
  if (contestClosed && !submitted) return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:"var(--bg)"}}>
      <div className="card" style={{maxWidth:500,textAlign:"center",padding:44,boxShadow:"var(--shadow2)"}}>
        <div style={{fontSize:48,marginBottom:16}}>📋</div>
        <div className="fw6" style={{fontSize:22,marginBottom:6}}>Contest has ended</div>
        <div className="sm m2" style={{marginBottom:24}}>Your answers have been automatically saved and submitted.</div>
        {score&&contest.mode==="assessment"&&(<>
          <div className="g3c" style={{marginBottom:24}}>
            {[{v:score.tasks_submitted,l:"Tasks done"},{v:`${score.correct_attributes}/${score.total_attributes}`,l:"Correct"},{v:`${score.accuracy_pct}%`,l:"Accuracy"}].map(x=>(
              <div className="sc" key={x.l}><div className="sv">{x.v}</div><div className="sl">{x.l}</div></div>
            ))}
          </div>
          <div className={`badge cert-${score.cert_level?.toLowerCase()}`} style={{fontSize:20,padding:"10px 24px",borderRadius:10,display:"inline-flex",marginBottom:8}}>
            {score.cert_level}
          </div>
          <div className="xs m3" style={{marginBottom:24}}>Based on {score.accuracy_pct}% accuracy</div>
        </>)}
        {submitting&&<div className="sm m2" style={{marginBottom:16}}><span className="sp"/> &nbsp;Submitting your answers...</div>}
        <button className="bg" onClick={onClose}>← Back to contests</button>
      </div>
    </div>
  );

  // Results screen
  if (submitted&&contest.mode==="assessment") return (
    <div style={{maxWidth:600,margin:"60px auto",padding:"0 20px"}}>
      <div className="card" style={{textAlign:"center",padding:44,boxShadow:"var(--shadow2)"}}>
        <div style={{fontSize:52,marginBottom:16}}>✓</div>
        <div style={{fontSize:24,fontWeight:700,marginBottom:8}}>Assessment submitted</div>
        <div className="sm m2" style={{marginBottom:28}}>{contest.name}</div>
        {score ? (
          <>
            <div className="g3c" style={{marginBottom:28}}>
              {[{v:score.tasks_submitted,l:"Tasks done"},{v:`${score.correct_attributes}/${score.total_attributes}`,l:"Correct attrs"},{v:`${score.accuracy_pct}%`,l:"Accuracy"}].map(x=>(
                <div className="sc" key={x.l}><div className="sv">{x.v}</div><div className="sl">{x.l}</div></div>
              ))}
            </div>
            <div className={`badge cert-${score.cert_level?.toLowerCase()}`} style={{fontSize:22,padding:"12px 28px",borderRadius:12,display:"inline-flex"}}>
              {score.cert_level}
            </div>
            <div className="xs m3" style={{marginTop:10}}>Certification based on {score.accuracy_pct}% accuracy</div>
          </>
        ) : (
          <div className="sm m2" style={{marginBottom:16}}>
            ✅ Your answers have been saved successfully.<br/>
            <span className="xs m3">Your score will be available shortly. You may close this window.</span>
          </div>
        )}
        <button className="bg" style={{marginTop:28}} onClick={onClose}>Back to contests</button>
      </div>
    </div>
  );

  const cur=items[idx];
  if (!cur) return <div style={{padding:40,textAlign:"center",color:"var(--text3)"}}>No items found in this contest.</div>;

  const di=cur.domain_items; const did=di?.domain_id; const af=fields[did]||[];
  // Smart split: handles field names containing commas e.g. "Cleaning, Care & Maintenance"
  const rawAttrs = di?.attributes_for_category||"";
  const allFieldNames = af.map(f=>f.field_name);
  let aa = [];
  if (rawAttrs) {
    let remaining = rawAttrs;
    const found = [];
    const sortedFields = [...allFieldNames].sort((a,b)=>b.length-a.length);
    while (remaining.length > 0) {
      remaining = remaining.replace(/^[,\s]+/, '');
      if (!remaining) break;
      const match = sortedFields.find(f=>remaining.startsWith(f));
      if (match) { found.push(match); remaining = remaining.slice(match.length); }
      else {
        const ci = remaining.indexOf(',');
        if (ci === -1) { found.push(remaining.trim()); break; }
        found.push(remaining.slice(0, ci).trim());
        remaining = remaining.slice(ci + 1);
      }
    }
    aa = found.filter(Boolean);
  }
  const task=tasks[cur.id]; const ta=answers[task?.id]||answers[cur.id]||{}; const isSub=task?.status==="submitted";
  const ctx=af.filter(f=>f.field_role==="context"); // always show all context fields
  const imgs=af.filter(f=>f.field_role==="image");
  const cure=af.filter(f=>f.field_role==="curate"&&(aa.length===0||aa.includes(f.field_name))&&String(di?.json_value?.[f.field_name]||"").trim()!=="");

  function tSt(item) { const t=tasks[item.id]; if(!t)return"ns"; if(t.status==="submitted")return"sb"; return"ip"; }
  const inc=items.filter(i=>!tasks[i.id]||tasks[i.id].status!=="submitted").length;

  return (
    <div style={{display:"flex",height:"100vh",overflow:"hidden",background:"var(--bg)"}}>

      {/* Task sidebar */}
      <div className="tsb">
        <div style={{padding:"14px 8px 10px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <span className="xs fw6 m2">Tasks</span>
          <button className="bg bxs" onClick={onClose} style={{fontSize:11}}>✕ Exit</button>
        </div>
        <div style={{display:"flex",flexWrap:"wrap",gap:6,padding:"0 4px"}}>
          {items.map((item,i)=>{
            const st=tSt(item);
            return (
              <div key={item.id} className={`td td-${st} ${i===idx?"td-act":""}`}
                onClick={async()=>{
                  if(isSaving.current){pendingNav.current=i;return;}
                  const saved=await flushTask(items[idx]);
                  if(saved){setIdx(i);setShowVal(false);}
                  else{pendingNav.current=i;} // queue nav for after retry resolves
                }} title={`Task ${i+1} — ${st}`}>
                {i+1}
              </div>
            );
          })}
        </div>
        <div style={{marginTop:16,padding:"0 6px"}}>
          {[["ns","Not started"],["ip","In progress"],["sb","Submitted"]].map(([st,lbl])=>(
            <div key={st} className="fx g2 ac xs m3" style={{marginBottom:5}}>
              <div className={`td td-${st}`} style={{width:12,height:12,fontSize:0}}/>
              {lbl}
            </div>
          ))}
        </div>
        {contest.mode==="assessment"&&(
          <div style={{padding:"12px 8px",marginTop:"auto",borderTop:"1px solid var(--border)"}}>
            <button className="bp wf" style={{fontSize:12}} onClick={submitAll} disabled={submitting}>
              {submitting?<><span className="sp"/> &nbsp;Submitting...</>:`Submit all${inc>0?` (${inc} pending)`:""}`}
            </button>
          </div>
        )}
      </div>

      {/* Main task area */}
      <div style={{flex:1,overflowY:"auto",padding:"20px 26px",background:"var(--bg)"}}>
        <div className="fx ac jb" style={{marginBottom:16}}>
          <div>
            <button className="bg bxs" onClick={onClose} style={{marginBottom:6}}>← Back to contests</button>
            <div className="xs m3" style={{marginBottom:2}}>{contest.name} · {contest.mode}</div>
            <div className="fw6" style={{fontSize:16}}>Task {idx+1} of {items.length}{di?.category?` — ${di.category}`:""}</div>
          </div>
          <div className="fx ac g3">
            {saveStatus==='saving'&&(
              <div className="fx ac g1" style={{fontSize:12,color:'#6b7280'}}>
                <span className="sp" style={{width:12,height:12}}/> Saving...
              </div>
            )}
            {saveStatus==='saved'&&(
              <div style={{fontSize:12,color:'#16a34a',fontWeight:600}}>✓ Saved</div>
            )}
            {saveStatus==='failed'&&(
              <div className="fx ac g1">
                <div style={{fontSize:12,color:'#dc2626',fontWeight:600}}>⚠️ Save failed</div>
                <button style={{fontSize:11,padding:'2px 8px',borderRadius:4,border:'1px solid #dc2626',background:'white',color:'#dc2626',cursor:'pointer'}}
                  onClick={async()=>{ const saved=await flushTask(items[idx]); if(saved&&pendingNav.current!==null){const t=pendingNav.current;pendingNav.current=null;setIdx(t);setShowVal(false);} }}>
                  Retry
                </button>
                <button style={{fontSize:11,padding:'2px 8px',borderRadius:4,border:'1px solid #9ca3af',background:'white',color:'#6b7280',cursor:'pointer'}}
                  onClick={()=>{ setSaveStatus('idle'); if(pendingNav.current!==null){const t=pendingNav.current;pendingNav.current=null;setIdx(t);setShowVal(false);} }}>
                  Skip
                </button>
              </div>
            )}
            <div className="fx g2">
              <button className="bg bsm" disabled={idx===0} onClick={async()=>{
                if(isSaving.current){pendingNav.current=idx-1;return;}
                const saved=await flushTask(items[idx]);
                if(saved){setIdx(i=>i-1);setShowVal(false);}
              }}>← Prev</button>
              <button className="bg bsm" disabled={idx===items.length-1} onClick={async()=>{
                if(isSaving.current){pendingNav.current=idx+1;return;}
                const saved=await flushTask(items[idx]);
                if(saved){setIdx(i=>i+1);setShowVal(false);}
              }}>Next →</button>
            </div>
          </div>
        </div>

        {/* Context card */}
        <div className="card" style={{marginBottom:16,padding:16}}>
          {imgs.length>0&&(
            <div className="fx g3 wrap" style={{marginBottom:12}}>
              {imgs.slice(0,6).map(f=>{
                const url=di?.json_value?.[f.field_name];
                const fullUrl = url ? url.replace(/[?&](odnHeight|odnWidth|odnBg)=[^&]*/g,"").replace(/[?&]$/,"") : null;
                return url?<img key={f.field_name} src={url} alt="" loading="lazy" className="imt"
                  style={{cursor:"pointer"}}
                  onClick={()=>window.open(fullUrl,"_blank")}
                  onMouseEnter={e=>e.target.style.opacity=".8"}
                  onMouseLeave={e=>e.target.style.opacity="1"}
                  onError={e=>e.target.style.display="none"}/>:null;
              })}
            </div>
          )}
          <div className="g2c">
            {ctx.map(f=>(
              <div key={f.field_name}>
                <div className="xs m3" style={{marginBottom:2}}>{f.field_name}</div>
                <div className="sm fw5">{String(di?.json_value?.[f.field_name]||"—")}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Validation results (practice) */}
        {showVal&&isSub&&(
          <div className="card" style={{marginBottom:16}}>
            <div className="fw6" style={{marginBottom:12}}>Validation results</div>
            {cure.map(f=>{
              const uv=ta[f.field_name]||"";
              const gv=String(di?.json_value?.[f.field_name]||"");
              const isMulti=f.comparison_type==="multiselect"||f.comparison_type==="list";
              const sc2=scoreF(uv,gv,f.comparison_type,contest.semantic_correct_threshold||0.7);
              const ok=f.comparison_type==="semantic"?sc2>=(contest.semantic_correct_threshold||0.7):f.comparison_type==="multiselect"||f.comparison_type==="list"?sc2>=0.99:sc2===1;
              return (
                <div key={f.field_name} className={ok?"fsc":sc2>0?"fsp":"fsw"}>
                  <div className="fx ac jb">
                    <span className="xs fw6">{f.field_name}</span>
                    <span className="xs mono fw6">{Math.round(sc2*100)}%</span>
                  </div>
                  <div className="fx g4 xs" style={{marginTop:4,flexWrap:"wrap"}}>
                    <span><span style={{opacity:.6}}>Your answer: </span>
                      {isMulti && uv ? (
                        uv.split(",").map(s=>s.trim()).filter(Boolean).map(s=>(
                          <span key={s} style={{marginRight:4,padding:"2px 8px",borderRadius:4,fontSize:12,
                            background:gv.toLowerCase().split(",").map(x=>x.trim()).includes(s.toLowerCase())?"var(--green-bg)":"var(--red-bg)",
                            color:gv.toLowerCase().split(",").map(x=>x.trim()).includes(s.toLowerCase())?"var(--green)":"var(--red)"}}>
                            {s}
                          </span>
                        ))
                      ) : <strong>{uv||"(blank)"}</strong>}
                    </span>
                    <span><span style={{opacity:.6}}>Correct: </span><strong>{gv||"(blank)"}</strong></span>
                  </div>
                </div>
              );
            })}
            <div className="fx g3 jb" style={{marginTop:16}}>
              <button className="bg bsm" disabled={idx===0} onClick={async()=>{
                if(isSaving.current){pendingNav.current=idx-1;return;}
                await flushTask(items[idx]);setIdx(i=>i-1);setShowVal(false);
              }}>← Prev task</button>
              {idx<items.length-1&&<button className="bg bsm" onClick={async()=>{
                if(isSaving.current){pendingNav.current=idx+1;return;}
                await flushTask(items[idx]);setIdx(i=>i+1);setShowVal(false);
              }}>Next task →</button>}
            </div>
          </div>
        )}

        {/* Curate fields */}
        {!showVal&&(
          <div className="card">
            <div className="fx ac jb" style={{marginBottom:16}}>
              <div className="fw6">Fill in attributes</div>
              {isSub&&contest.mode==="practice"&&<span className="badge b-green">✓ Submitted</span>}
              {isSub&&contest.mode==="assessment"&&<span className="badge b-blue">Saved</span>}
            </div>
            {cure.length===0&&<div className="sm m3">No curate fields configured for this item.</div>}
            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              {cure.map(f=>(
                <CurateField
                  key={`${cur.id}-${f.field_name}`}
                  fieldDef={f}
                  initialValue={ta[f.field_name]||""}
                  disabled={(isSub&&contest.mode==="practice")||contestClosed}
                  onSave={(fieldName, val)=>saveAns(cur, fieldName, val)}
                />
              ))}
            </div>
            <div className="fx g3 jb" style={{marginTop:20}}>
              <button className="bg bsm" disabled={idx===0} onClick={async()=>{
  if(isSaving.current){pendingNav.current=idx-1;return;}
  const saved=await flushTask(cur);
  if(!saved){return;} // blocked — save failed, user sees indicator
  const open=await checkContestOpen();
  if(open){setIdx(i=>i-1);setShowVal(false);}
}}>← Previous</button>
              <div className="fx g2">
                {contest.mode==="practice"&&!isSub&&(
                  <button className="bp bsm" onClick={async()=>{ const open=await checkContestOpen(); if(open) submitPractice(cur); }}>Validate ✓</button>
                )}
                {contest.mode==="practice"&&isSub&&(
                  <button className="bg bsm" onClick={()=>setShowVal(true)}>See answers</button>
                )}
                {idx<items.length-1&&(
                  <button className="bg bsm" onClick={async()=>{
                    if(isSaving.current){pendingNav.current=idx+1;return;}
                    const saved=await flushTask(cur);
                    if(!saved){return;} // blocked — save failed
                    const open=await checkContestOpen();
                    if(open){setIdx(i=>i+1);setShowVal(false);}
                  }}>Next →</button>
                )}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── ROOT ─────────────────────────────────────────────────────
export default function App() {
  const [prof,setProf]=useState(null);
  const [loading,setLoading]=useState(true);

  useEffect(()=>{
    const timeout = setTimeout(()=>setLoading(false), 6000);
    sb.auth.getSession().then(async({data:{session}})=>{
      clearTimeout(timeout);
      if (session?.user) {
        const {data:prof} = await sb.rpc("get_my_profile");
        setProf(prof||null);
      }
      setLoading(false);
    }).catch(()=>{ clearTimeout(timeout); setLoading(false); });

    const {data:{subscription}}=sb.auth.onAuthStateChange((event)=>{
      if (event==="SIGNED_OUT") setProf(null);
    });
    return ()=>subscription.unsubscribe();
  },[]);

  async function logout() { await sb.auth.signOut(); setProf(null); }

  if (loading) return (
    <>
      <style>{css}</style>
      <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:"var(--bg)"}}>
        <span className="sp" style={{width:32,height:32}}/>
      </div>
    </>
  );

  return (
    <>
      <style>{css}</style>
      {!prof
        ? <LoginPage onLogin={p=>setProf(p)}/>
        : prof.role==="admin"
          ? <AdminShell user={prof} onLogout={logout}/>
          : <UserShell user={prof} onLogout={logout}/>
      }
    </>
  );
}

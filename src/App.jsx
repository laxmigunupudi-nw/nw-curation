import { useState, useEffect, useRef } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";

const SUPABASE_URL = "https://hyoiwwmhrhkhjhrueuon.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imh5b2l3d21ocmhraGpocnVldW9uIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzkwMzgzNDksImV4cCI6MjA5NDYxNDM0OX0.DUvfR7q0qPpIe5gkDFVQf40JUhrNPEg98GVmb3HoMpg";
const sb = createClient(SUPABASE_URL, SUPABASE_ANON);

const css = `
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500;9..40,600&family=DM+Mono:wght@400;500&display=swap');
  *{box-sizing:border-box;margin:0;padding:0}
  :root{
    --bg:#0f1117;--bg2:#181c27;--bg3:#1f2435;--bg4:#252b3b;
    --border:#2e3548;--border2:#3d4560;
    --text:#e8eaf2;--text2:#9aa0b8;--text3:#5c6280;
    --accent:#4f7fff;--accent2:#3d6bef;--ag:rgba(79,127,255,.18);
    --green:#34d399;--gbg:rgba(52,211,153,.12);
    --red:#f87171;--rbg:rgba(248,113,113,.12);
    --amber:#fbbf24;--abg:rgba(251,191,36,.12);
    --purple:#a78bfa;--pbg:rgba(167,139,250,.12);
    --r:10px;--r2:14px;--r3:20px;
    --f:'DM Sans',sans-serif;--mono:'DM Mono',monospace;
  }
  html,body,#root{background:var(--bg);color:var(--text);font-family:var(--f);font-size:14px;line-height:1.6;min-height:100vh}
  input,textarea,select{background:var(--bg3);border:1px solid var(--border);border-radius:var(--r);color:var(--text);font-family:var(--f);font-size:14px;padding:8px 12px;width:100%;outline:none;transition:border .2s}
  input:focus,textarea:focus,select:focus{border-color:var(--accent)}
  input::placeholder{color:var(--text3)}
  button{cursor:pointer;font-family:var(--f);font-size:14px;border:none;border-radius:var(--r);padding:8px 16px;transition:all .18s;font-weight:500}
  .bp{background:var(--accent);color:#fff}.bp:hover{background:var(--accent2)}
  .bp:disabled{opacity:.6;cursor:not-allowed;transform:none}
  .bg{background:transparent;color:var(--text2);border:1px solid var(--border)}.bg:hover{border-color:var(--border2);color:var(--text);background:var(--bg3)}
  .bg:disabled{opacity:.5;cursor:not-allowed}
  .bd{background:var(--rbg);color:var(--red);border:1px solid rgba(248,113,113,.25)}.bd:hover{background:rgba(248,113,113,.2)}
  .bsm{padding:5px 11px;font-size:13px}.bxs{padding:3px 8px;font-size:12px}
  .card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r2);padding:18px}
  .badge{display:inline-flex;align-items:center;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:500}
  .b-green{background:var(--gbg);color:var(--green)}.b-red{background:var(--rbg);color:var(--red)}
  .b-amber{background:var(--abg);color:var(--amber)}.b-blue{background:var(--ag);color:var(--accent)}
  .b-purple{background:var(--pbg);color:var(--purple)}.b-gray{background:var(--bg4);color:var(--text2)}
  table{width:100%;border-collapse:collapse}
  th{text-align:left;padding:9px 13px;font-size:11px;font-weight:500;color:var(--text3);text-transform:uppercase;letter-spacing:.05em;border-bottom:1px solid var(--border)}
  td{padding:11px 13px;border-bottom:1px solid var(--border);color:var(--text2);font-size:13px}
  tr:last-child td{border-bottom:none}
  tr:hover td{background:var(--bg3);color:var(--text)}
  .tag{display:inline-block;padding:2px 7px;border-radius:6px;font-size:11px;font-weight:500;background:var(--bg4);color:var(--text2)}
  .div{height:1px;background:var(--border);margin:16px 0}
  .mo{position:fixed;inset:0;background:rgba(0,0,0,.72);display:flex;align-items:center;justify-content:center;z-index:1000;padding:16px}
  .md{background:var(--bg2);border:1px solid var(--border2);border-radius:var(--r3);padding:24px;width:100%;max-width:520px;max-height:92vh;overflow-y:auto}
  .md-lg{max-width:760px}
  .fg{margin-bottom:14px}
  .fl{display:block;font-size:11px;font-weight:500;color:var(--text2);margin-bottom:5px;text-transform:uppercase;letter-spacing:.04em}
  .row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
  .row3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}
  .fx{display:flex}.g2{gap:8px}.g3{gap:12px}.g4{gap:16px}
  .ac{align-items:center}.jb{justify-content:space-between}.f1{flex:1}.wrap{flex-wrap:wrap}
  .g2c{display:grid;grid-template-columns:1fr 1fr;gap:14px}
  .g3c{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px}
  .g4c{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
  .sm{font-size:13px}.xs{font-size:12px}.xl{font-size:22px}
  .m2{color:var(--text2)}.m3{color:var(--text3)}
  .mono{font-family:var(--mono)}
  .fw5{font-weight:500}.fw6{font-weight:600}
  .cg{color:var(--green)}.cr{color:var(--red)}.ca{color:var(--amber)}.cb{color:var(--accent)}
  .wf{width:100%}
  .mt2{margin-top:8px}.mt3{margin-top:12px}.mt4{margin-top:16px}
  .mb2{margin-bottom:8px}.mb3{margin-bottom:12px}.mb4{margin-bottom:16px}
  .sp{width:18px;height:18px;border:2px solid var(--border2);border-top-color:var(--accent);border-radius:50%;animation:spin .7s linear infinite;display:inline-block;vertical-align:middle}
  @keyframes spin{to{transform:rotate(360deg)}}
  .fi{animation:fi .25s ease}
  @keyframes fi{from{opacity:0;transform:translateY(5px)}to{opacity:1;transform:translateY(0)}}
  .sidebar{width:210px;min-width:210px;background:var(--bg2);border-right:1px solid var(--border);height:100vh;display:flex;flex-direction:column;position:fixed;left:0;top:0;z-index:10}
  .mc{margin-left:210px;min-height:100vh;padding:26px}
  .ni{display:flex;align-items:center;gap:9px;padding:9px 14px;border-radius:var(--r);color:var(--text2);cursor:pointer;transition:all .15s;font-size:13px;margin:1px 7px}
  .ni:hover{background:var(--bg3);color:var(--text)}
  .ni.act{background:var(--ag);color:var(--accent);font-weight:500}
  .ns{font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:.08em;padding:14px 14px 5px}
  .pt{font-size:21px;font-weight:600;color:var(--text);margin-bottom:3px}
  .ps{font-size:13px;color:var(--text2);margin-bottom:22px}
  .sc{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r2);padding:16px 18px}
  .sv{font-size:27px;font-weight:600;color:var(--text);line-height:1.2}
  .sl{font-size:11px;color:var(--text3);margin-top:3px;text-transform:uppercase;letter-spacing:.04em}
  .pb{height:6px;background:var(--bg4);border-radius:3px;overflow:hidden}
  .pf{height:100%;border-radius:3px;background:var(--accent);transition:width .4s}
  .tabs{display:flex;gap:3px;background:var(--bg3);padding:3px;border-radius:var(--r2);margin-bottom:18px}
  .tab{padding:7px 16px;border-radius:var(--r);cursor:pointer;font-size:13px;color:var(--text2);transition:all .15s}
  .tab.act{background:var(--bg2);color:var(--text);font-weight:500;box-shadow:0 1px 4px rgba(0,0,0,.3)}
  .tsb{width:190px;min-width:190px;background:var(--bg2);border-right:1px solid var(--border);height:100vh;overflow-y:auto;padding:14px 7px;display:flex;flex-direction:column}
  .td{width:27px;height:27px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:600;cursor:pointer;transition:all .15s;flex-shrink:0}
  .td-ns{background:var(--bg4);color:var(--text3)}
  .td-ip{background:var(--abg);color:var(--amber);border:1px solid var(--amber)}
  .td-sb{background:var(--gbg);color:var(--green)}
  .td-act{box-shadow:0 0 0 2px var(--accent);transform:scale(1.1)}
  .fsc{background:var(--gbg);border:1px solid rgba(52,211,153,.25);border-radius:var(--r);padding:9px 11px;margin-bottom:7px}
  .fsw{background:var(--rbg);border:1px solid rgba(248,113,113,.25);border-radius:var(--r);padding:9px 11px;margin-bottom:7px}
  .fsp{background:var(--abg);border:1px solid rgba(251,191,36,.25);border-radius:var(--r);padding:9px 11px;margin-bottom:7px}
  .cert-l3{background:linear-gradient(135deg,rgba(167,139,250,.2),rgba(79,127,255,.2));border:1px solid rgba(167,139,250,.4);color:var(--purple)}
  .cert-l2{background:var(--gbg);border:1px solid rgba(52,211,153,.3);color:var(--green)}
  .cert-l1{background:var(--abg);border:1px solid rgba(251,191,36,.3);color:var(--amber)}
  .cert-l0{background:var(--rbg);border:1px solid rgba(248,113,113,.3);color:var(--red)}
  .imt{width:60px;height:60px;object-fit:cover;border-radius:var(--r);border:1px solid var(--border);background:var(--bg4)}
  .cb2{background:var(--bg3);border:1px solid var(--border2);border-radius:var(--r);padding:9px 11px;font-family:var(--mono);font-size:11px;color:var(--accent);word-break:break-all;cursor:pointer}
  .cb2:hover{background:var(--bg4)}
  .al{padding:11px 14px;border-radius:var(--r);margin-bottom:14px;font-size:13px}
  .al-e{background:var(--rbg);color:var(--red);border:1px solid rgba(248,113,113,.25)}
  .al-s{background:var(--gbg);color:var(--green);border:1px solid rgba(52,211,153,.25)}
  .al-i{background:var(--ag);color:var(--accent);border:1px solid rgba(79,127,255,.25)}
  .hc{padding:6px 10px;border-radius:7px;font-size:11px;text-align:center;font-weight:500;min-width:100px}
  ::-webkit-scrollbar{width:4px;height:4px}
  ::-webkit-scrollbar-thumb{background:var(--border2);border-radius:2px}
  select option{background:var(--bg3)}
  input[type=file]{padding:6px 11px;font-size:12px}
  input:disabled,select:disabled,textarea:disabled{opacity:.5;cursor:not-allowed}
  .toast-wrap{position:fixed;bottom:20px;right:20px;z-index:2000;animation:fi .2s ease}
`;

// ── Scoring ──────────────────────────────────────────────────
const SW = new Set(["a","an","the","and","or","of","in","on","with","for","to","is","are","was","were","be","been","by","at","from"]);
function tok(s) { if (!s) return new Set(); return new Set(s.toLowerCase().replace(/[^a-z0-9\s]/g," ").split(/\s+/).filter(t=>t.length>1&&!SW.has(t))); }
const SYN = {"round neck":["crew neck"],"crew neck":["round neck"],"hand wash":["handwash"],"machine wash":["machine washable"],"pure cotton":["100% cotton","cotton"],"waterproof":["water resistant"],"antibacterial":["antimicrobial"]};
function exSyn(s) { let r=s.toLowerCase(); for (const [k,vs] of Object.entries(SYN)) { if(r.includes(k)) vs.forEach(v=>{r+=" "+v;}); vs.forEach(v=>{if(r.includes(v))r+=" "+k;}); } return r; }
function scoreF(uv, gv, ct, st=0.7) {
  if (!uv&&!gv) return 1; if (!uv||!gv) return 0;
  const u=String(uv).trim(), g=String(gv).trim();
  if (ct==="as_is") return u.toLowerCase()===g.toLowerCase()?1:0;
  if (ct==="list") { const us=new Set(u.toLowerCase().split(",").map(x=>x.trim()).filter(Boolean)); const gs=new Set(g.toLowerCase().split(",").map(x=>x.trim()).filter(Boolean)); if(us.size===0&&gs.size===0)return 1; const i=[...us].filter(x=>gs.has(x)).length; const un=new Set([...us,...gs]).size; return un===0?0:i/un; }
  if (ct==="numeric") { const un=parseFloat(u),gn=parseFloat(g); if(isNaN(un)||isNaN(gn))return 0; return Math.abs(un-gn)<0.001?1:0; }
  if (ct==="url") { const n=s=>s.toLowerCase().replace(/\/$/,"").replace(/^https?:\/\//,""); return n(u)===n(g)?1:0; }
  if (ct==="semantic") { const ue=exSyn(u),ge=exSyn(g); const ut=tok(ue),gt=tok(ge); if(ut.size===0&&gt.size===0)return 1; const i=[...ut].filter(x=>gt.has(x)).length; const un=new Set([...ut,...gt]).size; return un===0?0:i/un; }
  return 0;
}

function fmt(d) { if (!d) return "—"; return new Date(d).toLocaleString("en-IN",{dateStyle:"medium",timeStyle:"short"}); }

function useToast() {
  const [t, st] = useState(null);
  const show = (msg, type="success") => { st({msg,type}); setTimeout(()=>st(null),3500); };
  return [t, show];
}
function Toast({ t }) {
  if (!t) return null;
  return <div className="toast-wrap"><div className={`al al-${t.type==="error"?"e":"s"}`} style={{minWidth:240,boxShadow:"0 4px 20px rgba(0,0,0,.5)"}}>{t.msg}</div></div>;
}

// ── LOGIN ────────────────────────────────────────────────────
function LoginPage({ onLogin }) {
  const [email,setEmail] = useState(""); const [pw,setPw] = useState("");
  const [mode,setMode] = useState("login"); const [np,setNp] = useState(""); const [cf,setCf] = useState("");
  const [load,setLoad] = useState(false); const [err,setErr] = useState("");

  async function doLogin(e) {
    e.preventDefault(); setLoad(true); setErr("");
    const {data,error} = await sb.auth.signInWithPassword({email,password:pw});
    if (error) { setErr(error.message); setLoad(false); return; }
    const {data:u} = await sb.from("users").select("*").eq("id",data.user.id).single();
    if (u?.status==="disabled") { await sb.auth.signOut(); setErr("Account disabled. Contact admin."); setLoad(false); return; }
    onLogin(u); setLoad(false);
  }

  async function doSetPw(e) {
    e.preventDefault();
    if (np!==cf) { setErr("Passwords do not match"); return; }
    if (np.length<6) { setErr("Min 6 characters"); return; }
    setLoad(true); setErr("");
    const {data:u} = await sb.from("users").select("id,status").eq("email",email.toLowerCase()).single();
    if (!u) { setErr("Email not found. Ask admin to add you first."); setLoad(false); return; }
    if (u.status==="disabled") { setErr("Account disabled."); setLoad(false); return; }
    // Try updating password (works if already signed in or via signUp)
    const {error:upErr} = await sb.auth.updateUser({password:np});
    if (upErr) {
      // User not in auth yet — sign them up
      const {error:suErr} = await sb.auth.signUp({email,password:np});
      if (suErr) { setErr(suErr.message); setLoad(false); return; }
    }
    const {data:si,error:sie} = await sb.auth.signInWithPassword({email,password:np});
    if (sie) { setErr(sie.message); setLoad(false); return; }
    const {data:prof} = await sb.from("users").select("*").eq("id",si.user.id).single();
    onLogin(prof); setLoad(false);
  }

  return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"var(--bg)"}}>
      <div style={{width:"100%",maxWidth:380,padding:"0 20px"}}>
        <div style={{textAlign:"center",marginBottom:28}}>
          <div style={{fontSize:28,fontWeight:700,letterSpacing:"-0.5px"}}>NW Curation</div>
          <div style={{color:"var(--text3)",marginTop:5,fontSize:13}}>Retail Item Quality Assessment</div>
        </div>
        <div className="card" style={{borderRadius:"var(--r3)",padding:26}}>
          <div className="tabs" style={{marginBottom:22}}>
            <div className={`tab ${mode==="login"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>{setMode("login");setErr("");}}>Sign in</div>
            <div className={`tab ${mode==="set"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>{setMode("set");setErr("");}}>First time / Reset</div>
          </div>
          {err && <div className="al al-e">{err}</div>}
          {mode==="login" ? (
            <form onSubmit={doLogin}>
              <div className="fg"><label className="fl">Email</label><input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="you@example.com" required /></div>
              <div className="fg"><label className="fl">Password</label><input type="password" value={pw} onChange={e=>setPw(e.target.value)} placeholder="••••••••" required /></div>
              <button type="submit" className="bp wf" style={{marginTop:6}} disabled={load}>{load?<span className="sp"/>:null} Sign in</button>
            </form>
          ) : (
            <form onSubmit={doSetPw}>
              <div className="al al-i" style={{fontSize:12}}>Your email must be added by an admin before you can set a password.</div>
              <div className="fg"><label className="fl">Email</label><input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="you@example.com" required /></div>
              <div className="fg"><label className="fl">New password</label><input type="password" value={np} onChange={e=>setNp(e.target.value)} placeholder="Min 6 characters" required /></div>
              <div className="fg"><label className="fl">Confirm</label><input type="password" value={cf} onChange={e=>setCf(e.target.value)} placeholder="Repeat password" required /></div>
              <button type="submit" className="bp wf" disabled={load}>{load?<span className="sp"/>:null} Set password & sign in</button>
            </form>
          )}
        </div>
      </div>
    </div>
  );
}

// ── ADMIN SHELL ──────────────────────────────────────────────
function AdminShell({ user, onLogout }) {
  const [page, setPage] = useState("dash");
  const [toast, showToast] = useToast();
  const nav = [{id:"dash",icon:"⬡",label:"Dashboard"},{id:"users",icon:"◈",label:"Users"},{id:"domains",icon:"◉",label:"Domains"},{id:"contests",icon:"◎",label:"Contests"},{id:"progress",icon:"◐",label:"Live Progress"}];
  return (
    <div style={{display:"flex"}}>
      <Toast t={toast} />
      <div className="sidebar">
        <div style={{padding:"18px 14px 10px"}}>
          <div style={{fontSize:14,fontWeight:700,letterSpacing:"-0.3px"}}>NW Curation</div>
          <div style={{fontSize:10,color:"var(--text3)",marginTop:2}}>Admin Panel</div>
        </div>
        <div className="div" style={{margin:"0 10px"}} />
        <div style={{flex:1,paddingTop:6}}>
          <div className="ns">Menu</div>
          {nav.map(n=>(
            <div key={n.id} className={`ni ${page===n.id?"act":""}`} onClick={()=>setPage(n.id)}>
              <span style={{fontSize:15}}>{n.icon}</span><span>{n.label}</span>
            </div>
          ))}
        </div>
        <div style={{padding:"10px 14px 18px"}}>
          <div style={{fontSize:11,color:"var(--text3)",marginBottom:7,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{user.email}</div>
          <button className="bg bsm wf" onClick={onLogout}>Sign out</button>
        </div>
      </div>
      <div className="mc fi">
        {page==="dash" && <AdminDash />}
        {page==="users" && <AdminUsers showToast={showToast} />}
        {page==="domains" && <AdminDomains showToast={showToast} />}
        {page==="contests" && <AdminContests showToast={showToast} />}
        {page==="progress" && <AdminProgress />}
      </div>
    </div>
  );
}

function AdminDash() {
  const [s, setS] = useState({u:0,d:0,c:0,a:0});
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
      <div className="ps">Platform overview</div>
      <div className="g4c" style={{marginBottom:24}}>
        {[{label:"Participants",v:s.u,col:"var(--accent)"},{label:"Domains",v:s.d,col:"var(--purple)"},{label:"Contests",v:s.c,col:"var(--green)"},{label:"Active now",v:s.a,col:"var(--amber)"}].map(x=>(
          <div className="sc" key={x.label}><div className="sv" style={{color:x.col}}>{x.v}</div><div className="sl">{x.label}</div></div>
        ))}
      </div>
      <div className="card">
        <div className="fw6" style={{marginBottom:10}}>Quick start</div>
        {[["1","Add users","Users → add emails or upload CSV"],["2","Upload domains","Domains → golden dataset + config xlsx"],["3","Create a contest","Contests → mode, tasks, domains, assign users"],["4","Monitor live","Live Progress → real-time scores"]].map(([n,t,d])=>(
          <div key={n} className="fx g3 ac" style={{padding:"9px 0",borderBottom:"1px solid var(--border)"}}>
            <div style={{width:24,height:24,borderRadius:"50%",background:"var(--ag)",color:"var(--accent)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,flexShrink:0}}>{n}</div>
            <div><div className="fw5 sm">{t}</div><div className="xs m3">{d}</div></div>
          </div>
        ))}
      </div>
    </div>
  );
}

function AdminUsers({ showToast }) {
  const [users, setUsers] = useState([]); const [load, setLoad] = useState(true);
  const [showAdd, setShowAdd] = useState(false); const [showReset, setShowReset] = useState(null);
  const [resetLink, setResetLink] = useState(""); const [newEmail, setNewEmail] = useState("");
  const [newName, setNewName] = useState(""); const [csvTxt, setCsvTxt] = useState("");
  const [addMode, setAddMode] = useState("single"); const [saving, setSaving] = useState(false);
  const [copied, setCopied] = useState(false); const fRef = useRef();

  async function load2() { setLoad(true); const {data} = await sb.from("users").select("*").eq("role","participant").order("created_at",{ascending:false}); setUsers(data||[]); setLoad(false); }
  useEffect(()=>{load2();},[]);

  async function addUser(email, name) {
    const {error} = await sb.from("users").upsert({email:email.toLowerCase().trim(),full_name:name,role:"participant",status:"active"},{onConflict:"email",ignoreDuplicates:false});
    return !error;
  }

  async function handleSingle(e) {
    e.preventDefault(); setSaving(true);
    const ok = await addUser(newEmail.trim(), newName.trim());
    if (ok) { showToast(`User ${newEmail} added`); setNewEmail(""); setNewName(""); setShowAdd(false); load2(); }
    else showToast("Failed — email may already exist","error");
    setSaving(false);
  }

  async function handleCSV(e) {
    e.preventDefault(); setSaving(true); let added=0;
    const lines = csvTxt.trim().split("\n").filter(l=>l.trim());
    for (const line of lines) { const [email,name=""] = line.split(",").map(s=>s.trim()); if (!email||!email.includes("@")) continue; if (await addUser(email,name)) added++; }
    showToast(`${added} user(s) added`); setCsvTxt(""); setShowAdd(false); load2(); setSaving(false);
  }

  async function toggle(u) {
    const ns = u.status==="active"?"disabled":"active";
    await sb.from("users").update({status:ns}).eq("id",u.id);
    showToast(`User ${ns}`); load2();
  }

  async function genReset(u) {
    const token = Array.from(crypto.getRandomValues(new Uint8Array(20))).map(b=>b.toString(16).padStart(2,"0")).join("");
    await sb.from("password_reset_tokens").insert({user_id:u.id,token,expires_at:new Date(Date.now()+864e5).toISOString()});
    setResetLink(`${location.origin}${location.pathname}?reset=${token}`);
    setShowReset(u);
  }

  return (
    <div>
      <div className="fx ac jb" style={{marginBottom:18}}>
        <div><div className="pt">Users</div><div className="ps">{users.length} participant(s)</div></div>
        <button className="bp" onClick={()=>setShowAdd(true)}>+ Add users</button>
      </div>
      <div className="card" style={{padding:0}}>
        {load ? <div style={{padding:28,textAlign:"center"}}><span className="sp"/></div> : (
          <table>
            <thead><tr>{["Name","Email","Status","Added","Actions"].map(c=><th key={c}>{c}</th>)}</tr></thead>
            <tbody>
              {users.length===0 && <tr><td colSpan={5} style={{textAlign:"center",color:"var(--text3)",padding:28}}>No users yet.</td></tr>}
              {users.map(u=>(
                <tr key={u.id}>
                  <td className="fw5">{u.full_name||"—"}</td>
                  <td className="mono xs">{u.email}</td>
                  <td><span className={`badge b-${u.status==="active"?"green":"red"}`}>{u.status}</span></td>
                  <td>{fmt(u.created_at)}</td>
                  <td><div className="fx g2">
                    <button className="bg bxs" onClick={()=>genReset(u)}>Reset pwd</button>
                    <button className={`bxs ${u.status==="active"?"bd":"bg"}`} onClick={()=>toggle(u)}>{u.status==="active"?"Disable":"Enable"}</button>
                  </div></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {showAdd && (
        <div className="mo" onClick={()=>setShowAdd(false)}>
          <div className="md" onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:17,fontWeight:600,marginBottom:18}}>Add users</div>
            <div className="tabs">
              <div className={`tab ${addMode==="single"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setAddMode("single")}>Single</div>
              <div className={`tab ${addMode==="csv"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setAddMode("csv")}>Bulk CSV</div>
            </div>
            {addMode==="single" ? (
              <form onSubmit={handleSingle}>
                <div className="fg"><label className="fl">Full name</label><input value={newName} onChange={e=>setNewName(e.target.value)} placeholder="Jane Doe"/></div>
                <div className="fg"><label className="fl">Email *</label><input type="email" value={newEmail} onChange={e=>setNewEmail(e.target.value)} placeholder="jane@co.com" required/></div>
                <div className="fx g3 jb" style={{marginTop:18}}>
                  <button type="button" className="bg" onClick={()=>setShowAdd(false)}>Cancel</button>
                  <button type="submit" className="bp" disabled={saving}>{saving?<span className="sp"/>:null} Add user</button>
                </div>
              </form>
            ) : (
              <form onSubmit={handleCSV}>
                <div className="al al-i" style={{fontSize:12}}>Format: email, full_name (one per line)</div>
                <div className="fg"><label className="fl">Paste CSV</label><textarea value={csvTxt} onChange={e=>setCsvTxt(e.target.value)} placeholder={"jane@co.com, Jane Doe\nbob@co.com, Bob Smith"} rows={5} style={{fontFamily:"var(--mono)",fontSize:12}}/></div>
                <div className="fx g3 ac" style={{marginBottom:14}}>
                  <button type="button" className="bg bsm" onClick={()=>fRef.current.click()}>Upload .csv</button>
                  <input ref={fRef} type="file" accept=".csv,.txt" style={{display:"none"}} onChange={ev=>{const r=new FileReader();r.onload=e=>setCsvTxt(e.target.result);r.readAsText(ev.target.files[0]);}}/>
                </div>
                <div className="fx g3 jb">
                  <button type="button" className="bg" onClick={()=>setShowAdd(false)}>Cancel</button>
                  <button type="submit" className="bp" disabled={saving||!csvTxt.trim()}>{saving?<span className="sp"/>:null} Add all</button>
                </div>
              </form>
            )}
          </div>
        </div>
      )}

      {showReset && (
        <div className="mo" onClick={()=>{setShowReset(null);setResetLink("");}}>
          <div className="md" onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:17,fontWeight:600,marginBottom:14}}>Password reset link</div>
            <div className="sm m2" style={{marginBottom:14}}>Share with <strong style={{color:"var(--text)"}}>{showReset.email}</strong>. Expires in 24 hours.</div>
            <div className="cb2" onClick={()=>{navigator.clipboard.writeText(resetLink);setCopied(true);setTimeout(()=>setCopied(false),2000);}}>{resetLink}</div>
            <div className="xs m3" style={{marginTop:7}}>{copied?"✓ Copied!":"Click to copy"}</div>
            <div style={{marginTop:18,textAlign:"right"}}><button className="bg" onClick={()=>{setShowReset(null);setResetLink("");}}>Close</button></div>
          </div>
        </div>
      )}
    </div>
  );
}

function AdminDomains({ showToast }) {
  const [domains, setDomains] = useState([]); const [load, setLoad] = useState(true);
  const [show, setShow] = useState(false); const [uploading, setUploading] = useState(false);
  const [name, setName] = useState(""); const [gFile, setGFile] = useState(null); const [cFile, setCFile] = useState(null);
  const [prog, setProg] = useState(""); const [curUser, setCurUser] = useState(null);

  useEffect(()=>{ sb.auth.getUser().then(({data})=>setCurUser(data?.user)); load2(); },[]);

  async function load2() { setLoad(true); const {data} = await sb.from("domains").select("*, domain_items(count), domain_fields(count)").order("created_at",{ascending:false}); setDomains(data||[]); setLoad(false); }

  async function doUpload(e) {
    e.preventDefault(); if (!gFile||!cFile||!name.trim()) return;
    setUploading(true);
    try {
      setProg("Parsing config...");
      const cBuf = await new Promise(r=>{const fr=new FileReader();fr.onload=ev=>r(ev.target.result);fr.readAsArrayBuffer(cFile);});
      const cWb = XLSX.read(cBuf,{type:"array"});
      const cRows = XLSX.utils.sheet_to_json(cWb.Sheets[cWb.SheetNames[0]]);

      setProg("Parsing golden dataset...");
      const gBuf = await new Promise(r=>{const fr=new FileReader();fr.onload=ev=>r(ev.target.result);fr.readAsArrayBuffer(gFile);});
      const gWb = XLSX.read(gBuf,{type:"array"});
      const gRows = XLSX.utils.sheet_to_json(gWb.Sheets[gWb.SheetNames[0]]);

      setProg("Creating domain...");
      const {data:dom,error:de} = await sb.from("domains").insert({name:name.trim(),created_by:curUser.id}).select().single();
      if (de) throw new Error(de.message);

      setProg("Saving field config...");
      const fields = cRows.filter(r=>r["Field Name"]&&r["Field Role"]).map(r=>({domain_id:dom.id,field_name:String(r["Field Name"]).trim(),field_role:String(r["Field Role"]).trim(),input_type:String(r["Input Type"]||"as_is").trim(),comparison_type:String(r["Comparison Type"]||"as_is").trim(),dropdown_values:r["Dropdown Values"]?String(r["Dropdown Values"]).trim():null,display_order:parseInt(r["Display Order"]||0)}));
      if (fields.length>0) { const {error:fe} = await sb.from("domain_fields").insert(fields); if (fe) throw new Error(fe.message); }

      const keys = Object.keys(gRows[0]||{});
      const ikc = keys.find(k=>k==="Item ID")||keys.find(k=>k.toLowerCase().includes("item")&&k.toLowerCase().includes("id"))||keys[0];
      const catc = keys.find(k=>k==="Category")||keys.find(k=>k==="Product Type")||keys.find(k=>k.toLowerCase().includes("categor")||k.toLowerCase().includes("type"))||keys[0];
      const atrc = keys.find(k=>k.toLowerCase().includes("all attributes"));

      const items = gRows.map((row,i)=>({domain_id:dom.id,item_key:String(row[ikc]||`item-${i}`).trim(),category:String(row[catc]||"Unknown").trim(),json_value:row,attributes_for_category:atrc?String(row[atrc]||"").trim():""}));

      for (let i=0;i<items.length;i+=50) {
        const {error:ie} = await sb.from("domain_items").insert(items.slice(i,i+50));
        if (ie) throw new Error(ie.message);
        setProg(`Saved ${Math.min(i+50,items.length)}/${items.length} items...`);
      }

      showToast(`Domain "${name}" created with ${items.length} items`);
      setShow(false); setName(""); setGFile(null); setCFile(null); setProg(""); load2();
    } catch(err) { showToast(err.message,"error"); }
    setUploading(false);
  }

  async function del(d) {
    if (!confirm(`Delete domain "${d.name}"?`)) return;
    await sb.from("domains").delete().eq("id",d.id);
    showToast("Deleted"); load2();
  }

  return (
    <div>
      <div className="fx ac jb" style={{marginBottom:18}}>
        <div><div className="pt">Domains</div><div className="ps">Upload golden datasets and config files</div></div>
        <button className="bp" onClick={()=>setShow(true)}>+ Add domain</button>
      </div>
      <div className="g2c">
        {load ? <div style={{padding:28}}><span className="sp"/></div> :
         domains.length===0 ? <div className="card" style={{gridColumn:"1/-1",textAlign:"center",color:"var(--text3)",padding:36}}>No domains yet.</div> :
         domains.map(d=>(
          <div className="card" key={d.id}>
            <div className="fx ac jb" style={{marginBottom:10}}><div className="fw6" style={{fontSize:15}}>{d.name}</div><button className="bd bxs" onClick={()=>del(d)}>Delete</button></div>
            <div className="fx g4 sm m2"><span>{d.domain_items?.[0]?.count||0} items</span><span>{d.domain_fields?.[0]?.count||0} fields</span></div>
            <div className="xs m3" style={{marginTop:7}}>{fmt(d.created_at)}</div>
          </div>
        ))}
      </div>

      {show && (
        <div className="mo" onClick={()=>!uploading&&setShow(false)}>
          <div className="md" onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:17,fontWeight:600,marginBottom:18}}>Add new domain</div>
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
              {prog && <div className="al al-i" style={{fontSize:12}}>{prog}</div>}
              <div className="fx g3 jb" style={{marginTop:18}}>
                <button type="button" className="bg" onClick={()=>setShow(false)} disabled={uploading}>Cancel</button>
                <button type="submit" className="bp" disabled={uploading}>{uploading?<span className="sp"/>:null} Upload & create</button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

function AdminContests({ showToast }) {
  const [contests,setContests] = useState([]); const [domains,setDomains] = useState([]);
  const [users,setUsers] = useState([]); const [load,setLoad] = useState(true);
  const [showCreate,setShowCreate] = useState(false); const [showAssign,setShowAssign] = useState(null);
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

  async function toggle(c) {
    const next = c.status==="active"?"closed":c.status==="draft"?"active":"active";
    await sb.from("contests").update({status:next}).eq("id",c.id);
    showToast(`Contest ${next}`); load2();
  }

  async function del(c) {
    if (!confirm(`Delete "${c.name}"?`)) return;
    await sb.from("contests").delete().eq("id",c.id);
    showToast("Deleted"); load2();
  }

  const sc = {draft:"gray",active:"green",closed:"red"};

  return (
    <div>
      <div className="fx ac jb" style={{marginBottom:18}}>
        <div><div className="pt">Contests</div><div className="ps">{contests.length} total</div></div>
        <button className="bp" onClick={()=>setShowCreate(true)}>+ Create contest</button>
      </div>
      <div className="card" style={{padding:0}}>
        {load ? <div style={{padding:28,textAlign:"center"}}><span className="sp"/></div> : (
          <table>
            <thead><tr>{["Name","Mode","Domains","Tasks","Users","Status","Actions"].map(c=><th key={c}>{c}</th>)}</tr></thead>
            <tbody>
              {contests.length===0 && <tr><td colSpan={7} style={{textAlign:"center",color:"var(--text3)",padding:28}}>No contests yet.</td></tr>}
              {contests.map(c=>(
                <tr key={c.id}>
                  <td className="fw5">{c.name}</td>
                  <td><span className={`badge b-${c.mode==="practice"?"purple":"blue"}`}>{c.mode}</span></td>
                  <td className="xs">{c.contest_domains?.map(cd=>cd.domains?.name).join(", ")||"—"}</td>
                  <td>{c.task_count}</td>
                  <td>{c.contest_users?.[0]?.count||0}</td>
                  <td><span className={`badge b-${sc[c.status]}`}>{c.status}</span></td>
                  <td><div className="fx g2">
                    <button className="bg bxs" onClick={()=>setShowAssign(c)}>Users</button>
                    <button className="bg bxs" onClick={()=>toggle(c)}>{c.status==="active"?"Close":c.status==="draft"?"Activate":"Reopen"}</button>
                    <button className="bd bxs" onClick={()=>del(c)}>Del</button>
                  </div></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
      {showCreate && <CreateContestModal domains={domains} users={users} uid={curUser?.id} onClose={()=>setShowCreate(false)} onDone={()=>{setShowCreate(false);load2();showToast("Contest created!");}}/>}
      {showAssign && <AssignModal contest={showAssign} users={users} onClose={()=>setShowAssign(null)} onSaved={()=>{setShowAssign(null);load2();showToast("Users updated");}}/>}
    </div>
  );
}

function CreateContestModal({ domains, users, uid, onClose, onDone }) {
  const [name,setName]=useState(""); const [mode,setMode]=useState("assessment");
  const [tc,setTc]=useState(10); const [sa,setSa]=useState(""); const [ea,setEa]=useState("");
  const [l0,setL0]=useState(70); const [l1,setL1]=useState(80); const [l2,setL2]=useState(90); const [st,setSt]=useState(0.7);
  const [allocs,setAllocs]=useState([{did:"",tc:10}]);
  const [selU,setSelU]=useState([]); const [csvU,setCsvU]=useState(""); const [amode,setAmode]=useState("sel");
  const [saving,setSaving]=useState(false); const [err,setErr]=useState("");
  const fRef=useRef();

  async function doCreate(e) {
    e.preventDefault(); setErr("");
    const tot = allocs.reduce((s,a)=>s+parseInt(a.tc||0),0);
    if (tot!==parseInt(tc)) { setErr(`Domain tasks must sum to ${tc}. Currently ${tot}.`); return; }
    if (allocs.some(a=>!a.did)) { setErr("Select a domain for each row"); return; }
    setSaving(true);
    try {
      const {data:con,error:ce} = await sb.from("contests").insert({name:name.trim(),mode,task_count:parseInt(tc),start_at:sa||null,end_at:ea||null,l0_threshold:l0,l1_threshold:l1,l2_threshold:l2,semantic_correct_threshold:st,status:"draft",created_by:uid}).select().single();
      if (ce) throw new Error(ce.message);

      let off=0;
      for (const a of allocs) {
        await sb.from("contest_domains").insert({contest_id:con.id,domain_id:a.did,task_count:parseInt(a.tc)});
        const {data:items} = await sb.from("domain_items").select("id").eq("domain_id",a.did);
        if (!items||!items.length) continue;
        const sh = items.sort(()=>Math.random()-.5).slice(0,parseInt(a.tc));
        await sb.from("contest_items").insert(sh.map((it,i)=>({contest_id:con.id,domain_item_id:it.id,item_order:off+i+1})));
        off+=parseInt(a.tc);
      }

      let uids = [...selU];
      if (amode==="csv"&&csvU.trim()) { const em=csvU.trim().split("\n").map(l=>l.split(",")[0].trim()).filter(Boolean); const {data:fu}=await sb.from("users").select("id").in("email",em); uids=[...new Set([...uids,...(fu||[]).map(u=>u.id)])]; }
      if (uids.length>0) await sb.from("contest_users").insert(uids.map(uid2=>({contest_id:con.id,user_id:uid2})));
      onDone();
    } catch(er) { setErr(er.message); }
    setSaving(false);
  }

  return (
    <div className="mo" onClick={onClose}>
      <div className="md md-lg" onClick={e=>e.stopPropagation()} style={{maxHeight:"92vh",overflowY:"auto"}}>
        <div style={{fontSize:17,fontWeight:600,marginBottom:18}}>Create contest</div>
        <form onSubmit={doCreate}>
          {err && <div className="al al-e">{err}</div>}
          <div className="row2">
            <div className="fg"><label className="fl">Name *</label><input value={name} onChange={e=>setName(e.target.value)} placeholder="Apparel Assessment — June 2025" required/></div>
            <div className="fg"><label className="fl">Mode *</label><select value={mode} onChange={e=>setMode(e.target.value)}><option value="assessment">Assessment</option><option value="practice">Practice</option></select></div>
          </div>
          <div className="row3">
            <div className="fg"><label className="fl">Total tasks *</label><input type="number" min={1} max={100} value={tc} onChange={e=>setTc(e.target.value)} required/></div>
            <div className="fg"><label className="fl">Start time</label><input type="datetime-local" value={sa} onChange={e=>setSa(e.target.value)}/></div>
            <div className="fg"><label className="fl">End time</label><input type="datetime-local" value={ea} onChange={e=>setEa(e.target.value)}/></div>
          </div>

          <div style={{marginBottom:14}}>
            <div className="fx ac jb" style={{marginBottom:7}}>
              <label className="fl" style={{margin:0}}>Domain allocation *</label>
              <button type="button" className="bg bxs" onClick={()=>setAllocs(p=>[...p,{did:"",tc:5}])}>+ Domain</button>
            </div>
            {allocs.map((a,i)=>(
              <div key={i} className="fx g3 ac" style={{marginBottom:7}}>
                <select style={{flex:2}} value={a.did} onChange={e=>setAllocs(p=>p.map((x,idx)=>idx===i?{...x,did:e.target.value}:x))}>
                  <option value="">Select domain</option>
                  {domains.map(d=><option key={d.id} value={d.id}>{d.name}</option>)}
                </select>
                <input type="number" min={1} style={{flex:1}} value={a.tc} onChange={e=>setAllocs(p=>p.map((x,idx)=>idx===i?{...x,tc:e.target.value}:x))} placeholder="Tasks"/>
                {allocs.length>1 && <button type="button" className="bd bxs" onClick={()=>setAllocs(p=>p.filter((_,idx)=>idx!==i))}>✕</button>}
              </div>
            ))}
            <div className="xs m3">Sum must equal {tc}</div>
          </div>

          {mode==="assessment" && (
            <div style={{background:"var(--bg3)",borderRadius:"var(--r)",padding:12,marginBottom:14}}>
              <div className="sm fw5" style={{marginBottom:9}}>Certification thresholds</div>
              <div className="row3" style={{marginBottom:8}}>
                {[["L1 min %",l0,setL0],["L2 min %",l1,setL1],["L3 min %",l2,setL2]].map(([lbl,val,setter])=>(
                  <div key={lbl}><label className="fl">{lbl}</label><input type="number" min={0} max={100} value={val} onChange={e=>setter(Number(e.target.value))}/></div>
                ))}
              </div>
              <div><label className="fl">Semantic threshold (0–1)</label><input type="number" min={0} max={1} step={0.05} value={st} onChange={e=>setSt(Number(e.target.value))}/><div className="xs m3" style={{marginTop:3}}>Score ≥ this = correct. Default 0.7</div></div>
            </div>
          )}

          <div style={{marginBottom:14}}>
            <label className="fl">Assign users</label>
            <div className="tabs" style={{marginBottom:9}}>
              <div className={`tab ${amode==="sel"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setAmode("sel")}>Select</div>
              <div className={`tab ${amode==="csv"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setAmode("csv")}>By email</div>
            </div>
            {amode==="sel" ? (
              <div style={{maxHeight:160,overflowY:"auto",background:"var(--bg3)",borderRadius:"var(--r)",padding:7}}>
                {users.map(u=>(
                  <label key={u.id} className="fx g3 ac" style={{padding:"5px 7px",cursor:"pointer"}}>
                    <input type="checkbox" checked={selU.includes(u.id)} onChange={ev=>setSelU(p=>ev.target.checked?[...p,u.id]:p.filter(id=>id!==u.id))}/>
                    <span className="sm">{u.full_name||u.email}</span>
                    <span className="xs m3">{u.full_name?u.email:""}</span>
                  </label>
                ))}
                {users.length===0 && <div className="xs m3" style={{padding:7}}>No participants yet.</div>}
              </div>
            ) : (
              <div>
                <textarea value={csvU} onChange={e=>setCsvU(e.target.value)} placeholder={"jane@co.com\nbob@co.com"} rows={4} style={{fontFamily:"var(--mono)",fontSize:12}}/>
                <input ref={fRef} type="file" accept=".csv,.txt" style={{display:"none"}} onChange={ev=>{const r=new FileReader();r.onload=e=>setCsvU(e.target.result);r.readAsText(ev.target.files[0]);}}/>
                <button type="button" className="bg bxs" style={{marginTop:5}} onClick={()=>fRef.current.click()}>Upload CSV</button>
              </div>
            )}
          </div>

          <div className="fx g3 jb" style={{marginTop:18}}>
            <button type="button" className="bg" onClick={onClose}>Cancel</button>
            <button type="submit" className="bp" disabled={saving}>{saving?<span className="sp"/>:null} Create contest</button>
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
  useEffect(()=>{ sb.from("contest_users").select("user_id").eq("contest_id",contest.id).then(({data})=>{ const ids=(data||[]).map(r=>r.user_id); setAssigned(ids); setSel(ids); }); },[]);
  async function save() {
    setSaving(true); let toAdd=[...sel];
    if (mode==="csv"&&csv.trim()) { const em=csv.trim().split("\n").map(l=>l.split(",")[0].trim()).filter(Boolean); const {data}=await sb.from("users").select("id").in("email",em); toAdd=[...new Set([...toAdd,...(data||[]).map(u=>u.id)])]; }
    const newIds = toAdd.filter(id=>!assigned.includes(id));
    if (newIds.length>0) await sb.from("contest_users").insert(newIds.map(uid=>({contest_id:contest.id,user_id:uid})));
    const rem = assigned.filter(id=>!toAdd.includes(id));
    if (rem.length>0) await sb.from("contest_users").delete().eq("contest_id",contest.id).in("user_id",rem);
    onSaved(); setSaving(false);
  }
  return (
    <div className="mo" onClick={onClose}>
      <div className="md" onClick={e=>e.stopPropagation()}>
        <div style={{fontSize:17,fontWeight:600,marginBottom:16}}>Assign users — {contest.name}</div>
        <div className="tabs">
          <div className={`tab ${mode==="sel"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setMode("sel")}>Select</div>
          <div className={`tab ${mode==="csv"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setMode("csv")}>By email</div>
        </div>
        {mode==="sel" ? (
          <div style={{maxHeight:260,overflowY:"auto",background:"var(--bg3)",borderRadius:"var(--r)",padding:7,marginBottom:14}}>
            {users.map(u=>(
              <label key={u.id} className="fx g3 ac" style={{padding:"6px 7px",cursor:"pointer"}}>
                <input type="checkbox" checked={sel.includes(u.id)} onChange={ev=>setSel(p=>ev.target.checked?[...p,u.id]:p.filter(id=>id!==u.id))}/>
                <span className="sm">{u.full_name||u.email}</span>
                <span className="xs m3">{u.full_name?u.email:""}</span>
              </label>
            ))}
          </div>
        ) : (
          <div style={{marginBottom:14}}>
            <textarea value={csv} onChange={e=>setCsv(e.target.value)} placeholder={"jane@co.com\nbob@co.com"} rows={5} style={{fontFamily:"var(--mono)",fontSize:12}}/>
            <input ref={fRef} type="file" accept=".csv,.txt" style={{display:"none"}} onChange={ev=>{const r=new FileReader();r.onload=e=>setCsv(e.target.result);r.readAsText(ev.target.files[0]);}}/>
            <button type="button" className="bg bxs" style={{marginTop:5}} onClick={()=>fRef.current.click()}>Upload CSV</button>
          </div>
        )}
        <div className="sm m2" style={{marginBottom:14}}>{sel.length} selected</div>
        <div className="fx g3 jb">
          <button className="bg" onClick={onClose}>Cancel</button>
          <button className="bp" onClick={save} disabled={saving}>{saving?<span className="sp"/>:null} Save</button>
        </div>
      </div>
    </div>
  );
}

function AdminProgress() {
  const [contests,setContests]=useState([]); const [sel,setSel]=useState("");
  const [progress,setProgress]=useState([]); const [fa,setFa]=useState([]);
  const [load,setLoad]=useState(false); const [tab,setTab]=useState("users");

  useEffect(()=>{ sb.from("contests").select("id,name,mode,status").order("created_at",{ascending:false}).then(({data})=>setContests(data||[])); },[]);
  useEffect(()=>{
    if (!sel) return;
    loadP();
    const sub = sb.channel("ap").on("postgres_changes",{event:"*",schema:"public",table:"responses"},loadP).subscribe();
    return ()=>sb.removeChannel(sub);
  },[sel]);

  async function loadP() {
    if (!sel) return; setLoad(true);
    const con = contests.find(c=>c.id===sel);
    if (con?.mode==="assessment") {
      const [{data:p},{data:f}] = await Promise.all([
        sb.from("v_user_contest_accuracy").select("*, users(email,full_name)").eq("contest_id",sel),
        sb.from("v_field_accuracy").select("*").eq("contest_id",sel).order("field_accuracy_pct"),
      ]);
      setProgress(p||[]); setFa(f||[]);
    } else {
      const {data:p} = await sb.from("v_practice_progress").select("*, users(email,full_name)").eq("contest_id",sel);
      setProgress(p||[]); setFa([]);
    }
    setLoad(false);
  }

  const con = contests.find(c=>c.id===sel);

  return (
    <div>
      <div className="pt">Live Progress</div>
      <div className="ps">Real-time user performance</div>
      <div className="fg" style={{maxWidth:380,marginBottom:22}}>
        <label className="fl">Select contest</label>
        <select value={sel} onChange={e=>{setSel(e.target.value);setProgress([]);setFa([]);}}>
          <option value="">Choose a contest...</option>
          {contests.map(c=><option key={c.id} value={c.id}>[{c.mode}] {c.name} — {c.status}</option>)}
        </select>
      </div>
      {!sel && <div className="card" style={{textAlign:"center",color:"var(--text3)",padding:36}}>Select a contest to see progress</div>}
      {sel && (
        <>
          {con?.mode==="assessment" && (
            <div className="tabs" style={{maxWidth:280}}>
              <div className={`tab ${tab==="users"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setTab("users")}>Per user</div>
              <div className={`tab ${tab==="fields"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setTab("fields")}>Field heatmap</div>
            </div>
          )}
          {load ? <div style={{textAlign:"center",padding:36}}><span className="sp"/></div> : (
            <>
              {(tab==="users"||con?.mode==="practice") && (
                <div className="card" style={{padding:0}}>
                  <table>
                    <thead><tr>
                      <th>User</th><th>Tasks done</th>
                      {con?.mode==="assessment" && <><th>Total</th><th>Correct</th><th>Accuracy</th><th>Cert</th></>}
                    </tr></thead>
                    <tbody>
                      {progress.length===0 && <tr><td colSpan={7} style={{textAlign:"center",color:"var(--text3)",padding:22}}>No submissions yet</td></tr>}
                      {progress.map((p,i)=>(
                        <tr key={i}>
                          <td className="fw5">{p.users?.full_name||p.users?.email||"—"}</td>
                          <td>{p.tasks_submitted||p.tasks_completed||0}</td>
                          {con?.mode==="assessment" && <>
                            <td>{p.total_attributes}</td>
                            <td>{p.correct_attributes}</td>
                            <td><div className="fx ac g2"><div className="pb" style={{width:72}}><div className="pf" style={{width:`${p.accuracy_pct||0}%`,background:p.accuracy_pct>=90?"var(--green)":p.accuracy_pct>=70?"var(--amber)":"var(--red)"}}/></div><span className="mono xs">{p.accuracy_pct||0}%</span></div></td>
                            <td><span className={`badge cert-${p.cert_level?.toLowerCase()}`}>{p.cert_level}</span></td>
                          </>}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
              {tab==="fields" && con?.mode==="assessment" && (
                <div className="card">
                  <div className="fw5" style={{marginBottom:14}}>Field accuracy — lowest first</div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:7}}>
                    {fa.map(f=>{const pct=parseFloat(f.field_accuracy_pct)||0;const bg=pct>=80?"var(--gbg)":pct>=60?"var(--abg)":"var(--rbg)";const col=pct>=80?"var(--green)":pct>=60?"var(--amber)":"var(--red)";
                      return <div key={f.field_name} className="hc" style={{background:bg,color:col}}><div style={{fontSize:10,marginBottom:2}}>{f.field_name}</div><div style={{fontSize:15,fontWeight:700}}>{pct}%</div><div style={{fontSize:9,opacity:.7}}>{f.correct_responses}/{f.total_responses}</div></div>;
                    })}
                    {fa.length===0 && <div className="sm m2">No data yet</div>}
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

function UserShell({ user, onLogout }) {
  const [page,setPage] = useState("contests"); const [ac,setAc] = useState(null);
  const [toast,showToast] = useToast();
  return (
    <div style={{minHeight:"100vh"}}>
      <Toast t={toast}/>
      {page!=="task" && (
        <div style={{background:"var(--bg2)",borderBottom:"1px solid var(--border)",padding:"13px 26px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div style={{fontSize:15,fontWeight:700,letterSpacing:"-0.3px"}}>NW Curation</div>
          <div className="fx g3 ac">
            <span className="sm m2">{user.full_name||user.email}</span>
            <button className="bg bsm" onClick={()=>setPage(page==="profile"?"contests":"profile")}>Profile</button>
            <button className="bg bsm" onClick={onLogout}>Sign out</button>
          </div>
        </div>
      )}
      <div style={{padding:page==="task"?0:26}} className="fi">
        {page==="contests" && <UserContests user={user} onOpen={c=>{setAc(c);setPage("task");}}/>}
        {page==="profile" && <UserProfile user={user} showToast={showToast}/>}
        {page==="task" && ac && <ContestTaskView contest={ac} user={user} onClose={()=>{setAc(null);setPage("contests");}} showToast={showToast}/>}
      </div>
    </div>
  );
}

function UserContests({ user, onOpen }) {
  const [live,setLive]=useState([]); const [past,setPast]=useState([]); const [load,setLoad]=useState(true); const [tab,setTab]=useState("live");
  useEffect(()=>{load2();},[]);
  async function load2() {
    setLoad(true);
    const {data:cu} = await sb.from("contest_users").select("contest_id").eq("user_id",user.id);
    if (!cu||!cu.length) { setLoad(false); return; }
    const ids = cu.map(r=>r.contest_id);
    const [{data:all},{data:tasks},{data:scores}] = await Promise.all([
      sb.from("contests").select("*, contest_domains(domain_id,task_count,domains(name))").in("id",ids).order("created_at",{ascending:false}),
      sb.from("tasks").select("contest_id,status").eq("user_id",user.id),
      sb.from("v_user_contest_accuracy").select("*").eq("user_id",user.id),
    ]);
    const tm={};(tasks||[]).forEach(t=>{if(!tm[t.contest_id])tm[t.contest_id]={total:0,submitted:0};tm[t.contest_id].total++;if(t.status==="submitted")tm[t.contest_id].submitted++;});
    const sm={};(scores||[]).forEach(s=>{sm[s.contest_id]=s;});
    setLive((all||[]).filter(c=>c.status!=="closed").map(c=>({...c,progress:tm[c.id]||null})));
    setPast((all||[]).filter(c=>c.status==="closed").map(c=>({...c,progress:tm[c.id]||null,score:sm[c.id]||null})));
    setLoad(false);
  }
  if (load) return <div style={{textAlign:"center",padding:56}}><span className="sp"/></div>;
  return (
    <div style={{maxWidth:860,margin:"0 auto"}}>
      <div style={{marginBottom:22}}><div className="xl fw6">My contests</div><div className="sm m2" style={{marginTop:3}}>Welcome back, {user.full_name||user.email.split("@")[0]}</div></div>
      <div className="tabs" style={{maxWidth:280}}>
        <div className={`tab ${tab==="live"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setTab("live")}>Live ({live.length})</div>
        <div className={`tab ${tab==="past"?"act":""}`} style={{flex:1,textAlign:"center"}} onClick={()=>setTab("past")}>Past ({past.length})</div>
      </div>
      {tab==="live" && (
        <div>
          {live.length===0 && <div className="card" style={{textAlign:"center",color:"var(--text3)",padding:36}}>No live contests assigned to you.</div>}
          {["practice","assessment"].map(m=>{ const fl=live.filter(c=>c.mode===m); if(!fl.length) return null; return (
            <div key={m}><div className="xs m3" style={{marginBottom:8,textTransform:"uppercase",letterSpacing:".06em",marginTop:8}}>{m}</div><div className="g2c">{fl.map(c=><ContestCard key={c.id} contest={c} onOpen={()=>onOpen(c)}/>)}</div></div>
          );})}
        </div>
      )}
      {tab==="past" && (
        <div>
          {past.length===0 && <div className="card" style={{textAlign:"center",color:"var(--text3)",padding:36}}>No past contests yet.</div>}
          {past.map(c=>(
            <div className="card" key={c.id} style={{marginBottom:10}}>
              <div className="fx ac jb">
                <div><div className="fw5">{c.name}</div><div className="xs m3" style={{marginTop:2}}>{c.contest_domains?.map(cd=>cd.domains?.name).join(", ")} · {c.task_count} tasks · {c.mode}</div></div>
                <div className="fx g3 ac">
                  {c.mode==="assessment"&&c.score
                    ?<><div style={{textAlign:"right"}}><div className="mono fw6" style={{fontSize:17}}>{c.score.accuracy_pct}%</div><div className="xs m3">accuracy</div></div><span className={`badge cert-${c.score.cert_level?.toLowerCase()}`} style={{fontSize:13,padding:"5px 12px"}}>{c.score.cert_level}</span></>
                    :<div className="sm m2">{c.progress?.submitted||0}/{c.task_count} completed</div>
                  }
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
  const p = contest.progress; const pct = p&&contest.task_count?Math.round((p.submitted/contest.task_count)*100):0; const done = p?.submitted===contest.task_count;
  return (
    <div className="card" style={{cursor:"pointer",transition:"border-color .15s"}} onClick={onOpen} onMouseEnter={e=>e.currentTarget.style.borderColor="var(--border2)"} onMouseLeave={e=>e.currentTarget.style.borderColor="var(--border)"}>
      <div className="fx ac jb" style={{marginBottom:9}}><span className={`badge b-${contest.mode==="practice"?"purple":"blue"}`}>{contest.mode}</span>{done&&<span className="badge b-green">✓ Complete</span>}</div>
      <div className="fw6" style={{marginBottom:3}}>{contest.name}</div>
      <div className="xs m3" style={{marginBottom:10}}>{contest.contest_domains?.map(cd=>cd.domains?.name).join(", ")} · {contest.task_count} tasks</div>
      {contest.end_at && <div className="xs m2" style={{marginBottom:9}}>Ends: {fmt(contest.end_at)}</div>}
      <div className="pb" style={{marginBottom:5}}><div className="pf" style={{width:`${pct}%`}}/></div>
      <div className="xs m3">{p?.submitted||0}/{contest.task_count} submitted</div>
    </div>
  );
}

function UserProfile({ user, showToast }) {
  const [np,setNp]=useState(""); const [cf,setCf]=useState(""); const [fn,setFn]=useState(user.full_name||""); const [saving,setSaving]=useState(false);
  async function save(e) {
    e.preventDefault(); setSaving(true);
    if (fn!==user.full_name) await sb.from("users").update({full_name:fn}).eq("id",user.id);
    if (np) { if(np!==cf){showToast("Passwords do not match","error");setSaving(false);return;} const{error}=await sb.auth.updateUser({password:np}); if(error){showToast(error.message,"error");setSaving(false);return;} }
    showToast("Profile updated"); setSaving(false);
  }
  return (
    <div style={{maxWidth:460,margin:"0 auto"}}>
      <div className="xl fw6" style={{marginBottom:22}}>Profile</div>
      <div className="card"><form onSubmit={save}>
        <div className="fg"><label className="fl">Full name</label><input value={fn} onChange={e=>setFn(e.target.value)}/></div>
        <div className="fg"><label className="fl">Email</label><input value={user.email} disabled/></div>
        <div className="div"/>
        <div className="sm fw5" style={{marginBottom:10}}>Change password</div>
        <div className="fg"><label className="fl">New password</label><input type="password" value={np} onChange={e=>setNp(e.target.value)} placeholder="Leave blank to keep current"/></div>
        <div className="fg"><label className="fl">Confirm</label><input type="password" value={cf} onChange={e=>setCf(e.target.value)}/></div>
        <button type="submit" className="bp" disabled={saving}>{saving?<span className="sp"/>:null} Save changes</button>
      </form></div>
    </div>
  );
}

function ContestTaskView({ contest, user, onClose, showToast }) {
  const [items,setItems]=useState([]); const [tasks,setTasks]=useState({}); const [fields,setFields]=useState({});
  const [idx,setIdx]=useState(0); const [answers,setAnswers]=useState({});
  const [loading,setLoading]=useState(true); const [submitting,setSubmitting]=useState(false);
  const [showVal,setShowVal]=useState(false); const [submitted,setSubmitted]=useState(false); const [score,setScore]=useState(null);

  useEffect(()=>{ loadC(); },[]);

  async function loadC() {
    await sb.rpc("close_expired_contests").catch(()=>{});
    const {data:ci} = await sb.from("contest_items").select("*, domain_items(*, domains(id,name))").eq("contest_id",contest.id).order("item_order");
    setItems(ci||[]);
    const dids = [...new Set((ci||[]).map(i=>i.domain_items?.domains?.id).filter(Boolean))];
    const fm={};
    for (const did of dids) { const {data:df}=await sb.from("domain_fields").select("*").eq("domain_id",did).order("display_order"); fm[did]=df||[]; }
    setFields(fm);
    const {data:et} = await sb.from("tasks").select("*, responses(*)").eq("contest_id",contest.id).eq("user_id",user.id);
    const tm={}; const am={};
    (et||[]).forEach(t=>{ tm[t.contest_item_id]=t; am[t.id]={}; (t.responses||[]).forEach(r=>{am[t.id][r.field_name]=r.user_value;}); });
    setTasks(tm); setAnswers(am);
    const allSub = (ci||[]).every(item=>tm[item.id]?.status==="submitted");
    if (allSub&&contest.mode==="assessment") { setSubmitted(true); const {data:s}=await sb.from("v_user_contest_accuracy").select("*").eq("contest_id",contest.id).eq("user_id",user.id).single(); setScore(s); }
    setLoading(false);
  }

  async function ensureTask(item) {
    if (tasks[item.id]) return tasks[item.id];
    const {data:t} = await sb.from("tasks").insert({contest_id:contest.id,user_id:user.id,contest_item_id:item.id,status:"in_progress",started_at:new Date().toISOString()}).select().single();
    if (t) setTasks(prev=>({...prev,[item.id]:t}));
    return t;
  }

  async function saveAns(item, fn2, val) {
    const task = await ensureTask(item); if (!task||task.status==="submitted") return;
    if (contest.mode==="practice") { setAnswers(prev=>({...prev,[task.id]:{...(prev[task.id]||{}),[fn2]:val}})); return; }
    const golden = String(item.domain_items?.json_value?.[fn2]||"");
    const did = item.domain_items?.domains?.id;
    const fd = (fields[did]||[]).find(f=>f.field_name===fn2);
    const ct = fd?.comparison_type||"as_is";
    const sc = scoreF(val,golden,ct,contest.semantic_correct_threshold||0.7);
    await sb.from("responses").upsert({task_id:task.id,field_name:fn2,user_value:val,golden_value:golden,score:sc,comparison_type:ct,is_draft:true},{onConflict:"task_id,field_name"});
    setAnswers(prev=>({...prev,[task.id]:{...(prev[task.id]||{}),[fn2]:val}}));
  }

  async function submitPractice(item) {
    let task = tasks[item.id];
    if (!task) { const {data:t}=await sb.from("tasks").insert({contest_id:contest.id,user_id:user.id,contest_item_id:item.id,status:"submitted",started_at:new Date().toISOString(),submitted_at:new Date().toISOString()}).select().single(); task=t; if(task) setTasks(prev=>({...prev,[item.id]:task})); }
    else { await sb.from("tasks").update({status:"submitted",submitted_at:new Date().toISOString()}).eq("id",task.id); setTasks(prev=>({...prev,[item.id]:{...prev[item.id],status:"submitted"}})); }
    setShowVal(true);
  }

  async function submitAll() {
    const inc = items.filter(i=>!tasks[i.id]||tasks[i.id].status!=="submitted");
    if (inc.length>0&&!confirm(`${inc.length} task(s) empty. Submit all anyway?`)) return;
    setSubmitting(true);
    for (const item of items) {
      const task = tasks[item.id];
      if (!task) { await sb.from("tasks").insert({contest_id:contest.id,user_id:user.id,contest_item_id:item.id,status:"submitted",started_at:new Date().toISOString(),submitted_at:new Date().toISOString()}); }
      else if (task.status==="in_progress") { await sb.from("responses").update({is_draft:false}).eq("task_id",task.id); await sb.from("tasks").update({status:"submitted",submitted_at:new Date().toISOString()}).eq("id",task.id); }
    }
    const {data:s} = await sb.from("v_user_contest_accuracy").select("*").eq("contest_id",contest.id).eq("user_id",user.id).single();
    setScore(s); setSubmitted(true); setSubmitting(false);
  }

  if (loading) return <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh"}}><span className="sp"/></div>;

  if (submitted&&contest.mode==="assessment") return (
    <div style={{maxWidth:580,margin:"56px auto",padding:"0 18px"}}>
      <div className="card" style={{textAlign:"center",padding:36}}>
        <div style={{fontSize:42,marginBottom:14}}>✓</div>
        <div className="xl fw6" style={{marginBottom:7}}>Assessment submitted</div>
        <div className="sm m2" style={{marginBottom:24}}>{contest.name}</div>
        {score ? (
          <>
            <div className="g3c" style={{marginBottom:22}}>
              {[{v:score.tasks_submitted,l:"Tasks done"},{v:`${score.correct_attributes}/${score.total_attributes}`,l:"Correct attrs"},{v:`${score.accuracy_pct}%`,l:"Accuracy"}].map(x=><div className="sc" key={x.l}><div className="sv">{x.v}</div><div className="sl">{x.l}</div></div>)}
            </div>
            <div className={`badge cert-${score.cert_level?.toLowerCase()}`} style={{fontSize:20,padding:"10px 24px",borderRadius:10}}>{score.cert_level}</div>
            <div className="xs m3" style={{marginTop:8}}>Based on {score.accuracy_pct}% accuracy</div>
          </>
        ) : <span className="sp"/>}
        <button className="bg" style={{marginTop:24}} onClick={onClose}>Back to contests</button>
      </div>
    </div>
  );

  const cur = items[idx]; if (!cur) return <div style={{padding:36,textAlign:"center",color:"var(--text3)"}}>No items found.</div>;
  const di = cur.domain_items; const did = di?.domains?.id; const af = fields[did]||[];
  const aa = (di?.attributes_for_category||"").split(",").map(s=>s.trim()).filter(Boolean);
  const task = tasks[cur.id]; const ta = answers[task?.id]||{}; const isSub = task?.status==="submitted";
  const ctx = af.filter(f=>f.field_role==="context"&&(aa.length===0||aa.includes(f.field_name)));
  const imgs = af.filter(f=>f.field_role==="image");
  const cure = af.filter(f=>f.field_role==="curate"&&(aa.length===0||aa.includes(f.field_name)));
  const tSt = item => { const t=tasks[item.id]; if(!t||t.status==="not_started")return"ns"; if(t.status==="in_progress")return"ip"; return"sb"; };
  const inc = items.filter(i=>!tasks[i.id]||tasks[i.id].status!=="submitted").length;

  return (
    <div style={{display:"flex",height:"100vh",overflow:"hidden"}}>
      <div className="tsb">
        <div style={{padding:"14px 6px 9px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <span className="xs m2">Tasks</span>
          <button className="bg bxs" onClick={onClose} style={{fontSize:10}}>✕ Exit</button>
        </div>
        <div style={{display:"flex",flexWrap:"wrap",gap:5,padding:"0 3px"}}>
          {items.map((item,i)=>{ const st=tSt(item); return (
            <div key={item.id} className={`td td-${st} ${i===idx?"td-act":""}`} onClick={()=>{setIdx(i);setShowVal(false);}} title={`Task ${i+1}`}>{i+1}</div>
          );})}
        </div>
        <div style={{marginTop:14,padding:"0 6px"}}>
          {[["ns","Not started"],["ip","In progress"],["sb","Submitted"]].map(([st,lbl])=>(
            <div key={st} className="fx g2 ac xs m3" style={{marginBottom:3}}><div className={`td td-${st}`} style={{width:10,height:10}}/>{lbl}</div>
          ))}
        </div>
        {contest.mode==="assessment" && (
          <div style={{padding:"10px 6px",marginTop:"auto"}}>
            <button className="bp wf" style={{fontSize:11}} onClick={submitAll} disabled={submitting}>
              {submitting?<span className="sp"/>:`Submit all${inc>0?` (${inc} empty)`:""}`}
            </button>
          </div>
        )}
      </div>

      <div style={{flex:1,overflowY:"auto",padding:"18px 22px"}}>
        <div className="fx ac jb" style={{marginBottom:14}}>
          <div>
            <div className="xs m3" style={{marginBottom:2}}>{contest.name} · {contest.mode}</div>
            <div className="fw6" style={{fontSize:15}}>Task {idx+1} of {items.length} — {di?.category||""}</div>
          </div>
          <div className="fx g2">
            <button className="bg bsm" disabled={idx===0} onClick={()=>{setIdx(i=>i-1);setShowVal(false);}}>← Prev</button>
            <button className="bg bsm" disabled={idx===items.length-1} onClick={()=>{setIdx(i=>i+1);setShowVal(false);}}>Next →</button>
          </div>
        </div>

        <div className="card" style={{marginBottom:14,padding:14}}>
          <div className="fx g3 wrap" style={{marginBottom:10}}>
            {imgs.slice(0,5).map(f=>{ const url=di?.json_value?.[f.field_name]; return url?<img key={f.field_name} src={url} alt="" className="imt" onError={e=>e.target.style.display="none"}/>:null; })}
          </div>
          <div className="g2c">
            {ctx.map(f=><div key={f.field_name}><div className="xs m3">{f.field_name}</div><div className="sm fw5">{String(di?.json_value?.[f.field_name]||"—")}</div></div>)}
          </div>
        </div>

        {showVal && isSub && (
          <div className="card" style={{marginBottom:14}}>
            <div className="fw5" style={{marginBottom:10}}>Validation results</div>
            {cure.map(f=>{
              const uv=ta[f.field_name]||""; const gv=String(di?.json_value?.[f.field_name]||"");
              const sc2=scoreF(uv,gv,f.comparison_type,contest.semantic_correct_threshold||0.7);
              const ok=f.comparison_type==="semantic"?sc2>=(contest.semantic_correct_threshold||0.7):sc2===1;
              return (
                <div key={f.field_name} className={ok?"fsc":sc2>0?"fsp":"fsw"}>
                  <div className="fx ac jb"><span className="xs fw5">{f.field_name}</span><span className="xs mono">{Math.round(sc2*100)}%</span></div>
                  <div className="fx g4 xs" style={{marginTop:3}}><span><span style={{opacity:.6}}>You: </span>{uv||"—"}</span><span><span style={{opacity:.6}}>Answer: </span>{gv||"—"}</span></div>
                </div>
              );
            })}
            <div className="fx g3 jb" style={{marginTop:14}}>
              <button className="bg bsm" disabled={idx===0} onClick={()=>{setIdx(i=>i-1);setShowVal(false);}}>← Prev</button>
              {idx<items.length-1 && <button className="bg bsm" onClick={()=>{setIdx(i=>i+1);setShowVal(false);}}>Next task →</button>}
            </div>
          </div>
        )}

        {!showVal && (
          <div className="card">
            <div className="fx ac jb" style={{marginBottom:14}}>
              <div className="fw5">Fill in attributes</div>
              {isSub&&contest.mode==="practice"&&<span className="badge b-green">✓ Submitted</span>}
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:12}}>
              {cure.map(f=>{
                const val=ta[f.field_name]||"";
                const opts=f.dropdown_values?f.dropdown_values.split(";").map(s=>s.trim()).filter(Boolean):[];
                const dis=isSub&&contest.mode==="practice";
                return (
                  <div key={f.field_name}>
                    <label className="fl">{f.field_name} <span className="tag">{f.input_type}</span></label>
                    {f.input_type==="dropdown"
                      ?<select value={val} disabled={dis} onChange={async e=>saveAns(cur,f.field_name,e.target.value)}><option value="">Select...</option>{opts.map(o=><option key={o} value={o}>{o}</option>)}</select>
                      :f.input_type==="numeric"
                      ?<input type="number" value={val} disabled={dis} onChange={async e=>saveAns(cur,f.field_name,e.target.value)}/>
                      :<input type="text" value={val} disabled={dis} placeholder={`Enter ${f.field_name.toLowerCase()}...`} onChange={async e=>saveAns(cur,f.field_name,e.target.value)}/>
                    }
                  </div>
                );
              })}
            </div>
            <div className="fx g3 jb" style={{marginTop:18}}>
              <button className="bg bsm" disabled={idx===0} onClick={()=>setIdx(i=>i-1)}>← Previous</button>
              <div className="fx g2">
                {contest.mode==="practice"&&!isSub&&<button className="bp bsm" onClick={()=>submitPractice(cur)}>Validate</button>}
                {contest.mode==="practice"&&isSub&&<button className="bg bsm" onClick={()=>setShowVal(true)}>See answers</button>}
                {idx<items.length-1&&<button className="bg bsm" onClick={()=>setIdx(i=>i+1)}>Next →</button>}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default function App() {
  const [prof,setProf] = useState(null); const [loading,setLoading] = useState(true);

  useEffect(()=>{
    sb.auth.getSession().then(async({data:{session}})=>{
      if (session?.user) { const {data}=await sb.from("users").select("*").eq("id",session.user.id).single(); setProf(data); }
      setLoading(false);
    });
    const {data:{subscription}} = sb.auth.onAuthStateChange(async(_,session)=>{
      if (session?.user) { const {data}=await sb.from("users").select("*").eq("id",session.user.id).single(); setProf(data); }
      else setProf(null);
    });
    return ()=>subscription.unsubscribe();
  },[]);

  async function logout() { await sb.auth.signOut(); setProf(null); }

  if (loading) return <><style>{css}</style><div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:"var(--bg)"}}><span className="sp" style={{width:28,height:28}}/></div></>;

  return (
    <>
      <style>{css}</style>
      {!prof ? <LoginPage onLogin={p=>setProf(p)}/> : prof.role==="admin" ? <AdminShell user={prof} onLogout={logout}/> : <UserShell user={prof} onLogout={logout}/>}
    </>
  );
}

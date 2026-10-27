import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import ExcelJS from 'exceljs';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { supabase } from './supabase';
import './styles.css';

const monthNow = new Date().toISOString().slice(0, 7);
const todayISO = () => new Date().toISOString().slice(0, 10);
const brDate = s => {
  if (!s) return '';
  const [y,m,d] = s.split('-');
  return `${d}/${m}/${y}`;
};
const monthLabel = value => {
  if(!value) return 'Todos os períodos';
  const [year,month] = value.split('-');
  const names = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  return `${names[Number(month)-1]} de ${year}`;
};

const normalizePhone = phone => String(phone || '').replace(/\D/g, '');
const whatsappUrl = phone => {
  let n = normalizePhone(phone);
  if (!n) return '';
  if (!n.startsWith('55')) n = `55${n}`;
  return `https://wa.me/${n}`;
};
const whatsappDisplay = phone => {
  let n = normalizePhone(phone);
  if(n.startsWith('55') && n.length>11) n=n.slice(2);
  if(n.length===11) return `(${n.slice(0,2)}) ${n.slice(2,7)}-${n.slice(7)}`;
  if(n.length===10) return `(${n.slice(0,2)}) ${n.slice(2,6)}-${n.slice(6)}`;
  return phone || 'Não informado';
};
const followupWhatsappUrl = (phone,name) => {
  let n = normalizePhone(phone);
  if(!n) return '';
  if(!n.startsWith('55')) n=`55${n}`;
  const firstName=String(name||'').trim().split(/\s+/)[0] || 'Paciente';
  const message=`Olá, ${firstName}! Tudo bem? 😊 Aqui é da equipe da OftalmoCastro. Estou entrando em contato para acompanhar seu atendimento conosco. Podemos dar continuidade por aqui?`;
  return `https://wa.me/${n}?text=${encodeURIComponent(message)}`;
};
const countBy = (arr, key) => arr.reduce((o, x) => {
  const k = x[key] || 'Não informado';
  o[k] = (o[k] || 0) + 1;
  return o;
}, {});

const countManyBy = (arr, key) => arr.reduce((o, x) => {
  const values = Array.isArray(x[key]) && x[key].length ? x[key] : ['Não informado'];
  values.forEach(k => { o[k] = (o[k] || 0) + 1; });
  return o;
}, {});

const totalItems = (arr, key) => arr.reduce((total, item) => total + (Array.isArray(item[key]) ? item[key].length : 0), 0);
const totalItemsByStatus = (arr, key, status) => arr.reduce((total, item) =>
  total + (item.status === status && Array.isArray(item[key]) ? item[key].length : 0), 0);

const countWeightedBy = (arr, key, itemsKey) => arr.reduce((o, item) => {
  const label = item[key] || 'Não informado';
  const weight = Array.isArray(item[itemsKey]) && item[itemsKey].length ? item[itemsKey].length : 0;
  o[label] = (o[label] || 0) + weight;
  return o;
}, {});

const percentOf = (value,total) => total ? Math.round((value/total)*100) : 0;

const shiftMonth = (value,delta) => {
  const base = value || monthNow;
  const [year,month] = base.split('-').map(Number);
  const d = new Date(year,month-1+delta,1);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
};

const shortMonthLabel = value => {
  if(!value) return '';
  const [year,month] = value.split('-');
  const names = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  return `${names[Number(month)-1]}/${String(year).slice(-2)}`;
};

const variationPct = (current,previous) => {
  if(previous===0) return current>0 ? 100 : 0;
  return Math.round(((current-previous)/previous)*100);
};

const groupLaunchRows = (rows, kind) => {
  const groups = new Map();

  rows.forEach(r => {
    const groupId = r.launch_group_id || r.id;
    if(!groups.has(groupId)){
      groups.set(groupId,{
        group_id:groupId,
        db_ids:[],
        patient_id:r.patients?.id||'',
        date:kind==='exam' ? r.exam_date : r.surgery_date,
        patient:r.patients?.full_name||'',
        whatsapp:r.patients?.whatsapp||'',
        doctor:r.doctors?.name||'',
        doctor_id:r.doctors?.id||'',
        indicatedBy:r.indicated_by?.name||'',
        indicatedById:r.indicated_by?.id||'',
        status:r.status,
        obs:r.observation||'',
        launchedBy:r.launched_by,
        launchedByName:r.launched_by_name||''
      });
    }

    const g=groups.get(groupId);
    g.db_ids.push(r.id);

    if(kind==='exam'){
      g.exam_type_ids ||= [];
      g.exam_names ||= [];
      if(r.exam_types?.id) g.exam_type_ids.push(r.exam_types.id);
      if(r.exam_types?.name) g.exam_names.push(r.exam_types.name);
    }else{
      g.procedure_ids ||= [];
      g.procedure_names ||= [];
      if(r.procedures?.id) g.procedure_ids.push(r.procedures.id);
      if(r.procedures?.name) g.procedure_names.push(r.procedures.name);
      g.eye=r.eye;
      g.insurance=r.insurances?.name||'';
      g.insurance_id=r.insurances?.id||'';
      g.arrival=r.arrival_time||'';
      g.time=r.surgery_time||'';
      g.payment=r.payment_status||'';
    }
  });

  return [...groups.values()].map(g=>({
    ...g,
    id:g.group_id,
    ...(kind==='exam'
      ? {
          exam:g.exam_names.join(' • '),
          exam_type_id:g.exam_type_ids[0]||''
        }
      : {
          procedure:g.procedure_names.join(' • '),
          procedure_id:g.procedure_ids[0]||''
        })
  }));
};

function WhatsAppIcon(){
  return <svg viewBox="0 0 32 32" aria-hidden="true"><path fill="currentColor" d="M19.11 17.47c-.26-.13-1.52-.75-1.76-.84-.24-.09-.41-.13-.59.13-.17.26-.67.84-.82 1.02-.15.17-.3.2-.56.07-.26-.13-1.09-.4-2.08-1.29-.77-.68-1.29-1.53-1.44-1.79-.15-.26-.02-.4.11-.53.12-.12.26-.3.39-.45.13-.15.17-.26.26-.43.09-.17.04-.32-.02-.45-.07-.13-.59-1.42-.8-1.94-.21-.51-.43-.44-.59-.45h-.5c-.17 0-.45.07-.69.32-.24.26-.91.89-.91 2.17 0 1.27.93 2.5 1.06 2.67.13.17 1.82 2.78 4.41 3.9.62.27 1.1.43 1.47.55.62.2 1.18.17 1.62.1.49-.07 1.52-.62 1.73-1.22.21-.6.21-1.12.15-1.22-.07-.11-.24-.17-.5-.3Z"/><path fill="currentColor" d="M16.03 3C8.84 3 3 8.75 3 15.85c0 2.26.6 4.47 1.74 6.42L3 29l6.93-1.8a13.16 13.16 0 0 0 6.09 1.48h.01c7.18 0 13.02-5.76 13.02-12.84C29.05 8.75 23.21 3 16.03 3Zm0 23.51h-.01a10.94 10.94 0 0 1-5.57-1.51l-.4-.24-4.11 1.07 1.1-3.96-.26-.41a10.58 10.58 0 0 1-1.65-5.61C5.13 9.95 10.02 5.16 16.03 5.16S26.92 9.95 26.92 15.85c0 5.88-4.89 10.66-10.89 10.66Z"/></svg>;
}

function Bars({data}){
  const entries = Object.entries(data).sort((a,b)=>b[1]-a[1]);
  if(!entries.length) return <div className="empty">Sem dados para este período.</div>;
  const max = Math.max(...entries.map(([,v])=>v), 1);
  return <>{entries.slice(0,15).map(([k,v])=>
    <div className="bar-row" key={k}>
      <div className="ellipsis" title={k}>{k}</div>
      <div className="bar-bg"><div className="bar" style={{width:`${v/max*100}%`}} /></div>
      <b>{v}</b>
    </div>
  )}</>;
}


function DonutChart({title,data,centerLabel='Total'}){
  const entries=Object.entries(data||{}).filter(([,v])=>Number(v)>0);
  const total=entries.reduce((s,[,v])=>s+Number(v||0),0);
  const colors=['#4f8a6b','#315f72','#b9802b','#b85c5c','#7a6f9b','#6b8f9b','#9a7f56'];
  const radius=48;
  const circumference=2*Math.PI*radius;
  let offset=0;

  if(!total) return <div className="viz-card chart-empty-card"><div className="viz-title">{title}</div><div className="empty">Sem dados para este período.</div></div>;

  return <div className="viz-card donut-card">
    <div className="viz-title">{title}</div>
    <div className="donut-layout">
      <svg className="donut-svg" viewBox="0 0 130 130" role="img" aria-label={`${title}. Total ${total}`}>
        <circle cx="65" cy="65" r={radius} fill="none" stroke="#edf1f3" strokeWidth="18"/>
        {entries.map(([label,value],i)=>{
          const length=(Number(value)/total)*circumference;
          const dashoffset=-offset;
          offset+=length;
          return <circle
            key={label}
            cx="65" cy="65" r={radius}
            fill="none"
            stroke={colors[i%colors.length]}
            strokeWidth="18"
            strokeDasharray={`${length} ${circumference-length}`}
            strokeDashoffset={dashoffset}
            transform="rotate(-90 65 65)"
            strokeLinecap="butt"
          />;
        })}
        <text x="65" y="59" textAnchor="middle" className="donut-total">{total}</text>
        <text x="65" y="76" textAnchor="middle" className="donut-label">{centerLabel}</text>
      </svg>
      <div className="donut-legend">
        {entries.map(([label,value],i)=><div key={label}>
          <span className="legend-dot" style={{background:colors[i%colors.length]}}/>
          <span className="legend-name">{label}</span>
          <b>{value}</b>
          <small>{percentOf(value,total)}%</small>
        </div>)}
      </div>
    </div>
  </div>;
}

function TrendLineChart({title,labels,series,percent=false}){
  const width=760, height=260;
  const pad={l:52,r:22,t:28,b:44};
  const values=series.flatMap(s=>s.values||[]);
  const max=Math.max(...values,percent?100:1);
  const min=0;
  const plotW=width-pad.l-pad.r;
  const plotH=height-pad.t-pad.b;
  const x=i=>labels.length<=1?pad.l+plotW/2:pad.l+(i/(labels.length-1))*plotW;
  const y=v=>pad.t+plotH-((Number(v)-min)/(max-min||1))*plotH;
  const grid=[0,.25,.5,.75,1];

  return <div className="viz-card line-chart-card">
    <div className="viz-title-row">
      <div><div className="viz-title">{title}</div><small>{percent?'Percentual de conclusão por mês':'Comparação mensal de volume'}</small></div>
      <div className="line-legend">{series.map(s=><span key={s.name}><i style={{background:s.color}}/>{s.name}</span>)}</div>
    </div>
    <svg className="trend-svg" viewBox={`0 0 ${width} ${height}`} role="img" aria-label={title}>
      {grid.map(g=>{
        const val=Math.round(min+(max-min)*g);
        const gy=y(val);
        return <g key={g}>
          <line x1={pad.l} x2={width-pad.r} y1={gy} y2={gy} stroke="#e7edef" strokeWidth="1"/>
          <text x={pad.l-8} y={gy+4} textAnchor="end" className="trend-axis">{val}{percent?'%':''}</text>
        </g>;
      })}
      {labels.map((label,i)=><text key={label} x={x(i)} y={height-14} textAnchor="middle" className="trend-axis">{label}</text>)}
      {series.map(s=>{
        const points=(s.values||[]).map((v,i)=>`${x(i)},${y(v)}`).join(' ');
        return <g key={s.name}>
          <polyline points={points} fill="none" stroke={s.color} strokeWidth="3.2" strokeLinejoin="round" strokeLinecap="round"/>
          {(s.values||[]).map((v,i)=><g key={i}>
            <circle cx={x(i)} cy={y(v)} r="4.2" fill="#fff" stroke={s.color} strokeWidth="2.5"/>
            <text x={x(i)} y={y(v)-9} textAnchor="middle" className="trend-value">{v}{percent?'%':''}</text>
          </g>)}
        </g>;
      })}
    </svg>
  </div>;
}

function StatusTypeMatrix({title,subtitle,statuses,data,totalByStatus,labelName='Tipo'}){
  return <div className="card status-matrix-card">
    <div className="status-matrix-head">
      <div><span>DETALHAMENTO POR STATUS</span><h3>{title}</h3><p>{subtitle}</p></div>
    </div>
    <div className="status-matrix-grid">
      {statuses.map(status=>{
        const rows=Object.entries(data?.[status]||{}).sort((a,b)=>b[1]-a[1]);
        const total=totalByStatus?.[status]||0;
        return <section className={`status-matrix-section ${statusClass(status)}`} key={status}>
          <div className="status-matrix-title">
            <div><span className={`status ${statusClass(status)}`}>{status}</span><small>{labelName}s dentro deste status</small></div>
            <b>{total}</b>
          </div>
          {rows.length?<div className="status-type-list">
            {rows.map(([name,value])=><div key={name}>
              <span title={name}>{name}</span>
              <div className="mini-bar"><i style={{width:`${total?Math.max(4,(value/total)*100):0}%`}}/></div>
              <b>{value}</b>
              <small>{percentOf(value,total)}%</small>
            </div>)}
          </div>:<div className="status-type-empty">Nenhum registro neste status.</div>}
        </section>;
      })}
    </div>
  </div>;
}


function DoctorTypeMatrix({title,subtitle,rows,itemLabel='Procedimento'}){
  return <div className="card doctor-type-card">
    <div className="doctor-type-head">
      <div>
        <span>ANÁLISE POR MÉDICO</span>
        <h3>{title}</h3>
        <p>{subtitle}</p>
      </div>
    </div>

    {rows?.length ? <div className="doctor-type-grid">
      {rows.map(row=><section className="doctor-type-section" key={row.doctor}>
        <div className="doctor-type-title">
          <div>
            <b>{row.doctor}</b>
            <small>{row.total} {row.total===1?'registro':'registros'}</small>
          </div>
          <strong>{row.total}</strong>
        </div>

        <div className="doctor-type-list">
          {row.items.map(([name,value])=><div key={name}>
            <span title={name}>{name}</span>
            <div className="doctor-mini-bar"><i style={{width:`${row.total?Math.max(5,(value/row.total)*100):0}%`}}/></div>
            <b>{value}</b>
            <small>{percentOf(value,row.total)}%</small>
          </div>)}
        </div>
      </section>)}
    </div> : <div className="doctor-type-empty">
      Nenhum dado de {itemLabel.toLowerCase()} disponível para os filtros atuais.
    </div>}
  </div>;
}

function Modal({open,onClose,title,subtitle,children,onSave,saveText}){
  if(!open) return null;
  return <div className="modal" onMouseDown={e=>{if(e.target===e.currentTarget) onClose();}}>
    <div className="card modal-card">
      <div className="modal-head">
        <div><h2>{title}</h2>{subtitle&&<p>{subtitle}</p>}</div>
        <button className="icon-close" onClick={onClose}>×</button>
      </div>
      {children}
      <div className="form-actions">
        <button className="btn btn-light" onClick={onClose}>Cancelar</button>
        <button className="btn btn-primary" onClick={onSave}>{saveText}</button>
      </div>
    </div>
  </div>;
}

function PatientCell({name,phone}){
  const url = whatsappUrl(phone);
  return <div className="patient-cell">
    <div><b>{name}</b>{phone&&<small>{phone}</small>}</div>
    {url&&<a className="wa-btn" href={url} target="_blank" rel="noreferrer" title={`Abrir WhatsApp de ${name}`}><WhatsAppIcon/></a>}
  </div>;
}
function ObservationCell({text}) {
  return text ? <div className="obs-cell" title={text}>{text}</div> : <span className="muted">—</span>;
}
function Kpi({label,value}){return <div className="kpi"><small>{label}</small><strong>{value}</strong></div>;}

const statusClass = status => String(status||'')
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g,'')
  .toLowerCase()
  .replace(/\s+/g,'-');

function Select({value,setValue,options,first}) {
  return <select value={value} onChange={e=>setValue(e.target.value)}>
    {first&&<option value="">{first}</option>}
    {options.map(o=>{
      const value = typeof o === 'string' ? o : o.id;
      const label = typeof o === 'string' ? o : o.name;
      return <option key={value} value={value}>{label}</option>
    })}
  </select>;
}
function Field({label,children,full}){return <div className={full?'full':''}><label>{label}</label>{children}</div>;}

function accessLabel(profile){
  if(!profile) return 'Usuário';
  return profile.role === 'admin' ? 'Administrador' : 'Orientadora';
}

function Login({onLogged}) {
  const [email,setEmail] = useState('');
  const [password,setPassword] = useState('');
  const [showPassword,setShowPassword] = useState(false);
  const [loading,setLoading] = useState(false);
  const [error,setError] = useState('');

  const login = async e => {
    e.preventDefault();
    setLoading(true); setError('');
    const { data, error } = await supabase.auth.signInWithPassword({ email, password });
    if(error){ setError('E-mail ou senha inválidos.'); setLoading(false); return; }
    onLogged(data.session);
    setLoading(false);
  };

  return <div className="login-page">
    <form className="login-card" onSubmit={login}>
      <div className="brand-mark">OC</div>
      <h1>Oftalmocastro</h1>
      <p>Gestão de exames e cirurgias</p>
      <Field label="E-mail"><input type="email" value={email} onChange={e=>setEmail(e.target.value)} required /></Field>
      <Field label="Senha">
        <div className="password-field">
          <input type={showPassword ? 'text' : 'password'} value={password} onChange={e=>setPassword(e.target.value)} required />
          <button type="button" className="password-toggle" onClick={()=>setShowPassword(v=>!v)} aria-label={showPassword ? 'Ocultar senha' : 'Mostrar senha'}>
            {showPassword ? 'Ocultar' : 'Ver'}
          </button>
        </div>
      </Field>
      {error&&<div className="login-error">{error}</div>}
      <button className="btn btn-primary login-btn" disabled={loading}>{loading?'Entrando...':'Entrar'}</button>
    </form>
  </div>;
}

function App(){
  const [session,setSession] = useState(null);
  const [profile,setProfile] = useState(null);
  const [loading,setLoading] = useState(true);
  const [page,setPage] = useState('exames');
  const [launcherName,setLauncherName] = useState(()=>sessionStorage.getItem('oc_launcher_name') || '');

  const [doctors,setDoctors] = useState([]);
  const [examTypes,setExamTypes] = useState([]);
  const [procedures,setProcedures] = useState([]);
  const [insurances,setInsurances] = useState([]);
  const [profiles,setProfiles] = useState([]);
  const [records,setRecords] = useState([]);
  const [surgeries,setSurgeries] = useState([]);
  const [followups,setFollowups] = useState([]);

  const [examModal,setExamModal] = useState(false);
  const [surgeryModal,setSurgeryModal] = useState(false);
  const [catalogModal,setCatalogModal] = useState(false);
  const [editingExamId,setEditingExamId] = useState(null);
  const [editingSurgeryId,setEditingSurgeryId] = useState(null);
  const [followupModal,setFollowupModal] = useState(false);
  const [followupTarget,setFollowupTarget] = useState(null);

  const [filterMonth,setFilterMonth] = useState(monthNow);
  const [filterDoctor,setFilterDoctor] = useState('');
  const [filterExam,setFilterExam] = useState('');
  const [sMonth,setSMonth] = useState(monthNow);
  const [sDoctor,setSDoctor] = useState('');
  const [sProcedure,setSProcedure] = useState('');
  const [dashMonth,setDashMonth] = useState(monthNow);
  const [dashDoctor,setDashDoctor] = useState('');
  const [dashUser,setDashUser] = useState('');
  const [dashStatus,setDashStatus] = useState('');
  const [followDay,setFollowDay] = useState('');
  const [followType,setFollowType] = useState('');
  const [followDoctor,setFollowDoctor] = useState('');
  const [followSourceStatus,setFollowSourceStatus] = useState('');
  const [followStatus,setFollowStatus] = useState('');
  const [followSearch,setFollowSearch] = useState('');
  const [exportingPdf,setExportingPdf] = useState(false);
  const [exportingPanelPdf,setExportingPanelPdf] = useState(false);
  const resultsPageRef = useRef(null);
  const pdfPage1Ref = useRef(null);
  const pdfPage2Ref = useRef(null);
  const pdfPage3Ref = useRef(null);

  const [examForm,setExamForm] = useState({
    date:todayISO(),patient:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',exam_type_ids:[],status:'Agendado',launcher:'',obs:''
  });
  const [sForm,setSForm] = useState({
    date:todayISO(),patient:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',procedure_ids:[],eye:'Não se aplica',
    insurance_id:'',status:'Solicitação',arrival:'',time:'',payment:'Não informado',launcher:'',obs:''
  });
  const [followupForm,setFollowupForm] = useState({
    followup_date:todayISO(),status:'A fazer',next_contact_date:'',responsible:'',note:''
  });

  useEffect(()=>{
    supabase.auth.getSession().then(({data})=>{
      setSession(data.session);
      setLoading(false);
    });
    const {data:listener} = supabase.auth.onAuthStateChange((_event,session)=>{
      setSession(session);
    });
    return ()=>listener.subscription.unsubscribe();
  },[]);

  useEffect(()=>{
    if(session?.user?.id) loadAll();
  },[session?.user?.id]);

  const launcherOptions = ['Juliana','Luiza','Elisangela'];
  const launcherOptionsAll = ['Juliana','Luiza','Elisangela','ADM'];

  const currentLaunchName = profile?.role === 'admin'
    ? 'ADM'
    : launcherName;

  const examLaunchName = examForm.launcher || currentLaunchName;
  const surgeryLaunchName = sForm.launcher || currentLaunchName;

  const selectLauncher = name => {
    setLauncherName(name);
    sessionStorage.setItem('oc_launcher_name', name);
  };

  const loadAll = async ()=>{
    setLoading(true);
    const userId = session.user.id;
    const [
      profileRes, doctorsRes, examTypesRes, proceduresRes,
      insurancesRes, profilesRes, examsRes, surgeriesRes, followupsRes
    ] = await Promise.all([
      supabase.from('profiles').select('*').eq('id',userId).single(),
      supabase.from('doctors').select('*').eq('active',true).order('name'),
      supabase.from('exam_types').select('*').eq('active',true).order('name'),
      supabase.from('procedures').select('*').eq('active',true).order('name'),
      supabase.from('insurances').select('*').eq('active',true).order('name'),
      supabase.from('profiles').select('id,full_name,role,active').eq('active',true).order('full_name'),
      supabase.from('exams').select(`
        id,launch_group_id,exam_date,status,observation,launched_by,launched_by_name,
        patients(id,full_name,whatsapp),
        doctors:doctors!exams_doctor_id_fkey(id,name),
        indicated_by:doctors!exams_indicated_by_doctor_id_fkey(id,name),
        exam_types(id,name)
      `).order('exam_date',{ascending:false}),
      supabase.from('surgeries').select(`
        id,launch_group_id,surgery_date,eye,status,arrival_time,surgery_time,payment_status,observation,launched_by,launched_by_name,
        patients(id,full_name,whatsapp),
        doctors:doctors!surgeries_doctor_id_fkey(id,name),
        indicated_by:doctors!surgeries_indicated_by_doctor_id_fkey(id,name),
        procedures(id,name),
        insurances(id,name)
      `).order('surgery_date',{ascending:false}),
      supabase.from('followups').select('*').order('updated_at',{ascending:false})
    ]);

    const errors=[profileRes,doctorsRes,examTypesRes,proceduresRes,insurancesRes,profilesRes,examsRes,surgeriesRes,followupsRes]
      .map(x=>x.error).filter(Boolean);
    if(errors.length) console.error(errors);

    setProfile(profileRes.data || null);
    setDoctors(doctorsRes.data || []);
    setExamTypes(examTypesRes.data || []);
    setProcedures(proceduresRes.data || []);
    setInsurances(insurancesRes.data || []);
    setProfiles(profilesRes.data || []);

    setRecords(groupLaunchRows(examsRes.data||[],'exam'));
    setSurgeries(groupLaunchRows(surgeriesRes.data||[],'surgery'));
    setFollowups(followupsRes?.data || []);
    setLoading(false);
  };

  const getOrCreatePatient = async (name,whatsapp)=>{
    const cleanName=name.trim();
    const cleanPhone=whatsapp.trim();

    if(cleanPhone){
      const {data:byPhone}=await supabase.from('patients').select('id').eq('whatsapp',cleanPhone).limit(1).maybeSingle();
      if(byPhone?.id) return byPhone.id;
    }
    const {data:byName}=await supabase.from('patients').select('id,whatsapp').ilike('full_name',cleanName).limit(1).maybeSingle();
    if(byName?.id){
      if(cleanPhone && normalizePhone(byName.whatsapp)!==normalizePhone(cleanPhone)){
        const {error:updatePhoneError}=await supabase.from('patients').update({whatsapp:cleanPhone}).eq('id',byName.id);
        if(updatePhoneError) throw updatePhoneError;
      }
      return byName.id;
    }

    const {data,error}=await supabase.from('patients').insert({
      full_name:cleanName,
      whatsapp:cleanPhone||null,
      created_by:session.user.id
    }).select('id').single();
    if(error) throw error;
    return data.id;
  };

  const addExam = async ()=>{
    if(!examLaunchName) return alert('Selecione quem está fazendo o lançamento.');
    if(!examForm.date||!examForm.patient.trim()||!examForm.doctor_id||!examForm.exam_type_ids?.length)
      return alert('Preencha data, paciente, médico e selecione pelo menos um exame.');
    try{
      const patient_id=await getOrCreatePatient(examForm.patient,examForm.whatsapp);
      const launch_group_id=crypto.randomUUID();
      const rows=examForm.exam_type_ids.map(examTypeId=>({
        launch_group_id,
        exam_date:examForm.date, patient_id, doctor_id:examForm.doctor_id,
        indicated_by_doctor_id:examForm.indicated_by_doctor_id||null,
        exam_type_id:examTypeId, status:examForm.status,
        observation:examForm.obs||null, launched_by:session.user.id,
        launched_by_name:examLaunchName
      }));
      const {error}=await supabase.from('exams').insert(rows);
      if(error) throw error;
      setExamModal(false);
      setExamForm({date:todayISO(),patient:'',whatsapp:'',doctor_id:'',exam_type_ids:[],status:'Agendado',launcher:currentLaunchName||'',obs:''});
      await loadAll();
    }catch(e){ console.error(e); alert('Não foi possível salvar os exames.'); }
  };

  const addSurgery = async ()=>{
    if(!surgeryLaunchName) return alert('Selecione quem está fazendo o lançamento.');
    if(!sForm.date||!sForm.patient.trim()||!sForm.doctor_id||!sForm.procedure_ids?.length)
      return alert('Preencha data, paciente, cirurgião e selecione pelo menos um procedimento.');

    try{
      const patient_id=await getOrCreatePatient(sForm.patient,sForm.whatsapp);
      const launch_group_id=crypto.randomUUID();

      const rows=sForm.procedure_ids.map(procedureId=>({
        launch_group_id,
        surgery_date:sForm.date,
        patient_id,
        doctor_id:sForm.doctor_id,
        indicated_by_doctor_id:sForm.indicated_by_doctor_id||null,
        procedure_id:procedureId,
        eye:sForm.eye,
        insurance_id:sForm.insurance_id||null,
        status:sForm.status,
        arrival_time:sForm.arrival||null,
        surgery_time:sForm.time||null,
        payment_status:sForm.payment,
        observation:sForm.obs||null,
        launched_by:session.user.id,
        launched_by_name:surgeryLaunchName
      }));

      const {error}=await supabase.from('surgeries').insert(rows);
      if(error) throw error;

      setSurgeryModal(false);
      setSForm({
        date:todayISO(),patient:'',whatsapp:'',doctor_id:'',procedure_ids:[],
        eye:'Não se aplica',insurance_id:'',status:'Solicitação',
        arrival:'',time:'',payment:'Não informado',launcher:currentLaunchName||'',obs:''
      });
      await loadAll();
    }catch(e){
      console.error(e);
      alert('Não foi possível salvar os procedimentos cirúrgicos.');
    }
  };


  const openEditExam = r => {
    setEditingExamId(r.id);
    setExamForm({
      date:r.date,
      patient:r.patient,
      whatsapp:r.whatsapp || '',
      doctor_id:r.doctor_id,
      indicated_by_doctor_id:r.indicatedById||'',
      exam_type_ids:r.exam_type_ids||[],
      status:r.status,
      launcher:r.launchedByName || currentLaunchName || '',
      obs:r.obs || ''
    });
    setExamModal(true);
  };

  const openEditSurgery = r => {
    setEditingSurgeryId(r.id);
    setSForm({
      date:r.date,
      patient:r.patient,
      whatsapp:r.whatsapp || '',
      doctor_id:r.doctor_id,
      indicated_by_doctor_id:r.indicatedById||'',
      procedure_ids:r.procedure_ids||[],
      eye:r.eye || 'Não se aplica',
      insurance_id:r.insurance_id || '',
      status:r.status,
      arrival:r.arrival || '',
      time:r.time || '',
      payment:r.payment || 'Não informado',
      launcher:r.launchedByName || currentLaunchName || '',
      obs:r.obs || ''
    });
    setSurgeryModal(true);
  };

  const updatePatient = async (patientId,name,whatsapp) => {
    if(!patientId) return;
    const {error} = await supabase.from('patients').update({
      full_name:name.trim(),
      whatsapp:whatsapp.trim() || null
    }).eq('id',patientId);
    if(error) throw error;
  };

  const updateExam = async ()=>{
    const original = records.find(r=>r.id===editingExamId);
    if(!original) return;
    if(!examLaunchName) return alert('Selecione quem está fazendo o lançamento.');
    if(!examForm.date||!examForm.patient.trim()||!examForm.doctor_id||!examForm.exam_type_ids?.length)
      return alert('Preencha data, paciente, médico e selecione pelo menos um exame.');

    try{
      await updatePatient(original.patient_id, examForm.patient, examForm.whatsapp);

      const {error:deleteError}=await supabase.from('exams').delete().in('id',original.db_ids);
      if(deleteError) throw deleteError;

      const rows=examForm.exam_type_ids.map(examTypeId=>({
        launch_group_id:original.group_id,
        exam_date:examForm.date,
        patient_id:original.patient_id,
        doctor_id:examForm.doctor_id,
        indicated_by_doctor_id:examForm.indicated_by_doctor_id||null,
        exam_type_id:examTypeId,
        status:examForm.status,
        observation:examForm.obs||null,
        launched_by:session.user.id,
        launched_by_name:examLaunchName
      }));

      const {error}=await supabase.from('exams').insert(rows);
      if(error) throw error;

      setExamModal(false);
      setEditingExamId(null);
      setExamForm({date:todayISO(),patient:'',whatsapp:'',doctor_id:'',exam_type_ids:[],status:'Agendado',launcher:currentLaunchName||'',obs:''});
      await loadAll();
    }catch(e){
      console.error(e);
      alert('Não foi possível atualizar o lançamento de exames.');
    }
  };

  const updateSurgery = async ()=>{
    const original = surgeries.find(r=>r.id===editingSurgeryId);
    if(!original) return;
    if(!surgeryLaunchName) return alert('Selecione quem está fazendo o lançamento.');
    if(!sForm.date||!sForm.patient.trim()||!sForm.doctor_id||!sForm.procedure_ids?.length)
      return alert('Preencha data, paciente, cirurgião e selecione pelo menos um procedimento.');

    try{
      await updatePatient(original.patient_id, sForm.patient, sForm.whatsapp);

      const {error:deleteError}=await supabase.from('surgeries').delete().in('id',original.db_ids);
      if(deleteError) throw deleteError;

      const rows=sForm.procedure_ids.map(procedureId=>({
        launch_group_id:original.group_id,
        surgery_date:sForm.date,
        patient_id:original.patient_id,
        doctor_id:sForm.doctor_id,
        procedure_id:procedureId,
        eye:sForm.eye,
        insurance_id:sForm.insurance_id||null,
        status:sForm.status,
        arrival_time:sForm.arrival||null,
        surgery_time:sForm.time||null,
        payment_status:sForm.payment,
        observation:sForm.obs||null,
        launched_by:session.user.id,
        launched_by_name:surgeryLaunchName
      }));

      const {error}=await supabase.from('surgeries').insert(rows);
      if(error) throw error;

      setSurgeryModal(false);
      setEditingSurgeryId(null);
      setSForm({date:todayISO(),patient:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',procedure_ids:[],eye:'Não se aplica',insurance_id:'',status:'Solicitação',arrival:'',time:'',payment:'Não informado',launcher:currentLaunchName||'',obs:''});
      await loadAll();
    }catch(e){
      console.error(e);
      alert('Não foi possível atualizar o lançamento cirúrgico.');
    }
  };

  const closeExamModal = ()=>{
    setExamModal(false);
    setEditingExamId(null);
    setExamForm({date:todayISO(),patient:'',whatsapp:'',doctor_id:'',exam_type_ids:[],status:'Agendado',launcher:currentLaunchName||'',obs:''});
  };

  const closeSurgeryModal = ()=>{
    setSurgeryModal(false);
    setEditingSurgeryId(null);
    setSForm({date:todayISO(),patient:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',procedure_ids:[],eye:'Não se aplica',insurance_id:'',status:'Solicitação',arrival:'',time:'',payment:'Não informado',launcher:currentLaunchName||'',obs:''});
  };

  const delExam = async id=>{
    if(!confirm('Tem certeza que deseja excluir este lançamento de exames? Essa ação não poderá ser desfeita.')) return;
    const item=records.find(r=>r.id===id);
    const {error}=await supabase.from('exams').delete().in('id',item?.db_ids||[id]);
    if(error) return alert('Não foi possível excluir. Verifique as permissões do seu usuário no Supabase.');
    await loadAll();
  };
  const delSurgery = async id=>{
    if(!confirm('Tem certeza que deseja excluir este lançamento cirúrgico? Essa ação não poderá ser desfeita.')) return;
    const item=surgeries.find(r=>r.id===id);
    const {error}=await supabase.from('surgeries').delete().in('id',item?.db_ids||[id]);
    if(error) return alert('Não foi possível excluir. Verifique as permissões do seu usuário no Supabase.');
    await loadAll();
  };

  const addCatalog = async (table,name,extra={})=>{
    const clean=name.trim();
    if(!clean) return;
    const {error}=await supabase.from(table).insert({name:clean,...extra});
    if(error) return alert(error.message);
    await loadAll();
  };
  const deactivateCatalog = async (table,id)=>{
    if(!confirm('Remover este item da lista? Os lançamentos antigos serão preservados.')) return;
    const {error}=await supabase.from(table).update({active:false}).eq('id',id);
    if(error) return alert(error.message);
    await loadAll();
  };

  const filteredExams = useMemo(()=>records.filter(r=>
    (!filterMonth||r.date.startsWith(filterMonth)) &&
    (!filterDoctor||r.doctor_id===filterDoctor) &&
    (!filterExam||r.exam_type_ids?.includes(filterExam))
  ),[records,filterMonth,filterDoctor,filterExam]);

  const filteredSurgeries = useMemo(()=>surgeries.filter(r=>
    (!sMonth||r.date.startsWith(sMonth)) &&
    (!sDoctor||r.doctor_id===sDoctor) &&
    (!sProcedure||r.procedure_ids?.includes(sProcedure))
  ),[surgeries,sMonth,sDoctor,sProcedure]);

  const deleteFilteredExams = async ()=>{
    if(!filteredExams.length) return alert('Nenhum exame encontrado com os filtros atuais.');

    const message = `Você está prestes a excluir ${filteredExams.length} lançamento(s) de exames exibidos pelos filtros atuais. Essa ação não poderá ser desfeita. Deseja continuar?`;
    if(!confirm(message)) return;

    const ids = [...new Set(filteredExams.flatMap(item=>item.db_ids||[]))];
    if(!ids.length) return alert('Não foi possível identificar os exames para exclusão.');

    const {error}=await supabase.from('exams').delete().in('id',ids);
    if(error){
      console.error(error);
      return alert('Não foi possível excluir os exames filtrados.');
    }

    await loadAll();
    alert('Exames filtrados excluídos com sucesso.');
  };

  const deleteFilteredSurgeries = async ()=>{
    if(!filteredSurgeries.length) return alert('Nenhuma cirurgia encontrada com os filtros atuais.');

    const message = `Você está prestes a excluir ${filteredSurgeries.length} lançamento(s) de cirurgias exibidos pelos filtros atuais. Essa ação não poderá ser desfeita. Deseja continuar?`;
    if(!confirm(message)) return;

    const ids = [...new Set(filteredSurgeries.flatMap(item=>item.db_ids||[]))];
    if(!ids.length) return alert('Não foi possível identificar as cirurgias para exclusão.');

    const {error}=await supabase.from('surgeries').delete().in('id',ids);
    if(error){
      console.error(error);
      return alert('Não foi possível excluir as cirurgias filtradas.');
    }

    await loadAll();
    alert('Cirurgias filtradas excluídas com sucesso.');
  };

  const followupSourceRows = useMemo(()=>{
    const byKey=new Map(followups.map(f=>[`${f.source_type}:${f.source_group_id}`,f]));
    const examRows=records.map(r=>({
      key:`exam:${r.group_id||r.id}`,
      source_type:'exam',
      source_group_id:r.group_id||r.id,
      sourceRecord:r,
      patient_id:r.patient_id,
      patient:r.patient,
      whatsapp:r.whatsapp,
      source_date:r.date,
      doctor:r.doctor,
      doctor_id:r.doctor_id,
      item:r.exam,
      source_status:r.status,
      followup:byKey.get(`exam:${r.group_id||r.id}`)||null
    }));
    const surgeryRows=surgeries.map(r=>({
      key:`surgery:${r.group_id||r.id}`,
      source_type:'surgery',
      source_group_id:r.group_id||r.id,
      sourceRecord:r,
      patient_id:r.patient_id,
      patient:r.patient,
      whatsapp:r.whatsapp,
      source_date:r.date,
      doctor:r.doctor,
      doctor_id:r.doctor_id,
      item:r.procedure,
      source_status:r.status,
      followup:byKey.get(`surgery:${r.group_id||r.id}`)||null
    }));
    return [...examRows,...surgeryRows].sort((a,b)=>String(b.source_date||'').localeCompare(String(a.source_date||'')));
  },[records,surgeries,followups]);

  const filteredFollowups = useMemo(()=>{
    const q=followSearch.trim().toLowerCase();
    return followupSourceRows.filter(row=>
      (!followDay||row.source_date===followDay) &&
      (!followType||row.source_type===followType) &&
      (!followDoctor||row.doctor_id===followDoctor) &&
      (!followSourceStatus||row.source_status===followSourceStatus) &&
      (!followStatus||(row.followup?.status||'A fazer')===followStatus) &&
      (!q||String(row.patient||'').toLowerCase().includes(q)||String(row.item||'').toLowerCase().includes(q)||String(row.doctor||'').toLowerCase().includes(q))
    );
  },[followupSourceRows,followDay,followType,followDoctor,followSourceStatus,followStatus,followSearch]);

  const followupStats = useMemo(()=>{
    const total=filteredFollowups.length;
    const registered=filteredFollowups.filter(r=>Boolean(r.followup)).length;
    const notStarted=filteredFollowups.filter(r=>!r.followup).length;
    const inContact=filteredFollowups.filter(r=>r.followup?.status==='Em contato').length;
    const waiting=filteredFollowups.filter(r=>r.followup?.status==='Aguardando paciente').length;
    const noReturn=filteredFollowups.filter(r=>r.followup?.status==='Sem retorno').length;
    const done=filteredFollowups.filter(r=>r.followup?.status==='Concluído').length;
    const dueToday=filteredFollowups.filter(r=>r.followup?.next_contact_date===todayISO()).length;
    const followupRate=total?Math.round((registered/total)*100):0;
    const completionRate=registered?Math.round((done/registered)*100):0;
    return {total,registered,notStarted,inContact,waiting,noReturn,done,dueToday,followupRate,completionRate};
  },[filteredFollowups]);

  const openFollowup = row=>{
    const f=row.followup;
    setFollowupTarget(row);
    setFollowupForm({
      followup_date:f?.followup_date||todayISO(),
      status:f?.status||'A fazer',
      next_contact_date:f?.next_contact_date||'',
      responsible:f?.responsible_name||currentLaunchName||'',
      note:f?.note||''
    });
    setFollowupModal(true);
  };

  const saveFollowup = async()=>{
    if(!followupTarget) return;
    if(!followupForm.responsible) return alert('Selecione o responsável pelo follow-up.');
    try{
      const payload={
        source_type:followupTarget.source_type,
        source_group_id:followupTarget.source_group_id,
        patient_id:followupTarget.patient_id,
        followup_date:followupForm.followup_date||todayISO(),
        status:followupForm.status,
        next_contact_date:followupForm.next_contact_date||null,
        responsible_name:followupForm.responsible,
        note:followupForm.note||null,
        created_by:session.user.id,
        updated_at:new Date().toISOString()
      };
      const {error}=await supabase.from('followups').upsert(payload,{onConflict:'source_type,source_group_id'});
      if(error) throw error;
      setFollowupModal(false);
      setFollowupTarget(null);
      await loadAll();
    }catch(e){
      console.error(e);
      alert('Não foi possível salvar o follow-up. Confirme se a tabela followups foi criada no Supabase.');
    }
  };

  const openOriginalFromFollowup = row=>{
    if(row.source_type==='exam'){
      openEditExam(row.sourceRecord);
      setPage('exames');
    }else{
      openEditSurgery(row.sourceRecord);
      setPage('cirurgias');
    }
  };

  const dashboardExams = useMemo(()=>records.filter(r=>
    (!dashMonth||r.date.startsWith(dashMonth)) &&
    (!dashDoctor||r.doctor_id===dashDoctor) &&
    (!dashUser||r.launchedByName===dashUser) &&
    (!dashStatus||r.status===dashStatus)
  ),[records,dashMonth,dashDoctor,dashUser,dashStatus]);

  const dashboardSurgeries = useMemo(()=>surgeries.filter(r=>
    (!dashMonth||r.date.startsWith(dashMonth)) &&
    (!dashDoctor||r.doctor_id===dashDoctor) &&
    (!dashUser||r.launchedByName===dashUser) &&
    (!dashStatus||r.status===dashStatus)
  ),[surgeries,dashMonth,dashDoctor,dashUser,dashStatus]);

  const launchedName = record => record.launchedByName || (profiles.find(p=>p.id===record.launchedBy)?.role==='admin' ? 'ADM' : 'Orientadora');

  const userRows = ['Juliana','Luiza','Elisangela','ADM'].map(name=>({
    key:name,
    name,
    exams:totalItems(dashboardExams.filter(x=>launchedName(x)===name),'exam_type_ids'),
    surgeries:totalItems(dashboardSurgeries.filter(x=>launchedName(x)===name),'procedure_ids')
  }));

  const currentAccessName = currentLaunchName || 'Selecione quem está lançando';

  const resultSummary=useMemo(()=>{
    const totalEx=totalItems(dashboardExams,'exam_type_ids');
    const totalSu=totalItems(dashboardSurgeries,'procedure_ids');

    const exAgendado=totalItemsByStatus(dashboardExams,'exam_type_ids','Agendado');
    const exPendente=totalItemsByStatus(dashboardExams,'exam_type_ids','Pendente');
    const realEx=totalItemsByStatus(dashboardExams,'exam_type_ids','Realizado');
    const exCancelado=totalItemsByStatus(dashboardExams,'exam_type_ids','Cancelado');
    const exOpen=exAgendado+exPendente;

    const suSolicitacao=totalItemsByStatus(dashboardSurgeries,'procedure_ids','Solicitação');
    const suPendente=totalItemsByStatus(dashboardSurgeries,'procedure_ids','Pendente');
    const suAutorizada=totalItemsByStatus(dashboardSurgeries,'procedure_ids','Autorizada');
    const suAgendada=totalItemsByStatus(dashboardSurgeries,'procedure_ids','Agendada');
    const realSu=totalItemsByStatus(dashboardSurgeries,'procedure_ids','Realizada');
    const suCancelada=totalItemsByStatus(dashboardSurgeries,'procedure_ids','Cancelada');
    const suOpen=suSolicitacao+suPendente+suAutorizada+suAgendada;

    const paymentCounts=countWeightedBy(dashboardSurgeries,'payment','procedure_ids');
    const paymentPending=(paymentCounts['Pendente']||0)+(paymentCounts['20% pago']||0)+(paymentCounts['Não informado']||0);
    const paymentClosed=(paymentCounts['Pago']||0)+(paymentCounts['Convênio']||0);

    const examTop=Object.entries(countManyBy(dashboardExams,'exam_names')).sort((a,b)=>b[1]-a[1])[0];
    const surgTop=Object.entries(countManyBy(dashboardSurgeries,'procedure_names')).sort((a,b)=>b[1]-a[1])[0];
    const doctorTop=Object.entries(countWeightedBy(
      [...dashboardExams.map(x=>({...x,_items:x.exam_type_ids})),...dashboardSurgeries.map(x=>({...x,_items:x.procedure_ids}))],
      'doctor',
      '_items'
    )).sort((a,b)=>b[1]-a[1])[0];

    const insuranceTop=Object.entries(
      countWeightedBy(
        dashboardSurgeries.map(x=>({...x,insuranceLabel:x.insurance||'Particular'})),
        'insuranceLabel',
        'procedure_ids'
      )
    ).sort((a,b)=>b[1]-a[1])[0];

    const uniquePatients=new Set(
      [...dashboardExams,...dashboardSurgeries]
        .map(x=>`${String(x.patient||'').trim().toLowerCase()}|${normalizePhone(x.whatsapp)}`)
        .filter(x=>x!=='|')
    ).size;

    return {
      totalEx,totalSu,realEx,realSu,
      exAgendado,exPendente,exCancelado,exOpen,
      suSolicitacao,suPendente,suAutorizada,suAgendada,suCancelada,suOpen,
      paymentCounts,paymentPending,paymentClosed,
      examTop,surgTop,doctorTop,insuranceTop,uniquePatients,
      exRate:percentOf(realEx,totalEx),
      suRate:percentOf(realSu,totalSu),
      exOpenRate:percentOf(exOpen,totalEx),
      suOpenRate:percentOf(suOpen,totalSu),
      paymentPendingRate:percentOf(paymentPending,totalSu)
    };
  },[dashboardExams,dashboardSurgeries]);

  const advancedAnalytics=useMemo(()=>{
    const baseMonth=dashMonth||monthNow;
    const previousMonth=shiftMonth(baseMonth,-1);
    const months=Array.from({length:6},(_,i)=>shiftMonth(baseMonth,i-5));

    const baseExamRows=records.filter(r=>
      (!dashDoctor||r.doctor_id===dashDoctor) &&
      (!dashUser||r.launchedByName===dashUser)
    );

    const baseSurgeryRows=surgeries.filter(r=>
      (!dashDoctor||r.doctor_id===dashDoctor) &&
      (!dashUser||r.launchedByName===dashUser)
    );

    const byMonth=(arr,month)=>arr.filter(r=>r.date?.startsWith(month));

    const examTrend=months.map(month=>totalItems(byMonth(baseExamRows,month),'exam_type_ids'));
    const surgeryTrend=months.map(month=>totalItems(byMonth(baseSurgeryRows,month),'procedure_ids'));
    const examConversion=months.map(month=>{
      const rows=byMonth(baseExamRows,month);
      const total=totalItems(rows,'exam_type_ids');
      const done=totalItemsByStatus(rows,'exam_type_ids','Realizado');
      return percentOf(done,total);
    });
    const surgeryConversion=months.map(month=>{
      const rows=byMonth(baseSurgeryRows,month);
      const total=totalItems(rows,'procedure_ids');
      const done=totalItemsByStatus(rows,'procedure_ids','Realizada');
      return percentOf(done,total);
    });

    const currentExamRows=byMonth(baseExamRows,baseMonth);
    const previousExamRows=byMonth(baseExamRows,previousMonth);
    const currentSurgeryRows=byMonth(baseSurgeryRows,baseMonth);
    const previousSurgeryRows=byMonth(baseSurgeryRows,previousMonth);

    const currentExamTotal=totalItems(currentExamRows,'exam_type_ids');
    const previousExamTotal=totalItems(previousExamRows,'exam_type_ids');
    const currentSurgeryTotal=totalItems(currentSurgeryRows,'procedure_ids');
    const previousSurgeryTotal=totalItems(previousSurgeryRows,'procedure_ids');

    const currentExamDone=totalItemsByStatus(currentExamRows,'exam_type_ids','Realizado');
    const previousExamDone=totalItemsByStatus(previousExamRows,'exam_type_ids','Realizado');
    const currentSurgeryDone=totalItemsByStatus(currentSurgeryRows,'procedure_ids','Realizada');
    const previousSurgeryDone=totalItemsByStatus(previousSurgeryRows,'procedure_ids','Realizada');

    const examStatuses=['Realizado','Agendado','Pendente','Cancelado'];
    const surgeryStatuses=['Realizada','Agendada','Autorizada','Pendente','Solicitação','Cancelada'];

    const examByStatusType=Object.fromEntries(examStatuses.map(status=>[
      status,
      countManyBy(dashboardExams.filter(r=>r.status===status),'exam_names')
    ]));

    const surgeryByStatusType=Object.fromEntries(surgeryStatuses.map(status=>[
      status,
      countManyBy(dashboardSurgeries.filter(r=>r.status===status),'procedure_names')
    ]));

    const examStatusTotals=Object.fromEntries(examStatuses.map(status=>[
      status,totalItemsByStatus(dashboardExams,'exam_type_ids',status)
    ]));

    const surgeryStatusTotals=Object.fromEntries(surgeryStatuses.map(status=>[
      status,totalItemsByStatus(dashboardSurgeries,'procedure_ids',status)
    ]));

    const examStatusPie=examStatusTotals;
    const surgeryStatusPie=surgeryStatusTotals;

    const buildDoctorRows=(arr,namesKey,doctorKey='doctor')=>{
      const grouped=new Map();

      arr.forEach(record=>{
        const doctor=record[doctorKey]||'Não informado';
        if(!grouped.has(doctor)) grouped.set(doctor,{doctor,total:0,items:{}});
        const row=grouped.get(doctor);
        const names=Array.isArray(record[namesKey])?record[namesKey]:[];
        names.forEach(name=>{
          row.total+=1;
          row.items[name]=(row.items[name]||0)+1;
        });
      });

      return [...grouped.values()]
        .map(row=>({...row,items:Object.entries(row.items).sort((a,b)=>b[1]-a[1])}))
        .sort((a,b)=>b.total-a.total);
    };

    const examByDoctor=buildDoctorRows(dashboardExams,'exam_names','doctor');
    const surgeryByDoctor=buildDoctorRows(dashboardSurgeries,'procedure_names','doctor');

    const examByIndication=buildDoctorRows(
      dashboardExams.filter(r=>r.indicatedBy),
      'exam_names',
      'indicatedBy'
    );

    const surgeryByIndication=buildDoctorRows(
      dashboardSurgeries.filter(r=>r.indicatedBy),
      'procedure_names',
      'indicatedBy'
    );

    // Conversão por médico do próprio lançamento (doctor_id / doctor).
    // Não depende do campo opcional "Indicado por", então todos os lançamentos
    // com médico/cirurgião preenchido entram na análise.
    const conversionMap=new Map();
    const ensureConversionDoctor=doctor=>{
      const doctorName=doctor||'Não informado';
      if(!conversionMap.has(doctorName)) conversionMap.set(doctorName,{
        doctor:doctorName,
        examTotal:0,examDone:0,examScheduled:0,examPending:0,examCanceled:0,
        surgeryTotal:0,surgeryDone:0,surgeryScheduled:0,surgeryAuthorized:0,
        surgeryPending:0,surgeryRequested:0,surgeryCanceled:0
      });
      return conversionMap.get(doctorName);
    };

    dashboardExams.forEach(record=>{
      const row=ensureConversionDoctor(record.doctor);
      const weight=Array.isArray(record.exam_type_ids)?record.exam_type_ids.length:0;
      row.examTotal+=weight;
      if(record.status==='Realizado') row.examDone+=weight;
      else if(record.status==='Agendado') row.examScheduled+=weight;
      else if(record.status==='Pendente') row.examPending+=weight;
      else if(record.status==='Cancelado') row.examCanceled+=weight;
    });

    dashboardSurgeries.forEach(record=>{
      const row=ensureConversionDoctor(record.doctor);
      const weight=Array.isArray(record.procedure_ids)?record.procedure_ids.length:0;
      row.surgeryTotal+=weight;
      if(record.status==='Realizada') row.surgeryDone+=weight;
      else if(record.status==='Agendada') row.surgeryScheduled+=weight;
      else if(record.status==='Autorizada') row.surgeryAuthorized+=weight;
      else if(record.status==='Pendente') row.surgeryPending+=weight;
      else if(record.status==='Solicitação') row.surgeryRequested+=weight;
      else if(record.status==='Cancelada') row.surgeryCanceled+=weight;
    });

    const doctorConversion=[...conversionMap.values()]
      .map(row=>({
        ...row,
        examRate:percentOf(row.examDone,row.examTotal),
        surgeryRate:percentOf(row.surgeryDone,row.surgeryTotal),
        totalActivity:row.examTotal+row.surgeryTotal
      }))
      .sort((a,b)=>b.totalActivity-a.totalActivity || String(a.doctor).localeCompare(String(b.doctor),'pt-BR'));

    return{
      baseMonth,previousMonth,months,
      monthLabels:months.map(shortMonthLabel),
      examTrend,surgeryTrend,examConversion,surgeryConversion,
      examByStatusType,surgeryByStatusType,examStatusTotals,surgeryStatusTotals,
      examStatusPie,surgeryStatusPie,
      examByDoctor,surgeryByDoctor,examByIndication,surgeryByIndication,doctorConversion,
      currentExamTotal,previousExamTotal,currentSurgeryTotal,previousSurgeryTotal,
      currentExamRate:percentOf(currentExamDone,currentExamTotal),
      previousExamRate:percentOf(previousExamDone,previousExamTotal),
      currentSurgeryRate:percentOf(currentSurgeryDone,currentSurgeryTotal),
      previousSurgeryRate:percentOf(previousSurgeryDone,previousSurgeryTotal),
      examVolumeDelta:variationPct(currentExamTotal,previousExamTotal),
      surgeryVolumeDelta:variationPct(currentSurgeryTotal,previousSurgeryTotal),
      examRateDelta:percentOf(currentExamDone,currentExamTotal)-percentOf(previousExamDone,previousExamTotal),
      surgeryRateDelta:percentOf(currentSurgeryDone,currentSurgeryTotal)-percentOf(previousSurgeryDone,previousSurgeryTotal)
    };
  },[records,surgeries,dashboardExams,dashboardSurgeries,dashMonth,dashDoctor,dashUser]);

  const exportExcel = async ()=>{
    const month = dashMonth || 'todos-periodos';
    const periodLabel = monthLabel(dashMonth);
    const selectedDoctor = dashDoctor ? doctors.find(d=>d.id===dashDoctor)?.name || '-' : 'Todos';
    const selectedUser = dashUser || 'Todos';
    const selectedStatus = dashStatus || 'Todos';

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Oftalmocastro';
    workbook.lastModifiedBy = currentAccessName || 'Sistema';
    workbook.created = new Date();
    workbook.modified = new Date();

    const COLORS = {
      primary: '315F72',
      primary2: '4D8191',
      primarySoft: 'EAF2F5',
      text: '20313B',
      muted: '74838C',
      white: 'FFFFFF',
      green: '4F8A6B',
      greenSoft: 'EDF6F0',
      yellow: 'B9802B',
      yellowSoft: 'FFF5DF',
      red: 'B85C5C',
      redSoft: 'FBEEEE',
      border: 'DDE6EA',
      soft: 'F7FAFB'
    };

    const border = {
      top:{style:'thin',color:{argb:COLORS.border}},
      left:{style:'thin',color:{argb:COLORS.border}},
      bottom:{style:'thin',color:{argb:COLORS.border}},
      right:{style:'thin',color:{argb:COLORS.border}}
    };

    const styleTitle = (cell, size=20) => {
      cell.font = {name:'Aptos Display',size,bold:true,color:{argb:COLORS.white}};
      cell.fill = {type:'pattern',pattern:'solid',fgColor:{argb:COLORS.primary}};
      cell.alignment = {vertical:'middle',horizontal:'left'};
    };

    const styleSection = (cell) => {
      cell.font = {name:'Aptos',size:11,bold:true,color:{argb:COLORS.primary}};
      cell.fill = {type:'pattern',pattern:'solid',fgColor:{argb:COLORS.primarySoft}};
      cell.alignment = {vertical:'middle'};
      cell.border = border;
    };

    const statusStyle = status => {
      const s = String(status || '').toLowerCase();
      if(s.includes('realiz')) return {fill:COLORS.greenSoft,font:COLORS.green};
      if(s.includes('cancel')) return {fill:COLORS.redSoft,font:COLORS.red};
      if(s.includes('autoriz') || s.includes('pendent')) return {fill:COLORS.yellowSoft,font:COLORS.yellow};
      return {fill:COLORS.primarySoft,font:COLORS.primary};
    };

    const styleStatusCells = (ws, statusColumn, startRow, endRow) => {
      for(let r=startRow;r<=endRow;r++){
        const cell = ws.getCell(r,statusColumn);
        if(!cell.value) continue;
        const st = statusStyle(cell.value);
        cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:st.fill}};
        cell.font={name:'Aptos',size:10,bold:true,color:{argb:st.font}};
        cell.alignment={horizontal:'center',vertical:'middle'};
        cell.border=border;
      }
    };

    const downloadWorkbook = async()=>{
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer],{
        type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Oftalmocastro_Relatorio_${month}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(()=>URL.revokeObjectURL(url),1000);
    };

    const makeBarChartImage = async(title, data, width=900, height=420) => {
      const entries = Object.entries(data)
        .sort((a,b)=>b[1]-a[1])
        .slice(0,10);

      const canvas=document.createElement('canvas');
      canvas.width=width;
      canvas.height=height;
      const ctx=canvas.getContext('2d');

      ctx.fillStyle='#FFFFFF';
      ctx.fillRect(0,0,width,height);

      ctx.fillStyle='#20313B';
      ctx.font='700 26px Arial';
      ctx.fillText(title,34,42);

      if(!entries.length){
        ctx.fillStyle='#74838C';
        ctx.font='18px Arial';
        ctx.fillText('Sem dados para o período selecionado.',34,95);
        return canvas.toDataURL('image/png');
      }

      const max=Math.max(...entries.map(([,v])=>v),1);
      const top=75;
      const left=250;
      const right=70;
      const rowH=Math.min(31,(height-top-30)/entries.length);
      const barMax=width-left-right;

      ctx.font='14px Arial';
      entries.forEach(([label,value],i)=>{
        const y=top+i*rowH;
        const safeLabel=String(label).length>27 ? `${String(label).slice(0,26)}…` : String(label);

        ctx.fillStyle='#53636E';
        ctx.textAlign='right';
        ctx.fillText(safeLabel,left-14,y+16);

        ctx.fillStyle='#EAF2F5';
        ctx.fillRect(left,y+3,barMax,17);

        ctx.fillStyle='#315F72';
        ctx.fillRect(left,y+3,Math.max(4,(value/max)*barMax),17);

        ctx.fillStyle='#20313B';
        ctx.textAlign='left';
        ctx.font='700 14px Arial';
        ctx.fillText(String(value),left+Math.max(4,(value/max)*barMax)+8,y+16);
        ctx.font='14px Arial';
      });

      ctx.textAlign='left';
      ctx.strokeStyle='#DDE6EA';
      ctx.strokeRect(0,0,width,height);
      return canvas.toDataURL('image/png');
    };

    // =========================
    // RESUMO EXECUTIVO
    // =========================
    const wsSummary = workbook.addWorksheet('Resultados',{
      views:[{showGridLines:false}]
    });

    wsSummary.columns=[
      {width:4},{width:25},{width:19},{width:19},{width:19},{width:19},{width:19},{width:19},{width:4}
    ];

    wsSummary.mergeCells('B2:H3');
    styleTitle(wsSummary.getCell('B2'),22);
    wsSummary.getCell('B2').value='OFTALMOCASTRO • RELATÓRIO GERENCIAL';
    wsSummary.getRow(2).height=25;
    wsSummary.getRow(3).height=25;

    wsSummary.mergeCells('B4:H4');
    wsSummary.getCell('B4').value=`Período: ${periodLabel}`;
    wsSummary.getCell('B4').font={name:'Aptos',size:11,bold:true,color:{argb:COLORS.primary}};
    wsSummary.getCell('B4').alignment={vertical:'middle'};
    wsSummary.getRow(4).height=22;

    wsSummary.getCell('B6').value='FILTROS APLICADOS';
    wsSummary.mergeCells('B6:H6');
    styleSection(wsSummary.getCell('B6'));
    wsSummary.getRow(6).height=24;

    const filterRows=[
      ['Médico',selectedDoctor],
      ['Responsável',selectedUser],
      ['Status',selectedStatus]
    ];
    filterRows.forEach((row,index)=>{
      const r=7+index;
      wsSummary.getCell(r,2).value=row[0];
      wsSummary.getCell(r,2).font={bold:true,color:{argb:COLORS.muted}};
      wsSummary.getCell(r,3).value=row[1];
      wsSummary.mergeCells(r,3,r,8);
      wsSummary.getCell(r,3).font={color:{argb:COLORS.text}};
      wsSummary.getCell(r,2).border=border;
      wsSummary.getCell(r,3).border=border;
      wsSummary.getCell(r,3).alignment={vertical:'middle'};
      wsSummary.getRow(r).height=22;
    });

    wsSummary.getCell('B11').value='INDICADORES DO PERÍODO';
    wsSummary.mergeCells('B11:H11');
    styleSection(wsSummary.getCell('B11'));
    wsSummary.getRow(11).height=24;

    const kpis=[
      {range:'B13:C15',label:'Pacientes únicos',value:resultSummary.uniquePatients},
      {range:'D13:E15',label:'Total de exames',value:resultSummary.totalEx},
      {range:'F13:G15',label:'Exames realizados',value:resultSummary.realEx},
      {range:'B17:C19',label:'Total de cirurgias',value:resultSummary.totalSu},
      {range:'D17:E19',label:'Cirurgias realizadas',value:resultSummary.realSu},
      {range:'F17:G19',label:'Total de lançamentos',value:resultSummary.totalEx+resultSummary.totalSu}
    ];

    kpis.forEach(k=>{
      wsSummary.mergeCells(k.range);
      const cell=wsSummary.getCell(k.range.split(':')[0]);
      cell.value={richText:[
        {text:`${k.label}
`,font:{name:'Aptos',size:10,bold:true,color:{argb:COLORS.muted}}},
        {text:String(k.value),font:{name:'Aptos Display',size:24,bold:true,color:{argb:COLORS.primary}}}
      ]};
      cell.alignment={vertical:'middle',horizontal:'center',wrapText:true};
      cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:COLORS.soft}};
      cell.border=border;
    });

    wsSummary.getCell('B21').value='EFICIÊNCIA';
    wsSummary.mergeCells('B21:H21');
    styleSection(wsSummary.getCell('B21'));

    const efficiency=[
      ['Taxa de realização dos exames',resultSummary.exRate/100],
      ['Taxa de realização das cirurgias',resultSummary.suRate/100]
    ];
    efficiency.forEach((item,i)=>{
      const r=22+i;
      wsSummary.getCell(r,2).value=item[0];
      wsSummary.mergeCells(r,2,r,5);
      wsSummary.getCell(r,6).value=item[1];
      wsSummary.mergeCells(r,6,r,7);
      wsSummary.getCell(r,6).numFmt='0%';
      wsSummary.getCell(r,6).font={bold:true,color:{argb:COLORS.primary}};
      for(let c=2;c<=7;c++) wsSummary.getCell(r,c).border=border;
    });

    wsSummary.getCell('B26').value='DESTAQUES';
    wsSummary.mergeCells('B26:H26');
    styleSection(wsSummary.getCell('B26'));

    const highlights=[
      ['Exame com maior volume',resultSummary.examTop?.[0]||'Sem dados',resultSummary.examTop?.[1]||0],
      ['Cirurgia com maior volume',resultSummary.surgTop?.[0]||'Sem dados',resultSummary.surgTop?.[1]||0],
      ['Profissional com maior volume',resultSummary.doctorTop?.[0]||'Sem dados',resultSummary.doctorTop?.[1]||0]
    ];
    highlights.forEach((item,i)=>{
      const r=27+i;
      wsSummary.getCell(r,2).value=item[0];
      wsSummary.getCell(r,3).value=item[1];
      wsSummary.mergeCells(r,3,r,6);
      wsSummary.getCell(r,7).value=item[2];
      wsSummary.getCell(r,7).alignment={horizontal:'center'};
      wsSummary.getCell(r,2).font={bold:true,color:{argb:COLORS.muted}};
      wsSummary.getCell(r,3).font={bold:true,color:{argb:COLORS.text}};
      wsSummary.getCell(r,7).font={bold:true,color:{argb:COLORS.primary}};
      for(let c=2;c<=7;c++) wsSummary.getCell(r,c).border=border;
    });

    wsSummary.getCell('B32').value='Lançamentos por responsável';
    wsSummary.mergeCells('B32:H32');
    styleSection(wsSummary.getCell('B32'));

    ['Responsável','Exames','Cirurgias','Total','Participação'].forEach((h,i)=>{
      const c=2+i;
      wsSummary.getCell(33,c).value=h;
      wsSummary.getCell(33,c).font={bold:true,color:{argb:COLORS.white}};
      wsSummary.getCell(33,c).fill={type:'pattern',pattern:'solid',fgColor:{argb:COLORS.primary2}};
      wsSummary.getCell(33,c).alignment={horizontal:'center'};
      wsSummary.getCell(33,c).border=border;
    });

    const totalAll=resultSummary.totalEx+resultSummary.totalSu;
    userRows.forEach((u,i)=>{
      const r=34+i;
      const total=u.exams+u.surgeries;
      const values=[u.name,u.exams,u.surgeries,total,totalAll?total/totalAll:0];
      values.forEach((v,j)=>{
        const cell=wsSummary.getCell(r,2+j);
        cell.value=v;
        cell.border=border;
        cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:i%2===0?'FFFFFF':'F8FAFB'}};
        cell.alignment={horizontal:j===0?'left':'center'};
      });
      wsSummary.getCell(r,6).numFmt='0%';
    });

    wsSummary.pageSetup={
      orientation:'portrait',
      fitToPage:true,
      fitToWidth:1,
      fitToHeight:0,
      margins:{left:.3,right:.3,top:.5,bottom:.5,header:.2,footer:.2}
    };

    // =========================
    // ANÁLISE GRÁFICA
    // =========================
    const wsCharts=workbook.addWorksheet('Análise Gráfica',{
      views:[{showGridLines:false}]
    });
    wsCharts.columns=Array.from({length:12},()=>({width:12}));

    wsCharts.mergeCells('A1:L2');
    styleTitle(wsCharts.getCell('A1'),20);
    wsCharts.getCell('A1').value='ANÁLISE GRÁFICA • OFTALMOCASTRO';

    wsCharts.mergeCells('A3:L3');
    wsCharts.getCell('A3').value=`Período analisado: ${periodLabel} • Médico: ${selectedDoctor} • Responsável: ${selectedUser} • Status: ${selectedStatus}`;
    wsCharts.getCell('A3').font={size:10,color:{argb:COLORS.muted}};
    wsCharts.getCell('A3').alignment={vertical:'middle'};

    const charts=[
      ['Exames por tipo',countManyBy(dashboardExams,'exam_names'),'A5:F18'],
      ['Exames por médico',countBy(dashboardExams,'doctor'),'G5:L18'],
      ['Status dos exames',countBy(dashboardExams,'status'),'A20:F33'],
      ['Exames por responsável',countBy(dashboardExams.map(x=>({...x,launcher:launchedName(x)})),'launcher'),'G20:L33'],
      ['Cirurgias por procedimento',countManyBy(dashboardSurgeries,'procedure_names'),'A35:F48'],
      ['Cirurgias por cirurgião',countBy(dashboardSurgeries,'doctor'),'G35:L48'],
      ['Status das cirurgias',countBy(dashboardSurgeries,'status'),'A50:F63'],
      ['Cirurgias por convênio',countBy(dashboardSurgeries.map(x=>({...x,insuranceLabel:x.insurance||'Particular'})),'insuranceLabel'),'G50:L63']
    ];

    for(const [title,data,range] of charts){
      const png=await makeBarChartImage(title,data);
      const imageId=workbook.addImage({base64:png,extension:'png'});
      const [from,to]=range.split(':');
      wsCharts.addImage(imageId,{tl:{col:wsCharts.getCell(from).col-1,row:wsCharts.getCell(from).row-1},
                                 br:{col:wsCharts.getCell(to).col,row:wsCharts.getCell(to).row}});
    }

    wsCharts.mergeCells('A65:L65');
    wsCharts.getCell('A65').value='Leitura rápida';
    styleSection(wsCharts.getCell('A65'));
    wsCharts.mergeCells('A66:L69');
    wsCharts.getCell('A66').value =
      `Neste período foram registrados ${resultSummary.totalEx} exames e ${resultSummary.totalSu} cirurgias, `+
      `com ${resultSummary.uniquePatients} pacientes únicos. A taxa de realização foi de ${resultSummary.exRate}% nos exames `+
      `e ${resultSummary.suRate}% nas cirurgias. O exame com maior volume foi ${resultSummary.examTop?.[0]||'—'} `+
      `e o procedimento cirúrgico com maior volume foi ${resultSummary.surgTop?.[0]||'—'}.`;
    wsCharts.getCell('A66').alignment={wrapText:true,vertical:'top'};
    wsCharts.getCell('A66').font={size:11,color:{argb:COLORS.text}};
    wsCharts.getCell('A66').fill={type:'pattern',pattern:'solid',fgColor:{argb:COLORS.soft}};
    wsCharts.getCell('A66').border=border;

    // =========================
    // EXAMES
    // =========================
    const wsExams=workbook.addWorksheet('Exames',{
      views:[{state:'frozen',ySplit:3,showGridLines:false}]
    });

    wsExams.mergeCells('A1:H1');
    styleTitle(wsExams.getCell('A1'),18);
    wsExams.getCell('A1').value=`EXAMES • ${periodLabel}`;
    wsExams.mergeCells('A2:H2');
    wsExams.getCell('A2').value=`Filtros: Médico ${selectedDoctor} | Responsável ${selectedUser} | Status ${selectedStatus}`;
    wsExams.getCell('A2').font={size:10,color:{argb:COLORS.muted}};

    const examHeaders=['Data','Paciente','WhatsApp','Médico','Exame','Status','Lançado por','Observação'];
    const examRows=dashboardExams.map(r=>[
      brDate(r.date),r.patient,r.whatsapp,r.doctor,r.exam,r.status,launchedName(r),r.obs
    ]);

    wsExams.addTable({
      name:'TabelaExames',
      ref:'A3',
      headerRow:true,
      totalsRow:false,
      style:{theme:'TableStyleMedium2',showRowStripes:true},
      columns:examHeaders.map(name=>({name})),
      rows:examRows.length?examRows:[['—','Nenhum exame no período','','','','','','']]
    });

    [13,28,18,24,30,18,18,45].forEach((w,i)=>wsExams.getColumn(i+1).width=w);
    wsExams.getColumn(8).alignment={wrapText:true,vertical:'top'};
    styleStatusCells(wsExams,6,4,3+Math.max(1,examRows.length));
    wsExams.autoFilter={from:'A3',to:'H3'};
    wsExams.pageSetup={orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0};

    // =========================
    // CIRURGIAS
    // =========================
    const wsSurgeries=workbook.addWorksheet('Cirurgias',{
      views:[{state:'frozen',ySplit:3,showGridLines:false}]
    });

    wsSurgeries.mergeCells('A1:M1');
    styleTitle(wsSurgeries.getCell('A1'),18);
    wsSurgeries.getCell('A1').value=`CIRURGIAS • ${periodLabel}`;
    wsSurgeries.mergeCells('A2:M2');
    wsSurgeries.getCell('A2').value=`Filtros: Médico ${selectedDoctor} | Responsável ${selectedUser} | Status ${selectedStatus}`;
    wsSurgeries.getCell('A2').font={size:10,color:{argb:COLORS.muted}};

    const surgeryHeaders=[
      'Data','Paciente','WhatsApp','Cirurgião','Procedimento','Olho','Convênio',
      'Status','Horário chegada','Horário cirurgia','Pagamento','Lançado por','Observação'
    ];
    const surgeryRows=dashboardSurgeries.map(r=>[
      brDate(r.date),r.patient,r.whatsapp,r.doctor,r.procedure,r.eye,r.insurance||'Particular',
      r.status,r.arrival,r.time,r.payment,launchedName(r),r.obs
    ]);

    wsSurgeries.addTable({
      name:'TabelaCirurgias',
      ref:'A3',
      headerRow:true,
      totalsRow:false,
      style:{theme:'TableStyleMedium2',showRowStripes:true},
      columns:surgeryHeaders.map(name=>({name})),
      rows:surgeryRows.length?surgeryRows:[['—','Nenhuma cirurgia no período','','','','','','','','','','','']]
    });

    [13,28,18,24,30,18,22,22,18,18,18,18,45].forEach((w,i)=>wsSurgeries.getColumn(i+1).width=w);
    wsSurgeries.getColumn(13).alignment={wrapText:true,vertical:'top'};
    styleStatusCells(wsSurgeries,8,4,3+Math.max(1,surgeryRows.length));
    wsSurgeries.autoFilter={from:'A3',to:'M3'};
    wsSurgeries.pageSetup={orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0};

    // =========================
    // LANÇAMENTOS
    // =========================
    const wsUsers=workbook.addWorksheet('Lançamentos',{
      views:[{showGridLines:false}]
    });
    wsUsers.columns=[{width:24},{width:16},{width:16},{width:16},{width:16}];

    wsUsers.mergeCells('A1:E1');
    styleTitle(wsUsers.getCell('A1'),18);
    wsUsers.getCell('A1').value=`LANÇAMENTOS POR RESPONSÁVEL • ${periodLabel}`;

    const userData=userRows.map(u=>{
      const total=u.exams+u.surgeries;
      return [u.name,u.exams,u.surgeries,total,totalAll?total/totalAll:0];
    });

    wsUsers.addTable({
      name:'TabelaLancamentos',
      ref:'A3',
      headerRow:true,
      totalsRow:true,
      style:{theme:'TableStyleMedium2',showRowStripes:true},
      columns:[
        {name:'Responsável',totalsRowLabel:'TOTAL'},
        {name:'Exames',totalsRowFunction:'sum'},
        {name:'Cirurgias',totalsRowFunction:'sum'},
        {name:'Total',totalsRowFunction:'sum'},
        {name:'Participação',totalsRowFunction:'none'}
      ],
      rows:userData
    });

    for(let r=4;r<4+userData.length;r++) wsUsers.getCell(r,5).numFmt='0%';

    wsUsers.mergeCells('A11:E11');
    wsUsers.getCell('A11').value='Indicador de produtividade';
    styleSection(wsUsers.getCell('A11'));
    wsUsers.mergeCells('A12:E14');
    wsUsers.getCell('A12').value='Use esta aba para acompanhar o volume de registros de Juliana, Luiza, Elisangela e ADM. A participação considera o total de lançamentos do período selecionado.';
    wsUsers.getCell('A12').alignment={wrapText:true,vertical:'top'};
    wsUsers.getCell('A12').fill={type:'pattern',pattern:'solid',fgColor:{argb:COLORS.soft}};
    wsUsers.getCell('A12').border=border;

    // Aba inicial
    workbook.views=[{activeTab:0}];

    await downloadWorkbook();
  };


  const exportResultsPagePDF = async ()=>{
    if(exportingPanelPdf) return;
    setExportingPanelPdf(true);

    try{
      const selectedDoctor=dashDoctor ? doctors.find(d=>d.id===dashDoctor)?.name || 'Não identificado' : 'Todos os médicos';
      const selectedUser=dashUser || 'Todos os responsáveis';
      const selectedStatus=dashStatus || 'Todos os status';
      const periodLabel=monthLabel(dashMonth);
      const generatedAt=new Date().toLocaleString('pt-BR');

      // Exportação nativa em A4: não captura a tela e não corta cards.
      // O PDF contém somente os indicadores exibidos no modelo aprovado.
      const pdf=new jsPDF({orientation:'portrait',unit:'mm',format:'a4',compress:true});
      const pageWidth=210;
      const pageHeight=297;
      const margin=10;
      const contentWidth=pageWidth-(margin*2);
      const footerY=289;

      const C={
        bg:[253,247,243],
        paper:[255,255,255],
        primary:[44,96,151],
        primaryDark:[38,80,120],
        teal:[53,111,128],
        teal2:[104,159,173],
        lime:[201,210,139],
        text:[35,58,68],
        muted:[111,129,137],
        border:[218,229,233],
        soft:[246,250,251],
        blueSoft:[236,245,248],
        green:[73,137,102],
        greenSoft:[236,247,240],
        yellow:[184,128,35],
        yellowSoft:[255,247,228],
        red:[181,82,82],
        redSoft:[252,238,238]
      };

      const safe=value=>String(value??'')
        .replace(/•/g,'-')
        .replace(/[“”]/g,'"')
        .replace(/[‘’]/g,"'")
        .replace(/…/g,'...')
        .replace(/[^\x00-\xFF]/g,'');

      const split=(value,width,size=7)=>{
        pdf.setFontSize(size);
        return pdf.splitTextToSize(safe(value||'—'),width);
      };

      // Mantém textos sempre dentro da área disponível.
      // Reduz levemente a fonte e, se ainda necessário, aplica reticências.
      const fitText=(value,maxWidth,size=6,minSize=4.2)=>{
        const raw=safe(value||'—');
        let fontSize=size;
        pdf.setFontSize(fontSize);
        while(fontSize>minSize && pdf.getTextWidth(raw)>maxWidth){
          fontSize-=0.2;
          pdf.setFontSize(fontSize);
        }
        if(pdf.getTextWidth(raw)<=maxWidth) return {text:raw,size:fontSize};
        let text=raw;
        while(text.length>1 && pdf.getTextWidth(`${text}…`)>maxWidth){
          text=text.slice(0,-1);
        }
        return {text:`${text}…`,size:fontSize};
      };

      const statusPalette=status=>{
        const s=String(status||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
        if(s.includes('realiz')) return {fill:C.greenSoft,text:C.green};
        if(s.includes('cancel')) return {fill:C.redSoft,text:C.red};
        if(s.includes('pendent')) return {fill:C.yellowSoft,text:C.yellow};
        if(s.includes('solicit')) return {fill:[244,244,247],text:[105,106,120]};
        if(s.includes('autoriz')) return {fill:C.blueSoft,text:C.primary};
        if(s.includes('agend')) return {fill:C.blueSoft,text:C.teal};
        return {fill:C.soft,text:C.primary};
      };

      const drawDots=()=>{
        pdf.setFillColor(...C.lime);
        for(let row=0;row<8;row++){
          for(let col=0;col<3;col++){
            pdf.circle(3.5+col*3,64+row*5,0.45,'F');
            pdf.circle(pageWidth-3.5-col*3,64+row*5,0.45,'F');
          }
        }
      };

      const pageHeader=(title,subtitle='')=>{
        pdf.setFillColor(...C.bg);
        pdf.rect(0,0,pageWidth,pageHeight,'F');
        drawDots();

        // Marca fixa no canto esquerdo.
        pdf.setTextColor(...C.primaryDark);
        pdf.setFont('helvetica','bold');
        pdf.setFontSize(11);
        pdf.text('OFTALMOCASTRO',margin,16);
        pdf.setTextColor(...C.muted);
        pdf.setFont('helvetica','normal');
        pdf.setFontSize(6.1);
        pdf.text('GESTÃO DE EXAMES E CIRURGIAS',margin,21);

        // Faixa de título sempre contida no A4, sem estourar ou deformar.
        const boxW=92;
        const boxX=pageWidth-margin-boxW;
        pdf.setFillColor(...C.lime);
        pdf.roundedRect(boxX,8,boxW,23,7,7,'F');

        let titleSize=13.2;
        let titleLines=[];
        while(titleSize>=10.4){
          pdf.setFont('helvetica','bold');
          pdf.setFontSize(titleSize);
          titleLines=pdf.splitTextToSize(safe(title),boxW-12);
          if(titleLines.length<=2) break;
          titleSize-=0.4;
        }
        titleLines=titleLines.slice(0,2);
        pdf.setTextColor(...C.primary);
        pdf.setFont('helvetica','bold');
        pdf.setFontSize(titleSize);
        const titleStartY=titleLines.length===1?20:15.5;
        pdf.text(titleLines,boxX+boxW-6,titleStartY,{align:'right',lineHeightFactor:1.0});

        pdf.setFillColor(...C.primary);
        pdf.rect(boxX,27,boxW,8,'F');
        pdf.setTextColor(255,255,255);
        pdf.setFont('helvetica','bold');
        pdf.setFontSize(6.6);
        pdf.text('OFTALMOCASTRO',boxX+boxW-6,32.3,{align:'right'});

        if(subtitle){
          pdf.setTextColor(...C.muted);
          pdf.setFont('helvetica','normal');
          pdf.setFontSize(6.0);
          const subtitleLines=split(subtitle,contentWidth,6.0).slice(0,2);
          pdf.text(subtitleLines,margin,40.8,{lineHeightFactor:1.05});
        }
      };

      const footer=(page,total)=>{
        pdf.setDrawColor(...C.border);
        pdf.line(margin,footerY-2,pageWidth-margin,footerY-2);
        pdf.setTextColor(...C.muted);
        pdf.setFont('helvetica','normal');
        pdf.setFontSize(5.4);
        pdf.text(`Período: ${safe(periodLabel)}  |  Gerado em ${safe(generatedAt)}`,margin,footerY+1.5);
        pdf.text(`Página ${page} de ${total}`,pageWidth-margin,footerY+1.5,{align:'right'});
      };

      const drawMetricCard=(x,y,w,h,label,value,detail='',accent='blue')=>{
        const accentColor=accent==='green'?C.green:accent==='yellow'?C.yellow:C.primary;
        const fill=accent==='green'?C.greenSoft:accent==='yellow'?C.yellowSoft:C.paper;
        pdf.setFillColor(...fill);
        pdf.setDrawColor(...C.border);
        pdf.roundedRect(x,y,w,h,3,3,'FD');

        // Pequena faixa lateral ajuda a leitura e mantém todos os cards visualmente alinhados.
        pdf.setFillColor(...accentColor);
        pdf.roundedRect(x+1.4,y+3,1.2,h-6,0.6,0.6,'F');

        pdf.setTextColor(...C.muted);
        pdf.setFont('helvetica','bold');
        const labelFit=fitText(String(label).toUpperCase(),w-10,5.0,4.1);
        pdf.setFontSize(labelFit.size);
        pdf.text(labelFit.text,x+4.2,y+5.7);

        pdf.setTextColor(...accentColor);
        pdf.setFont('helvetica','bold');
        const valueFit=fitText(value,w-10,14.8,9.2);
        pdf.setFontSize(valueFit.size);
        pdf.text(valueFit.text,x+4.2,y+14.3);

        if(detail){
          pdf.setTextColor(...C.muted);
          pdf.setFont('helvetica','normal');
          pdf.setFontSize(4.8);
          const lines=split(detail,w-9,4.8).slice(0,2);
          pdf.text(lines,x+4.2,y+18.8,{lineHeightFactor:1.05});
        }
      };

      const drawStatusCard=(x,y,w,h,title,rate,items,openText)=>{
        pdf.setFillColor(...C.paper);
        pdf.setDrawColor(...C.border);
        pdf.roundedRect(x,y,w,h,3,3,'FD');

        pdf.setTextColor(...C.primary);
        pdf.setFont('helvetica','bold');
        pdf.setFontSize(5.0);
        pdf.text(safe(title.split(' ')[0]).toUpperCase(),x+4,y+5.4);

        const titleFit=fitText(title,w-29,8.4,6.2);
        pdf.setTextColor(...C.text);
        pdf.setFont('helvetica','bold');
        pdf.setFontSize(titleFit.size);
        pdf.text(titleFit.text,x+4,y+11.2);

        pdf.setFillColor(...C.greenSoft);
        pdf.setDrawColor(215,233,221);
        pdf.roundedRect(x+w-20,y+3.5,16,11.5,2,2,'FD');
        pdf.setTextColor(...C.green);
        pdf.setFont('helvetica','bold');
        const rateFit=fitText(`${rate}%`,13,9.0,7.0);
        pdf.setFontSize(rateFit.size);
        pdf.text(rateFit.text,x+w-12,y+8.7,{align:'center'});
        pdf.setFont('helvetica','normal');
        pdf.setFontSize(3.9);
        pdf.text('REALIZAÇÃO',x+w-12,y+12.1,{align:'center'});

        // Reserva uma faixa exclusiva para a observação inferior. Assim os status
        // nunca encostam ou passam por cima do texto de "em aberto".
        const noteH=openText?10:3;
        const bodyTop=y+18;
        const bodyBottom=y+h-noteH-2;
        const bodyH=Math.max(12,bodyBottom-bodyTop);
        const cols=items.length>4?3:2;
        const rows=Math.ceil(items.length/cols);
        const gapX=2;
        const gapY=2;
        const innerW=w-8;
        const cellW=(innerW-gapX*(cols-1))/cols;
        const cellH=(bodyH-gapY*(rows-1))/Math.max(rows,1);

        items.forEach((item,i)=>{
          const row=Math.floor(i/cols);
          const col=i%cols;
          const cx=x+4+col*(cellW+gapX);
          const cy=bodyTop+row*(cellH+gapY);
          const pal=statusPalette(item.label);
          pdf.setFillColor(...pal.fill);
          pdf.setDrawColor(...C.border);
          pdf.roundedRect(cx,cy,cellW,cellH,1.8,1.8,'FD');

          const labelFit=fitText(String(item.label).toUpperCase(),cellW-5,4.6,3.8);
          pdf.setTextColor(...pal.text);
          pdf.setFont('helvetica','bold');
          pdf.setFontSize(labelFit.size);
          pdf.text(labelFit.text,cx+2.4,cy+4.5);

          pdf.setFontSize(Math.max(8.2,Math.min(10.2,cellH*.62)));
          pdf.text(String(item.value),cx+2.4,cy+10.3);

          if(item.detail && cellH>=13){
            const detailFit=fitText(item.detail,cellW-5,4.0,3.5);
            pdf.setTextColor(...C.muted);
            pdf.setFont('helvetica','normal');
            pdf.setFontSize(detailFit.size);
            pdf.text(detailFit.text,cx+2.4,cy+cellH-2.4);
          }
        });

        if(openText){
          const noteY=y+h-noteH;
          pdf.setDrawColor(...C.border);
          pdf.line(x+4,noteY,x+w-4,noteY);
          pdf.setTextColor(...C.muted);
          pdf.setFont('helvetica','normal');
          const noteFit=fitText(openText,w-8,4.8,4.0);
          pdf.setFontSize(noteFit.size);
          pdf.text(noteFit.text,x+4,noteY+5.5);
        }
      };

      const makeDonut=(data,centerLabel)=>{
        const canvas=document.createElement('canvas');
        canvas.width=720;
        canvas.height=420;
        const ctx=canvas.getContext('2d');
        ctx.fillStyle='#FFFFFF';
        ctx.fillRect(0,0,canvas.width,canvas.height);

        const entries=Object.entries(data||{}).filter(([,v])=>Number(v)>0).slice(0,6);
        const total=entries.reduce((s,[,v])=>s+Number(v||0),0);
        const colors=['#4F8966','#386C7D','#C48926','#B85C5C','#7A6F9B','#6B8F9B'];

        if(!total){
          ctx.fillStyle='#74838C';
          ctx.font='18px Arial';
          ctx.textAlign='center';
          ctx.fillText('Sem dados para o período.',360,210);
          return canvas.toDataURL('image/png');
        }

        const cx=168,cy=210,r=108,inner=68;
        let start=-Math.PI/2;
        entries.forEach(([,value],i)=>{
          const angle=(Number(value)/total)*Math.PI*2;
          ctx.beginPath();
          ctx.moveTo(cx,cy);
          ctx.arc(cx,cy,r,start,start+angle);
          ctx.closePath();
          ctx.fillStyle=colors[i%colors.length];
          ctx.fill();
          start+=angle;
        });

        ctx.beginPath();
        ctx.arc(cx,cy,inner,0,Math.PI*2);
        ctx.fillStyle='#fff';
        ctx.fill();
        ctx.fillStyle='#20313B';
        ctx.font='700 40px Arial';
        ctx.textAlign='center';
        ctx.fillText(String(total),cx,cy-4);
        ctx.fillStyle='#74838C';
        ctx.font='14px Arial';
        ctx.fillText(centerLabel,cx,cy+23);

        const legendX=350;
        const rightX=685;
        const startY=74;
        const rowStep=entries.length<=4?64:52;
        entries.forEach(([label,value],i)=>{
          const ly=startY+i*rowStep;
          ctx.fillStyle=colors[i%colors.length];
          ctx.beginPath();
          ctx.arc(legendX,ly,7,0,Math.PI*2);
          ctx.fill();

          const raw=String(label);
          const legendLabel=raw.length>23?`${raw.slice(0,22)}…`:raw;
          ctx.fillStyle='#53636E';
          ctx.font='15px Arial';
          ctx.textAlign='left';
          ctx.fillText(legendLabel,legendX+19,ly+5);

          ctx.fillStyle='#20313B';
          ctx.font='700 15px Arial';
          ctx.textAlign='right';
          ctx.fillText(String(value),rightX-50,ly+5);
          ctx.fillStyle='#7A8A90';
          ctx.font='13px Arial';
          ctx.fillText(`${percentOf(value,total)}%`,rightX,ly+5);
        });

        return canvas.toDataURL('image/png');
      };

      const drawDonutCard=(x,y,w,h,title,data,centerLabel)=>{
        pdf.setFillColor(...C.paper);
        pdf.setDrawColor(...C.border);
        pdf.roundedRect(x,y,w,h,3,3,'FD');

        const titleFit=fitText(title,w-8,7.1,5.6);
        pdf.setTextColor(...C.text);
        pdf.setFont('helvetica','bold');
        pdf.setFontSize(titleFit.size);
        pdf.text(titleFit.text,x+4,y+6.6);

        const img=makeDonut(data,centerLabel);
        const sourceRatio=720/420;
        const availableW=w-8;
        const availableH=h-15;
        let imgW=availableW;
        let imgH=imgW/sourceRatio;
        if(imgH>availableH){
          imgH=availableH;
          imgW=imgH*sourceRatio;
        }
        const imgX=x+(w-imgW)/2;
        const imgY=y+11+(availableH-imgH)/2;
        // Mantém a proporção original do gráfico. Nada fica esticado ou "torto".
        pdf.addImage(img,'PNG',imgX,imgY,imgW,imgH,undefined,'FAST');
      };

      const drawHighlight=(x,y,w,h,label,name,value)=>{
        pdf.setFillColor(...C.paper);
        pdf.setDrawColor(...C.border);
        pdf.roundedRect(x,y,w,h,3,3,'FD');
        pdf.setFillColor(...C.teal2);
        pdf.roundedRect(x+1.5,y+3,1.2,h-6,0.6,0.6,'F');

        pdf.setTextColor(...C.muted);
        pdf.setFont('helvetica','bold');
        const labelFit=fitText(String(label).toUpperCase(),w-10,4.6,3.9);
        pdf.setFontSize(labelFit.size);
        pdf.text(labelFit.text,x+4.2,y+5.4);

        pdf.setTextColor(...C.text);
        pdf.setFont('helvetica','bold');
        const nameFit=fitText(name||'Sem dados',w-10,7.1,5.2);
        pdf.setFontSize(nameFit.size);
        pdf.text(nameFit.text,x+4.2,y+11.5);

        const valueText=`${value||0} ${Number(value)===1?'registro':'registros'}`;
        const valueFit=fitText(valueText,w-10,5.3,4.3);
        pdf.setTextColor(...C.primary);
        pdf.setFont('helvetica','bold');
        pdf.setFontSize(valueFit.size);
        pdf.text(valueFit.text,x+4.2,y+17.5);
      };

      const drawBarRow=(x,y,w,label,value,max,total=null,fontSize=5.2,rowH=4.3)=>{
        // Layout em uma única linha: nome | barra | quantidade/percentual.
        // Evita que percentuais ou nomes invadam a linha seguinte.
        const labelW=w*0.48;
        const valueW=total ? 18 : 10;
        const gap=3;
        const barX=x+labelW+gap;
        const barW=Math.max(8,w-labelW-valueW-gap*2);
        const textY=y+rowH*0.67;

        pdf.setTextColor(...C.text);
        pdf.setFont('helvetica','normal');
        const labelFit=fitText(label,labelW-2,fontSize,3.8);
        pdf.setFontSize(labelFit.size);
        pdf.text(labelFit.text,x,textY);

        const barY=y+Math.max(0.7,rowH*0.37);
        pdf.setFillColor(232,239,241);
        pdf.roundedRect(barX,barY,barW,2.2,1,1,'F');
        if(Number(value)>0){
          pdf.setFillColor(...C.teal);
          pdf.roundedRect(barX,barY,Math.max(1,(Number(value)/Math.max(Number(max)||1,1))*barW),2.2,1,1,'F');
        }

        const valueText=total ? `${value} · ${percentOf(value,total)}%` : String(value);
        pdf.setTextColor(...C.primaryDark);
        pdf.setFont('helvetica','bold');
        const valueFit=fitText(valueText,valueW,fontSize,3.8);
        pdf.setFontSize(valueFit.size);
        pdf.text(valueFit.text,x+w,textY,{align:'right'});
      };

      // Compacta rankings sem perder o total: quando houver muitas linhas,
      // mantém as principais e soma o restante em "Outros".
      // Isso evita que conteúdo ultrapasse os cards no A4.
      const compactEntries=(entries,maxRows=5,otherLabel='Outros')=>{
        const ordered=[...(entries||[])].sort((a,b)=>Number(b[1]||0)-Number(a[1]||0));
        if(ordered.length<=maxRows) return ordered;
        const keep=Math.max(1,maxRows-1);
        const visible=ordered.slice(0,keep);
        const remainder=ordered.slice(keep);
        const remainderTotal=remainder.reduce((sum,[,value])=>sum+Number(value||0),0);
        return [...visible,[`${otherLabel} (${remainder.length})`,remainderTotal]];
      };

      const compactDoctorRows=(rows,maxDoctors=8,maxItems=5)=>{
        const ordered=[...(rows||[])].sort((a,b)=>Number(b.total||0)-Number(a.total||0));
        const normalize=row=>({
          ...row,
          items:compactEntries(row.items||[],maxItems,'Outros')
        });
        if(ordered.length<=maxDoctors) return ordered.map(normalize);

        const keep=Math.max(1,maxDoctors-1);
        const visible=ordered.slice(0,keep).map(normalize);
        const remainder=ordered.slice(keep);
        const itemTotals={};
        let total=0;
        remainder.forEach(row=>{
          total+=Number(row.total||0);
          (row.items||[]).forEach(([name,value])=>{
            itemTotals[name]=(itemTotals[name]||0)+Number(value||0);
          });
        });
        visible.push({
          doctor:`Outros médicos (${remainder.length})`,
          total,
          items:compactEntries(Object.entries(itemTotals),maxItems,'Outros')
        });
        return visible;
      };

      const doctorColumnHeight=(rows,rowH=3.7,headH=6.7,gap=1.4)=>
        10+(rows||[]).reduce((sum,row)=>sum+headH+Math.max(1,row.items?.length||0)*rowH+gap,0);

      const drawDoctorColumn=(x,y,w,h,title,rows,itemLabel)=>{
        pdf.setFillColor(...C.paper);pdf.setDrawColor(...C.border);pdf.roundedRect(x,y,w,h,3,3,'FD');
        pdf.setTextColor(...C.primary);pdf.setFont('helvetica','bold');
        const titleFit=fitText(title,w-8,7.5,5.8);
        pdf.setFontSize(titleFit.size);pdf.text(titleFit.text,x+4,y+7);

        const list=compactDoctorRows(rows,8,5);
        if(!list.length){
          pdf.setTextColor(...C.muted);pdf.setFont('helvetica','normal');pdf.setFontSize(6);pdf.text('Sem dados para os filtros atuais.',x+4,y+16);return;
        }

        const available=h-15;
        const raw=doctorColumnHeight(list)-10;
        const scale=Math.min(1,available/Math.max(raw,1));
        const rowH=Math.max(3.15,3.9*scale);
        const headH=Math.max(5.7,7.0*scale);
        const gap=Math.max(1.0,1.5*scale);
        const font=Math.max(4.2,5.0*scale);
        let cy=y+11;

        list.forEach(row=>{
          const items=row.items||[];
          const groupH=headH+Math.max(1,items.length)*rowH+gap;
          pdf.setFillColor(...C.soft);pdf.setDrawColor(...C.border);pdf.roundedRect(x+3,cy,w-6,groupH-gap,2,2,'FD');

          const doctorName=fitText(row.doctor,w-27,font+0.7,4.2);
          pdf.setTextColor(...C.text);pdf.setFont('helvetica','bold');pdf.setFontSize(doctorName.size);
          pdf.text(doctorName.text,x+6,cy+headH*.64);
          pdf.setTextColor(...C.primary);pdf.setFontSize(font+1.1);
          pdf.text(String(row.total),x+w-6,cy+headH*.64,{align:'right'});

          let iy=cy+headH;
          if(items.length){
            const max=Math.max(...items.map(([,v])=>Number(v)||0),1);
            items.forEach(([name,value])=>{
              drawBarRow(x+6,iy,w-12,name,value,max,row.total,font,rowH);
              iy+=rowH;
            });
          }else{
            pdf.setTextColor(...C.muted);pdf.setFont('helvetica','normal');pdf.setFontSize(font);
            pdf.text(`Sem ${itemLabel.toLowerCase()}`,x+6,iy+2.5);
          }
          cy+=groupH;
        });
      };

      const statusColumnHeight=(statuses,byStatus,rowH=3.65,headH=6.6,gap=1.35)=>
        10+statuses.reduce((sum,status)=>{
          const compact=compactEntries(Object.entries(byStatus?.[status]||{}),5,'Outros');
          return sum+headH+Math.max(1,compact.length)*rowH+gap;
        },0);

      const drawStatusColumn=(x,y,w,h,title,statuses,byStatus,totalByStatus,itemLabel)=>{
        pdf.setFillColor(...C.paper);pdf.setDrawColor(...C.border);pdf.roundedRect(x,y,w,h,3,3,'FD');
        pdf.setTextColor(...C.primary);pdf.setFont('helvetica','bold');
        const titleFit=fitText(title,w-8,7.5,5.8);
        pdf.setFontSize(titleFit.size);pdf.text(titleFit.text,x+4,y+7);

        const available=h-15;
        const raw=statusColumnHeight(statuses,byStatus)-10;
        const scale=Math.min(1,available/Math.max(raw,1));
        const rowH=Math.max(3.1,3.85*scale);
        const headH=Math.max(5.6,6.9*scale);
        const gap=Math.max(1.0,1.45*scale);
        const font=Math.max(4.15,4.95*scale);
        let cy=y+11;

        statuses.forEach(status=>{
          const entries=compactEntries(Object.entries(byStatus?.[status]||{}),5,'Outros');
          const total=totalByStatus?.[status]||0;
          const groupH=headH+Math.max(1,entries.length)*rowH+gap;
          const pal=statusPalette(status);
          pdf.setFillColor(...pal.fill);pdf.setDrawColor(...C.border);pdf.roundedRect(x+3,cy,w-6,groupH-gap,2,2,'FD');

          const statusFit=fitText(status,w-27,font+0.7,4.2);
          pdf.setTextColor(...pal.text);pdf.setFont('helvetica','bold');pdf.setFontSize(statusFit.size);
          pdf.text(statusFit.text,x+6,cy+headH*.64);
          pdf.setFontSize(font+1.1);pdf.text(String(total),x+w-6,cy+headH*.64,{align:'right'});

          let iy=cy+headH;
          if(entries.length){
            const max=Math.max(...entries.map(([,v])=>Number(v)||0),1);
            entries.forEach(([name,value])=>{
              drawBarRow(x+6,iy,w-12,name,value,max,total,font,rowH);
              iy+=rowH;
            });
          }else{
            pdf.setTextColor(...C.muted);pdf.setFont('helvetica','normal');pdf.setFontSize(font);
            pdf.text(`Nenhum ${itemLabel.toLowerCase()}`,x+6,iy+2.5);
          }
          cy+=groupH;
        });
      };

      const drawRankingBox=(x,y,w,h,title,data,maxItems=6)=>{
        const source=Array.isArray(data)?data:Object.entries(data||{});
        const entries=compactEntries(source,maxItems,'Outros');
        pdf.setFillColor(...C.paper);pdf.setDrawColor(...C.border);pdf.roundedRect(x,y,w,h,3,3,'FD');

        pdf.setTextColor(...C.text);
        pdf.setFont('helvetica','bold');
        const titleFit=fitText(title,w-8,6.7,5.2);
        pdf.setFontSize(titleFit.size);
        pdf.text(titleFit.text,x+4,y+6);

        if(!entries.length){
          pdf.setTextColor(...C.muted);pdf.setFont('helvetica','normal');pdf.setFontSize(5.2);pdf.text('Sem dados.',x+4,y+14);return;
        }

        const usableH=h-13;
        const rowH=Math.min(5.0,usableH/entries.length);
        const font=Math.max(4.4,Math.min(5.2,rowH+0.3));
        const max=Math.max(...entries.map(([,v])=>Number(v)||0),1);
        let ry=y+10;
        entries.forEach(([label,value],index)=>{
          if(index>0){
            pdf.setDrawColor(238,242,243);
            pdf.line(x+4,ry-0.25,x+w-4,ry-0.25);
          }
          drawBarRow(x+4,ry,w-8,label,value,max,null,font,rowH);
          ry+=rowH;
        });
      };

      const drawResponsibleTable=(x,y,w,h)=>{
        const totalAll=resultSummary.totalEx+resultSummary.totalSu;
        pdf.setFillColor(...C.paper);
        pdf.setDrawColor(...C.border);
        pdf.roundedRect(x,y,w,h,3,3,'FD');

        pdf.setTextColor(...C.text);
        pdf.setFont('helvetica','bold');
        pdf.setFontSize(7);
        pdf.text('Lançamentos por responsável',x+4,y+6);
        pdf.setTextColor(...C.muted);
        pdf.setFont('helvetica','normal');
        pdf.setFontSize(4.7);
        pdf.text('Exames, cirurgias, total e participação.',x+4,y+10.2);

        const cols=[0.30,0.15,0.15,0.15,0.25];
        const headers=['Responsável','Exames','Cirurgias','Total','Participação'];
        const innerX=x+4;
        const innerW=w-8;
        const headerY=y+13;
        const headerH=7;
        const tableBottom=y+h-3;

        pdf.setFillColor(...C.blueSoft);
        pdf.rect(innerX,headerY,innerW,headerH,'F');

        let cx=innerX;
        headers.forEach((head,i)=>{
          const cw=innerW*cols[i];
          pdf.setTextColor(...C.primaryDark);
          pdf.setFont('helvetica','bold');
          const headFit=fitText(String(head).toUpperCase(),cw-4,4.5,3.8);
          pdf.setFontSize(headFit.size);
          pdf.text(headFit.text,i===0?cx+2:cx+cw/2,headerY+4.6,{align:i===0?'left':'center'});
          if(i>0){
            pdf.setDrawColor(...C.border);
            pdf.line(cx,headerY,cx,tableBottom);
          }
          cx+=cw;
        });

        const bodyTop=headerY+headerH;
        const bodyH=Math.max(8,tableBottom-bodyTop);
        const rowH=bodyH/Math.max(userRows.length,1);

        userRows.forEach((u,idx)=>{
          const ry=bodyTop+idx*rowH;
          if(idx%2===0){
            pdf.setFillColor(...C.soft);
            pdf.rect(innerX,ry,innerW,rowH,'F');
          }
          if(idx>0){
            pdf.setDrawColor(...C.border);
            pdf.line(innerX,ry,innerX+innerW,ry);
          }

          const values=[
            u.name,
            u.exams,
            u.surgeries,
            u.exams+u.surgeries,
            totalAll?`${Math.round(((u.exams+u.surgeries)/totalAll)*100)}%`:'0%'
          ];

          let px=innerX;
          values.forEach((value,i)=>{
            const cw=innerW*cols[i];
            pdf.setTextColor(...(i===0?C.text:C.primaryDark));
            pdf.setFont('helvetica',i===0||i===3?'bold':'normal');
            const cellFit=fitText(value,cw-4,5.0,4.0);
            pdf.setFontSize(cellFit.size);
            const baseline=ry+rowH/2+cellFit.size*0.18;
            pdf.text(cellFit.text,i===0?px+2:px+cw/2,baseline,{align:i===0?'left':'center'});
            px+=cw;
          });
        });
      };

      const drawDoctorConversionTable=(x,y,w,h,title,rows,kind)=>{
        pdf.setFillColor(...C.paper);pdf.setDrawColor(...C.border);pdf.roundedRect(x,y,w,h,3,3,'FD');
        pdf.setTextColor(...C.primary);pdf.setFont('helvetica','bold');
        const titleFit=fitText(title,w-8,8.2,6.3);pdf.setFontSize(titleFit.size);pdf.text(titleFit.text,x+4,y+7);

        const source=(rows||[]).filter(row=>kind==='exam'?row.examTotal>0:row.surgeryTotal>0);
        if(!source.length){
          pdf.setTextColor(...C.muted);pdf.setFont('helvetica','normal');pdf.setFontSize(5.4);
          pdf.text('Sem indicações registradas para os filtros atuais.',x+4,y+17);
          return;
        }

        const isExam=kind==='exam';
        const columns=isExam
          ? [
              {label:'Médico',key:'doctor',share:.28,align:'left'},
              {label:'Total',key:'examTotal',share:.12},
              {label:'Realizados',key:'examDone',share:.13},
              {label:'Agendados',key:'examScheduled',share:.13},
              {label:'Pendentes',key:'examPending',share:.12},
              {label:'Cancelados',key:'examCanceled',share:.11},
              {label:'Conversão',key:'examRate',share:.11,rate:true}
            ]
          : [
              {label:'Médico',key:'doctor',share:.23,align:'left'},
              {label:'Total',key:'surgeryTotal',share:.10},
              {label:'Realizadas',key:'surgeryDone',share:.10},
              {label:'Agendadas',key:'surgeryScheduled',share:.10},
              {label:'Autorizadas',key:'surgeryAuthorized',share:.11},
              {label:'Pendentes',key:'surgeryPending',share:.10},
              {label:'Solicitações',key:'surgeryRequested',share:.11},
              {label:'Canceladas',key:'surgeryCanceled',share:.08},
              {label:'Conversão',key:'surgeryRate',share:.07,rate:true}
            ];

        const innerX=x+3,innerW=w-6;
        const headY=y+11,headH=9;
        const bodyBottom=y+h-4;
        pdf.setFillColor(...C.blueSoft);pdf.rect(innerX,headY,innerW,headH,'F');

        let cx=innerX;
        columns.forEach((col,index)=>{
          const cw=innerW*col.share;
          if(index>0){pdf.setDrawColor(...C.border);pdf.line(cx,headY,cx,bodyBottom);}
          pdf.setTextColor(...C.primaryDark);pdf.setFont('helvetica','bold');
          const fit=fitText(col.label,cw-3,4.6,3.5);pdf.setFontSize(fit.size);
          pdf.text(fit.text,col.align==='left'?cx+1.6:cx+cw/2,headY+5.8,{align:col.align==='left'?'left':'center'});
          cx+=cw;
        });

        const bodyTop=headY+headH;
        const rowH=Math.min(12,(bodyBottom-bodyTop)/Math.max(source.length,1));
        source.forEach((row,index)=>{
          const ry=bodyTop+index*rowH;
          if(index%2===0){pdf.setFillColor(...C.soft);pdf.rect(innerX,ry,innerW,rowH,'F');}
          if(index>0){pdf.setDrawColor(...C.border);pdf.line(innerX,ry,innerX+innerW,ry);}
          let px=innerX;
          columns.forEach(col=>{
            const cw=innerW*col.share;
            let value=row[col.key]??0;
            if(col.rate) value=`${value}%`;
            pdf.setTextColor(...(col.rate?C.green:(col.key==='doctor'?C.text:C.primaryDark)));
            pdf.setFont('helvetica',col.key==='doctor'||col.rate?'bold':'normal');
            const fit=fitText(value,cw-3,col.key==='doctor'?5.2:5.0,3.6);pdf.setFontSize(fit.size);
            const baseline=ry+rowH/2+fit.size*.18;
            pdf.text(fit.text,col.align==='left'?px+1.6:px+cw/2,baseline,{align:col.align==='left'?'left':'center'});
            px+=cw;
          });
        });
      };

      const topEntries=data=>Object.entries(data||{}).sort((a,b)=>b[1]-a[1]);
      const examLaunchCounts=countBy(dashboardExams.map(x=>({...x,launcher:launchedName(x)})),'launcher');
      const surgeryInsuranceCounts=countBy(dashboardSurgeries.map(x=>({...x,insuranceLabel:x.insurance||'Particular'})),'insuranceLabel');

      // =====================================================
      // PÁGINA 1 — DADOS E INDICADORES
      // =====================================================
      pageHeader('DADOS E INDICADORES',`Período: ${periodLabel}  |  Médico: ${selectedDoctor}  |  Responsável: ${selectedUser}  |  Status: ${selectedStatus}`);
      let y=49;
      const cardGap=3;
      const metricW=(contentWidth-cardGap*3)/4;
      drawMetricCard(margin,y,metricW,25,'Pacientes únicos',resultSummary.uniquePatients,'pacientes no período');
      drawMetricCard(margin+(metricW+cardGap),y,metricW,25,'Total de exames',resultSummary.totalEx,`${resultSummary.realEx} realizados - ${resultSummary.exRate}%`);
      drawMetricCard(margin+(metricW+cardGap)*2,y,metricW,25,'Total de cirurgias',resultSummary.totalSu,`${resultSummary.realSu} realizadas - ${resultSummary.suRate}%`,'green');
      drawMetricCard(margin+(metricW+cardGap)*3,y,metricW,25,'Total de lançamentos',resultSummary.totalEx+resultSummary.totalSu,'exames + cirurgias');
      y+=29;

      const half=(contentWidth-4)/2;
      drawStatusCard(margin,y,half,63,'Situação dos exames',resultSummary.exRate,[
        {label:'Realizados',value:resultSummary.realEx,detail:'concluídos'},
        {label:'Agendados',value:resultSummary.exAgendado,detail:'programados'},
        {label:'Pendentes',value:resultSummary.exPendente,detail:'precisam atenção'},
        {label:'Cancelados',value:resultSummary.exCancelado,detail:'fora do fluxo'}
      ],`${resultSummary.exOpen} exame(s) ainda em aberto - ${resultSummary.exOpenRate}% do total`);
      drawStatusCard(margin+half+4,y,half,63,'Situação das cirurgias',resultSummary.suRate,[
        {label:'Realizadas',value:resultSummary.realSu,detail:'concluídas'},
        {label:'Agendadas',value:resultSummary.suAgendada,detail:'com data'},
        {label:'Autorizadas',value:resultSummary.suAutorizada,detail:'liberadas'},
        {label:'Pendentes',value:resultSummary.suPendente,detail:'precisam atenção'},
        {label:'Solicitações',value:resultSummary.suSolicitacao,detail:'em análise'},
        {label:'Canceladas',value:resultSummary.suCancelada,detail:'fora do fluxo'}
      ],`${resultSummary.suOpen} cirurgia(s) ainda em fluxo - ${resultSummary.suOpenRate}% do total`);
      y+=68;

      const hiW=(contentWidth-6)/3;
      drawHighlight(margin,y,hiW,22,'Exame com maior volume',resultSummary.examTop?.[0]||'Sem dados',resultSummary.examTop?.[1]||0);
      drawHighlight(margin+hiW+3,y,hiW,22,'Cirurgia com maior volume',resultSummary.surgTop?.[0]||'Sem dados',resultSummary.surgTop?.[1]||0);
      drawHighlight(margin+(hiW+3)*2,y,hiW,22,'Profissional com maior volume',resultSummary.doctorTop?.[0]||'Sem dados',resultSummary.doctorTop?.[1]||0);
      y+=27;

      drawDonutCard(margin,y,half,77,'Distribuição dos status dos exames',advancedAnalytics.examStatusPie,'EXAMES');
      drawDonutCard(margin+half+4,y,half,77,'Distribuição dos status das cirurgias',advancedAnalytics.surgeryStatusPie,'CIRURGIAS');

      // =====================================================
      // PÁGINA 2 — ANÁLISE POR MÉDICO
      // =====================================================
      pdf.addPage('a4','portrait');
      pageHeader('ANÁLISE POR MÉDICO','Exames por médico e tipo + cirurgias por cirurgião e procedimento.');
      y=49;
      drawDoctorColumn(margin,y,half,231,'Exames por médico e tipo',advancedAnalytics.examByDoctor,'Exame');
      drawDoctorColumn(margin+half+4,y,half,231,'Cirurgias por cirurgião e procedimento',advancedAnalytics.surgeryByDoctor,'Procedimento');

      // =====================================================
      // PÁGINA 3 — DETALHAMENTO POR STATUS
      // =====================================================
      pdf.addPage('a4','portrait');
      pageHeader('DETALHAMENTO POR STATUS','Tipos de exame e procedimentos dentro de cada status do período.');
      y=49;
      drawStatusColumn(margin,y,half,231,'Exames por status e tipo',
        ['Realizado','Pendente','Agendado','Cancelado'],advancedAnalytics.examByStatusType,advancedAnalytics.examStatusTotals,'Exame');
      drawStatusColumn(margin+half+4,y,half,231,'Cirurgias por status e procedimento',
        ['Pendente','Solicitação','Autorizada','Agendada','Realizada','Cancelada'],advancedAnalytics.surgeryByStatusType,advancedAnalytics.surgeryStatusTotals,'Procedimento');

      // =====================================================
      // PÁGINA 4 — ANÁLISES GERAIS + LANÇAMENTOS
      // =====================================================
      pdf.addPage('a4','portrait');
      pageHeader('ANÁLISES GERAIS','Distribuição dos exames, cirurgias e lançamentos por responsável.');
      y=49;
      const smallGap=3;
      const boxW=(contentWidth-smallGap)/2;
      const boxH=39;

      pdf.setTextColor(...C.primary);pdf.setFont('helvetica','bold');pdf.setFontSize(7);pdf.text('ANÁLISE DE EXAMES',margin,y);
      y+=4;
      drawRankingBox(margin,y,boxW,boxH,'Exames por tipo',topEntries(countManyBy(dashboardExams,'exam_names')));
      drawRankingBox(margin+boxW+smallGap,y,boxW,boxH,'Exames por médico',topEntries(countBy(dashboardExams,'doctor')));
      y+=boxH+3;
      drawRankingBox(margin,y,boxW,boxH,'Status dos exames',topEntries(countBy(dashboardExams,'status')));
      drawRankingBox(margin+boxW+smallGap,y,boxW,boxH,'Exames por responsável',topEntries(examLaunchCounts));
      y+=boxH+6;

      pdf.setTextColor(...C.primary);pdf.setFont('helvetica','bold');pdf.setFontSize(7);pdf.text('ANÁLISE DE CIRURGIAS',margin,y);
      y+=4;
      drawRankingBox(margin,y,boxW,boxH,'Cirurgias por procedimento',topEntries(countManyBy(dashboardSurgeries,'procedure_names')));
      drawRankingBox(margin+boxW+smallGap,y,boxW,boxH,'Cirurgias por cirurgião',topEntries(countBy(dashboardSurgeries,'doctor')));
      y+=boxH+3;
      drawRankingBox(margin,y,boxW,boxH,'Status das cirurgias',topEntries(countBy(dashboardSurgeries,'status')));
      drawRankingBox(margin+boxW+smallGap,y,boxW,boxH,'Cirurgias por convênio',topEntries(surgeryInsuranceCounts));
      y+=boxH+5;

      drawResponsibleTable(margin,y,contentWidth,47);

      // =====================================================
      // PÁGINA 5 — CONVERSÃO POR MÉDICO
      // =====================================================
      pdf.addPage('a4','portrait');
      pageHeader('CONVERSÃO POR MÉDICO','Status dos exames e cirurgias de cada médico + taxa efetivamente realizada.');
      y=49;

      pdf.setFillColor(...C.soft);pdf.setDrawColor(...C.border);pdf.roundedRect(margin,y,contentWidth,18,3,3,'FD');
      pdf.setTextColor(...C.text);pdf.setFont('helvetica','bold');pdf.setFontSize(6.4);
      pdf.text('Como ler',margin+4,y+6);
      pdf.setTextColor(...C.muted);pdf.setFont('helvetica','normal');pdf.setFontSize(5.1);
      const conversionNote='A conversão representa o percentual de itens realizados sobre o total de exames ou cirurgias vinculados a cada médico. As colunas mostram exatamente quantos estão realizados, agendados, pendentes, autorizados, em solicitação ou cancelados.';
      pdf.text(pdf.splitTextToSize(safe(conversionNote),contentWidth-8),margin+4,y+11);
      y+=23;

      drawDoctorConversionTable(margin,y,contentWidth,93,'EXAMES - taxa de conversão e status por médico',advancedAnalytics.doctorConversion,'exam');
      y+=99;
      drawDoctorConversionTable(margin,y,contentWidth,93,'CIRURGIAS - taxa de conversão e status por médico',advancedAnalytics.doctorConversion,'surgery');

      const totalPages=pdf.getNumberOfPages();
      for(let p=1;p<=totalPages;p++){
        pdf.setPage(p);
        footer(p,totalPages);
      }

      pdf.save(`Oftalmocastro_Resultados_${dashMonth||'todos-periodos'}.pdf`);
    }catch(error){
      console.error(error);
      alert('Não foi possível gerar o PDF de Resultados.');
    }finally{
      setExportingPanelPdf(false);
    }
  };


  const exportResultsPDF = async ()=>{
    if(exportingPdf) return;
    setExportingPdf(true);

    try{
      const selectedDoctor=dashDoctor ? doctors.find(d=>d.id===dashDoctor)?.name || 'Não identificado' : 'Todos os médicos';
      const selectedUser=dashUser || 'Todos os responsáveis';
      const selectedStatus=dashStatus || 'Todos os status';
      const periodLabel=monthLabel(dashMonth);
      const generatedAt=new Date().toLocaleString('pt-BR');

      const pdf=new jsPDF({orientation:'landscape',unit:'mm',format:'a4',compress:true});
      const margin=12;
      const pageWidth=297;
      const pageHeight=210;
      const contentWidth=pageWidth-(margin*2);
      const footerY=202;
      let y=0;
      let sectionTitle='';

      const safe=value=>String(value??'')
        .replace(/•/g,'-')
        .replace(/[“”]/g,'"')
        .replace(/[‘’]/g,"'")
        .replace(/…/g,'...')
        .replace(/[^\x00-\xFF]/g,'');

      const split=(value,width,size=8)=>{
        pdf.setFontSize(size);
        return pdf.splitTextToSize(safe(value||'—'),width);
      };

      const makeCanvas=(width=1100,height=500)=>{
        const canvas=document.createElement('canvas');
        canvas.width=width;
        canvas.height=height;
        const ctx=canvas.getContext('2d');
        ctx.fillStyle='#FFFFFF';
        ctx.fillRect(0,0,width,height);
        return {canvas,ctx};
      };

      const makeLineChartImage=(title,labels,series,percent=false)=>{
        const {canvas,ctx}=makeCanvas(1200,500);
        const W=canvas.width,H=canvas.height;
        const pad={l:82,r:45,t:80,b:70};
        const plotW=W-pad.l-pad.r,plotH=H-pad.t-pad.b;
        const values=series.flatMap(s=>s.values);
        const max=Math.max(...values,percent?100:1);
        const x=i=>labels.length<=1?pad.l+plotW/2:pad.l+(i/(labels.length-1))*plotW;
        const y=v=>pad.t+plotH-(Number(v)/(max||1))*plotH;

        ctx.fillStyle='#20313B';
        ctx.font='700 28px Arial';
        ctx.fillText(title,36,42);

        [0,.25,.5,.75,1].forEach(g=>{
          const val=Math.round(max*g);
          const gy=y(val);
          ctx.strokeStyle='#E6ECEE';
          ctx.lineWidth=1;
          ctx.beginPath();ctx.moveTo(pad.l,gy);ctx.lineTo(W-pad.r,gy);ctx.stroke();
          ctx.fillStyle='#7B8A90';
          ctx.font='14px Arial';
          ctx.textAlign='right';
          ctx.fillText(`${val}${percent?'%':''}`,pad.l-12,gy+5);
        });

        labels.forEach((label,i)=>{
          ctx.fillStyle='#667A82';
          ctx.font='14px Arial';
          ctx.textAlign='center';
          ctx.fillText(label,x(i),H-34);
        });

        series.forEach(s=>{
          ctx.strokeStyle=s.color;
          ctx.fillStyle=s.color;
          ctx.lineWidth=4;
          ctx.beginPath();
          s.values.forEach((v,i)=>{
            if(i===0) ctx.moveTo(x(i),y(v));
            else ctx.lineTo(x(i),y(v));
          });
          ctx.stroke();

          s.values.forEach((v,i)=>{
            ctx.beginPath();
            ctx.arc(x(i),y(v),6,0,Math.PI*2);
            ctx.fill();
            ctx.fillStyle='#20313B';
            ctx.font='700 13px Arial';
            ctx.textAlign='center';
            ctx.fillText(`${v}${percent?'%':''}`,x(i),y(v)-13);
            ctx.fillStyle=s.color;
          });
        });

        let lx=pad.l;
        series.forEach(s=>{
          ctx.fillStyle=s.color;
          ctx.fillRect(lx,58,18,5);
          ctx.fillStyle='#53636E';
          ctx.font='14px Arial';
          ctx.textAlign='left';
          ctx.fillText(s.name,lx+26,64);
          lx+=170;
        });

        ctx.strokeStyle='#DDE6EA';
        ctx.strokeRect(0,0,W,H);
        return canvas.toDataURL('image/png');
      };

      const makeDonutImage=(title,data)=>{
        const {canvas,ctx}=makeCanvas(700,500);
        const entries=Object.entries(data||{}).filter(([,v])=>Number(v)>0);
        const total=entries.reduce((s,[,v])=>s+Number(v),0);
        const colors=['#4F8A6B','#315F72','#B9802B','#B85C5C','#7A6F9B','#6B8F9B','#9A7F56'];

        ctx.fillStyle='#20313B';
        ctx.font='700 27px Arial';
        ctx.fillText(title,30,42);

        if(!total){
          ctx.fillStyle='#74838C';ctx.font='18px Arial';
          ctx.fillText('Sem dados para o período.',30,90);
          return canvas.toDataURL('image/png');
        }

        const cx=205,cy=265,r=120,inner=72;
        let start=-Math.PI/2;
        entries.forEach(([label,value],i)=>{
          const angle=(Number(value)/total)*Math.PI*2;
          ctx.beginPath();
          ctx.moveTo(cx,cy);
          ctx.arc(cx,cy,r,start,start+angle);
          ctx.closePath();
          ctx.fillStyle=colors[i%colors.length];
          ctx.fill();
          start+=angle;
        });

        ctx.beginPath();ctx.arc(cx,cy,inner,0,Math.PI*2);ctx.fillStyle='#fff';ctx.fill();
        ctx.fillStyle='#20313B';ctx.font='700 34px Arial';ctx.textAlign='center';ctx.fillText(String(total),cx,cy-3);
        ctx.fillStyle='#74838C';ctx.font='15px Arial';ctx.fillText('Total',cx,cy+22);

        let ly=118;
        entries.forEach(([label,value],i)=>{
          ctx.fillStyle=colors[i%colors.length];ctx.fillRect(395,ly-11,16,16);
          ctx.fillStyle='#53636E';ctx.font='15px Arial';ctx.textAlign='left';
          const short=String(label).length>22?String(label).slice(0,21)+'…':String(label);
          ctx.fillText(short,422,ly+1);
          ctx.fillStyle='#20313B';ctx.font='700 15px Arial';
          ctx.fillText(`${value} (${percentOf(value,total)}%)`,575,ly+1);
          ly+=39;
        });

        ctx.strokeStyle='#DDE6EA';ctx.strokeRect(0,0,canvas.width,canvas.height);
        return canvas.toDataURL('image/png');
      };

      const statusPalette=status=>{
        const s=String(status||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
        if(s.includes('realiz')||s==='pago'||s==='convenio') return {fill:[237,246,240],text:[62,118,86],border:[208,229,216]};
        if(s.includes('cancel')) return {fill:[251,238,238],text:[166,76,76],border:[238,213,213]};
        if(s.includes('pendent')||s.includes('solicit')||s.includes('autoriz')||s.includes('20%')) return {fill:[255,247,229],text:[151,104,32],border:[239,220,179]};
        return {fill:[234,242,245],text:[49,95,114],border:[211,228,233]};
      };

      const pageHeader=(title,subtitle='')=>{
        pdf.setFillColor(35,78,94);
        pdf.rect(0,0,pageWidth,24,'F');

        pdf.setTextColor(255,255,255);
        pdf.setFont('helvetica','bold');
        pdf.setFontSize(15);
        pdf.text('OFTALMOCASTRO',margin,9);

        pdf.setFont('helvetica','normal');
        pdf.setFontSize(7.6);
        pdf.text('Gestao de exames e cirurgias - Relatorio gerencial',margin,15);

        pdf.setFont('helvetica','bold');
        pdf.setFontSize(11);
        pdf.text(safe(title),pageWidth-margin,9,{align:'right'});

        pdf.setFont('helvetica','normal');
        pdf.setFontSize(7);
        if(subtitle) pdf.text(safe(subtitle),pageWidth-margin,15,{align:'right'});

        y=31;
        sectionTitle=title;
      };

      const newPage=(title=sectionTitle,subtitle=periodLabel)=>{
        pdf.addPage('a4','landscape');
        pageHeader(title,subtitle);
      };

      const ensure=(height,title=sectionTitle)=>{
        if(y+height>footerY-6) newPage(title,periodLabel);
      };

      const drawSection=(title,subtitle='')=>{
        ensure(18,title);
        pdf.setFillColor(242,247,248);
        pdf.setDrawColor(219,230,233);
        pdf.roundedRect(margin,y,contentWidth,14,2,2,'FD');
        pdf.setTextColor(39,78,88);
        pdf.setFont('helvetica','bold');
        pdf.setFontSize(10);
        pdf.text(safe(title),margin+4,y+5.5);
        if(subtitle){
          pdf.setTextColor(112,130,136);
          pdf.setFont('helvetica','normal');
          pdf.setFontSize(6.7);
          pdf.text(safe(subtitle),margin+4,y+10.5);
        }
        y+=18;
      };

      const drawMetricCards=(metrics,cols=4)=>{
        const gap=4;
        const cardW=(contentWidth-gap*(cols-1))/cols;
        const cardH=24;
        const rows=Math.ceil(metrics.length/cols);

        ensure(rows*(cardH+4)+2);
        metrics.forEach((m,i)=>{
          const row=Math.floor(i/cols);
          const col=i%cols;
          const x=margin+col*(cardW+gap);
          const cy=y+row*(cardH+4);

          pdf.setFillColor(...(m.fill||[250,252,252]));
          pdf.setDrawColor(...(m.border||[223,231,232]));
          pdf.roundedRect(x,cy,cardW,cardH,2.5,2.5,'FD');

          pdf.setTextColor(116,132,136);
          pdf.setFont('helvetica','bold');
          pdf.setFontSize(6.5);
          pdf.text(safe(String(m.label||'').toUpperCase()),x+4,cy+5.5);

          pdf.setTextColor(...(m.valueColor||[41,82,91]));
          pdf.setFont('helvetica','bold');
          pdf.setFontSize(15);
          pdf.text(safe(m.value),x+4,cy+13.8);

          if(m.detail){
            pdf.setTextColor(129,143,146);
            pdf.setFont('helvetica','normal');
            pdf.setFontSize(6.2);
            const lines=split(m.detail,cardW-8,6.2).slice(0,2);
            pdf.text(lines,x+4,cy+18.2);
          }
        });
        y+=rows*(cardH+4)+1;
      };

      const drawStatusPanel=(title,total,items)=>{
        const x=margin;
        const panelW=contentWidth;
        const rowH=10;
        const h=16+(items.length*rowH);
        ensure(h+5,title);

        pdf.setFillColor(255,255,255);
        pdf.setDrawColor(221,229,231);
        pdf.roundedRect(x,y,panelW,h,2.5,2.5,'FD');

        pdf.setTextColor(45,72,78);
        pdf.setFont('helvetica','bold');
        pdf.setFontSize(9.5);
        pdf.text(safe(title),x+4,y+6);
        pdf.setTextColor(116,130,134);
        pdf.setFont('helvetica','normal');
        pdf.setFontSize(6.8);
        pdf.text(`Total analisado: ${total}`,x+4,y+11);

        let rowY=y+16;
        items.forEach((item,i)=>{
          const palette=statusPalette(item.label);
          if(i%2===0){
            pdf.setFillColor(249,251,251);
            pdf.rect(x+1,rowY-1,panelW-2,rowH,'F');
          }

          pdf.setFillColor(...palette.fill);
          pdf.setDrawColor(...palette.border);
          pdf.roundedRect(x+4,rowY+1,34,6,2,2,'FD');
          pdf.setTextColor(...palette.text);
          pdf.setFont('helvetica','bold');
          pdf.setFontSize(6.2);
          pdf.text(safe(item.label),x+21,rowY+5,{align:'center'});

          pdf.setTextColor(54,72,77);
          pdf.setFont('helvetica','bold');
          pdf.setFontSize(8);
          pdf.text(String(item.value),x+48,rowY+5);

          const pct=percentOf(item.value,total);
          pdf.setTextColor(117,130,134);
          pdf.setFont('helvetica','normal');
          pdf.setFontSize(6.7);
          pdf.text(`${pct}% do total`,x+60,rowY+5);

          const barX=x+91;
          const barW=panelW-101;
          pdf.setFillColor(236,241,242);
          pdf.roundedRect(barX,rowY+2,barW,4,2,2,'F');
          if(pct>0){
            pdf.setFillColor(...palette.text);
            pdf.roundedRect(barX,rowY+2,Math.max(2,(barW*pct)/100),4,2,2,'F');
          }
          rowY+=rowH;
        });

        y+=h+6;
      };

      const drawSimpleRanking=(title,data,maxItems=8)=>{
        const entries=Object.entries(data||{}).sort((a,b)=>b[1]-a[1]).slice(0,maxItems);
        if(!entries.length) return;

        const h=16+(entries.length*8);
        ensure(h+4,title);

        pdf.setFillColor(255,255,255);
        pdf.setDrawColor(221,229,231);
        pdf.roundedRect(margin,y,contentWidth,h,2.5,2.5,'FD');

        pdf.setTextColor(45,72,78);
        pdf.setFont('helvetica','bold');
        pdf.setFontSize(9);
        pdf.text(safe(title),margin+4,y+6);

        const max=Math.max(...entries.map(([,v])=>v),1);
        let ry=y+13;

        entries.forEach(([label,value])=>{
          const labelLines=split(label,72,6.5);
          pdf.setTextColor(82,101,106);
          pdf.setFont('helvetica','normal');
          pdf.setFontSize(6.5);
          pdf.text(labelLines[0],margin+4,ry+3);

          const bx=margin+82;
          const bw=contentWidth-102;
          pdf.setFillColor(237,242,243);
          pdf.roundedRect(bx,ry,bw,4,2,2,'F');
          pdf.setFillColor(49,95,114);
          pdf.roundedRect(bx,ry,Math.max(2,(value/max)*bw),4,2,2,'F');

          pdf.setTextColor(39,76,86);
          pdf.setFont('helvetica','bold');
          pdf.setFontSize(6.7);
          pdf.text(String(value),margin+contentWidth-4,ry+3,{align:'right'});
          ry+=8;
        });

        y+=h+6;
      };

      const drawTable=(title,columns,rows,options={})=>{
        drawSection(title,options.subtitle||'');

        const headerH=9;
        const padX=1.8;
        const widths=columns.map(c=>c.width);
        const totalW=widths.reduce((a,b)=>a+b,0);
        if(Math.abs(totalW-contentWidth)>0.6){
          throw new Error(`Tabela "${title}" com largura invalida: ${totalW}mm`);
        }

        const drawHeader=()=>{
          ensure(headerH+6,title);
          let x=margin;
          pdf.setFillColor(49,95,114);
          pdf.setTextColor(255,255,255);
          pdf.setFont('helvetica','bold');
          pdf.setFontSize(6.4);
          columns.forEach((c,i)=>{
            pdf.rect(x,y,widths[i],headerH,'F');
            const lines=split(c.label,widths[i]-padX*2,6.4).slice(0,2);
            pdf.text(lines,x+padX,y+3.7);
            x+=widths[i];
          });
          y+=headerH;
        };

        drawHeader();

        if(!rows.length){
          pdf.setFillColor(250,252,252);
          pdf.setDrawColor(228,234,235);
          pdf.rect(margin,y,contentWidth,12,'FD');
          pdf.setTextColor(126,138,141);
          pdf.setFont('helvetica','normal');
          pdf.setFontSize(7.5);
          pdf.text('Nenhum registro encontrado para os filtros selecionados.',margin+4,y+7);
          y+=16;
          return;
        }

        rows.forEach((row,rowIndex)=>{
          const cellLines=columns.map(c=>{
            const value=typeof c.value==='function' ? c.value(row) : row[c.key];
            return split(value,c.width-padX*2,6.3);
          });

          const maxLines=Math.max(...cellLines.map(lines=>Math.min(lines.length,options.maxLines||4)),1);
          const rowH=Math.max(8,maxLines*3.2+3);

          if(y+rowH>footerY-6){
            newPage(title,periodLabel);
            drawHeader();
          }

          let x=margin;
          if(rowIndex%2===1){
            pdf.setFillColor(249,251,251);
            pdf.rect(margin,y,contentWidth,rowH,'F');
          }

          columns.forEach((c,i)=>{
            pdf.setDrawColor(228,234,235);
            pdf.rect(x,y,widths[i],rowH);

            if(c.key==='status'){
              const raw=typeof c.value==='function' ? c.value(row) : row[c.key];
              const palette=statusPalette(raw);
              const badgeW=Math.min(widths[i]-3,Math.max(13,safe(raw).length*2.25+5));
              pdf.setFillColor(...palette.fill);
              pdf.setDrawColor(...palette.border);
              pdf.roundedRect(x+1.5,y+2,badgeW,5.5,1.8,1.8,'FD');
              pdf.setTextColor(...palette.text);
              pdf.setFont('helvetica','bold');
              pdf.setFontSize(5.8);
              pdf.text(safe(raw||'—'),x+1.5+badgeW/2,y+5.8,{align:'center'});
            }else{
              pdf.setTextColor(58,77,82);
              pdf.setFont('helvetica',c.bold?'bold':'normal');
              pdf.setFontSize(6.3);
              pdf.text(cellLines[i].slice(0,options.maxLines||4),x+padX,y+4);
            }
            x+=widths[i];
          });

          y+=rowH;
        });

        y+=5;
      };

      const drawDoctorTypeBreakdown=(title,rows,itemLabel)=>{
        const flatRows=[];
        (rows||[]).forEach(row=>{
          row.items.forEach(([name,value],index)=>{
            flatRows.push({
              doctor:index===0?row.doctor:'',
              item:name,
              qty:value,
              share:`${percentOf(value,row.total)}%`,
              total:index===0?row.total:''
            });
          });
        });

        drawTable(
          title,
          [
            {key:'doctor',label:'Médico',width:70,bold:true},
            {key:'item',label:itemLabel,width:120},
            {key:'qty',label:'Quantidade',width:30},
            {key:'share',label:'Participação',width:28},
            {key:'total',label:'Total médico',width:25}
          ],
          flatRows,
          {
            subtitle:`Distribuição por médico e ${itemLabel.toLowerCase()}, de forma simples para leitura gerencial.`,
            maxLines:2
          }
        );
      };

      const drawStatusTypeBreakdown=(groupTitle,statuses,byStatus,totalByStatus,itemLabel)=>{
        statuses.forEach(status=>{
          const rows=Object.entries(byStatus?.[status]||{})
            .sort((a,b)=>b[1]-a[1])
            .map(([name,value])=>({
              name,
              qty:value,
              share:`${percentOf(value,totalByStatus?.[status]||0)}%`
            }));

          drawTable(
            `${groupTitle.toUpperCase()} - ${safe(status).toUpperCase()} - POR ${itemLabel.toUpperCase()}`,
            [
              {key:'name',label:itemLabel,width:190,bold:true},
              {key:'qty',label:'Quantidade',width:42},
              {key:'share',label:'Participação',width:41}
            ],
            rows,
            {
              subtitle:`Total no status ${safe(status)}: ${totalByStatus?.[status]||0}. Distribuição detalhada por ${itemLabel.toLowerCase()}.`,
              maxLines:2
            }
          );
        });
      };

      // ======================================================
      // PÁGINA 1 — RESUMO EXECUTIVO
      // ======================================================
      pageHeader('RELATORIO GERENCIAL COMPLETO',periodLabel);

      pdf.setTextColor(38,58,65);
      pdf.setFont('helvetica','bold');
      pdf.setFontSize(15);
      pdf.text('Resultados e indicadores operacionais',margin,y);

      pdf.setTextColor(112,128,133);
      pdf.setFont('helvetica','normal');
      pdf.setFontSize(7.3);
      pdf.text('Visao consolidada com situacao dos exames, cirurgias, pendencias e registros detalhados.',margin,y+5);
      y+=11;

      const filterBoxH=18;
      pdf.setFillColor(248,250,250);
      pdf.setDrawColor(222,230,232);
      pdf.roundedRect(margin,y,contentWidth,filterBoxH,2,2,'FD');

      const filterCols=[
        ['Periodo',periodLabel],
        ['Medico / cirurgiao',selectedDoctor],
        ['Responsavel',selectedUser],
        ['Status',selectedStatus]
      ];
      filterCols.forEach((f,i)=>{
        const x=margin+4+i*(contentWidth/4);
        pdf.setTextColor(126,139,143);
        pdf.setFont('helvetica','bold');
        pdf.setFontSize(5.8);
        pdf.text(safe(f[0].toUpperCase()),x,y+5);
        pdf.setTextColor(54,72,77);
        pdf.setFont('helvetica','normal');
        pdf.setFontSize(7);
        pdf.text(split(f[1],contentWidth/4-9,7).slice(0,2),x,y+10);
      });
      y+=23;

      drawMetricCards([
        {label:'Pacientes unicos',value:resultSummary.uniquePatients,detail:'Pacientes identificados no periodo'},
        {label:'Total de exames',value:resultSummary.totalEx,detail:`${resultSummary.realEx} realizados - ${resultSummary.exRate}% concluido`},
        {label:'Total de cirurgias',value:resultSummary.totalSu,detail:`${resultSummary.realSu} realizadas - ${resultSummary.suRate}% concluido`},
        {label:'Total operacional',value:resultSummary.totalEx+resultSummary.totalSu,detail:'Exames + procedimentos cirurgicos'}
      ],4);

      drawSection('SITUACAO OPERACIONAL','Separacao clara entre concluido, em aberto e cancelado.');

      drawMetricCards([
        {label:'Exames concluidos',value:resultSummary.realEx,detail:`${resultSummary.exRate}% do total`,fill:[237,246,240],border:[211,231,218],valueColor:[62,118,86]},
        {label:'Exames em aberto',value:resultSummary.exOpen,detail:`${resultSummary.exAgendado} agendados + ${resultSummary.exPendente} pendentes`,fill:[255,247,229],border:[239,220,179],valueColor:[151,104,32]},
        {label:'Cirurgias concluidas',value:resultSummary.realSu,detail:`${resultSummary.suRate}% do total`,fill:[237,246,240],border:[211,231,218],valueColor:[62,118,86]},
        {label:'Cirurgias em aberto',value:resultSummary.suOpen,detail:`Solicitacao, pendente, autorizada ou agendada`,fill:[255,247,229],border:[239,220,179],valueColor:[151,104,32]}
      ],4);

      drawMetricCards([
        {label:'Exames cancelados',value:resultSummary.exCancelado,detail:'Mantidos para leitura gerencial',fill:[251,238,238],border:[238,213,213],valueColor:[166,76,76]},
        {label:'Cirurgias canceladas',value:resultSummary.suCancelada,detail:'Mantidas para leitura gerencial',fill:[251,238,238],border:[238,213,213],valueColor:[166,76,76]},
        {label:'Financeiro pendente',value:resultSummary.paymentPending,detail:`${resultSummary.paymentPendingRate}% das cirurgias/procedimentos`,fill:[255,247,229],border:[239,220,179],valueColor:[151,104,32]},
        {label:'Pago / convenio',value:resultSummary.paymentClosed,detail:'Procedimentos com situacao financeira fechada',fill:[237,246,240],border:[211,231,218],valueColor:[62,118,86]}
      ],4);

      const quickText=
        `No periodo foram registrados ${resultSummary.totalEx} exames e ${resultSummary.totalSu} procedimentos cirurgicos para ${resultSummary.uniquePatients} pacientes unicos. `+
        `Nos exames, ${resultSummary.realEx} foram realizados, ${resultSummary.exOpen} permanecem em aberto e ${resultSummary.exCancelado} foram cancelados. `+
        `Nas cirurgias, ${resultSummary.realSu} foram realizadas, ${resultSummary.suOpen} permanecem em fluxo operacional e ${resultSummary.suCancelada} foram canceladas. `+
        `O exame de maior volume foi ${resultSummary.examTop?.[0]||'nao identificado'} e o procedimento cirurgico mais frequente foi ${resultSummary.surgTop?.[0]||'nao identificado'}.`;

      ensure(24);
      pdf.setFillColor(244,248,249);
      pdf.setDrawColor(218,229,232);
      const quickLines=split(quickText,contentWidth-8,7.3);
      const quickH=Math.max(20,quickLines.length*3.5+10);
      pdf.roundedRect(margin,y,contentWidth,quickH,2.5,2.5,'FD');
      pdf.setTextColor(49,95,114);
      pdf.setFont('helvetica','bold');
      pdf.setFontSize(7);
      pdf.text('LEITURA RAPIDA',margin+4,y+5.5);
      pdf.setTextColor(77,96,102);
      pdf.setFont('helvetica','normal');
      pdf.setFontSize(7.3);
      pdf.text(quickLines,margin+4,y+10.5);
      y+=quickH+5;

      // ======================================================
      // TENDENCIA / CONVERSAO / PIZZA
      // ======================================================
      newPage('TENDENCIA, CONVERSAO E DISTRIBUICAO',periodLabel);

      drawSection('COMPARATIVO COM O MES ANTERIOR',`Comparação de ${periodLabel} com ${monthLabel(advancedAnalytics.previousMonth)}.`);

      drawMetricCards([
        {
          label:'Volume de exames',
          value:`${advancedAnalytics.examVolumeDelta>=0?'+':''}${advancedAnalytics.examVolumeDelta}%`,
          detail:`${advancedAnalytics.currentExamTotal} atual vs ${advancedAnalytics.previousExamTotal} anterior`,
          fill:advancedAnalytics.examVolumeDelta>=0?[237,246,240]:[251,238,238],
          valueColor:advancedAnalytics.examVolumeDelta>=0?[62,118,86]:[166,76,76]
        },
        {
          label:'Volume de cirurgias',
          value:`${advancedAnalytics.surgeryVolumeDelta>=0?'+':''}${advancedAnalytics.surgeryVolumeDelta}%`,
          detail:`${advancedAnalytics.currentSurgeryTotal} atual vs ${advancedAnalytics.previousSurgeryTotal} anterior`,
          fill:advancedAnalytics.surgeryVolumeDelta>=0?[237,246,240]:[251,238,238],
          valueColor:advancedAnalytics.surgeryVolumeDelta>=0?[62,118,86]:[166,76,76]
        },
        {
          label:'Conversao de exames',
          value:`${advancedAnalytics.currentExamRate}%`,
          detail:`${advancedAnalytics.examRateDelta>=0?'+':''}${advancedAnalytics.examRateDelta} p.p. vs mês anterior`
        },
        {
          label:'Conversao de cirurgias',
          value:`${advancedAnalytics.currentSurgeryRate}%`,
          detail:`${advancedAnalytics.surgeryRateDelta>=0?'+':''}${advancedAnalytics.surgeryRateDelta} p.p. vs mês anterior`
        }
      ],4);

      const volumeLine=makeLineChartImage(
        'Evolucao mensal de volume',
        advancedAnalytics.monthLabels,
        [
          {name:'Exames',values:advancedAnalytics.examTrend,color:'#315F72'},
          {name:'Cirurgias',values:advancedAnalytics.surgeryTrend,color:'#4F8A6B'}
        ],
        false
      );
      const conversionLine=makeLineChartImage(
        'Taxa de conversao / realizacao',
        advancedAnalytics.monthLabels,
        [
          {name:'Exames',values:advancedAnalytics.examConversion,color:'#315F72'},
          {name:'Cirurgias',values:advancedAnalytics.surgeryConversion,color:'#B9802B'}
        ],
        true
      );

      ensure(72);
      pdf.addImage(volumeLine,'PNG',margin,y,132,59,undefined,'FAST');
      pdf.addImage(conversionLine,'PNG',margin+141,y,132,59,undefined,'FAST');
      y+=65;

      const examDonut=makeDonutImage('Status dos exames',advancedAnalytics.examStatusPie);
      const surgeryDonut=makeDonutImage('Status das cirurgias',advancedAnalytics.surgeryStatusPie);

      ensure(68);
      pdf.addImage(examDonut,'PNG',margin,y,132,62,undefined,'FAST');
      pdf.addImage(surgeryDonut,'PNG',margin+141,y,132,62,undefined,'FAST');
      y+=68;

      // ======================================================
      // PÁGINA 2 — EXAMES
      // ======================================================
      newPage('EXAMES - SITUACAO E DETALHAMENTO',periodLabel);

      drawStatusPanel('Status dos exames',resultSummary.totalEx,[
        {label:'Realizado',value:resultSummary.realEx},
        {label:'Agendado',value:resultSummary.exAgendado},
        {label:'Pendente',value:resultSummary.exPendente},
        {label:'Cancelado',value:resultSummary.exCancelado}
      ]);

      drawSimpleRanking('Exames por tipo',countManyBy(dashboardExams,'exam_names'),7);

      drawStatusTypeBreakdown(
        'Exames',
        ['Realizado','Agendado','Pendente','Cancelado'],
        advancedAnalytics.examByStatusType,
        advancedAnalytics.examStatusTotals,
        'Tipo de exame'
      );

      drawDoctorTypeBreakdown(
        'EXAMES POR MEDICO E TIPO',
        advancedAnalytics.examByDoctor,
        'Tipo de exame'
      );

      if(advancedAnalytics.examByIndication.length){
        drawDoctorTypeBreakdown(
          'INDICACOES DE EXAMES POR MEDICO',
          advancedAnalytics.examByIndication,
          'Tipo de exame indicado'
        );
      }

      drawTable(
        'DETALHAMENTO DOS EXAMES',
        [
          {key:'date',label:'Data',width:18,value:r=>brDate(r.date)},
          {key:'patient',label:'Paciente',width:43,bold:true},
          {key:'doctor',label:'Medico',width:31},
          {key:'indicatedBy',label:'Indicado por',width:29,value:r=>r.indicatedBy||'—'},
          {key:'exam',label:'Exame(s)',width:46},
          {key:'status',label:'Status',width:22},
          {key:'responsible',label:'Responsavel',width:23,value:r=>launchedName(r)},
          {key:'obs',label:'Observacao',width:61}
        ],
        dashboardExams,
        {subtitle:'Lista completa dos exames encontrados nos filtros atuais.',maxLines:3}
      );

      // ======================================================
      // CIRURGIAS
      // ======================================================
      newPage('CIRURGIAS - SITUACAO E DETALHAMENTO',periodLabel);

      drawStatusPanel('Status das cirurgias / procedimentos',resultSummary.totalSu,[
        {label:'Realizada',value:resultSummary.realSu},
        {label:'Agendada',value:resultSummary.suAgendada},
        {label:'Autorizada',value:resultSummary.suAutorizada},
        {label:'Pendente',value:resultSummary.suPendente},
        {label:'Solicitacao',value:resultSummary.suSolicitacao},
        {label:'Cancelada',value:resultSummary.suCancelada}
      ]);

      drawStatusPanel('Situacao financeira das cirurgias / procedimentos',resultSummary.totalSu,[
        {label:'Pago',value:resultSummary.paymentCounts['Pago']||0},
        {label:'Convenio',value:resultSummary.paymentCounts['Convênio']||0},
        {label:'20% pago',value:resultSummary.paymentCounts['20% pago']||0},
        {label:'Pendente',value:resultSummary.paymentCounts['Pendente']||0},
        {label:'Nao informado',value:resultSummary.paymentCounts['Não informado']||0}
      ]);

      drawStatusTypeBreakdown(
        'Cirurgias',
        ['Realizada','Agendada','Autorizada','Pendente','Solicitação','Cancelada'],
        advancedAnalytics.surgeryByStatusType,
        advancedAnalytics.surgeryStatusTotals,
        'Procedimento'
      );

      drawDoctorTypeBreakdown(
        'CIRURGIAS POR CIRURGIAO E PROCEDIMENTO',
        advancedAnalytics.surgeryByDoctor,
        'Procedimento'
      );

      if(advancedAnalytics.surgeryByIndication.length){
        drawDoctorTypeBreakdown(
          'INDICACOES DE CIRURGIA POR MEDICO',
          advancedAnalytics.surgeryByIndication,
          'Procedimento indicado'
        );
      }

      drawTable(
        'DETALHAMENTO DAS CIRURGIAS',
        [
          {key:'date',label:'Data',width:16,value:r=>brDate(r.date)},
          {key:'patient',label:'Paciente',width:35,bold:true},
          {key:'doctor',label:'Cirurgiao',width:24},
          {key:'indicatedBy',label:'Indicado por',width:24,value:r=>r.indicatedBy||'—'},
          {key:'procedure',label:'Procedimento(s)',width:34},
          {key:'eye',label:'Olho',width:14},
          {key:'insurance',label:'Convenio',width:21,value:r=>r.insurance||'Particular'},
          {key:'status',label:'Status',width:19},
          {key:'payment',label:'Pagamento',width:20},
          {key:'time',label:'Horario',width:14,value:r=>r.time||'—'},
          {key:'responsible',label:'Responsavel',width:19,value:r=>launchedName(r)},
          {key:'obs',label:'Observacao',width:33}
        ],
        dashboardSurgeries,
        {subtitle:'Lista completa com status, pagamento, horario e observacoes.',maxLines:3}
      );

      // ======================================================
      // PÁGINA FINAL — COMPLEMENTO GERENCIAL
      // ======================================================
      newPage('ANALISE COMPLEMENTAR',periodLabel);

      drawSection('DESTAQUES DO PERIODO','Principais volumes encontrados apos a aplicacao dos filtros.');

      drawMetricCards([
        {label:'Exame com maior volume',value:resultSummary.examTop?.[1]||0,detail:resultSummary.examTop?.[0]||'Sem dados'},
        {label:'Cirurgia com maior volume',value:resultSummary.surgTop?.[1]||0,detail:resultSummary.surgTop?.[0]||'Sem dados'},
        {label:'Profissional com maior volume',value:resultSummary.doctorTop?.[1]||0,detail:resultSummary.doctorTop?.[0]||'Sem dados'},
        {label:'Convenio com maior volume',value:resultSummary.insuranceTop?.[1]||0,detail:resultSummary.insuranceTop?.[0]||'Sem dados'}
      ],4);

      drawSimpleRanking(
        'Lançamentos por responsavel',
        Object.fromEntries(userRows.map(u=>[u.name,u.exams+u.surgeries])),
        8
      );

      drawSimpleRanking(
        'Cirurgias por convenio',
        countWeightedBy(
          dashboardSurgeries.map(x=>({...x,insuranceLabel:x.insurance||'Particular'})),
          'insuranceLabel',
          'procedure_ids'
        ),
        8
      );

      drawSimpleRanking('Cirurgias por procedimento',countManyBy(dashboardSurgeries,'procedure_names'),8);

      // Rodape e numeracao em todas as paginas
      const totalPages=pdf.getNumberOfPages();
      for(let p=1;p<=totalPages;p++){
        pdf.setPage(p);
        const w=pdf.internal.pageSize.getWidth();
        const h=pdf.internal.pageSize.getHeight();

        pdf.setDrawColor(220,228,230);
        pdf.line(margin,h-9,w-margin,h-9);

        pdf.setTextColor(128,140,143);
        pdf.setFont('helvetica','normal');
        pdf.setFontSize(6.2);
        pdf.text(`Emitido em ${safe(generatedAt)} - ${safe(currentAccessName||'Sistema')}`,margin,h-5);
        pdf.text(`Pagina ${p} de ${totalPages}`,w-margin,h-5,{align:'right'});
      }

      pdf.save(`Oftalmocastro_Relatorio_Completo_${dashMonth||'todos-periodos'}.pdf`);
    }catch(error){
      console.error(error);
      alert('Nao foi possivel gerar o PDF completo de resultados.');
    }finally{
      setExportingPdf(false);
    }
  };

  const logout=async()=>{ sessionStorage.removeItem('oc_launcher_name'); setLauncherName(''); await supabase.auth.signOut(); setProfile(null); };

  if(loading) return <div className="loading-screen">Carregando...</div>;
  if(!session) return <Login onLogged={setSession}/>;

  return <div className="app">
    <aside className="sidebar">
      <div className="brand"><div className="brand-mark">OC</div><div><b>Oftalmocastro</b><span>Gestão de exames e cirurgias</span></div></div>
      <div className="user-box">
        <b>{profile?.role==='admin' ? 'ADM' : (launcherName || 'Orientadora')}</b>
        <span>{session.user.email}</span>
        {profile?.role!=='admin'&&<select className="launcher-select" value={launcherName} onChange={e=>selectLauncher(e.target.value)}>
          <option value="">Quem está lançando?</option>
          {launcherOptions.map(name=><option key={name} value={name}>{name}</option>)}
        </select>}
      </div>
      <div className="nav">
        <button className={page==='exames'?'active':''} onClick={()=>setPage('exames')}>📋 <span>Exames</span></button>
        <button className={page==='cirurgias'?'active':''} onClick={()=>setPage('cirurgias')}>🏥 <span>Cirurgias</span></button>
        <button className={page==='followup'?'active':''} onClick={()=>setPage('followup')}>🔄 <span>Acompanhamentos</span></button>
        <button className={page==='resultados'?'active':''} onClick={()=>setPage('resultados')}>📊 <span>Resultados</span></button>
        <button className={page==='cadastros'?'active':''} onClick={()=>setPage('cadastros')}>⚙️ <span>Cadastros</span></button>
        <button className={page==='manual'?'active':''} onClick={()=>setPage('manual')}>📖 <span>Manual</span></button>
      </div>
      <button className="btn btn-light logout-btn" onClick={logout}>Sair</button>
    </aside>

    <main className="main">
      {page==='exames'&&<section>
        <div className="topbar">
          <div><div className="eyebrow">GESTÃO OPERACIONAL</div><h1>Controle mensal de exames</h1><div className="subtitle">Dados compartilhados entre os usuários da clínica.</div></div>
          <div className="top-actions">
            <button className="btn btn-light" onClick={()=>setPage('cadastros')}>⚙️ Cadastros</button>
            <button className="btn btn-primary" onClick={()=>{setEditingExamId(null);setExamForm({date:todayISO(),patient:'',whatsapp:'',doctor_id:'',exam_type_ids:[],status:'Agendado',launcher:currentLaunchName||'',obs:''});setExamModal(true)}}>+ Novo exame</button>
          </div>
        </div>
        <div className="kpis">
          <Kpi label="Exames no mês" value={totalItems(filteredExams,'exam_type_ids')}/>
          <Kpi label="Agendados" value={totalItemsByStatus(filteredExams,'exam_type_ids','Agendado')}/>
          <Kpi label="Pendentes" value={totalItemsByStatus(filteredExams,'exam_type_ids','Pendente')}/>
          <Kpi label="Realizados" value={totalItemsByStatus(filteredExams,'exam_type_ids','Realizado')}/>
          <Kpi label="Cancelados" value={totalItemsByStatus(filteredExams,'exam_type_ids','Cancelado')}/>
        </div>
        <div className="card">
          <div className="filters">
            <input type="month" value={filterMonth} onChange={e=>setFilterMonth(e.target.value)}/>
            <Select value={filterDoctor} setValue={setFilterDoctor} options={doctors} first="Todos os médicos"/>
            <Select value={filterExam} setValue={setFilterExam} options={examTypes} first="Todos os exames"/>
            <button className="btn btn-light" onClick={()=>{setFilterMonth('');setFilterDoctor('');setFilterExam('')}}>Limpar</button>
            <button
              className="btn btn-danger-soft bulk-delete-btn"
              onClick={deleteFilteredExams}
              disabled={!filteredExams.length}
              title="Excluir todos os lançamentos exibidos pelos filtros atuais"
            >
              🗑 Excluir filtrados ({filteredExams.length})
            </button>
          </div>
          <div className="table-wrap"><table><thead><tr><th>Data</th><th>Paciente</th><th>Médico</th><th>Indicado por</th><th>Exame</th><th>Status</th><th>Observação</th><th>Lançado por</th><th>Ações</th></tr></thead>
          <tbody>{filteredExams.length?filteredExams.map(r=><tr key={r.id}>
            <td>{brDate(r.date)}</td><td><PatientCell name={r.patient} phone={r.whatsapp}/></td>
            <td>{r.doctor}</td><td>{r.indicatedBy||<span className="muted">—</span>}</td><td>{r.exam}</td><td><span className={`status ${statusClass(r.status)}`}>{r.status}</span></td>
            <td><ObservationCell text={r.obs}/></td><td>{launchedName(r)}</td>
            <td><div className="row-actions">
              <button className="btn btn-edit-soft" onClick={()=>openEditExam(r)}>Editar</button>
              <button className="btn btn-danger-soft" onClick={()=>delExam(r.id)}>Excluir</button>
            </div></td>
          </tr>):<tr><td colSpan="9" className="empty">Nenhum exame encontrado.</td></tr>}</tbody></table></div>
        </div>
      </section>}

      {page==='cirurgias'&&<section>
        <div className="topbar"><div><div className="eyebrow">CENTRO CIRÚRGICO</div><h1>Controle de cirurgias</h1><div className="subtitle">Programação, autorização, contato e observações.</div></div>
          <div className="top-actions">
            <button className="btn btn-light" onClick={()=>setPage('cadastros')}>⚙️ Cadastros</button>
            <button className="btn btn-primary" onClick={()=>{setEditingSurgeryId(null);setSForm({date:todayISO(),patient:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',procedure_ids:[],eye:'Não se aplica',insurance_id:'',status:'Solicitação',arrival:'',time:'',payment:'Não informado',launcher:currentLaunchName||'',obs:''});setSurgeryModal(true)}}>+ Nova cirurgia</button>
          </div>
        </div>
        <div className="kpis"><Kpi label="Cirurgias no período" value={totalItems(filteredSurgeries,'procedure_ids')}/><Kpi label="Agendadas" value={totalItemsByStatus(filteredSurgeries,'procedure_ids','Agendada')}/><Kpi label="Pendentes" value={totalItemsByStatus(filteredSurgeries,'procedure_ids','Pendente')}/><Kpi label="Realizadas" value={totalItemsByStatus(filteredSurgeries,'procedure_ids','Realizada')}/></div>
        <div className="card">
          <div className="filters">
            <input type="month" value={sMonth} onChange={e=>setSMonth(e.target.value)}/>
            <Select value={sDoctor} setValue={setSDoctor} options={doctors} first="Todos os cirurgiões"/>
            <Select value={sProcedure} setValue={setSProcedure} options={procedures} first="Todos os procedimentos"/>
            <button className="btn btn-light" onClick={()=>{setSMonth('');setSDoctor('');setSProcedure('')}}>Limpar</button>
            <button
              className="btn btn-danger-soft bulk-delete-btn"
              onClick={deleteFilteredSurgeries}
              disabled={!filteredSurgeries.length}
              title="Excluir todos os lançamentos exibidos pelos filtros atuais"
            >
              🗑 Excluir filtrados ({filteredSurgeries.length})
            </button>
          </div>
          <div className="table-wrap"><table><thead><tr><th>Data</th><th>Paciente</th><th>Cirurgião</th><th>Indicado por</th><th>Procedimento</th><th>Olho</th><th>Convênio</th><th>Status</th><th>Observação</th><th>Lançado por</th><th>Ações</th></tr></thead>
          <tbody>{filteredSurgeries.length?filteredSurgeries.map(r=><tr key={r.id}>
            <td>{brDate(r.date)}{r.time&&<><br/><small>{r.time}</small></>}</td><td><PatientCell name={r.patient} phone={r.whatsapp}/></td>
            <td>{r.doctor}</td><td>{r.indicatedBy||<span className="muted">—</span>}</td><td>{r.procedure}</td><td>{r.eye}</td><td>{r.insurance}</td>
            <td><span className={`status ${statusClass(r.status)}`}>{r.status}</span></td>
            <td><ObservationCell text={r.obs}/></td><td>{launchedName(r)}</td>
            <td><div className="row-actions">
              <button className="btn btn-edit-soft" onClick={()=>openEditSurgery(r)}>Editar</button>
              <button className="btn btn-danger-soft" onClick={()=>delSurgery(r.id)}>Excluir</button>
            </div></td>
          </tr>):<tr><td colSpan="11" className="empty">Nenhuma cirurgia encontrada.</td></tr>}</tbody></table></div>
        </div>
      </section>}

      {page==='followup'&&<section className="followup-page">
        <div className="topbar followup-topbar">
          <div>
            <div className="eyebrow">ACOMPANHAMENTO OPERACIONAL</div>
            <h1>Follow-up de exames e cirurgias</h1>
            <div className="subtitle">Pacientes conectados aos lançamentos originais. Acompanhe contatos, próximos passos e o status atual sem duplicar cadastros.</div>
          </div>
        </div>

        <div className="kpis followup-kpis">
          <Kpi label="Pacientes na fila" value={followupStats.total}/>
          <Kpi label="Follow-ups registrados" value={followupStats.registered}/>
          <Kpi label="Ainda não iniciados" value={followupStats.notStarted}/>
          <Kpi label="Concluídos" value={followupStats.done}/>
        </div>

        <div className="card followup-metrics-card">
          <div className="followup-metrics-head">
            <div>
              <span>DESEMPENHO DO FOLLOW-UP</span>
              <h3>Como está o acompanhamento?</h3>
              <p>Os indicadores abaixo acompanham automaticamente os filtros selecionados.</p>
            </div>
            <div className="followup-rate-box">
              <strong>{followupStats.followupRate}%</strong>
              <span>já acompanhados</span>
            </div>
          </div>

          <div className="followup-progress-wrap">
            <div className="followup-progress-label"><span>Andamento geral</span><b>{followupStats.registered} de {followupStats.total}</b></div>
            <div className="followup-progress"><i style={{width:`${followupStats.followupRate}%`}}/></div>
          </div>

          <div className="followup-metrics-grid">
            <div><span>Em contato</span><b>{followupStats.inContact}</b><small>contato em andamento</small></div>
            <div><span>Aguardando paciente</span><b>{followupStats.waiting}</b><small>aguardando resposta</small></div>
            <div><span>Sem retorno</span><b>{followupStats.noReturn}</b><small>precisam nova tentativa</small></div>
            <div><span>Contato hoje</span><b>{followupStats.dueToday}</b><small>programados para hoje</small></div>
            <div className="success"><span>Concluídos</span><b>{followupStats.done}</b><small>{followupStats.completionRate}% dos registrados</small></div>
          </div>
        </div>

        <div className="card followup-toolbar">
          <div className="filter-title">Filtros do acompanhamento</div>
          <div className="followup-filters">
            <label><span>Dia do exame/cirurgia</span><input type="date" value={followDay} onChange={e=>setFollowDay(e.target.value)}/></label>
            <label><span>Origem</span><select value={followType} onChange={e=>setFollowType(e.target.value)}><option value="">Exames e cirurgias</option><option value="exam">Exames</option><option value="surgery">Cirurgias</option></select></label>
            <label><span>Médico</span><Select value={followDoctor} setValue={setFollowDoctor} options={doctors} first="Todos os médicos"/></label>
            <label><span>Status do lançamento</span><select value={followSourceStatus} onChange={e=>setFollowSourceStatus(e.target.value)}><option value="">Todos</option>{['Agendado','Pendente','Realizado','Cancelado','Solicitação','Autorizada','Agendada','Realizada','Cancelada'].map(x=><option key={x} value={x}>{x}</option>)}</select></label>
            <label><span>Status do follow-up</span><select value={followStatus} onChange={e=>setFollowStatus(e.target.value)}><option value="">Todos</option>{['A fazer','Em contato','Aguardando paciente','Sem retorno','Concluído'].map(x=><option key={x} value={x}>{x}</option>)}</select></label>
            <label className="followup-search"><span>Buscar paciente</span><input value={followSearch} onChange={e=>setFollowSearch(e.target.value)} placeholder="Nome, médico ou procedimento"/></label>
          </div>
          <div className="followup-filter-actions">
            <button className="btn btn-light" onClick={()=>{setFollowDay('');setFollowType('');setFollowDoctor('');setFollowSourceStatus('');setFollowStatus('');setFollowSearch('')}}>Limpar filtros</button>
            <span>{filteredFollowups.length} registro(s)</span>
          </div>
        </div>

        <div className="followup-list">
          {filteredFollowups.length ? filteredFollowups.map(row=>{
            const f=row.followup;
            const followLabel=f?.status||'A fazer';
            const isDue=f?.next_contact_date===todayISO();
            return <article className={`card followup-card ${isDue?'due-today':''}`} key={row.key}>
              <div className="followup-card-head">
                <div>
                  <div className="followup-badges">
                    <span className={`source-badge ${row.source_type}`}>{row.source_type==='exam'?'EXAME':'CIRURGIA'}</span>
                    <span className={`status ${statusClass(row.source_status)}`}>{row.source_status}</span>
                    <span className={`follow-status ${statusClass(followLabel)}`}>{followLabel}</span>
                    {isDue&&<span className="follow-due">CONTATO HOJE</span>}
                  </div>
                  <h3>{row.patient||'Paciente não informado'}</h3>
                  <p>{row.item||'—'}</p>
                </div>
              </div>

              <div className="followup-info-grid">
                <div><span>Data</span><b>{brDate(row.source_date)}</b></div>
                <div><span>Médico</span><b>{row.doctor||'—'}</b></div>
                <div><span>Último follow-up</span><b>{f?.followup_date?brDate(f.followup_date):'Não registrado'}</b></div>
                <div><span>Próximo contato</span><b>{f?.next_contact_date?brDate(f.next_contact_date):'—'}</b></div>
                <div><span>Responsável</span><b>{f?.responsible_name||'—'}</b></div>
              </div>

              <div className="followup-note">
                <span>Observação do follow-up</span>
                <p>{f?.note||'Nenhuma observação registrada.'}</p>
              </div>

              <div className="followup-actions">
                <button className="btn btn-primary" onClick={()=>openFollowup(row)}>{f?'Atualizar follow-up':'Registrar follow-up'}</button>
                <button className="btn btn-light" onClick={()=>openOriginalFromFollowup(row)}>{row.source_type==='exam'?'Abrir exame':'Abrir cirurgia'}</button>
              </div>
            </article>;
          }) : <div className="card empty followup-empty">Nenhum paciente encontrado com os filtros selecionados.</div>}
        </div>
      </section>}

      {page==='resultados'&&<section className="results-export-root" ref={resultsPageRef}>
        <div className="pdf-page-block" ref={pdfPage1Ref}>
          <div className="topbar results-header">
            <div>
              <div className="eyebrow">INTELIGÊNCIA OPERACIONAL</div>
              <h1>Resultados e análise mensal</h1>
              <div className="subtitle">Visão consolidada da produção de exames, cirurgias, pacientes e responsáveis pelos lançamentos.</div>
            </div>
            <div className="top-actions no-pdf">
              <button className="btn btn-light" onClick={exportExcel}>⬇ Excel</button>
              <button className="btn btn-panel-pdf" onClick={exportResultsPagePDF} disabled={exportingPanelPdf}>
                {exportingPanelPdf?'Gerando PDF...':'⬇ PDF resultados'}
              </button>
              <button className="btn btn-excel" onClick={exportResultsPDF} disabled={exportingPdf}>
                {exportingPdf?'Gerando PDF detalhado...':'⬇ PDF detalhado'}
              </button>
            </div>
          </div>

          <div className="card results-toolbar">
            <div className="filter-title">Filtros do relatório</div>
            <div className="filters results-filters">
              <input type="month" value={dashMonth} onChange={e=>setDashMonth(e.target.value)}/>
              <Select value={dashDoctor} setValue={setDashDoctor} options={doctors} first="Todos os médicos"/>
              <Select value={dashUser} setValue={setDashUser} options={['Juliana','Luiza','Elisangela','ADM']} first="Todos os responsáveis"/>
              <Select value={dashStatus} setValue={setDashStatus} options={['Agendado','Realizado','Cancelado','Solicitação','Pendente','Autorizada','Agendada','Realizada','Cancelada']} first="Todos os status"/>
            </div>
            <div className="active-period">Período analisado: <b>{monthLabel(dashMonth)}</b></div>
          </div>

          <div className="analysis-grid results-kpis">
            <div className="analysis-card highlight"><span>Pacientes únicos</span><strong>{resultSummary.uniquePatients}</strong><small>pacientes no período</small></div>
            <div className="analysis-card"><span>Total de exames</span><strong>{resultSummary.totalEx}</strong><small>{resultSummary.realEx} realizados • {resultSummary.exRate}% de realização</small></div>
            <div className="analysis-card"><span>Total de cirurgias</span><strong>{resultSummary.totalSu}</strong><small>{resultSummary.realSu} realizadas • {resultSummary.suRate}% de realização</small></div>
            <div className="analysis-card"><span>Total de lançamentos</span><strong>{resultSummary.totalEx+resultSummary.totalSu}</strong><small>exames + cirurgias</small></div>
          </div>

          <div className="results-status-overview">
            <div className="card status-summary-card">
              <div className="status-summary-head">
                <div>
                  <span>EXAMES</span>
                  <h3>Situação dos exames</h3>
                  <p>Veja claramente o que foi concluído e o que ainda exige acompanhamento.</p>
                </div>
                <div className="status-rate success"><b>{resultSummary.exRate}%</b><small>realização</small></div>
              </div>
              <div className="status-breakdown-grid">
                <div className="status-breakdown-item success"><span>Realizados</span><b>{resultSummary.realEx}</b><small>concluídos</small></div>
                <div className="status-breakdown-item info"><span>Agendados</span><b>{resultSummary.exAgendado}</b><small>programados</small></div>
                <div className="status-breakdown-item warning"><span>Pendentes</span><b>{resultSummary.exPendente}</b><small>precisam atenção</small></div>
                <div className="status-breakdown-item danger"><span>Cancelados</span><b>{resultSummary.exCancelado}</b><small>fora do fluxo</small></div>
              </div>
              <div className="status-summary-foot"><b>{resultSummary.exOpen}</b> exame(s) ainda em aberto • {resultSummary.exOpenRate}% do total</div>
            </div>

            <div className="card status-summary-card">
              <div className="status-summary-head">
                <div>
                  <span>CIRURGIAS</span>
                  <h3>Situação das cirurgias</h3>
                  <p>Acompanhe cada etapa até a realização ou cancelamento.</p>
                </div>
                <div className="status-rate success"><b>{resultSummary.suRate}%</b><small>realização</small></div>
              </div>
              <div className="status-breakdown-grid surgery">
                <div className="status-breakdown-item success"><span>Realizadas</span><b>{resultSummary.realSu}</b><small>concluídas</small></div>
                <div className="status-breakdown-item info"><span>Agendadas</span><b>{resultSummary.suAgendada}</b><small>com data</small></div>
                <div className="status-breakdown-item info"><span>Autorizadas</span><b>{resultSummary.suAutorizada}</b><small>liberadas</small></div>
                <div className="status-breakdown-item warning"><span>Pendentes</span><b>{resultSummary.suPendente}</b><small>precisam atenção</small></div>
                <div className="status-breakdown-item neutral"><span>Solicitações</span><b>{resultSummary.suSolicitacao}</b><small>em análise</small></div>
                <div className="status-breakdown-item danger"><span>Canceladas</span><b>{resultSummary.suCancelada}</b><small>fora do fluxo</small></div>
              </div>
              <div className="status-summary-foot"><b>{resultSummary.suOpen}</b> cirurgia(s) ainda em fluxo • {resultSummary.suOpenRate}% do total</div>
            </div>
          </div>

          <div className="results-operational-strip">
            <div><span>Financeiro em aberto</span><b>{resultSummary.paymentPending}</b><small>pendente, 20% pago ou não informado</small></div>
            <div><span>Financeiro fechado</span><b>{resultSummary.paymentClosed}</b><small>pago ou convênio</small></div>
            <div><span>Exame de maior volume</span><b className="text-value">{resultSummary.examTop?.[0]||'Sem dados'}</b><small>{resultSummary.examTop?.[1]||0} lançamento(s)</small></div>
            <div><span>Cirurgia de maior volume</span><b className="text-value">{resultSummary.surgTop?.[0]||'Sem dados'}</b><small>{resultSummary.surgTop?.[1]||0} lançamento(s)</small></div>
          </div>

          <div className="results-highlights">
            <div className="card insight-card"><span>Exame com maior volume</span><b>{resultSummary.examTop?.[0]||'Sem dados'}</b><small>{resultSummary.examTop?.[1]||0} lançamentos</small></div>
            <div className="card insight-card"><span>Cirurgia com maior volume</span><b>{resultSummary.surgTop?.[0]||'Sem dados'}</b><small>{resultSummary.surgTop?.[1]||0} lançamentos</small></div>
            <div className="card insight-card"><span>Profissional com maior volume</span><b>{resultSummary.doctorTop?.[0]||'Sem dados'}</b><small>{resultSummary.doctorTop?.[1]||0} registros</small></div>
          </div>

          <div className="results-section analytics-section">
            <div className="results-section-title">
              <div>
                <h3>Evolução e conversão</h3>
                <p>Comparativo dos últimos 6 meses e variação em relação ao mês anterior.</p>
              </div>
              <div className="pill">Base: {monthLabel(advancedAnalytics.baseMonth)}</div>
            </div>

            <div className="comparison-cards">
              <div className={`comparison-card ${advancedAnalytics.examVolumeDelta>=0?'positive':'negative'}`}>
                <span>Volume de exames vs mês anterior</span>
                <b>{advancedAnalytics.examVolumeDelta>=0?'↑':'↓'} {Math.abs(advancedAnalytics.examVolumeDelta)}%</b>
                <small>{advancedAnalytics.currentExamTotal} neste mês • {advancedAnalytics.previousExamTotal} no anterior</small>
              </div>
              <div className={`comparison-card ${advancedAnalytics.surgeryVolumeDelta>=0?'positive':'negative'}`}>
                <span>Volume de cirurgias vs mês anterior</span>
                <b>{advancedAnalytics.surgeryVolumeDelta>=0?'↑':'↓'} {Math.abs(advancedAnalytics.surgeryVolumeDelta)}%</b>
                <small>{advancedAnalytics.currentSurgeryTotal} neste mês • {advancedAnalytics.previousSurgeryTotal} no anterior</small>
              </div>
              <div className={`comparison-card ${advancedAnalytics.examRateDelta>=0?'positive':'negative'}`}>
                <span>Conversão de exames</span>
                <b>{advancedAnalytics.currentExamRate}%</b>
                <small>{advancedAnalytics.examRateDelta>=0?'+':''}{advancedAnalytics.examRateDelta} p.p. vs mês anterior</small>
              </div>
              <div className={`comparison-card ${advancedAnalytics.surgeryRateDelta>=0?'positive':'negative'}`}>
                <span>Conversão de cirurgias</span>
                <b>{advancedAnalytics.currentSurgeryRate}%</b>
                <small>{advancedAnalytics.surgeryRateDelta>=0?'+':''}{advancedAnalytics.surgeryRateDelta} p.p. vs mês anterior</small>
              </div>
            </div>

            <div className="analytics-chart-grid">
              <TrendLineChart
                title="Evolução mensal de volume"
                labels={advancedAnalytics.monthLabels}
                series={[
                  {name:'Exames',values:advancedAnalytics.examTrend,color:'#315f72'},
                  {name:'Cirurgias',values:advancedAnalytics.surgeryTrend,color:'#4f8a6b'}
                ]}
              />
              <TrendLineChart
                title="Taxa de conversão / realização"
                labels={advancedAnalytics.monthLabels}
                percent
                series={[
                  {name:'Exames',values:advancedAnalytics.examConversion,color:'#315f72'},
                  {name:'Cirurgias',values:advancedAnalytics.surgeryConversion,color:'#b9802b'}
                ]}
              />
            </div>

            <div className="analytics-chart-grid donut-grid">
              <DonutChart title="Distribuição dos status dos exames" data={advancedAnalytics.examStatusPie} centerLabel="Exames"/>
              <DonutChart title="Distribuição dos status das cirurgias" data={advancedAnalytics.surgeryStatusPie} centerLabel="Cirurgias"/>
            </div>
          </div>

          <div className="results-section doctor-analysis-section">
            <div className="results-section-title">
              <div>
                <h3>Produção por médico</h3>
                <p>Veja exatamente quais exames e procedimentos estão associados a cada médico.</p>
              </div>
              <div className="pill">Médico × tipo</div>
            </div>

            <div className="doctor-analysis-grid">
              <DoctorTypeMatrix
                title="Exames por médico e tipo"
                subtitle="Exemplo de leitura: Dr. Alexandre — 8 Mapeamentos, 5 Retinografias, 3 OCTs."
                rows={advancedAnalytics.examByDoctor}
                itemLabel="Exame"
              />

              <DoctorTypeMatrix
                title="Cirurgias por cirurgião e procedimento"
                subtitle="Exemplo de leitura: Dr. Daniel — 5 Cataratas, 7 Pterígios e demais procedimentos do período."
                rows={advancedAnalytics.surgeryByDoctor}
                itemLabel="Cirurgia"
              />
            </div>
          </div>

          <div className="results-section indication-analysis-section">
            <div className="results-section-title">
              <div>
                <h3>Indicações por médico</h3>
                <p>Quando o campo “Indicado por” estiver preenchido no lançamento, o sistema separa quem indicou de quem realizou.</p>
              </div>
              <div className="pill">Origem da indicação</div>
            </div>

            <div className="doctor-analysis-grid">
              <DoctorTypeMatrix
                title="Indicações de exames"
                subtitle="Quantidade e tipo de exame indicado por cada médico."
                rows={advancedAnalytics.examByIndication}
                itemLabel="Exame indicado"
              />

              <DoctorTypeMatrix
                title="Indicações de cirurgias"
                subtitle="Quantidade e procedimento cirúrgico indicado por cada médico."
                rows={advancedAnalytics.surgeryByIndication}
                itemLabel="Cirurgia indicada"
              />
            </div>
          </div>

          <StatusTypeMatrix
            title="Exames por status e por tipo"
            subtitle="Dentro de cada status você vê exatamente quantos exames de cada tipo existem — por exemplo, quantas retinografias ou mapeamentos estão pendentes, realizados, agendados ou cancelados."
            statuses={['Pendente','Realizado','Agendado','Cancelado']}
            data={advancedAnalytics.examByStatusType}
            totalByStatus={advancedAnalytics.examStatusTotals}
            labelName="Exame"
          />

          <StatusTypeMatrix
            title="Cirurgias por status e procedimento"
            subtitle="Mostra quantos procedimentos de cada tipo estão em solicitação, pendentes, autorizados, agendados, realizados ou cancelados."
            statuses={['Pendente','Solicitação','Autorizada','Agendada','Realizada','Cancelada']}
            data={advancedAnalytics.surgeryByStatusType}
            totalByStatus={advancedAnalytics.surgeryStatusTotals}
            labelName="Procedimento"
          />

          <div className="results-section results-section-first">
            <div className="results-section-title"><div><h3>Análise de exames</h3><p>Distribuição dos exames no período selecionado.</p></div><div className="pill">{resultSummary.totalEx} exames</div></div>
            <div className="grid">
              <div className="card chart-card"><h2>Exames por tipo</h2><Bars data={countManyBy(dashboardExams,'exam_names')}/></div>
              <div className="card chart-card"><h2>Exames por médico</h2><Bars data={countBy(dashboardExams,'doctor')}/></div>
            </div>
          </div>
        </div>

        <div className="pdf-page-block" ref={pdfPage2Ref}>
          <div className="grid pdf-continuation-grid">
            <div className="card chart-card"><h2>Status dos exames</h2><Bars data={countBy(dashboardExams,'status')}/></div>
            <div className="card chart-card"><h2>Exames por responsável</h2><Bars data={countBy(dashboardExams.map(x=>({...x,launcher:launchedName(x)})),'launcher')}/></div>
          </div>

          <div className="results-section">
            <div className="results-section-title"><div><h3>Análise de cirurgias</h3><p>Produção, procedimentos, status e convênios.</p></div><div className="pill">{resultSummary.totalSu} cirurgias</div></div>
            <div className="grid">
              <div className="card chart-card"><h2>Cirurgias por procedimento</h2><Bars data={countManyBy(dashboardSurgeries,'procedure_names')}/></div>
              <div className="card chart-card"><h2>Cirurgias por cirurgião</h2><Bars data={countBy(dashboardSurgeries,'doctor')}/></div>
            </div>
          </div>
        </div>

        <div className="pdf-page-block" ref={pdfPage3Ref}>
          <div className="grid pdf-continuation-grid">
            <div className="card chart-card"><h2>Status das cirurgias</h2><Bars data={countBy(dashboardSurgeries,'status')}/></div>
            <div className="card chart-card"><h2>Cirurgias por convênio</h2><Bars data={countBy(dashboardSurgeries.map(x=>({...x,insuranceLabel:x.insurance||'Particular'})),'insuranceLabel')}/></div>
          </div>

          <div className="card results-users-card">
            <div className="section-title">
              <div><h2>Lançamentos por responsável</h2><p className="section-description">Acompanhe quantos registros foram realizados por cada usuário.</p></div>
              <div className="pill">{resultSummary.totalEx+resultSummary.totalSu} registros</div>
            </div>
            <div className="table-wrap"><table><thead><tr><th>Responsável</th><th>Exames</th><th>Cirurgias</th><th>Total</th><th>Participação</th></tr></thead>
            <tbody>{userRows.map(r=>{
              const total=r.exams+r.surgeries;
              const all=resultSummary.totalEx+resultSummary.totalSu;
              return <tr key={r.key}><td><b>{r.name}</b></td><td>{r.exams}</td><td>{r.surgeries}</td><td><b>{total}</b></td><td>{all?Math.round(total/all*100):0}%</td></tr>
            })}</tbody></table></div>
          </div>
        </div>
      </section>}

      {page==='cadastros'&&<section className="catalog-page">
        <div className="topbar">
          <div>
            <div className="eyebrow">CONFIGURAÇÕES DO SISTEMA</div>
            <h1>Cadastros da clínica</h1>
            <div className="subtitle">ADM e Orientadora podem adicionar ou remover médicos, exames e cirurgias/procedimentos. As alterações ficam disponíveis automaticamente nos próximos lançamentos.</div>
          </div>
        </div>

        <div className="catalog-info card">
          <div>
            <b>Cadastros dinâmicos</b>
            <p>Quando surgir um novo médico, exame ou cirurgia na clínica, você pode cadastrar aqui sem precisar alterar o código do sistema.</p>
          </div>
          <span>ADM e Orientadora</span>
        </div>

        <div className="catalog-grid catalog-grid-main">
          <Catalog
            title="Médicos / profissionais"
            description="Cadastre os médicos e profissionais que poderão ser selecionados nos exames e cirurgias."
            placeholder="Ex.: Dra. Mariana Castro"
            items={doctors}
            onAdd={v=>addCatalog('doctors',v)}
            onRemove={id=>deactivateCatalog('doctors',id)}
          />
          <Catalog
            title="Tipos de exames"
            description="Cadastre todos os exames realizados pela clínica para disponibilizá-los no lançamento de exames."
            placeholder="Ex.: Mapeamento de Retina"
            items={examTypes}
            onAdd={v=>addCatalog('exam_types',v)}
            onRemove={id=>deactivateCatalog('exam_types',id)}
          />
          <Catalog
            title="Cirurgias / procedimentos"
            description="Cadastre novas cirurgias e procedimentos para disponibilizá-los no lançamento cirúrgico."
            placeholder="Ex.: Cirurgia de Catarata"
            items={procedures}
            onAdd={v=>addCatalog('procedures',v,{category:'Cirurgia'})}
            onRemove={id=>deactivateCatalog('procedures',id)}
          />
        </div>

        <div className="card catalog-tip">
          <b>Como funciona?</b>
          <p>Digite o nome no campo correspondente e clique em <strong>Adicionar</strong>. O novo item será salvo no Supabase e passará a aparecer nas listas do sistema. Ao remover um cadastro, os lançamentos antigos continuam preservados.</p>
        </div>
      </section>}

      {page==='manual'&&<section className="manual-page">
        <div className="topbar">
          <div>
            <div className="eyebrow">GUIA DO SISTEMA</div>
            <h1>Manual de uso</h1>
            <div className="subtitle">Um guia rápido para usar o sistema da Oftalmocastro no dia a dia.</div>
          </div>
        </div>

        <div className="manual-hero card">
          <div>
            <span className="manual-badge">COMECE AQUI</span>
            <h2>Como funciona o sistema?</h2>
            <p>O sistema centraliza os lançamentos de exames e cirurgias da clínica. As informações ficam salvas no Supabase e podem ser acompanhadas nos Resultados e exportadas para Excel.</p>
          </div>
          <div className="manual-flow">
            <span>1. Entrar</span><b>→</b><span>2. Selecionar responsável</span><b>→</b><span>3. Lançar</span><b>→</b><span>4. Acompanhar resultados</span>
          </div>
        </div>

        <div className="manual-grid">
          <article className="card manual-card">
            <div className="manual-number">01</div>
            <h3>Acesso ao sistema</h3>
            <p>Entre com o e-mail e a senha cadastrados. O acesso pode ser de <b>ADM</b> ou <b>Orientadora</b>.</p>
            <div className="manual-note">Os dois acessos podem utilizar o sistema completo. No acesso Orientadora, selecione Juliana, Luiza ou Elisangela para identificar quem realizou cada lançamento.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">02</div>
            <h3>Quem está lançando?</h3>
            <p>No acesso de Orientadora, selecione <b>Juliana, Luiza ou Elisangela</b> no campo da lateral antes de realizar os lançamentos.</p>
            <div className="manual-note">No acesso ADM, o responsável é preenchido automaticamente como ADM.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">03</div>
            <h3>Lançar um exame</h3>
            <p>Acesse <b>Exames → Novo exame</b>. Preencha data, paciente, médico, tipo de exame, status e observação. O WhatsApp é opcional e pode ser informado para facilitar o contato no Follow-up. Você pode selecionar vários exames no mesmo lançamento. Depois clique em <b>Salvar exame</b>.</p>
            <div className="manual-note">Use o campo Observação para registrar informações importantes que a equipe precise consultar depois.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">04</div>
            <h3>Lançar uma cirurgia</h3>
            <p>Acesse <b>Cirurgias → Nova cirurgia</b>. Informe paciente, cirurgião e selecione um ou vários procedimentos no mesmo lançamento. Depois complete olho, convênio, status, horários, pagamento e observações.</p>
            <div className="manual-note">Mantenha o status atualizado para que os indicadores mensais representem a situação real da cirurgia.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">05</div>
            <h3>Editar informações</h3>
            <p>Na tabela de Exames ou Cirurgias, localize o registro e clique em <b>Editar</b>. O formulário abrirá com as informações atuais. Faça a alteração e clique em <b>Salvar alterações</b>.</p>
          </article>

          <article className="card manual-card">
            <div className="manual-number">06</div>
            <h3>Excluir um lançamento</h3>
            <p>Clique em <b>Excluir</b> na linha desejada. O sistema solicitará uma confirmação antes de remover o registro.</p>
            <div className="manual-warning">A exclusão é definitiva. Confira o paciente e o procedimento antes de confirmar.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">07</div>
            <h3>Filtros</h3>
            <p>Use os filtros no topo das páginas para localizar informações por <b>mês, médico, exame ou procedimento</b>. Em Resultados também é possível filtrar por responsável e status.</p>
          </article>

          <article className="card manual-card">
            <div className="manual-number">08</div>
            <h3>Resultados</h3>
            <p>A página Resultados mostra pacientes únicos, volumes, taxas de realização, exames e cirurgias mais registrados, produção por médico e tipo, indicações por médico quando informadas, status, convênio e lançamentos por responsável.</p>
            <div className="manual-note">Os indicadores acompanham os filtros selecionados. Assim você pode analisar um mês, profissional, responsável ou status específico.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">09</div>
            <h3>Exportar resultados</h3>
            <p>Na página Resultados, aplique os filtros desejados e clique em <b>PDF completo</b>. O relatório mostra resumo executivo, indicadores, situação dos exames e cirurgias, pendências, cancelamentos, situação financeira, comparação com o mês anterior, gráficos de linha, gráficos de pizza e o detalhamento de cada tipo de exame/procedimento dentro de cada status.</p>
            <p>O PDF é paginado automaticamente para que todos os exames e cirurgias filtrados apareçam com <b>paciente, médico/cirurgião, procedimento, status, responsável e observações</b>. A exportação em Excel continua disponível separadamente para conferência dos dados.</p>
          </article>

          <article className="card manual-card">
            <div className="manual-number">10</div>
            <h3>Cadastros da clínica</h3>
            <p>Tanto no acesso <b>ADM</b> quanto no acesso <b>Orientadora</b>, abra <b>Cadastros</b> no menu lateral. Nessa página é possível adicionar ou remover <b>médicos, exames e cirurgias/procedimentos</b>. Os novos itens passam a aparecer automaticamente nos formulários de lançamento.</p><div className="manual-note">Itens removidos deixam de aparecer para novos lançamentos, mas os registros antigos permanecem preservados.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">11</div>
            <h3>WhatsApp do paciente</h3>
            <p>O WhatsApp é opcional ao lançar ou editar um exame/cirurgia. Quando informado, fica salvo no cadastro do paciente e pode ser usado no Follow-up. No Follow-up, o WhatsApp fica oculto na visão principal para manter a tela limpa. Ao clicar em <b>Registrar follow-up</b> ou <b>Atualizar follow-up</b>, o contato aparece junto com o botão <b>Enviar WhatsApp</b>, que abre a conversa com uma mensagem pronta alterando automaticamente o nome do paciente.</p>
          </article>

          <article className="card manual-card">
            <div className="manual-number">12</div>
            <h3>Boas práticas</h3>
            <p>Confira os dados antes de salvar, mantenha os status atualizados, registre observações importantes e evite criar lançamentos duplicados.</p>
            <div className="manual-note">Ao finalizar o uso em um computador compartilhado, clique sempre em <b>Sair</b>.</div>
          </article>
        </div>

        <div className="card manual-help">
          <h3>Fluxo recomendado no dia a dia</h3>
          <div className="manual-checklist">
            <span>✓ Selecione seu nome antes de começar</span>
            <span>✓ Lance exames e cirurgias conforme forem agendados</span>
            <span>✓ Atualize o status quando houver mudança</span>
            <span>✓ Revise os Resultados no fechamento do mês</span>
            <span>✓ Exporte o Excel para conferência e análise</span>
          </div>
        </div>
      </section>}

    </main>

    <style>{`
      .exam-multi-select{border:1px solid #e3e9ed;border-radius:12px;padding:12px;background:#fbfcfd}
      .exam-multi-head{display:flex;justify-content:space-between;gap:12px;align-items:center;margin-bottom:10px;color:#53636e;font-size:13px}
      .exam-multi-head b{color:#315f72}
      .exam-options{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;max-height:240px;overflow:auto}
      .exam-option{display:flex!important;align-items:center;gap:9px;border:1px solid #e3e9ed;border-radius:9px;padding:10px 11px;background:#fff;cursor:pointer;margin:0!important}
      .exam-option.selected{border-color:#477f91;background:#eaf2f5;color:#315f72;font-weight:700}
      .exam-option input{width:auto;margin:0;accent-color:#315f72}
      @media(max-width:700px){.exam-options{grid-template-columns:1fr}}
    `}</style>
    <style>{`
      .btn-panel-pdf{
        display:inline-flex;
        align-items:center;
        justify-content:center;
        gap:8px;
        color:#315f72;
        border:1px solid #d5e4e9;
        background:linear-gradient(180deg,#f3f8fa,#eaf3f6);
        box-shadow:0 5px 14px rgba(49,95,114,.08);
      }

      .btn-panel-pdf:hover{
        background:#dfedf2;
        box-shadow:0 8px 18px rgba(49,95,114,.12);
      }

      .doctor-analysis-section,
      .indication-analysis-section{
        margin-top:28px;
      }

      .doctor-analysis-grid{
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:15px;
      }

      .doctor-type-card{
        min-width:0;
        padding:19px;
      }

      .doctor-type-head{
        margin-bottom:14px;
      }

      .doctor-type-head>div>span{
        display:block;
        color:var(--primary-2);
        font-size:9px;
        font-weight:900;
        letter-spacing:.09em;
      }

      .doctor-type-head h3{
        margin:4px 0;
        color:var(--text);
        font-size:17px;
      }

      .doctor-type-head p{
        margin:0;
        color:var(--muted);
        font-size:10.5px;
        line-height:1.5;
      }

      .doctor-type-grid{
        display:grid;
        gap:10px;
      }

      .doctor-type-section{
        overflow:hidden;
        border:1px solid var(--border);
        border-radius:13px;
        background:#fbfcfc;
      }

      .doctor-type-title{
        display:flex;
        align-items:center;
        justify-content:space-between;
        gap:12px;
        padding:11px 12px;
        border-bottom:1px solid var(--border);
        background:#fff;
      }

      .doctor-type-title b{
        display:block;
        color:#334b56;
        font-size:12px;
      }

      .doctor-type-title small{
        display:block;
        margin-top:3px;
        color:var(--muted);
        font-size:8.5px;
      }

      .doctor-type-title strong{
        color:var(--primary);
        font-size:21px;
      }

      .doctor-type-list>div{
        display:grid;
        grid-template-columns:minmax(140px,1.15fr) minmax(80px,1fr) 30px 36px;
        align-items:center;
        gap:8px;
        min-height:38px;
        padding:7px 11px;
        border-bottom:1px solid #edf1f2;
      }

      .doctor-type-list>div:last-child{
        border-bottom:0;
      }

      .doctor-type-list span{
        overflow:hidden;
        color:#53636e;
        font-size:9.5px;
        font-weight:700;
        text-overflow:ellipsis;
        white-space:nowrap;
      }

      .doctor-type-list b{
        color:var(--primary);
        font-size:10.5px;
        text-align:right;
      }

      .doctor-type-list small{
        color:var(--muted);
        font-size:8.5px;
        text-align:right;
      }

      .doctor-mini-bar{
        height:6px;
        overflow:hidden;
        border-radius:999px;
        background:#e9eef0;
      }

      .doctor-mini-bar i{
        display:block;
        height:100%;
        border-radius:inherit;
        background:linear-gradient(90deg,#315f72,#77a3b0);
      }

      .doctor-type-empty{
        padding:26px 15px;
        border:1px dashed #dbe4e7;
        border-radius:12px;
        color:var(--muted);
        font-size:10.5px;
        text-align:center;
        background:#fafcfc;
      }

      @media(max-width:1100px){
        .doctor-analysis-grid{
          grid-template-columns:1fr;
        }
      }

      @media(max-width:650px){
        .doctor-type-list>div{
          grid-template-columns:minmax(115px,1fr) 60px 28px 32px;
        }
      }
    `}</style>

    <style>{`
      .analytics-section{
        margin-top:30px;
        padding-top:26px;
        border-top:1px solid var(--border);
      }

      .comparison-cards{
        display:grid;
        grid-template-columns:repeat(4,minmax(0,1fr));
        gap:11px;
        margin:0 0 15px;
      }

      .comparison-card{
        min-height:100px;
        padding:14px 15px;
        border:1px solid var(--border);
        border-radius:14px;
        background:#fff;
        box-shadow:var(--shadow-sm);
      }

      .comparison-card span,
      .comparison-card b,
      .comparison-card small{
        display:block;
      }

      .comparison-card span{
        color:var(--muted);
        font-size:9px;
        font-weight:900;
        text-transform:uppercase;
        letter-spacing:.045em;
      }

      .comparison-card b{
        margin:7px 0 4px;
        color:var(--primary);
        font-size:23px;
        letter-spacing:-.03em;
      }

      .comparison-card small{
        color:#87959a;
        font-size:9.5px;
        line-height:1.4;
      }

      .comparison-card.positive{
        border-color:#d8e8dd;
        background:linear-gradient(180deg,#fff,#f5faf7);
      }

      .comparison-card.positive b{color:var(--green)}

      .comparison-card.negative{
        border-color:#eedede;
        background:linear-gradient(180deg,#fff,#fff7f7);
      }

      .comparison-card.negative b{color:var(--red)}

      .analytics-chart-grid{
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:15px;
        margin-top:15px;
      }

      .viz-card{
        min-width:0;
        padding:18px;
        border:1px solid var(--border);
        border-radius:16px;
        background:#fff;
        box-shadow:var(--shadow-sm);
      }

      .viz-title{
        color:var(--text);
        font-size:14px;
        font-weight:850;
        letter-spacing:-.015em;
      }

      .viz-title-row{
        display:flex;
        align-items:flex-start;
        justify-content:space-between;
        gap:12px;
        margin-bottom:7px;
      }

      .viz-title-row small{
        display:block;
        margin-top:3px;
        color:var(--muted);
        font-size:9px;
      }

      .line-legend{
        display:flex;
        flex-wrap:wrap;
        justify-content:flex-end;
        gap:9px;
        color:var(--muted);
        font-size:9px;
      }

      .line-legend span{
        display:flex;
        align-items:center;
        gap:5px;
      }

      .line-legend i{
        width:16px;
        height:3px;
        display:block;
        border-radius:99px;
      }

      .trend-svg{
        width:100%;
        min-height:245px;
        overflow:visible;
      }

      .trend-axis{
        fill:#809096;
        font-size:11px;
      }

      .trend-value{
        fill:#53636e;
        font-size:10px;
        font-weight:800;
      }

      .donut-layout{
        display:grid;
        grid-template-columns:190px minmax(0,1fr);
        align-items:center;
        gap:15px;
        margin-top:10px;
      }

      .donut-svg{
        width:180px;
        max-width:100%;
        margin:auto;
      }

      .donut-total{
        fill:#20313b;
        font-size:21px;
        font-weight:850;
      }

      .donut-label{
        fill:#74838c;
        font-size:9px;
        font-weight:750;
        text-transform:uppercase;
      }

      .donut-legend{
        display:grid;
        gap:8px;
      }

      .donut-legend>div{
        display:grid;
        grid-template-columns:10px minmax(0,1fr) auto auto;
        align-items:center;
        gap:7px;
        padding:7px 0;
        border-bottom:1px solid #eef2f3;
      }

      .donut-legend>div:last-child{
        border-bottom:0;
      }

      .legend-dot{
        width:9px;
        height:9px;
        display:block;
        border-radius:50%;
      }

      .legend-name{
        overflow:hidden;
        color:#53636e;
        font-size:10px;
        text-overflow:ellipsis;
        white-space:nowrap;
      }

      .donut-legend b{
        color:var(--text);
        font-size:10px;
      }

      .donut-legend small{
        min-width:34px;
        color:var(--muted);
        font-size:9px;
        text-align:right;
      }

      .status-matrix-card{
        margin-top:18px;
        padding:20px;
      }

      .status-matrix-head{
        margin-bottom:15px;
      }

      .status-matrix-head span{
        display:block;
        color:var(--primary-2);
        font-size:9px;
        font-weight:900;
        letter-spacing:.09em;
      }

      .status-matrix-head h3{
        margin:4px 0;
        color:var(--text);
        font-size:18px;
      }

      .status-matrix-head p{
        max-width:900px;
        margin:0;
        color:var(--muted);
        font-size:11px;
        line-height:1.5;
      }

      .status-matrix-grid{
        display:grid;
        grid-template-columns:repeat(2,minmax(0,1fr));
        gap:12px;
      }

      .status-matrix-section{
        min-width:0;
        overflow:hidden;
        border:1px solid var(--border);
        border-radius:13px;
        background:#fbfcfc;
      }

      .status-matrix-title{
        display:flex;
        align-items:center;
        justify-content:space-between;
        gap:10px;
        padding:12px 13px;
        border-bottom:1px solid var(--border);
        background:#fff;
      }

      .status-matrix-title>div{
        min-width:0;
      }

      .status-matrix-title small{
        display:block;
        margin-top:5px;
        color:var(--muted);
        font-size:8.5px;
      }

      .status-matrix-title>b{
        min-width:38px;
        color:var(--primary);
        font-size:23px;
        text-align:right;
      }

      .status-type-list{
        display:grid;
      }

      .status-type-list>div{
        display:grid;
        grid-template-columns:minmax(135px,1.1fr) minmax(90px,1fr) 34px 38px;
        align-items:center;
        gap:8px;
        min-height:42px;
        padding:8px 12px;
        border-bottom:1px solid #edf1f2;
      }

      .status-type-list>div:last-child{
        border-bottom:0;
      }

      .status-type-list>div>span{
        overflow:hidden;
        color:#4d6069;
        font-size:10px;
        font-weight:700;
        text-overflow:ellipsis;
        white-space:nowrap;
      }

      .status-type-list b{
        color:var(--primary);
        font-size:11px;
        text-align:right;
      }

      .status-type-list small{
        color:var(--muted);
        font-size:9px;
        text-align:right;
      }

      .mini-bar{
        height:7px;
        overflow:hidden;
        border-radius:99px;
        background:#e9eef0;
      }

      .mini-bar i{
        display:block;
        height:100%;
        border-radius:inherit;
        background:linear-gradient(90deg,var(--primary),#7ca6b2);
      }

      .status-type-empty{
        padding:17px;
        color:var(--muted);
        font-size:10px;
        text-align:center;
      }

      @media(max-width:1150px){
        .comparison-cards{
          grid-template-columns:1fr 1fr;
        }

        .analytics-chart-grid{
          grid-template-columns:1fr;
        }
      }

      @media(max-width:900px){
        .status-matrix-grid{
          grid-template-columns:1fr;
        }
      }

      @media(max-width:650px){
        .comparison-cards{
          grid-template-columns:1fr;
        }

        .donut-layout{
          grid-template-columns:1fr;
        }

        .status-type-list>div{
          grid-template-columns:minmax(120px,1fr) 70px 30px 34px;
        }
      }
    `}</style>

    <style>{`
      .results-status-overview{
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:15px;
        margin:0 0 18px;
      }

      .status-summary-card{
        padding:19px;
        box-shadow:var(--shadow-sm);
      }

      .status-summary-head{
        display:flex;
        align-items:flex-start;
        justify-content:space-between;
        gap:15px;
        margin-bottom:15px;
      }

      .status-summary-head>div:first-child>span{
        display:block;
        margin-bottom:4px;
        color:var(--primary-2);
        font-size:9px;
        font-weight:900;
        letter-spacing:.1em;
      }

      .status-summary-head h3{
        margin:0;
        color:var(--text);
        font-size:17px;
      }

      .status-summary-head p{
        max-width:460px;
        margin:4px 0 0;
        color:var(--muted);
        font-size:11px;
        line-height:1.45;
      }

      .status-rate{
        min-width:78px;
        padding:9px 11px;
        border:1px solid var(--border);
        border-radius:12px;
        text-align:center;
        background:#f8fafb;
      }

      .status-rate b,
      .status-rate small{
        display:block;
      }

      .status-rate b{
        color:var(--primary);
        font-size:19px;
      }

      .status-rate small{
        margin-top:2px;
        color:var(--muted);
        font-size:8px;
        text-transform:uppercase;
        letter-spacing:.04em;
      }

      .status-rate.success{
        border-color:#d8e9de;
        background:var(--green-soft);
      }

      .status-rate.success b{
        color:var(--green);
      }

      .status-breakdown-grid{
        display:grid;
        grid-template-columns:repeat(4,minmax(0,1fr));
        gap:8px;
      }

      .status-breakdown-grid.surgery{
        grid-template-columns:repeat(3,minmax(0,1fr));
      }

      .status-breakdown-item{
        min-height:84px;
        padding:11px 12px;
        border:1px solid var(--border);
        border-radius:12px;
        background:#fafcfc;
      }

      .status-breakdown-item span,
      .status-breakdown-item b,
      .status-breakdown-item small{
        display:block;
      }

      .status-breakdown-item span{
        color:var(--muted);
        font-size:9px;
        font-weight:850;
        text-transform:uppercase;
        letter-spacing:.035em;
      }

      .status-breakdown-item b{
        margin:5px 0 2px;
        color:var(--text);
        font-size:22px;
      }

      .status-breakdown-item small{
        color:#8d999e;
        font-size:9px;
      }

      .status-breakdown-item.success{
        border-color:#dcecdf;
        background:#f4faf6;
      }

      .status-breakdown-item.success b{color:var(--green)}

      .status-breakdown-item.warning{
        border-color:#efdfb9;
        background:#fffaf0;
      }

      .status-breakdown-item.warning b{color:#a77423}

      .status-breakdown-item.danger{
        border-color:#efdede;
        background:#fff6f6;
      }

      .status-breakdown-item.danger b{color:#aa5555}

      .status-breakdown-item.info{
        border-color:#dae8ec;
        background:#f3f8fa;
      }

      .status-breakdown-item.info b{color:var(--primary)}

      .status-breakdown-item.neutral{
        border-color:#e4e7e9;
        background:#f7f8f9;
      }

      .status-summary-foot{
        margin-top:11px;
        padding-top:10px;
        border-top:1px solid #edf1f3;
        color:var(--muted);
        font-size:10.5px;
      }

      .status-summary-foot b{
        color:var(--primary);
      }

      .results-operational-strip{
        display:grid;
        grid-template-columns:repeat(4,minmax(0,1fr));
        gap:10px;
        margin:0 0 18px;
      }

      .results-operational-strip>div{
        min-height:82px;
        padding:13px 14px;
        border:1px solid var(--border);
        border-radius:13px;
        background:linear-gradient(180deg,#fff,#fafcfc);
      }

      .results-operational-strip span,
      .results-operational-strip b,
      .results-operational-strip small{
        display:block;
      }

      .results-operational-strip span{
        color:var(--muted);
        font-size:9px;
        font-weight:850;
        text-transform:uppercase;
      }

      .results-operational-strip b{
        margin:5px 0 2px;
        color:var(--primary);
        font-size:20px;
      }

      .results-operational-strip b.text-value{
        overflow:hidden;
        font-size:13px;
        line-height:1.3;
        text-overflow:ellipsis;
        white-space:nowrap;
      }

      .results-operational-strip small{
        color:#8c999d;
        font-size:9px;
        line-height:1.35;
      }

      @media(max-width:1100px){
        .results-status-overview{
          grid-template-columns:1fr;
        }

        .results-operational-strip{
          grid-template-columns:1fr 1fr;
        }
      }

      @media(max-width:700px){
        .status-breakdown-grid,
        .status-breakdown-grid.surgery,
        .results-operational-strip{
          grid-template-columns:1fr 1fr;
        }

        .status-summary-head{
          flex-direction:column;
        }
      }

      @media(max-width:520px){
        .status-breakdown-grid,
        .status-breakdown-grid.surgery,
        .results-operational-strip{
          grid-template-columns:1fr;
        }
      }
    `}</style>

    <style>{`
      .pdf-page-block{display:block}
      .pdf-page-block + .pdf-page-block{margin-top:0}
      .pdf-continuation-grid{margin-top:24px}
      .results-section-first{margin-bottom:24px}
      .btn:disabled{opacity:.65;cursor:not-allowed}
      @media(max-width:900px){
        .results-header .top-actions{width:100%;flex-wrap:wrap}
        .results-header .top-actions .btn{flex:1}
      }
    `}</style>
    <style>{`
      .bulk-delete-btn{
        white-space:nowrap;
        border-color:#efd0d0!important;
        background:#fff5f5!important;
        color:#a84f4f!important;
      }
      .bulk-delete-btn:hover:not(:disabled){
        background:#fbe9e9!important;
        border-color:#e4bcbc!important;
      }
      .bulk-delete-btn:disabled{
        opacity:.45;
        cursor:not-allowed;
      }
      @media(max-width:900px){
        .filters .bulk-delete-btn{
          width:100%;
        }
      }
    `}</style>

    <Modal open={examModal} onClose={closeExamModal} title={editingExamId?'Editar exame':'Novo exame'} subtitle="Selecione quem está fazendo este lançamento."  onSave={editingExamId?updateExam:addExam} saveText={editingExamId?'Salvar alterações':'Salvar exame'}>
      <div className="form-grid">
        <Field label="Data"><input type="date" value={examForm.date} onChange={e=>setExamForm({...examForm,date:e.target.value})}/></Field>
        <Field label="Paciente"><input value={examForm.patient} onChange={e=>setExamForm({...examForm,patient:e.target.value})}/></Field>
        <Field label="WhatsApp / contato (opcional)"><input type="tel" inputMode="tel" autoComplete="tel" value={examForm.whatsapp} onChange={e=>setExamForm({...examForm,whatsapp:e.target.value})} placeholder="(18) 99999-9999"/></Field>
        <Field label="Médico"><Select value={examForm.doctor_id} setValue={v=>setExamForm({...examForm,doctor_id:v})} options={doctors} first="Selecione"/></Field>
        <Field label="Indicado por (opcional)"><Select value={examForm.indicated_by_doctor_id||''} setValue={v=>setExamForm({...examForm,indicated_by_doctor_id:v})} options={doctors} first="Não informado"/></Field>
        <Field label={editingExamId ? "Exame" : "Exames"} full>
          <div className="exam-multi-select">
            <div className="exam-multi-head">
              <span>Selecione um ou vários exames para este lançamento</span>
              <b>{examForm.exam_type_ids?.length||0} selecionado(s)</b>
            </div>
            <div className="exam-options">
              {examTypes.map(exam=>{
                const checked=examForm.exam_type_ids?.includes(exam.id);
                return <label key={exam.id} className={`exam-option ${checked?'selected':''}`}>
                  <input type="checkbox" checked={checked} onChange={()=>{
                    const ids=examForm.exam_type_ids||[];
                    setExamForm({...examForm,exam_type_ids:checked?ids.filter(id=>id!==exam.id):[...ids,exam.id]});
                  }}/>
                  <span>{exam.name}</span>
                </label>;
              })}
            </div>
          </div>
        </Field>
        <Field label="Status"><Select value={examForm.status} setValue={v=>setExamForm({...examForm,status:v})} options={['Agendado','Pendente','Realizado','Cancelado']}/></Field>
        <Field label="Lançado por">
          <Select
            value={examForm.launcher || currentLaunchName || ''}
            setValue={v=>setExamForm({...examForm,launcher:v})}
            options={launcherOptionsAll}
            first="Selecione quem lançou"
          />
        </Field>
        <Field label="Observação" full><textarea value={examForm.obs} onChange={e=>setExamForm({...examForm,obs:e.target.value})} rows="4"/></Field>
      </div>
    </Modal>

    <style>{`
      .status.pendente{
        background:#fff3cd!important;
        color:#8a6518!important;
        border:1px solid #f0d98a!important;
      }
    `}</style>

    <style>{`
      .procedure-multi-select{
        border:1px solid #e3e9ed;
        border-radius:12px;
        padding:12px;
        background:#fbfcfd;
      }
      .procedure-multi-head{
        display:flex;
        justify-content:space-between;
        align-items:center;
        gap:12px;
        margin-bottom:10px;
        color:#53636e;
        font-size:13px;
      }
      .procedure-multi-head b{
        color:#315f72;
      }
      .procedure-options{
        display:grid;
        grid-template-columns:repeat(2,minmax(0,1fr));
        gap:8px;
        max-height:250px;
        overflow:auto;
      }
      .procedure-option{
        display:flex!important;
        align-items:center;
        gap:9px;
        margin:0!important;
        padding:10px 11px;
        border:1px solid #e3e9ed;
        border-radius:9px;
        background:#fff;
        cursor:pointer;
      }
      .procedure-option.selected{
        border-color:#477f91;
        background:#eaf2f5;
        color:#315f72;
        font-weight:700;
      }
      .procedure-option input{
        width:auto;
        margin:0;
        accent-color:#315f72;
      }
      @media(max-width:700px){
        .procedure-options{grid-template-columns:1fr}
      }
    `}</style>

    <Modal open={surgeryModal} onClose={closeSurgeryModal} title={editingSurgeryId?'Editar cirurgia':'Nova cirurgia'} subtitle="Selecione quem está fazendo este lançamento."  onSave={editingSurgeryId?updateSurgery:addSurgery} saveText={editingSurgeryId?'Salvar alterações':'Salvar cirurgia'}>
      <div className="form-grid">
        <Field label="Data"><input type="date" value={sForm.date} onChange={e=>setSForm({...sForm,date:e.target.value})}/></Field>
        <Field label="Paciente"><input value={sForm.patient} onChange={e=>setSForm({...sForm,patient:e.target.value})}/></Field>
        <Field label="WhatsApp / contato (opcional)"><input type="tel" inputMode="tel" autoComplete="tel" value={sForm.whatsapp} onChange={e=>setSForm({...sForm,whatsapp:e.target.value})} placeholder="(18) 99999-9999"/></Field>
        <Field label="Cirurgião"><Select value={sForm.doctor_id} setValue={v=>setSForm({...sForm,doctor_id:v})} options={doctors} first="Selecione"/></Field>
        <Field label="Indicado por (opcional)"><Select value={sForm.indicated_by_doctor_id||''} setValue={v=>setSForm({...sForm,indicated_by_doctor_id:v})} options={doctors} first="Não informado"/></Field>
        <Field label={editingSurgeryId ? "Procedimento" : "Procedimentos"} full>
          <div className="procedure-multi-select">
            <div className="procedure-multi-head">
              <span>Selecione um ou vários procedimentos para este lançamento</span>
              <b>{sForm.procedure_ids?.length||0} selecionado(s)</b>
            </div>

            <div className="procedure-options">
              {procedures.map(procedure=>{
                const checked=sForm.procedure_ids?.includes(procedure.id);

                return <label key={procedure.id} className={`procedure-option ${checked?'selected':''}`}>
                  <input
                    type="checkbox"
                    checked={checked}
                    onChange={()=>{
                      const ids=sForm.procedure_ids||[];
                      setSForm({
                        ...sForm,
                        procedure_ids:checked
                          ? ids.filter(id=>id!==procedure.id)
                          : [...ids,procedure.id]
                      });
                    }}
                  />
                  <span>{procedure.name}</span>
                </label>;
              })}
            </div>
          </div>
        </Field>
        <Field label="Olho"><Select value={sForm.eye} setValue={v=>setSForm({...sForm,eye:v})} options={['Não se aplica','Direito (OD)','Esquerdo (OE)','Ambos']}/></Field>
        <Field label="Convênio"><Select value={sForm.insurance_id} setValue={v=>setSForm({...sForm,insurance_id:v})} options={insurances} first="Particular / selecione"/></Field>
        <Field label="Status"><Select value={sForm.status} setValue={v=>setSForm({...sForm,status:v})} options={['Solicitação','Pendente','Autorizada','Agendada','Realizada','Cancelada']}/></Field>
        <Field label="Lançado por">
          <Select
            value={sForm.launcher || currentLaunchName || ''}
            setValue={v=>setSForm({...sForm,launcher:v})}
            options={launcherOptionsAll}
            first="Selecione quem lançou"
          />
        </Field>
        <Field label="Horário de chegada"><input type="time" value={sForm.arrival} onChange={e=>setSForm({...sForm,arrival:e.target.value})}/></Field>
        <Field label="Horário da cirurgia"><input type="time" value={sForm.time} onChange={e=>setSForm({...sForm,time:e.target.value})}/></Field>
        <Field label="Pagamento"><Select value={sForm.payment} setValue={v=>setSForm({...sForm,payment:v})} options={['Não informado','Pendente','20% pago','Pago','Convênio']}/></Field>
        <Field label="Observações" full><textarea value={sForm.obs} onChange={e=>setSForm({...sForm,obs:e.target.value})} rows="5"/></Field>
      </div>
    </Modal>

    <style>{`
      .followup-metrics-card{padding:20px}
      .followup-metrics-head{display:flex;align-items:flex-start;justify-content:space-between;gap:18px}
      .followup-metrics-head>div:first-child>span{display:block;color:#315f72;font-size:10px;font-weight:800;letter-spacing:.08em}
      .followup-metrics-head h3{margin:5px 0 3px;color:#20313b;font-size:18px}
      .followup-metrics-head p{margin:0;color:#74838c;font-size:12px}
      .followup-rate-box{flex:0 0 150px;padding:13px 16px;border-radius:12px;background:#edf6f0;text-align:center}
      .followup-rate-box strong{display:block;color:#417257;font-size:26px;line-height:1}
      .followup-rate-box span{display:block;margin-top:5px;color:#6f8277;font-size:10px;font-weight:700;text-transform:uppercase}
      .followup-progress-wrap{margin-top:18px}
      .followup-progress-label{display:flex;justify-content:space-between;gap:12px;margin-bottom:7px;color:#687a83;font-size:11px}
      .followup-progress-label b{color:#315f72}
      .followup-progress{height:9px;border-radius:999px;background:#eaf0f2;overflow:hidden}
      .followup-progress i{display:block;height:100%;min-width:0;border-radius:999px;background:#4f8a6b;transition:width .2s ease}
      .followup-metrics-grid{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px;margin-top:16px}
      .followup-metrics-grid>div{min-width:0;padding:12px 13px;border:1px solid #e5ecef;border-radius:11px;background:#fafcfc}
      .followup-metrics-grid>div.success{background:#f2f9f4;border-color:#d8ebde}
      .followup-metrics-grid span{display:block;color:#7d8c93;font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.04em}
      .followup-metrics-grid b{display:block;margin:4px 0 2px;color:#315f72;font-size:21px}
      .followup-metrics-grid small{display:block;color:#84939a;font-size:10px;line-height:1.3}
      .followup-modal-contact{display:flex;align-items:center;justify-content:space-between;gap:14px;padding:14px 15px;margin-bottom:14px;border:1px solid #cfe5d6;border-radius:12px;background:#f4fbf6}
      .followup-modal-contact>div{min-width:0}
      .followup-modal-contact span{display:block;color:#73877a;font-size:9px;font-weight:800;letter-spacing:.06em;text-transform:uppercase}
      .followup-modal-contact b{display:block;margin-top:4px;color:#2f5f45;font-size:14px;overflow-wrap:anywhere}
      .followup-modal-contact .btn{flex:0 0 auto}
      .followup-page{display:flex;flex-direction:column;gap:18px}
      .followup-kpis{grid-template-columns:repeat(4,minmax(0,1fr))}
      .followup-toolbar{padding:18px}
      .followup-filters{display:grid;grid-template-columns:repeat(5,minmax(150px,1fr));gap:12px;align-items:end}
      .followup-filters label{display:flex;flex-direction:column;gap:6px;margin:0;color:#687a83;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.04em}
      .followup-filters input,.followup-filters select{width:100%;min-width:0;text-transform:none;letter-spacing:0;font-weight:500}
      .followup-search{grid-column:span 2}
      .followup-filter-actions{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-top:14px;padding-top:14px;border-top:1px solid #e6ecef;color:#74838c;font-size:12px}
      .followup-list{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}
      .followup-card{padding:18px;min-width:0;overflow:hidden}
      .followup-card.due-today{box-shadow:0 0 0 2px #d9b45b inset}
      .followup-card-head{display:flex;align-items:flex-start;justify-content:space-between;gap:14px;padding-bottom:14px;border-bottom:1px solid #edf1f3}
      .followup-card-head h3{margin:8px 0 3px;color:#20313b;font-size:17px}
      .followup-card-head p{margin:0;color:#6e7f87;font-size:13px;line-height:1.4;overflow-wrap:anywhere}
      .followup-contact-box{flex:0 0 auto;min-width:150px;padding:10px 12px;border:1px solid #dce8e2;border-radius:11px;background:#f5fbf7;text-align:right}
      .followup-contact-box span{display:block;margin-bottom:3px;color:#6c8b79;font-size:8px;font-weight:800;letter-spacing:.06em}
      .followup-contact-box b{display:block;color:#2f5f45;font-size:12px;white-space:nowrap}
      .followup-badges{display:flex;flex-wrap:wrap;gap:6px;align-items:center}
      .source-badge,.follow-status,.follow-due{display:inline-flex;align-items:center;min-height:23px;padding:4px 8px;border-radius:999px;font-size:9px;font-weight:800;letter-spacing:.04em}
      .source-badge.exam{background:#eaf2f5;color:#315f72}.source-badge.surgery{background:#edf6f0;color:#417257}
      .follow-status{background:#f2f5f6;color:#5c6d75}.follow-status.concluido{background:#edf6f0;color:#417257}.follow-status.sem-retorno{background:#fbeeee;color:#a64c4c}.follow-status.aguardando-paciente{background:#fff5df;color:#976820}.follow-status.em-contato{background:#eaf2f5;color:#315f72}
      .follow-due{background:#fff5df;color:#8d641f}
      .followup-info-grid{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:8px;margin:14px 0}
      .followup-info-grid>div{min-width:0;padding:10px;border-radius:10px;background:#f8fafb}
      .followup-info-grid span,.followup-note span{display:block;margin-bottom:4px;color:#84939a;font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.04em}
      .followup-info-grid b{display:block;color:#31444d;font-size:12px;line-height:1.35;overflow-wrap:anywhere}
      .followup-note{padding:12px;border:1px solid #e5ecef;border-radius:10px;background:#fff}
      .followup-note p{margin:0;color:#53636e;font-size:12px;line-height:1.5;white-space:pre-wrap;overflow-wrap:anywhere}
      .followup-actions{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px}
      .btn-whatsapp-followup{display:inline-flex;align-items:center;justify-content:center;gap:7px;background:#eaf7ef!important;border-color:#bfe1cb!important;color:#247443!important;text-decoration:none!important;font-weight:800}
      .btn-whatsapp-followup:hover{background:#dcf1e4!important;border-color:#a9d5b8!important}
      .btn-whatsapp-followup svg{width:17px;height:17px;flex:0 0 17px}
      .followup-empty{grid-column:1/-1;padding:28px;text-align:center}
      .followup-form-grid textarea{resize:vertical;min-height:110px}
      @media(max-width:1180px){.followup-metrics-grid{grid-template-columns:repeat(3,minmax(0,1fr))}.followup-filters{grid-template-columns:repeat(3,minmax(0,1fr))}.followup-search{grid-column:span 1}.followup-info-grid{grid-template-columns:repeat(3,minmax(0,1fr))}}
      @media(max-width:900px){.followup-metrics-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.followup-list{grid-template-columns:1fr}.followup-kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.followup-filters{grid-template-columns:repeat(2,minmax(0,1fr))}.followup-info-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
      @media(max-width:620px){.followup-metrics-head{flex-direction:column}.followup-rate-box{width:100%;flex-basis:auto}.followup-metrics-grid,.followup-kpis,.followup-filters,.followup-info-grid{grid-template-columns:1fr}.followup-modal-contact{align-items:stretch;flex-direction:column}.followup-modal-contact .btn{width:100%}.followup-card{padding:14px}.followup-card-head{align-items:stretch;flex-direction:column}.followup-contact-box{text-align:left;width:100%;min-width:0}.followup-filter-actions{align-items:stretch;flex-direction:column}.followup-actions .btn,.followup-actions a.btn{flex:1 1 100%}}
    `}</style>

    <Modal open={followupModal} onClose={()=>{setFollowupModal(false);setFollowupTarget(null)}} title={followupTarget?.followup?'Atualizar follow-up':'Registrar follow-up'} subtitle={followupTarget ? `${followupTarget.patient} • ${followupTarget.source_type==='exam'?'Exame':'Cirurgia'} • ${brDate(followupTarget.source_date)}` : ''} onSave={saveFollowup} saveText="Salvar follow-up">
      {followupTarget&&<div className="followup-modal-contact">
        <div>
          <span>Contato do paciente</span>
          <b>{whatsappDisplay(followupTarget.whatsapp)}</b>
        </div>
        {followupTarget.whatsapp ? <a className="btn btn-whatsapp-followup" href={followupWhatsappUrl(followupTarget.whatsapp,followupTarget.patient)} target="_blank" rel="noreferrer" title={`Conversar com ${followupTarget.patient} no WhatsApp`}>
          <WhatsAppIcon/><span>Enviar WhatsApp</span>
        </a> : <button className="btn btn-light" disabled>WhatsApp não informado</button>}
      </div>}
      <div className="form-grid followup-form-grid">
        <Field label="Data do contato"><input type="date" value={followupForm.followup_date} onChange={e=>setFollowupForm({...followupForm,followup_date:e.target.value})}/></Field>
        <Field label="Status do follow-up"><Select value={followupForm.status} setValue={v=>setFollowupForm({...followupForm,status:v})} options={['A fazer','Em contato','Aguardando paciente','Sem retorno','Concluído']}/></Field>
        <Field label="Próximo contato"><input type="date" value={followupForm.next_contact_date} onChange={e=>setFollowupForm({...followupForm,next_contact_date:e.target.value})}/></Field>
        <Field label="Responsável"><Select value={followupForm.responsible} setValue={v=>setFollowupForm({...followupForm,responsible:v})} options={launcherOptionsAll} first="Selecione"/></Field>
        <Field label="Observação do follow-up" full><textarea rows="5" value={followupForm.note} onChange={e=>setFollowupForm({...followupForm,note:e.target.value})} placeholder="Ex.: paciente pediu retorno na sexta, aguarda autorização, não respondeu WhatsApp..."/></Field>
      </div>
    </Modal>

    <Modal open={catalogModal} onClose={()=>setCatalogModal(false)} title="Cadastros da clínica" subtitle="Disponível para ADM e Orientadora." onSave={()=>setCatalogModal(false)} saveText="Concluir">
      <Catalog title="Médicos / profissionais" items={doctors} onAdd={v=>addCatalog('doctors',v)} onRemove={id=>deactivateCatalog('doctors',id)} />
      <Catalog title="Exames" items={examTypes} onAdd={v=>addCatalog('exam_types',v)} onRemove={id=>deactivateCatalog('exam_types',id)} />
      <Catalog title="Cirurgias / procedimentos" items={procedures} onAdd={v=>addCatalog('procedures',v,{category:'Cirurgia'})} onRemove={id=>deactivateCatalog('procedures',id)} />
    </Modal>
  </div>;
}

function Catalog({title,description,placeholder='Adicionar novo',items,onAdd,onRemove}){
  const [value,setValue]=useState('');

  const submit = async ()=>{
    const clean=value.trim();
    if(!clean) return;
    await onAdd(clean);
    setValue('');
  };

  return <div className="catalog-card catalog-card-large">
    <div className="catalog-card-head">
      <div>
        <h3>{title}</h3>
        {description&&<p>{description}</p>}
      </div>
      <span className="catalog-count">{items.length}</span>
    </div>

    <div className="inline-add">
      <input
        value={value}
        onChange={e=>setValue(e.target.value)}
        onKeyDown={e=>{if(e.key==='Enter'){e.preventDefault();submit();}}}
        placeholder={placeholder}
      />
      <button className="btn btn-primary" onClick={submit} disabled={!value.trim()}>+ Adicionar</button>
    </div>

    <div className="catalog-list">
      {items.length ? items.map(x=>
        <div className="catalog-item" key={x.id}>
          <span>{x.name}</span>
          <button title={`Remover ${x.name}`} onClick={()=>onRemove(x.id)}>×</button>
        </div>
      ) : <div className="catalog-empty">Nenhum cadastro ativo.</div>}
    </div>
  </div>;
}

createRoot(document.getElementById('root')).render(<App/>);
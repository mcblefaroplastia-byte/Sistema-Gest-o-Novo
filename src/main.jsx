import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import ExcelJS from 'exceljs';
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

  const [examModal,setExamModal] = useState(false);
  const [surgeryModal,setSurgeryModal] = useState(false);
  const [catalogModal,setCatalogModal] = useState(false);
  const [editingExamId,setEditingExamId] = useState(null);
  const [editingSurgeryId,setEditingSurgeryId] = useState(null);

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

  const [examForm,setExamForm] = useState({
    date:todayISO(),patient:'',whatsapp:'',doctor_id:'',exam_type_ids:[],status:'Agendado',launcher:'',obs:''
  });
  const [sForm,setSForm] = useState({
    date:todayISO(),patient:'',whatsapp:'',doctor_id:'',procedure_ids:[],eye:'Não se aplica',
    insurance_id:'',status:'Solicitação',arrival:'',time:'',payment:'Não informado',launcher:'',obs:''
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
      insurancesRes, profilesRes, examsRes, surgeriesRes
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
        doctors(id,name),
        exam_types(id,name)
      `).order('exam_date',{ascending:false}),
      supabase.from('surgeries').select(`
        id,launch_group_id,surgery_date,eye,status,arrival_time,surgery_time,payment_status,observation,launched_by,launched_by_name,
        patients(id,full_name,whatsapp),
        doctors(id,name),
        procedures(id,name),
        insurances(id,name)
      `).order('surgery_date',{ascending:false})
    ]);

    const errors=[profileRes,doctorsRes,examTypesRes,proceduresRes,insurancesRes,profilesRes,examsRes,surgeriesRes]
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
    setLoading(false);
  };

  const getOrCreatePatient = async (name,whatsapp)=>{
    const cleanName=name.trim();
    const cleanPhone=whatsapp.trim();

    if(cleanPhone){
      const {data:byPhone}=await supabase.from('patients').select('id').eq('whatsapp',cleanPhone).limit(1).maybeSingle();
      if(byPhone?.id) return byPhone.id;
    }
    const {data:byName}=await supabase.from('patients').select('id').ilike('full_name',cleanName).limit(1).maybeSingle();
    if(byName?.id) return byName.id;

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
      setSForm({date:todayISO(),patient:'',whatsapp:'',doctor_id:'',procedure_ids:[],eye:'Não se aplica',insurance_id:'',status:'Solicitação',arrival:'',time:'',payment:'Não informado',launcher:currentLaunchName||'',obs:''});
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
    setSForm({date:todayISO(),patient:'',whatsapp:'',doctor_id:'',procedure_ids:[],eye:'Não se aplica',insurance_id:'',status:'Solicitação',arrival:'',time:'',payment:'Não informado',launcher:currentLaunchName||'',obs:''});
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
    const realEx=totalItemsByStatus(dashboardExams,'exam_type_ids','Realizado');
    const realSu=totalItemsByStatus(dashboardSurgeries,'procedure_ids','Realizada');
    const examTop=Object.entries(countManyBy(dashboardExams,'exam_names')).sort((a,b)=>b[1]-a[1])[0];
    const surgTop=Object.entries(countManyBy(dashboardSurgeries,'procedure_names')).sort((a,b)=>b[1]-a[1])[0];
    const doctorTop=Object.entries(countBy([...dashboardExams,...dashboardSurgeries],'doctor')).sort((a,b)=>b[1]-a[1])[0];
    const uniquePatients=new Set([...dashboardExams,...dashboardSurgeries].map(x=>`${x.patient}|${normalizePhone(x.whatsapp)}`)).size;
    return {totalEx,totalSu,realEx,realSu,examTop,surgTop,doctorTop,uniquePatients,
      exRate:totalEx?Math.round(realEx/totalEx*100):0,
      suRate:totalSu?Math.round(realSu/totalSu*100):0};
  },[dashboardExams,dashboardSurgeries]);

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
    const wsSummary = workbook.addWorksheet('Resumo Executivo',{
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
      ['Exames por tipo',countManyBy(dashboardExams,'exam_names'),'A5:F20'],
      ['Exames por médico',countBy(dashboardExams,'doctor'),'G5:L20'],
      ['Cirurgias por procedimento',countManyBy(dashboardSurgeries,'procedure_names'),'A22:F37'],
      ['Lançamentos por responsável',Object.fromEntries(userRows.map(u=>[u.name,u.exams+u.surgeries])),'G22:L37']
    ];

    for(const [title,data,range] of charts){
      const png=await makeBarChartImage(title,data);
      const imageId=workbook.addImage({base64:png,extension:'png'});
      const [from,to]=range.split(':');
      wsCharts.addImage(imageId,{tl:{col:wsCharts.getCell(from).col-1,row:wsCharts.getCell(from).row-1},
                                 br:{col:wsCharts.getCell(to).col,row:wsCharts.getCell(to).row}});
    }

    wsCharts.mergeCells('A39:L39');
    wsCharts.getCell('A39').value='Leitura rápida';
    styleSection(wsCharts.getCell('A39'));
    wsCharts.mergeCells('A40:L43');
    wsCharts.getCell('A40').value =
      `Neste período foram registrados ${resultSummary.totalEx} exames e ${resultSummary.totalSu} cirurgias, `+
      `com ${resultSummary.uniquePatients} pacientes únicos. A taxa de realização foi de ${resultSummary.exRate}% nos exames `+
      `e ${resultSummary.suRate}% nas cirurgias. O exame com maior volume foi ${resultSummary.examTop?.[0]||'—'} `+
      `e o procedimento cirúrgico com maior volume foi ${resultSummary.surgTop?.[0]||'—'}.`;
    wsCharts.getCell('A40').alignment={wrapText:true,vertical:'top'};
    wsCharts.getCell('A40').font={size:11,color:{argb:COLORS.text}};
    wsCharts.getCell('A40').fill={type:'pattern',pattern:'solid',fgColor:{argb:COLORS.soft}};
    wsCharts.getCell('A40').border=border;

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
          <div className="table-wrap"><table><thead><tr><th>Data</th><th>Paciente</th><th>Médico</th><th>Exame</th><th>Status</th><th>Observação</th><th>Lançado por</th><th>Ações</th></tr></thead>
          <tbody>{filteredExams.length?filteredExams.map(r=><tr key={r.id}>
            <td>{brDate(r.date)}</td><td><PatientCell name={r.patient} phone={r.whatsapp}/></td>
            <td>{r.doctor}</td><td>{r.exam}</td><td><span className={`status ${statusClass(r.status)}`}>{r.status}</span></td>
            <td><ObservationCell text={r.obs}/></td><td>{launchedName(r)}</td>
            <td><div className="row-actions">
              <button className="btn btn-edit-soft" onClick={()=>openEditExam(r)}>Editar</button>
              <button className="btn btn-danger-soft" onClick={()=>delExam(r.id)}>Excluir</button>
            </div></td>
          </tr>):<tr><td colSpan="8" className="empty">Nenhum exame encontrado.</td></tr>}</tbody></table></div>
        </div>
      </section>}

      {page==='cirurgias'&&<section>
        <div className="topbar"><div><div className="eyebrow">CENTRO CIRÚRGICO</div><h1>Controle de cirurgias</h1><div className="subtitle">Programação, autorização, contato e observações.</div></div>
          <div className="top-actions">
            <button className="btn btn-light" onClick={()=>setPage('cadastros')}>⚙️ Cadastros</button>
            <button className="btn btn-primary" onClick={()=>{setEditingSurgeryId(null);setSForm({date:todayISO(),patient:'',whatsapp:'',doctor_id:'',procedure_ids:[],eye:'Não se aplica',insurance_id:'',status:'Solicitação',arrival:'',time:'',payment:'Não informado',launcher:currentLaunchName||'',obs:''});setSurgeryModal(true)}}>+ Nova cirurgia</button>
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
          <div className="table-wrap"><table><thead><tr><th>Data</th><th>Paciente</th><th>Cirurgião</th><th>Procedimento</th><th>Olho</th><th>Convênio</th><th>Status</th><th>Observação</th><th>Lançado por</th><th>Ações</th></tr></thead>
          <tbody>{filteredSurgeries.length?filteredSurgeries.map(r=><tr key={r.id}>
            <td>{brDate(r.date)}{r.time&&<><br/><small>{r.time}</small></>}</td><td><PatientCell name={r.patient} phone={r.whatsapp}/></td>
            <td>{r.doctor}</td><td>{r.procedure}</td><td>{r.eye}</td><td>{r.insurance}</td>
            <td><span className={`status ${statusClass(r.status)}`}>{r.status}</span></td>
            <td><ObservationCell text={r.obs}/></td><td>{launchedName(r)}</td>
            <td><div className="row-actions">
              <button className="btn btn-edit-soft" onClick={()=>openEditSurgery(r)}>Editar</button>
              <button className="btn btn-danger-soft" onClick={()=>delSurgery(r.id)}>Excluir</button>
            </div></td>
          </tr>):<tr><td colSpan="10" className="empty">Nenhuma cirurgia encontrada.</td></tr>}</tbody></table></div>
        </div>
      </section>}

      {page==='resultados'&&<section>
        <div className="topbar results-header">
          <div>
            <div className="eyebrow">INTELIGÊNCIA OPERACIONAL</div>
            <h1>Resultados e análise mensal</h1>
            <div className="subtitle">Visão consolidada da produção de exames, cirurgias, pacientes e responsáveis pelos lançamentos.</div>
          </div>
          <button className="btn btn-excel" onClick={exportExcel}>⬇ Exportar relatório Excel</button>
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

        <div className="results-highlights">
          <div className="card insight-card"><span>Exame com maior volume</span><b>{resultSummary.examTop?.[0]||'Sem dados'}</b><small>{resultSummary.examTop?.[1]||0} lançamentos</small></div>
          <div className="card insight-card"><span>Cirurgia com maior volume</span><b>{resultSummary.surgTop?.[0]||'Sem dados'}</b><small>{resultSummary.surgTop?.[1]||0} lançamentos</small></div>
          <div className="card insight-card"><span>Profissional com maior volume</span><b>{resultSummary.doctorTop?.[0]||'Sem dados'}</b><small>{resultSummary.doctorTop?.[1]||0} registros</small></div>
        </div>

        <div className="results-section">
          <div className="results-section-title"><div><h3>Análise de exames</h3><p>Distribuição dos exames no período selecionado.</p></div><div className="pill">{resultSummary.totalEx} exames</div></div>
          <div className="grid">
            <div className="card chart-card"><h2>Exames por tipo</h2><Bars data={countManyBy(dashboardExams,'exam_names')}/></div>
            <div className="card chart-card"><h2>Exames por médico</h2><Bars data={countBy(dashboardExams,'doctor')}/></div>
          </div>
          <div className="grid">
            <div className="card chart-card"><h2>Status dos exames</h2><Bars data={countBy(dashboardExams,'status')}/></div>
            <div className="card chart-card"><h2>Exames por responsável</h2><Bars data={countBy(dashboardExams.map(x=>({...x,launcher:launchedName(x)})),'launcher')}/></div>
          </div>
        </div>

        <div className="results-section">
          <div className="results-section-title"><div><h3>Análise de cirurgias</h3><p>Produção, procedimentos, status e convênios.</p></div><div className="pill">{resultSummary.totalSu} cirurgias</div></div>
          <div className="grid">
            <div className="card chart-card"><h2>Cirurgias por procedimento</h2><Bars data={countManyBy(dashboardSurgeries,'procedure_names')}/></div>
            <div className="card chart-card"><h2>Cirurgias por cirurgião</h2><Bars data={countBy(dashboardSurgeries,'doctor')}/></div>
          </div>
          <div className="grid">
            <div className="card chart-card"><h2>Status das cirurgias</h2><Bars data={countBy(dashboardSurgeries,'status')}/></div>
            <div className="card chart-card"><h2>Cirurgias por convênio</h2><Bars data={countBy(dashboardSurgeries.map(x=>({...x,insuranceLabel:x.insurance||'Particular'})),'insuranceLabel')}/></div>
          </div>
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
            <p>Acesse <b>Exames → Novo exame</b>. Preencha data, paciente, WhatsApp, médico, tipo de exame, status e observação. Você pode selecionar vários exames no mesmo lançamento. Depois clique em <b>Salvar exame</b>.</p>
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
            <p>A página Resultados mostra pacientes únicos, volumes, taxas de realização, exames e cirurgias mais registrados, produção por médico, status, convênio e lançamentos por responsável.</p>
            <div className="manual-note">Os indicadores acompanham os filtros selecionados. Assim você pode analisar um mês, profissional, responsável ou status específico.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">09</div>
            <h3>Exportar para Excel</h3>
            <p>Na página Resultados, aplique os filtros desejados e clique em <b>Exportar relatório Excel</b>.</p>
            <p>O arquivo é gerado com cinco abas: <b>Resumo Executivo, Análise Gráfica, Exames, Cirurgias e Lançamentos</b>, com indicadores, cores, tabelas formatadas e gráficos para facilitar a análise mensal.</p>
          </article>

          <article className="card manual-card">
            <div className="manual-number">10</div>
            <h3>Cadastros da clínica</h3>
            <p>Tanto no acesso <b>ADM</b> quanto no acesso <b>Orientadora</b>, abra <b>Cadastros</b> no menu lateral. Nessa página é possível adicionar ou remover <b>médicos, exames e cirurgias/procedimentos</b>. Os novos itens passam a aparecer automaticamente nos formulários de lançamento.</p><div className="manual-note">Itens removidos deixam de aparecer para novos lançamentos, mas os registros antigos permanecem preservados.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">11</div>
            <h3>WhatsApp do paciente</h3>
            <p>Ao informar o WhatsApp no cadastro, o ícone aparece ao lado do paciente. Clique nele para abrir rapidamente a conversa no WhatsApp.</p>
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
        <Field label="WhatsApp"><input value={examForm.whatsapp} onChange={e=>setExamForm({...examForm,whatsapp:e.target.value})} placeholder="(18) 99999-9999"/></Field>
        <Field label="Médico"><Select value={examForm.doctor_id} setValue={v=>setExamForm({...examForm,doctor_id:v})} options={doctors} first="Selecione"/></Field>
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
        <Field label="WhatsApp"><input value={sForm.whatsapp} onChange={e=>setSForm({...sForm,whatsapp:e.target.value})}/></Field>
        <Field label="Cirurgião"><Select value={sForm.doctor_id} setValue={v=>setSForm({...sForm,doctor_id:v})} options={doctors} first="Selecione"/></Field>
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
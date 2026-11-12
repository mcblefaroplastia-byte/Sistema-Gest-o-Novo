import React, { useState } from 'react';
import { Download, Eye, Trash2, Upload } from 'lucide-react';
import { supabase } from '../supabase';
import { brDate, whatsappUrl } from '../utils/helpers';

export function WhatsAppIcon(){
  return <svg viewBox="0 0 32 32" aria-hidden="true"><path fill="currentColor" d="M19.11 17.47c-.26-.13-1.52-.75-1.76-.84-.24-.09-.41-.13-.59.13-.17.26-.67.84-.82 1.02-.15.17-.3.2-.56.07-.26-.13-1.09-.4-2.08-1.29-.77-.68-1.29-1.53-1.44-1.79-.15-.26-.02-.4.11-.53.12-.12.26-.3.39-.45.13-.15.17-.26.26-.43.09-.17.04-.32-.02-.45-.07-.13-.59-1.42-.8-1.94-.21-.51-.43-.44-.59-.45h-.5c-.17 0-.45.07-.69.32-.24.26-.91.89-.91 2.17 0 1.27.93 2.5 1.06 2.67.13.17 1.82 2.78 4.41 3.9.62.27 1.1.43 1.47.55.62.2 1.18.17 1.62.1.49-.07 1.52-.62 1.73-1.22.21-.6.21-1.12.15-1.22-.07-.11-.24-.17-.5-.3Z"/><path fill="currentColor" d="M16.03 3C8.84 3 3 8.75 3 15.85c0 2.26.6 4.47 1.74 6.42L3 29l6.93-1.8a13.16 13.16 0 0 0 6.09 1.48h.01c7.18 0 13.02-5.76 13.02-12.84C29.05 8.75 23.21 3 16.03 3Zm0 23.51h-.01a10.94 10.94 0 0 1-5.57-1.51l-.4-.24-4.11 1.07 1.1-3.96-.26-.41a10.58 10.58 0 0 1-1.65-5.61C5.13 9.95 10.02 5.16 16.03 5.16S26.92 9.95 26.92 15.85c0 5.88-4.89 10.66-10.89 10.66Z"/></svg>;
}

export function PeriodFilter({mode,setMode,month,setMonth,date,setDate}){
  return <div className="period-filter">
    <select value={mode} onChange={e=>{
      const next=e.target.value;setMode(next);
      if(next==='month') setDate(''); else setMonth('');
    }}>
      <option value="month">Mês</option><option value="date">Data</option>
    </select>
    {mode==='month'
      ? <input type="month" value={month} onChange={e=>setMonth(e.target.value)}/>
      : <input type="date" value={date} onChange={e=>setDate(e.target.value)}/>}
  </div>;
}

export function Modal({open,onClose,title,subtitle,children,onSave,saveText}){
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

export function PatientCell({name,phone,birthDate}){
  const url = whatsappUrl(phone);
  return <div className="patient-cell">
    <div><b>{name}</b>{birthDate&&<small>Nasc.: {brDate(birthDate)}</small>}{phone&&<small>{phone}</small>}</div>
    {url&&<a className="wa-btn" href={url} target="_blank" rel="noreferrer" title={`Abrir WhatsApp de ${name}`}><WhatsAppIcon/></a>}
  </div>;
}
export function ObservationCell({text}) {
  return text ? <div className="obs-cell" title={text}>{text}</div> : <span className="muted">—</span>;
}
export function LinkedProceduresCell({names=[]}){
  const items=(Array.isArray(names)?names:[]).filter(Boolean);
  if(!items.length) return <span className="muted">Sem procedimentos vinculados</span>;
  return <div className="linked-procedures-cell">
    <small>Vinculados ao pacote</small>
    <div>{items.map((name,index)=><span className="linked-procedure-chip" key={`${name}-${index}`}>{name}</span>)}</div>
  </div>;
}

export function ProcedureListCell({names=[]}){
  const items=(Array.isArray(names)?names:[]).filter(Boolean);
  if(!items.length) return <span className="muted">—</span>;
  return <div className="procedure-launch-cell">
    <small>{items.length} {items.length===1?'procedimento':'procedimentos'} neste lançamento</small>
    <div>{items.map((name,index)=><span className="linked-procedure-chip" key={`${name}-${index}`}>{name}</span>)}</div>
  </div>;
}
export function Kpi({label,value}){return <div className="kpi"><small>{label}</small><strong>{value}</strong></div>;}

export function Select({value,setValue,options,first}) {
  return <select value={value} onChange={e=>setValue(e.target.value)}>
    {first&&<option value="">{first}</option>}
    {options.map(o=>{
      const value = typeof o === 'string' ? o : o.id;
      const label = typeof o === 'string' ? o : o.name;
      return <option key={value} value={value}>{label}</option>
    })}
  </select>;
}
export function Field({label,children,full}){return <div className={full?'full':''}><label>{label}</label>{children}</div>;}

export function BrandMark(){
  const [failed,setFailed]=useState(false);
  return <div className={`brand-mark ${failed?'':'brand-logo-ready'}`} aria-label="OftalmoCastro">
    {failed
      ? <span>OC</span>
      : <img src="/logo-oftalmocastro.png" alt="OftalmoCastro" onError={()=>setFailed(true)}/>}
  </div>;
}

export function Login({onLogged}) {
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
      <BrandMark/>
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

export function FinancialDocumentsField({form,setForm,onOpen,onDownload}){
  const saved=Array.isArray(form.financial_documents)?form.financial_documents:[];
  const pending=Array.isArray(form.pending_documents)?form.pending_documents:[];
  const formatSize=bytes=>{
    const n=Number(bytes||0);
    if(!n) return '';
    if(n<1024*1024) return `${Math.max(1,Math.round(n/1024))} KB`;
    return `${(n/(1024*1024)).toFixed(1)} MB`;
  };
  const addFiles=e=>{
    const incoming=[...(e.target.files||[])];
    if(!incoming.length) return;
    const allowed=new Set(['application/pdf','image/png','image/jpeg','image/webp']);
    const maxSize=15*1024*1024;
    const valid=incoming.filter(file=>allowed.has(file.type)&&file.size<=maxSize);
    const invalidType=incoming.filter(file=>!allowed.has(file.type));
    const tooLarge=incoming.filter(file=>file.size>maxSize);
    if(invalidType.length) alert(`Alguns arquivos não foram adicionados. Formatos aceitos: PDF, PNG, JPG/JPEG e WEBP.`);
    if(tooLarge.length) alert(`Alguns arquivos ultrapassam 15 MB e não foram adicionados.`);
    setForm({...form,pending_documents:[...pending,...valid]});
    e.target.value='';
  };
  const removeSaved=doc=>{
    setForm({
      ...form,
      financial_documents:saved.filter(item=>item.path!==doc.path),
      removed_document_paths:[...new Set([...(form.removed_document_paths||[]),doc.path])]
    });
  };
  const removePending=index=>setForm({...form,pending_documents:pending.filter((_,i)=>i!==index)});
  return <div className="financial-documents full">
    <div className="documents-upload-zone">
      <div><b>Nota fiscal / comprovantes</b><small>Você pode anexar vários arquivos: PDF, PNG, JPG/JPEG ou WEBP.</small></div>
      <label className="btn btn-light document-picker"><Upload size={15}/> Selecionar arquivos<input type="file" multiple accept="application/pdf,image/png,image/jpeg,image/webp" onChange={addFiles}/></label>
    </div>
    {(saved.length||pending.length)?<div className="documents-list">
      {saved.map(doc=><div className="document-chip saved" key={doc.path}>
        <div className="document-icon">✓</div><div className="document-meta"><b>{doc.name||'Documento anexado'}</b><small>{formatSize(doc.size)}{doc.type?` • ${doc.type.replace('application/','').replace('image/','').toUpperCase()}`:''}</small></div>
        <div className="document-actions"><button type="button" className="btn btn-light" onClick={()=>onOpen(doc.path,doc.name)}><Eye size={14}/> Visualizar</button><button type="button" className="btn btn-light" onClick={()=>onDownload?.(doc.path,doc.name)}><Download size={14}/> Baixar</button><button type="button" className="btn btn-danger-soft" onClick={()=>removeSaved(doc)}><Trash2 size={14}/> Remover</button></div>
      </div>)}
      {pending.map((file,index)=><div className="document-chip pending" key={`${file.name}-${file.size}-${index}`}>
        <div className="document-icon">↑</div><div className="document-meta"><b>{file.name}</b><small>{formatSize(file.size)} • será enviado ao salvar</small></div>
        <div className="document-actions"><button type="button" className="btn btn-danger-soft" onClick={()=>removePending(index)}><Trash2 size={14}/> Remover</button></div>
      </div>)}
    </div>:<div className="documents-empty">Nenhum documento anexado.</div>}
  </div>;
}

export function Catalog({title,description,placeholder='Adicionar novo',items,onAdd,onRemove,onMove,moveLabel}){
  const [value,setValue]=useState('');

  const submit = async ()=>{
    const clean=value.trim();
    if(!clean) return;
    const saved=await onAdd(clean);
    if(saved!==false) setValue('');
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
          <div className="catalog-item-actions">
            {onMove&&<button className="catalog-move-btn" title={moveLabel||'Mover item'} onClick={()=>onMove(x.id)}>↔</button>}
            <button title={`Remover ${x.name}`} onClick={()=>onRemove(x.id)}>×</button>
          </div>
        </div>
      ) : <div className="catalog-empty">Nenhum cadastro ativo.</div>}
    </div>
  </div>;
}

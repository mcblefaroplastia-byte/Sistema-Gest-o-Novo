import React from 'react';
import { Field, Select } from './Common';
import {
  moneyBR,
  netFinancialAmount,
  normalizedBreakdown,
  numericValue,
  paymentTone,
  paymentToneLabel,
  todayISO
} from '../utils/helpers';

export function PaymentTrafficLight({record,onClick}){
  const tone=paymentTone(record);
  const label=paymentToneLabel(record);
  return <button type="button" className={`payment-light ${tone}`} onClick={onClick} title={`${label}. Clique para abrir o lançamento financeiro.`} aria-label={label}>
    <span className="payment-light-dot"/>
    <small>{label}</small>
  </button>;
}

export function PaymentBreakdownEditor({form,setForm,insurances}){
  const rows=normalizedBreakdown(form.payment_breakdown);
  const updateRows=next=>{
    const received=next.filter(x=>x.status==='Recebido').reduce((s,x)=>s+numericValue(x.amount),0);
    const payerTypes=[...new Set(next.map(x=>x.payer_type).filter(Boolean))];
    const methods=[...new Set(next.map(x=>x.method).filter(Boolean))];
    const net=Math.max(numericValue(form.total_amount)-numericValue(form.discount_amount),0);
    const receivedDates=next.filter(x=>x.status==='Recebido'&&x.received_date).map(x=>x.received_date).sort();
    const pendingDates=next.filter(x=>x.status!=='Recebido'&&x.expected_date).map(x=>x.expected_date).sort();
    const billed=next.filter(x=>x.payer_type==='Convênio').reduce((s,x)=>s+numericValue(x.amount),0);
    setForm({
      ...form,
      payment_breakdown:next,
      received_amount:String(received||''),
      payment_origin:payerTypes.length>1?'Misto':(payerTypes[0]||form.payment_origin||'Particular'),
      payment_method:methods.length>1?'Múltiplos':(methods[0]||form.payment_method||''),
      financial_status:net>0&&received>=net?'Recebido':received>0?'Parcial':(next.some(x=>x.payer_type==='Convênio')?'Aguardando convênio':'Pendente'),
      payment_confirmed:true,
      received_date:receivedDates.length?receivedDates[receivedDates.length-1]:(form.received_date||''),
      expected_receipt_date:pendingDates[0]||form.expected_receipt_date||'',
      billed_amount:String(billed||'')
    });
  };
  const addRow=()=>updateRows([...rows,{id:crypto.randomUUID(),payer_type:'Particular',insurance_id:'',method:'PIX',amount:'',installments:'1',status:'Pendente',expected_date:'',received_date:''}]);
  const setRow=(index,patch)=>updateRows(rows.map((r,i)=>i===index?{...r,...patch}:r));
  const removeRow=index=>updateRows(rows.filter((_,i)=>i!==index));
  return <div className="payment-breakdown full">
    <div className="payment-breakdown-head">
      <div><span>PAGAMENTOS</span><b>Como este valor será / foi recebido?</b><small>Adicione uma linha para cada pagamento. Em pagamento misto, registre separadamente convênio e particular.</small></div>
      <button type="button" className="btn btn-light" onClick={addRow}>+ Adicionar pagamento</button>
    </div>
    {rows.length?<div className="payment-breakdown-list">{rows.map((row,index)=><div className="payment-breakdown-row simplified" key={row.id||index}>
      <Field label="Pagador"><Select value={row.payer_type||'Particular'} setValue={v=>setRow(index,{payer_type:v,insurance_id:v==='Convênio'?row.insurance_id:''})} options={['Particular','Convênio']}/></Field>
      {row.payer_type==='Convênio'?<Field label="Convênio"><Select value={row.insurance_id||''} setValue={v=>setRow(index,{insurance_id:v})} options={insurances} first="Selecione"/></Field>:null}
      <Field label="Valor (R$)"><input type="number" min="0" step="0.01" value={row.amount||''} onChange={e=>setRow(index,{amount:e.target.value})}/></Field>
      <Field label="Forma"><Select value={row.method||''} setValue={v=>setRow(index,{method:v})} options={['PIX','Dinheiro','Débito','Crédito','Crédito parcelado','Boleto','Transferência','Promissória','Convênio']} first="Selecione"/></Field>
      <Field label="Parcelas"><Select value={String(row.installments||1)} setValue={v=>setRow(index,{installments:v})} options={Array.from({length:12},(_,i)=>String(i+1))}/></Field>
      <Field label="Pagamento"><Select value={row.status==='Recebido'?'Pago':'Não pago'} setValue={v=>setRow(index,{status:v==='Pago'?'Recebido':'Pendente',received_date:v==='Pago'?(row.received_date||todayISO()):''})} options={['Pago','Não pago']}/></Field>
      {row.status==='Recebido'
        ? <Field label="Data do pagamento"><input type="date" value={row.received_date||''} onChange={e=>setRow(index,{received_date:e.target.value,status:'Recebido'})}/></Field>
        : <Field label="Previsão / vencimento"><input type="date" value={row.expected_date||''} onChange={e=>setRow(index,{expected_date:e.target.value,status:'Pendente'})}/></Field>}
      <button type="button" className="btn btn-danger-soft payment-row-remove" onClick={()=>removeRow(index)}>Remover</button>
    </div>)}</div>:<div className="payment-breakdown-empty">Nenhum pagamento lançado. Clique em “Adicionar pagamento”.</div>}
  </div>;
}

export function RepassesEditor({form,setForm}){
  const rows=Array.isArray(form.repasses)?form.repasses:[];
  const setRows=repasses=>setForm({...form,repasses});
  const addRow=()=>setRows([...rows,{id:crypto.randomUUID(),professional:'',type:'Honorário médico',amount:''}]);
  const setRow=(index,patch)=>setRows(rows.map((r,i)=>i===index?{...r,...patch}:r));
  const removeRow=index=>setRows(rows.filter((_,i)=>i!==index));
  const total=rows.reduce((s,r)=>s+numericValue(r.amount),0);
  return <div className="repasses-box full">
    <div className="repasses-head"><div><span>REPASSES</span><b>Honorários e profissionais envolvidos</b><small>Deixe em branco quando não houver repasse.</small></div><button type="button" className="btn btn-light" onClick={addRow}>+ Adicionar repasse</button></div>
    {rows.length?<><div className="repasses-list">{rows.map((row,index)=><div className="repasse-row" key={row.id||index}>
      <Field label="Profissional"><input value={row.professional||''} onChange={e=>setRow(index,{professional:e.target.value})} placeholder="Ex.: Dra. Irene"/></Field>
      <Field label="Tipo"><Select value={row.type||'Honorário médico'} setValue={v=>setRow(index,{type:v})} options={['Honorário médico','Anestesista','Auxiliar','Outro']}/></Field>
      <Field label="Valor (R$)"><input type="number" min="0" step="0.01" value={row.amount||''} onChange={e=>setRow(index,{amount:e.target.value})}/></Field>
      <button type="button" className="btn btn-danger-soft" onClick={()=>removeRow(index)}>Remover</button>
    </div>)}</div><div className="repasse-total">Total previsto em repasses: <b>{moneyBR(total)}</b></div></>:<div className="payment-breakdown-empty">Nenhum repasse cadastrado.</div>}
  </div>;
}

export function FinanceQuickSummary({form,setForm}){
  const net=Math.max(numericValue(form.total_amount)-numericValue(form.discount_amount),0);
  const received=numericValue(form.received_amount);
  const status=String(form.financial_status||'').toLowerCase();
  const tone=status.includes('parcial')
    ? 'partial'
    : ((net>0&&received>=net)||status==='recebido')
      ? 'paid'
      : 'unpaid';
  const paymentValue=!form.payment_confirmed?'':tone==='paid'?'Pago':tone==='partial'?'Pago parcial':'Não pago';
  const balance=Math.max(net-received,0);

  const changePayment=value=>{
    if(!setForm) return;
    if(value==='Pago'){
      setForm({
        ...form,
        financial_status:'Recebido',
        payment_confirmed:true,
        received_amount:String(net||''),
        received_date:form.received_date||todayISO()
      });
      return;
    }
    if(value==='Pago parcial'){
      const current=received>0&&received<net?received:0;
      setForm({
        ...form,
        financial_status:'Parcial',
        payment_confirmed:true,
        received_amount:String(current||''),
        received_date:form.received_date||todayISO()
      });
      return;
    }
    setForm({
      ...form,
      financial_status:'Pendente',
      payment_confirmed:true,
      received_amount:'',
      received_date:''
    });
  };

  return <div className="finance-quick-summary full">
    <div><span>Valor líquido</span><b>{moneyBR(net)}</b></div>
    <div><span>Recebido</span><b>{moneyBR(received)}</b></div>
    <div><span>Saldo</span><b>{moneyBR(balance)}</b></div>
    <div className={`finance-summary-status ${tone}`}>
      <span>Pagamento *</span>
      <div className="finance-summary-payment-select">
        <i/>
        <select value={paymentValue} onChange={e=>changePayment(e.target.value)} aria-label="Pagamento" required>
          <option value="" disabled>Selecione</option>
          <option>Pago</option>
          <option>Pago parcial</option>
          <option>Não pago</option>
        </select>
      </div>
    </div>
  </div>;
}

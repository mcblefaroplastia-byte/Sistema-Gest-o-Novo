export const monthNow = new Date().toISOString().slice(0, 7);
export const todayISO = () => new Date().toISOString().slice(0, 10);
export const brDate = s => {
  if (!s) return '';
  const [y,m,d] = s.split('-');
  return `${d}/${m}/${y}`;
};
export const monthLabel = value => {
  if(!value) return 'Todos os períodos';
  const [year,month] = value.split('-');
  const names = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  return `${names[Number(month)-1]} de ${year}`;
};

export const normalizePhone = phone => String(phone || '').replace(/\D/g, '');
export const normalizeSearchDate = value => {
  const v=String(value||'').trim();
  const digits=v.replace(/\D/g,'');
  if(digits.length===8){
    const d=digits.slice(0,2),m=digits.slice(2,4),y=digits.slice(4);
    return `${y}-${m}-${d}`;
  }
  return v;
};
export const patientMatchesSearch = (record,query) => {
  const q=String(query||'').trim().toLowerCase();
  if(!q) return true;
  const birth=String(record?.patient_birth_date||'');
  const birthBr=birth?brDate(birth):'';
  const normalizedDate=normalizeSearchDate(q);
  return String(record?.patient||'').toLowerCase().includes(q) ||
    String(record?.whatsapp||'').includes(q) ||
    birth.includes(q) || birthBr.includes(q) || (normalizedDate && birth===normalizedDate);
};
export const whatsappUrl = phone => {
  let n = normalizePhone(phone);
  if (!n) return '';
  if (!n.startsWith('55')) n = `55${n}`;
  return `https://wa.me/${n}`;
};
export const whatsappDisplay = phone => {
  let n = normalizePhone(phone);
  if(n.startsWith('55') && n.length>11) n=n.slice(2);
  if(n.length===11) return `(${n.slice(0,2)}) ${n.slice(2,7)}-${n.slice(7)}`;
  if(n.length===10) return `(${n.slice(0,2)}) ${n.slice(2,6)}-${n.slice(6)}`;
  return phone || 'Não informado';
};
export const followupWhatsappUrl = (phone,name) => {
  let n = normalizePhone(phone);
  if(!n) return '';
  if(!n.startsWith('55')) n=`55${n}`;
  const firstName=String(name||'').trim().split(/\s+/)[0] || 'Paciente';
  const message=`Olá, ${firstName}! Tudo bem? 😊 Aqui é da equipe da OftalmoCastro. Estou entrando em contato para acompanhar seu atendimento conosco. Podemos dar continuidade por aqui?`;
  return `https://wa.me/${n}?text=${encodeURIComponent(message)}`;
};
export const countBy = (arr, key) => arr.reduce((o, x) => {
  const k = x[key] || 'Não informado';
  o[k] = (o[k] || 0) + 1;
  return o;
}, {});

export const countManyBy = (arr, key) => arr.reduce((o, x) => {
  const values = Array.isArray(x[key]) && x[key].length ? x[key] : ['Não informado'];
  values.forEach(k => { o[k] = (o[k] || 0) + 1; });
  return o;
}, {});

export const totalItems = (arr, key) => arr.reduce((total, item) => total + (Array.isArray(item[key]) ? item[key].length : 0), 0);
export const totalItemsByStatus = (arr, key, status) => arr.reduce((total, item) =>
  total + (item.status === status && Array.isArray(item[key]) ? item[key].length : 0), 0);

export const countRowsByStatus = (arr,status) => arr.filter(item=>item.status===status).length;

export const countWeightedBy = (arr, key, itemsKey) => arr.reduce((o, item) => {
  const label = item[key] || 'Não informado';
  const weight = Array.isArray(item[itemsKey]) && item[itemsKey].length ? item[itemsKey].length : 0;
  o[label] = (o[label] || 0) + weight;
  return o;
}, {});

export const percentOf = (value,total) => total ? Math.round((value/total)*100) : 0;

export const numericValue = value => {
  const n=Number(value||0);
  return Number.isFinite(n)?n:0;
};
export const moneyBR = value => numericValue(value).toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
export const financeOrigin = record => record?.payment_origin || (record?.insurance_id || record?.insurance ? 'Convênio' : 'Particular');
export const isCanceledOperational = status => String(status||'').toLowerCase().includes('cancel');
export const isRealizedOperational = status => String(status||'').toLowerCase().includes('realiz');
export const isPotentialOperational = status => !isCanceledOperational(status) && !isRealizedOperational(status);
export const netFinancialAmount = record => Math.max(numericValue(record?.total_amount ?? record?.total)-numericValue(record?.discount_amount),0);
export const normalizedBreakdown = value => Array.isArray(value)?value:[];
export const paymentTone = record => {
  const total=netFinancialAmount(record);
  const received=numericValue(record?.received_amount ?? record?.received);
  const status=String(record?.financial_status||'').toLowerCase();
  if((total>0 && received>=total) || status==='recebido') return 'paid';
  if(received>0 || status.includes('parcial') || status.includes('faturado') || status.includes('aguardando')) return 'partial';
  return 'unpaid';
};
export const paymentToneLabel = record => paymentTone(record)==='paid'?'Pago':paymentTone(record)==='partial'?'Pago parcial':'Não pago';
export const paymentAllocation = (record,type) => {
  const rows=normalizedBreakdown(record?.payment_breakdown);
  const total=netFinancialAmount(record);
  if(rows.length) return rows.filter(x=>x.payer_type===type).reduce((s,x)=>s+numericValue(x.amount),0);
  const origin=record?.payment_origin||financeOrigin(record);
  if(origin===type) return total;
  if(origin==='Misto') return 0;
  return 0;
};

export const normalizeCatalogLabel = value => String(value||'')
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g,'')
  .toLowerCase()
  .replace(/[^a-z0-9]+/g,' ')
  .trim();

export const knownCatalogKind = value => {
  const n=normalizeCatalogLabel(value);
  const surgery=
    n.includes('catarata') || n.includes('blefaroplast') || n.includes('ptose') ||
    n.includes('pterig') || n.includes('estrab') || n.includes('refrativ') ||
    n.includes('ectrop') || n.includes('entrop') ||
    ((n.includes('tumor'))&&(n.includes('palpebr')||n.includes('conjunt'))) ||
    n.includes('calazio') || (n.includes('nodul')&&n.includes('palpebr')) ||
    (n.includes('sondagem')&&n.includes('lacr')) || n.includes('outra cirurgia');
  if(surgery) return 'surgery';

  const procedure=
    n.includes('anti vegf') || n.includes('avastin') || n.includes('lucentis') || n.includes('eylia') || n.includes('vabysmo') ||
    n.includes('ozurdex') || (n.includes('vias lacrimais')&&!n.includes('sondagem')) ||
    (n.includes('les')&&n.includes('palpebr')&&!n.includes('tumor')&&!n.includes('nodul')) ||
    n.includes('laser co2') || (n.includes('laser')&&n.includes('mao')) ||
    n.includes('botox') || n.includes('toxina botulin') || n.includes('microfoc') || n.includes('ultrassom') || n.startsWith('us ') ||
    n.includes('bioestimul') || n.includes('profhilo') || n.includes('skinvive') || n.includes('harmoniz') ||
    n.includes('hialuron') || (n.includes('cicatriz')&&n.includes('acne')) || n.includes('terco inferior') ||
    (n.includes('lifting')&&n.includes('superc'));
  if(procedure) return 'procedure';
  return 'unknown';
};

export const shiftMonth = (value,delta) => {
  const base = value || monthNow;
  const [year,month] = base.split('-').map(Number);
  const d = new Date(year,month-1+delta,1);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
};

export const shortMonthLabel = value => {
  if(!value) return '';
  const [year,month] = value.split('-');
  const names = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  return `${names[Number(month)-1]}/${String(year).slice(-2)}`;
};

export const variationPct = (current,previous) => {
  if(previous===0) return current>0 ? 100 : 0;
  return Math.round(((current-previous)/previous)*100);
};

export const groupLaunchRows = (rows, kind) => {
  if(kind==='exam'){
    const groups=new Map();
    rows.forEach(r=>{
      const groupId=r.launch_group_id||r.id;
      if(!groups.has(groupId)){
        groups.set(groupId,{
          group_id:groupId,
          db_ids:[],
          patient_id:r.patients?.id||'',
          date:r.exam_date,
          realization_date:r.realization_date||'',
          eye:r.eye||'',
          patient:r.patients?.full_name||'',
          whatsapp:r.patients?.whatsapp||'',
          patient_birth_date:r.patients?.birth_date||'',
          doctor:r.doctors?.name||'',
          doctor_id:r.doctors?.id||'',
          indicatedBy:r.indicated_by?.name||'',
          indicatedById:r.indicated_by?.id||'',
          status:r.status,
          obs:r.observation||'',
          launchedBy:r.launched_by,
          launchedByName:r.launched_by_name||'',
          exam_type_ids:[],
          exam_names:[],
          insurance:r.insurances?.name||'',
          insurance_id:r.insurances?.id||'',
          payment_origin:r.payment_origin||'',
          total_amount:numericValue(r.total_amount),
          discount_amount:numericValue(r.discount_amount),
          received_amount:numericValue(r.received_amount),
          payment_method:r.payment_method||'',
          installments:Number(r.installments||1),
          financial_status:r.financial_status||'Pendente',
          expected_receipt_date:r.expected_receipt_date||'',
          received_date:r.received_date||'',
          billed_amount:numericValue(r.billed_amount),
          billing_competence:r.billing_competence||'',
          billing_date:r.billing_date||'',
          financial_notes:r.financial_notes||'',
          price_preset_key:r.price_preset_key||'',payment_breakdown:r.payment_breakdown||[],invoice_number:r.invoice_number||'',invoice_issued_date:r.invoice_issued_date||'',invoice_file_path:r.invoice_file_path||'',financial_documents:r.financial_documents||[],repasses:r.repasses||[]
        });
      }
      const g=groups.get(groupId);
      g.db_ids.push(r.id);
      if(r.exam_types?.id) g.exam_type_ids.push(r.exam_types.id);
      if(r.exam_types?.name) g.exam_names.push(r.exam_types.name);
    });
    return [...groups.values()].map(g=>({
      ...g,
      id:g.group_id,
      exam:g.exam_names.join(' • '),
      exam_type_id:g.exam_type_ids[0]||''
    }));
  }

  // V17: cada linha classificada como surgery/review representa UMA cirurgia.
  // Registros procedure_legacy permanecem preservados no banco, mas aparecem
  // somente no módulo Procedimentos.
  return rows
    .filter(r=>r.record_kind!=='procedure_legacy')
    .map(r=>({
      id:r.id,
      group_id:r.launch_group_id||r.id,
      db_ids:[r.id],
      patient_id:r.patients?.id||'',
      date:r.surgery_date,
      scheduled_on:r.scheduled_on||'',
      second_eye_date:r.second_eye_date||'',
      patient:r.patients?.full_name||'',
      whatsapp:r.patients?.whatsapp||'',
      patient_birth_date:r.patients?.birth_date||'',
      doctor:r.doctors?.name||'',
      doctor_id:r.doctors?.id||'',
      indicatedBy:r.indicated_by?.name||'',
      indicatedById:r.indicated_by?.id||'',
      status:r.status,
      obs:r.observation||'',
      launchedBy:r.launched_by,
      launchedByName:r.launched_by_name||'',
      surgery_type_id:r.surgery_type?.id||r.surgery_type_id||'',
      surgery_type:r.surgery_type?.name||'Cirurgia a classificar',
      surgery_type_names:[r.surgery_type?.name||'Cirurgia a classificar'],
      procedure_ids:[],
      procedure_names:[],
      procedure:'Sem procedimentos vinculados',
      eye:r.eye,
      insurance:r.insurances?.name||'',
      insurance_id:r.insurances?.id||'',
      arrival:r.arrival_time||'',
      time:r.surgery_time||'',
      payment:r.payment_status||'',
      payment_origin:r.payment_origin||'',
      total_amount:numericValue(r.total_amount),
      discount_amount:numericValue(r.discount_amount),
      received_amount:numericValue(r.received_amount),
      payment_method:r.payment_method||'',
      installments:r.installments||1,
      financial_status:r.financial_status||'Pendente',
      expected_receipt_date:r.expected_receipt_date||'',
      received_date:r.received_date||'',
      billed_amount:numericValue(r.billed_amount),
      billing_competence:r.billing_competence||'',
      billing_date:r.billing_date||'',
      financial_notes:r.financial_notes||'',payment_breakdown:r.payment_breakdown||[],invoice_number:r.invoice_number||'',invoice_issued_date:r.invoice_issued_date||'',invoice_file_path:r.invoice_file_path||'',financial_documents:r.financial_documents||[],repasses:r.repasses||[],lio_description:r.lio_description||'',
      procedure_finance:{},
      classification_review:r.record_kind==='review'
    }));
};

export const mapProcedureRows = (rows, surgeryById=new Map(), surgeryByGroup=new Map()) => rows.filter(r=>r.record_kind!=='surgery_legacy').map(r=>{
  const linkedSurgery=(r.surgery_id&&surgeryById.get(r.surgery_id)) || surgeryByGroup.get(r.surgery_group_id) || null;
  return {
    id:r.id,
    group_id:r.launch_group_id||r.id,
    patient_id:r.patients?.id||'',
    date:r.procedure_date,
    realization_date:r.realization_date||'',
    patient:r.patients?.full_name||'',
    whatsapp:r.patients?.whatsapp||'',
    patient_birth_date:r.patients?.birth_date||'',
    doctor:r.doctors?.name||'',
    doctor_id:r.doctors?.id||'',
    indicatedBy:r.indicated_by?.name||'',
    indicatedById:r.indicated_by?.id||'',
    procedure:r.procedures?.name||'',
    procedure_id:r.procedures?.id||r.procedure_id||'',
    eye:r.eye||'Não se aplica',
    insurance:r.insurances?.name||'',
    insurance_id:r.insurances?.id||'',
    status:r.status,
    time:r.procedure_time||'',
    payment:r.payment_status||'Não informado',
    payment_origin:r.payment_origin||'',
    total_amount:numericValue(r.total_amount),
    discount_amount:numericValue(r.discount_amount),
    received_amount:numericValue(r.received_amount),
    payment_method:r.payment_method||'',
    installments:r.installments||1,
    financial_status:r.financial_status||'Pendente',
    expected_receipt_date:r.expected_receipt_date||'',
    received_date:r.received_date||'',
    billed_amount:numericValue(r.billed_amount),
    billing_competence:r.billing_competence||'',
    billing_date:r.billing_date||'',
    financial_notes:r.financial_notes||'',
    price_preset_key:r.price_preset_key||'',
    payment_breakdown:r.payment_breakdown||[],
    invoice_number:r.invoice_number||'',
    invoice_issued_date:r.invoice_issued_date||'',
    invoice_file_path:r.invoice_file_path||'',
    financial_documents:r.financial_documents||[],
    repasses:r.repasses||[],
    included_in_package:r.included_in_package!==false,
    additional_amount:numericValue(r.additional_amount),
    obs:r.observation||'',
    launchedBy:r.launched_by,
    launchedByName:r.launched_by_name||'',
    surgery_id:r.surgery_id||'',
    surgery_group_id:r.surgery_group_id||'',
    linked_surgery:linkedSurgery?.surgery_type||'',
    linked_surgery_group:linkedSurgery||null,
    legacy_surgery_id:r.legacy_surgery_id||''
  };
});


// Agrupa apenas a VISUALIZAÇÃO dos procedimentos.
// Cada procedure_record continua existindo individualmente e continua contando
// como 1 procedimento nos Resultados. Procedimentos do mesmo pacote cirúrgico
// aparecem em uma única linha para facilitar a leitura operacional.
export const groupProcedureRows = rows => {
  const groups=new Map();
  (rows||[]).forEach(row=>{
    // Prefere surgery_group_id para que registros antigos do mesmo pacote
    // sejam unidos mesmo quando alguns têm surgery_id e outros têm apenas o grupo.
    const groupKey=row.surgery_group_id
      ? `surgery-group:${row.surgery_group_id}:${row.patient_id}:${row.date}`
      : row.surgery_id
        ? `surgery:${row.surgery_id}`
        : row.group_id
          ? `launch:${row.group_id}`
          : `row:${row.id}`;

    if(!groups.has(groupKey)){
      groups.set(groupKey,{
        ...row,
        id:groupKey,
        group_key:groupKey,
        db_ids:[],
        procedure_ids:[],
        procedure_names:[],
        item_count:0
      });
    }

    const group=groups.get(groupKey);
    group.db_ids.push(row.id);
    group.procedure_ids.push(row.procedure_id);
    group.procedure_names.push(row.procedure);
    group.item_count+=1;

    // Se algum item do grupo possuir o vínculo resolvido, preserva-o no grupo.
    if(!group.linked_surgery && row.linked_surgery) group.linked_surgery=row.linked_surgery;
    if(!group.linked_surgery_group && row.linked_surgery_group) group.linked_surgery_group=row.linked_surgery_group;
    if(!group.surgery_id && row.surgery_id) group.surgery_id=row.surgery_id;
    if(!group.surgery_group_id && row.surgery_group_id) group.surgery_group_id=row.surgery_group_id;
  });

  return [...groups.values()].map(group=>({
    ...group,
    procedure:group.procedure_names.join(' • ')
  }));
};

export const statusClass = status => String(status||'')
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g,'')
  .toLowerCase()
  .replace(/\s+/g,'-');

export function accessLabel(profile){
  if(!profile) return 'Usuário';
  return profile.role === 'admin' ? 'Administrador' : 'Orientadora';
}

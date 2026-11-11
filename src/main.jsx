import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import ExcelJS from 'exceljs';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { Activity, BarChart3, CalendarDays, ChevronRight, CircleHelp, ClipboardList, Download, Eye, FileDown, FileText, LayoutDashboard, ListChecks, LogOut, Plus, RefreshCw, Settings, Stethoscope, Trash2, Upload, WalletCards } from 'lucide-react';
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
const normalizeSearchDate = value => {
  const v=String(value||'').trim();
  const digits=v.replace(/\D/g,'');
  if(digits.length===8){
    const d=digits.slice(0,2),m=digits.slice(2,4),y=digits.slice(4);
    return `${y}-${m}-${d}`;
  }
  return v;
};
const patientMatchesSearch = (record,query) => {
  const q=String(query||'').trim().toLowerCase();
  if(!q) return true;
  const birth=String(record?.patient_birth_date||'');
  const birthBr=birth?brDate(birth):'';
  const normalizedDate=normalizeSearchDate(q);
  return String(record?.patient||'').toLowerCase().includes(q) ||
    String(record?.whatsapp||'').includes(q) ||
    birth.includes(q) || birthBr.includes(q) || (normalizedDate && birth===normalizedDate);
};
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

const countRowsByStatus = (arr,status) => arr.filter(item=>item.status===status).length;

const countWeightedBy = (arr, key, itemsKey) => arr.reduce((o, item) => {
  const label = item[key] || 'Não informado';
  const weight = Array.isArray(item[itemsKey]) && item[itemsKey].length ? item[itemsKey].length : 0;
  o[label] = (o[label] || 0) + weight;
  return o;
}, {});

const percentOf = (value,total) => total ? Math.round((value/total)*100) : 0;

const numericValue = value => {
  const n=Number(value||0);
  return Number.isFinite(n)?n:0;
};
const moneyBR = value => numericValue(value).toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
const financeOrigin = record => record?.payment_origin || (record?.insurance_id || record?.insurance ? 'Convênio' : 'Particular');
const isCanceledOperational = status => String(status||'').toLowerCase().includes('cancel');
const isRealizedOperational = status => String(status||'').toLowerCase().includes('realiz');
const isPotentialOperational = status => !isCanceledOperational(status) && !isRealizedOperational(status);
const netFinancialAmount = record => Math.max(numericValue(record?.total_amount ?? record?.total)-numericValue(record?.discount_amount),0);
const normalizedBreakdown = value => Array.isArray(value)?value:[];
const paymentTone = record => {
  const total=netFinancialAmount(record);
  const received=numericValue(record?.received_amount ?? record?.received);
  const status=String(record?.financial_status||'').toLowerCase();
  if((total>0 && received>=total) || status==='recebido') return 'paid';
  if(received>0 || status.includes('parcial') || status.includes('faturado') || status.includes('aguardando')) return 'partial';
  return 'unpaid';
};
const paymentToneLabel = record => paymentTone(record)==='paid'?'Pago':paymentTone(record)==='partial'?'Pago parcial':'Não pago';
const paymentAllocation = (record,type) => {
  const rows=normalizedBreakdown(record?.payment_breakdown);
  const total=netFinancialAmount(record);
  if(rows.length) return rows.filter(x=>x.payer_type===type).reduce((s,x)=>s+numericValue(x.amount),0);
  const origin=record?.payment_origin||financeOrigin(record);
  if(origin===type) return total;
  if(origin==='Misto') return 0;
  return 0;
};


// V24 - Tabela de valores padrão. O preço é apenas uma sugestão inicial:
// depois de aplicado ao lançamento, o campo Valor total continua totalmente editável.
const STANDARD_PRICE_PRESETS = [
  // EXAMES / CONSULTAS
  {key:'exam-consulta-particular',kind:'exam',match:['consulta'],label:'Consulta - Particular',amount:400,origin:'Particular'},
  {key:'exam-consulta-convenio',kind:'exam',match:['consulta'],label:'Consulta - Convênio',amount:220,origin:'Convênio'},
  {key:'exam-paquimetria-particular',kind:'exam',match:['paquimetria'],label:'Paquimetria - Particular',amount:250,origin:'Particular'},
  {key:'exam-paquimetria-convenio',kind:'exam',match:['paquimetria'],label:'Paquimetria - Convênio',amount:200,origin:'Convênio'},
  {key:'exam-paquimetria-sus',kind:'exam',match:['paquimetria'],label:'Paquimetria - SUS',amount:180,origin:'SUS'},
  {key:'exam-mapeamento-particular',kind:'exam',match:['mapeamento'],label:'Mapeamento de Retina - Particular',amount:250,origin:'Particular'},
  {key:'exam-mapeamento-convenio',kind:'exam',match:['mapeamento'],label:'Mapeamento de Retina - Convênio',amount:200,origin:'Convênio'},
  {key:'exam-mapeamento-sus',kind:'exam',match:['mapeamento'],label:'Mapeamento de Retina - SUS',amount:180,origin:'SUS'},
  {key:'exam-retinografia-particular',kind:'exam',match:['retinografia'],label:'Retinografia - Particular',amount:280,origin:'Particular'},
  {key:'exam-retinografia-convenio',kind:'exam',match:['retinografia'],label:'Retinografia - Convênio',amount:220,origin:'Convênio'},
  {key:'exam-retinografia-sus',kind:'exam',match:['retinografia'],label:'Retinografia - SUS',amount:180,origin:'SUS'},
  {key:'exam-campimetria-particular',kind:'exam',match:['campimetria'],label:'Campimetria - Particular',amount:380,origin:'Particular'},
  {key:'exam-campimetria-convenio',kind:'exam',match:['campimetria'],label:'Campimetria - Convênio',amount:320,origin:'Convênio'},
  {key:'exam-campimetria-sus',kind:'exam',match:['campimetria'],label:'Campimetria - SUS',amount:180,origin:'SUS'},
  {key:'exam-topografia-particular',kind:'exam',match:['topografia'],label:'Topografia - Particular',amount:280,origin:'Particular'},
  {key:'exam-topografia-convenio',kind:'exam',match:['topografia'],label:'Topografia - Convênio',amount:240,origin:'Convênio'},
  {key:'exam-topografia-sus',kind:'exam',match:['topografia'],label:'Topografia - SUS',amount:180,origin:'SUS'},
  {key:'exam-microscopia-particular',kind:'exam',match:['microscopia'],label:'Microscopia - Particular',amount:320,origin:'Particular'},
  {key:'exam-microscopia-convenio',kind:'exam',match:['microscopia'],label:'Microscopia - Convênio',amount:280,origin:'Convênio'},
  {key:'exam-microscopia-sus',kind:'exam',match:['microscopia'],label:'Microscopia - SUS',amount:180,origin:'SUS'},
  {key:'exam-biometria-particular',kind:'exam',match:['biometria'],label:'Biometria - Particular',amount:250,origin:'Particular'},
  {key:'exam-biometria-convenio',kind:'exam',match:['biometria'],label:'Biometria - Convênio',amount:200,origin:'Convênio'},
  {key:'exam-oct-particular',kind:'exam',match:['oct'],label:'OCT - Particular',amount:750,origin:'Particular'},
  {key:'exam-oct-convenio',kind:'exam',match:['oct'],label:'OCT - Convênio',amount:650,origin:'Convênio'},
  {key:'exam-oct-sus',kind:'exam',match:['oct'],label:'OCT - SUS (AO)',amount:480,origin:'SUS'},
  {key:'exam-ortoptico-convenio',kind:'exam',match:['ortopt'],label:'Teste Ortóptico - Convênio',amount:250,origin:'Convênio'},
  {key:'exam-ishihara-convenio',kind:'exam',match:['ishihara'],label:'Teste de Ishihara - Convênio',amount:180,origin:'Convênio'},
  {key:'exam-titmus-convenio',kind:'exam',match:['titmus'],label:'Teste de Titmus - Convênio',amount:180,origin:'Convênio'},
  {key:'exam-curva-particular',kind:'exam',match:['curva tensional'],label:'Curva Tensional - Particular',amount:200,origin:'Particular'},
  {key:'exam-curva-convenio',kind:'exam',match:['curva tensional'],label:'Curva Tensional - Convênio',amount:180,origin:'Convênio'},
  {key:'exam-lentes-particular',kind:'exam',match:['teste de lentes','lentes'],label:'Teste de Lentes - Particular',amount:100,origin:'Particular'},
  {key:'exam-lentes-convenio',kind:'exam',match:['teste de lentes','lentes'],label:'Teste de Lentes - Convênio',amount:100,origin:'Convênio'},
  {key:'exam-lentes-sus',kind:'exam',match:['teste de lentes','lentes'],label:'Teste de Lentes - SUS',amount:100,origin:'SUS'},
  {key:'exam-schirmer-particular',kind:'exam',match:['schirmer'],label:'Teste de Schirmer - Particular',amount:200,origin:'Particular'},
  {key:'exam-schirmer-convenio',kind:'exam',match:['schirmer'],label:'Teste de Schirmer - Convênio',amount:180,origin:'Convênio'},
  {key:'exam-oculografia-particular',kind:'exam',match:['oculografia'],label:'Oculografia - Particular',amount:220,origin:'Particular'},
  {key:'exam-yag-1-particular',kind:'exam',match:['yag'],label:'YAG Laser - 1 olho - Particular',amount:520,origin:'Particular'},
  {key:'exam-yag-1-convenio',kind:'exam',match:['yag'],label:'YAG Laser - 1 olho - Convênio',amount:450,origin:'Convênio'},
  {key:'exam-yag-1-sus',kind:'exam',match:['yag'],label:'YAG Laser - 1 olho - SUS',amount:180,origin:'SUS'},
  {key:'exam-yag-ao-particular',kind:'exam',match:['yag'],label:'YAG Laser - AO - Particular',amount:900,origin:'Particular'},
  {key:'exam-yag-ao-convenio',kind:'exam',match:['yag'],label:'YAG Laser - AO - Convênio',amount:820,origin:'Convênio'},
  {key:'exam-yag-ao-sus',kind:'exam',match:['yag'],label:'YAG Laser - AO - SUS',amount:350,origin:'SUS'},
  {key:'exam-iridotomia-1-particular',kind:'exam',match:['iridotomia'],label:'Iridotomia - 1 olho - Particular',amount:520,origin:'Particular'},
  {key:'exam-iridotomia-1-convenio',kind:'exam',match:['iridotomia'],label:'Iridotomia - 1 olho - Convênio',amount:450,origin:'Convênio'},
  {key:'exam-iridotomia-ao-particular',kind:'exam',match:['iridotomia'],label:'Iridotomia - AO - Particular',amount:900,origin:'Particular'},
  {key:'exam-iridotomia-ao-convenio',kind:'exam',match:['iridotomia'],label:'Iridotomia - AO - Convênio',amount:820,origin:'Convênio'},
  {key:'exam-check-retina-particular',kind:'exam',match:['check up retina','check-up retina'],label:'Check-up Retina - Particular',amount:1150,origin:'Particular'},
  {key:'exam-check-retina-convenio',kind:'exam',match:['check up retina','check-up retina'],label:'Check-up Retina - Convênio',amount:980,origin:'Convênio'},
  {key:'exam-check-glaucoma-particular',kind:'exam',match:['check up glaucoma','check-up glaucoma'],label:'Check-up Glaucoma - Particular',amount:1250,origin:'Particular'},
  {key:'exam-check-glaucoma-convenio',kind:'exam',match:['check up glaucoma','check-up glaucoma'],label:'Check-up Glaucoma - Convênio',amount:1100,origin:'Convênio'},
  {key:'exam-check-glaucoma-sus',kind:'exam',match:['check up glaucoma','check-up glaucoma'],label:'Check-up Glaucoma - SUS',amount:900,origin:'SUS'},
  {key:'exam-kit-refrativa-particular',kind:'exam',match:['kit refrativa'],label:'Kit Refrativa - Particular',amount:700,origin:'Particular'},
  {key:'exam-kit-refrativa-convenio',kind:'exam',match:['kit refrativa'],label:'Kit Refrativa - Convênio',amount:600,origin:'Convênio'},
  {key:'exam-us-1-particular',kind:'exam',match:['ultrassom'],label:'Ultrassom - 1 olho - Particular',amount:260,origin:'Particular'},
  {key:'exam-us-ao-particular',kind:'exam',match:['ultrassom'],label:'Ultrassom - AO - Particular',amount:420,origin:'Particular'},
  {key:'exam-us-ao-convenio',kind:'exam',match:['ultrassom'],label:'Ultrassom - AO - Convênio',amount:380,origin:'Convênio'},

  // CIRURGIAS
  {key:'cat-part-visiontech-1',kind:'surgery',match:['catarata'],label:'Catarata Particular - Visiontech - 1 olho',amount:3830,origin:'Particular'},
  {key:'cat-part-visiontech-2',kind:'surgery',match:['catarata'],label:'Catarata Particular - Visiontech - 2 olhos',amount:6880,origin:'Particular'},
  {key:'cat-part-asferica-1',kind:'surgery',match:['catarata'],label:'Catarata Particular - Asférica - 1 olho',amount:7180,origin:'Particular'},
  {key:'cat-part-asferica-2',kind:'surgery',match:['catarata'],label:'Catarata Particular - Asférica - 2 olhos',amount:13380,origin:'Particular'},
  {key:'cat-part-trifocal-1',kind:'surgery',match:['catarata'],label:'Catarata Particular - Trifocal - 1 olho',amount:14580,origin:'Particular'},
  {key:'cat-part-trifocal-2',kind:'surgery',match:['catarata'],label:'Catarata Particular - Trifocal - 2 olhos',amount:25940,origin:'Particular'},
  {key:'cat-part-monotorica-1',kind:'surgery',match:['catarata'],label:'Catarata Particular - Mono Tórica - 1 olho',amount:8280,origin:'Particular'},
  {key:'cat-part-monotorica-2',kind:'surgery',match:['catarata'],label:'Catarata Particular - Mono Tórica - 2 olhos',amount:11280,origin:'Particular'},
  {key:'cat-cassi-visiontech-1',kind:'surgery',match:['catarata'],label:'Catarata CASSI/IAMSPE/SulAmérica - Visiontech - 1 olho',amount:2280,origin:'Convênio'},
  {key:'cat-cassi-visiontech-2',kind:'surgery',match:['catarata'],label:'Catarata CASSI/IAMSPE/SulAmérica - Visiontech - 2 olhos',amount:3880,origin:'Convênio'},
  {key:'cat-cassi-esferica-1',kind:'surgery',match:['catarata'],label:'Catarata CASSI/IAMSPE/SulAmérica - Esférica - 1 olho',amount:4280,origin:'Convênio'},
  {key:'cat-cassi-esferica-2',kind:'surgery',match:['catarata'],label:'Catarata CASSI/IAMSPE/SulAmérica - Esférica - 2 olhos',amount:7480,origin:'Convênio'},
  {key:'cat-cassi-asferica-1',kind:'surgery',match:['catarata'],label:'Catarata CASSI/IAMSPE/SulAmérica - Asférica - 1 olho',amount:6080,origin:'Convênio'},
  {key:'cat-cassi-asferica-2',kind:'surgery',match:['catarata'],label:'Catarata CASSI/IAMSPE/SulAmérica - Asférica - 2 olhos',amount:10280,origin:'Convênio'},
  {key:'cat-cassi-trifocal-1',kind:'surgery',match:['catarata'],label:'Catarata CASSI/IAMSPE/SulAmérica - Trifocal - 1 olho',amount:13580,origin:'Convênio'},
  {key:'cat-cassi-trifocal-2',kind:'surgery',match:['catarata'],label:'Catarata CASSI/IAMSPE/SulAmérica - Trifocal - 2 olhos',amount:23580,origin:'Convênio'},
  {key:'cat-cassi-monotorica-1',kind:'surgery',match:['catarata'],label:'Catarata CASSI/IAMSPE/SulAmérica - Mono Tórica - 1 olho',amount:6280,origin:'Convênio'},
  {key:'cat-cassi-monotorica-2',kind:'surgery',match:['catarata'],label:'Catarata CASSI/IAMSPE/SulAmérica - Mono Tórica - 2 olhos',amount:11180,origin:'Convênio'},
  {key:'cat-cassi-tritorica-1',kind:'surgery',match:['catarata'],label:'Catarata CASSI/IAMSPE/SulAmérica - Tri Tórica - 1 olho',amount:13580,origin:'Convênio'},
  {key:'cat-cassi-tritorica-2',kind:'surgery',match:['catarata'],label:'Catarata CASSI/IAMSPE/SulAmérica - Tri Tórica - 2 olhos',amount:23580,origin:'Convênio'},
  {key:'cat-unimed-asferica-1',kind:'surgery',match:['catarata'],label:'Catarata Unimed/Oeste - Asférica - 1 olho',amount:3330,origin:'Convênio'},
  {key:'cat-unimed-asferica-2',kind:'surgery',match:['catarata'],label:'Catarata Unimed/Oeste - Asférica - 2 olhos',amount:5880,origin:'Convênio'},
  {key:'cat-unimed-trifocal-1',kind:'surgery',match:['catarata'],label:'Catarata Unimed/Oeste - Trifocal - 1 olho',amount:9880,origin:'Convênio'},
  {key:'cat-unimed-trifocal-2',kind:'surgery',match:['catarata'],label:'Catarata Unimed/Oeste - Trifocal - 2 olhos',amount:17480,origin:'Convênio'},
  {key:'faco-trec-1',kind:'surgery',match:['faco trec'],label:'FACO + TREC - 1 olho',amount:9580,origin:'Convênio'},
  {key:'faco-trec-2',kind:'surgery',match:['faco trec'],label:'FACO + TREC - 2 olhos',amount:18280,origin:'Convênio'},
  {key:'trec-1',kind:'surgery',match:['trec'],label:'TREC - 1 olho',amount:5850,origin:'Convênio'},
  {key:'trec-2',kind:'surgery',match:['trec'],label:'TREC - 2 olhos',amount:11480,origin:'Convênio'},
  {key:'refrativa-1',kind:'surgery',match:['refrativ'],label:'Cirurgia refrativa - 1 olho',amount:4900,origin:'Particular'},
  {key:'refrativa-2',kind:'surgery',match:['refrativ'],label:'Cirurgia refrativa - 2 olhos',amount:7900,origin:'Particular'},
  {key:'crosslinking-1',kind:'surgery',match:['crosslink'],label:'Crosslinking - 1 olho',amount:3200,origin:'Particular'},
  {key:'crosslinking-2',kind:'surgery',match:['crosslink'],label:'Crosslinking - 2 olhos',amount:6400,origin:'Particular'},
  {key:'pterigio-part-1',kind:'surgery',match:['pterig'],label:'Pterígio Particular/Convênio - 1 olho',amount:3600,origin:'Particular'},
  {key:'pterigio-part-2',kind:'surgery',match:['pterig'],label:'Pterígio Particular/Convênio - 2 olhos',amount:6250,origin:'Particular'},
  {key:'pterigio-iamspe-1',kind:'surgery',match:['pterig'],label:'Pterígio IAMSPE - 1 olho',amount:2250,origin:'Convênio'},
  {key:'pterigio-iamspe-2',kind:'surgery',match:['pterig'],label:'Pterígio IAMSPE - 2 olhos',amount:3950,origin:'Convênio'},
  {key:'cisto-moll',kind:'surgery',match:['cisto','moll'],label:'Cisto de Moll',amount:850,origin:'Particular'},
  {key:'calazio-1',kind:'surgery',match:['calazio'],label:'Calázio - 1',amount:850,origin:'Particular'},
  {key:'calazio-2',kind:'surgery',match:['calazio'],label:'Calázio - 2',amount:1600,origin:'Particular'},
  {key:'tumor-palpebral',kind:'surgery',match:['tumor palpebr','t u palpebr'],label:'T.U. Palpebral',amount:1800,origin:'Particular'},
  {key:'lobuloplastia',kind:'surgery',match:['lobuloplast'],label:'Lobuloplastia - 2 lados',amount:1600,origin:'Particular'},
  {key:'cauterizacao',kind:'surgery',match:['cauteriza','cilio'],label:'Cauterização de cílios - cada olho/sessão',amount:480,origin:'Particular'},
  {key:'estrabismo',kind:'surgery',match:['estrab'],label:'Cirurgia de estrabismo - com sedação',amount:15500,origin:'Particular'},
  {key:'blefaro-superior',kind:'surgery',match:['blefaroplastia superior'],exclude:['inferior'],label:'Blefaroplastia superior',amount:6200,origin:'Particular'},
  {key:'blefaro-inferior',kind:'surgery',match:['blefaroplastia inferior'],label:'Blefaroplastia inferior',amount:6800,origin:'Particular'},
  {key:'blefaro-sup-inf',kind:'surgery',match:['blefaroplastia superior inferior','blefaroplastia superior + inferior'],label:'Blefaroplastia superior + inferior',amount:12800,origin:'Particular'},
  {key:'blefaro-sup-lifting',kind:'surgery',match:['blefaroplastia superior','lifting'],label:'Blefaroplastia superior + lifting de supercílio',amount:8600,origin:'Particular'},
  {key:'blefaro-completa-lifting',kind:'surgery',match:['blefaroplastia','inferior','lifting'],label:'Blefaroplastia superior + inferior + lifting',amount:13300,origin:'Particular'},
  {key:'supercilio',kind:'surgery',match:['supercilio'],label:'Supercílio',amount:6200,origin:'Particular'},
  {key:'entropio-1',kind:'surgery',match:['entrop'],label:'Entrópio/Ectrópio - 1 olho',amount:6300,origin:'Particular'},
  {key:'entropio-2',kind:'surgery',match:['entrop'],label:'Entrópio - ambos os olhos',amount:9800,origin:'Particular'},
  {key:'ectropio-2',kind:'surgery',match:['ectrop'],label:'Ectrópio - ambos os olhos',amount:9800,origin:'Particular'},
  {key:'ptose-1',kind:'surgery',match:['ptose'],label:'Ptose - 1 olho',amount:6300,origin:'Particular'},
  {key:'ptose-2',kind:'surgery',match:['ptose'],label:'Ptose - ambos os olhos',amount:9800,origin:'Particular'},
  {key:'ptose-pacote',kind:'surgery',match:['ptose'],label:'Ptose na cirurgia - pacote',amount:2600,origin:'Particular'},

  // PROCEDIMENTOS
  {key:'botox-unica',kind:'procedure',match:['botox','toxina'],label:'Botox - aplicação com retoque',amount:1600,origin:'Particular'},
  {key:'botox-3',kind:'procedure',match:['botox','toxina'],label:'Botox - pacote 3 aplicações / plano anual',amount:4200,origin:'Particular'},
  {key:'botox-cirurgia',kind:'procedure',match:['botox','toxina'],label:'Botox na cirurgia',amount:1400,origin:'Particular'},
  {key:'laser-face',kind:'procedure',match:['laser co2','laser de face','laser face'],label:'Laser de face',amount:3300,origin:'Particular'},
  {key:'laser-face-cirurgia',kind:'procedure',match:['laser co2','laser de face','laser face'],label:'Laser de face - pacote na cirurgia',amount:2700,origin:'Particular'},
  {key:'laser-area',kind:'procedure',match:['laser'],label:'Laser por área',amount:960,origin:'Particular'},
  {key:'laser-area-cirurgia',kind:'procedure',match:['laser'],label:'Laser por área na cirurgia',amount:900,origin:'Particular'},
  {key:'laser-cicatriz-1',kind:'procedure',match:['laser','cicatriz'],label:'Laser cicatriz - 1 sessão',amount:750,origin:'Particular'},
  {key:'laser-cicatriz-3',kind:'procedure',match:['laser','cicatriz'],label:'Laser cicatriz - 3 sessões',amount:1900,origin:'Particular'},
  {key:'laser-cool-peel',kind:'procedure',match:['laser'],label:'Laser frio (Cool Peel)',amount:1600,origin:'Particular'},
  {key:'laser-combo',kind:'procedure',match:['laser'],label:'Laser Combo - Face + Pescoço + Colo',amount:5000,origin:'Particular'},
  {key:'us-olhos',kind:'procedure',match:['microfoc','ultrassom'],label:'Ultrassom microfocado - área dos olhos',amount:1600,origin:'Particular'},
  {key:'us-face',kind:'procedure',match:['microfoc','ultrassom'],label:'Ultrassom microfocado - face',amount:2400,origin:'Particular'},
  {key:'us-face-full',kind:'procedure',match:['microfoc','ultrassom'],label:'Ultrassom microfocado - face full',amount:3500,origin:'Particular'},
  {key:'us-papada',kind:'procedure',match:['microfoc','ultrassom'],label:'Ultrassom microfocado - papada',amount:1600,origin:'Particular'},
  {key:'us-pescoco',kind:'procedure',match:['microfoc','ultrassom'],label:'Ultrassom microfocado - pescoço',amount:1600,origin:'Particular'},
  {key:'us-adicional',kind:'procedure',match:['microfoc','ultrassom'],label:'Ultrassom microfocado - área adicional',amount:1200,origin:'Particular'},
  {key:'us-full-papada',kind:'procedure',match:['microfoc','ultrassom'],label:'Ultrassom microfocado - full face + papada',amount:4100,origin:'Particular'},
  {key:'us-corporal',kind:'procedure',match:['microfoc','ultrassom'],label:'Ultrassom microfocado - área corporal',amount:3500,origin:'Particular'},
  {key:'us-corporal-bi',kind:'procedure',match:['microfoc','ultrassom'],label:'Ultrassom microfocado - área corporal bilateral',amount:4900,origin:'Particular'},
  {key:'profhilo',kind:'procedure',match:['profhilo'],label:'Profhilo - 1 seringa',amount:2200,origin:'Particular'},
  {key:'bioestimulador',kind:'procedure',match:['bioestimul'],label:'Bioestimulador de colágeno',amount:2280,origin:'Particular'},
  {key:'bioestimulador-cirurgia',kind:'procedure',match:['bioestimul'],label:'Bioestimulador - na cirurgia',amount:1900,origin:'Particular'},
  {key:'skinvive',kind:'procedure',match:['skinvive','skinbooster'],label:'Skinbooster / Skinvive',amount:900,origin:'Particular'},
  {key:'pdo',kind:'procedure',match:['pdo'],label:'Fios de PDO',amount:450,origin:'Particular'},
  {key:'preenchimento-labial',kind:'procedure',match:['preenchimento labial'],label:'Preenchimento labial',amount:1500,origin:'Particular'},
  {key:'acido-hialuronico',kind:'procedure',match:['hialuron'],label:'Ácido hialurônico - 1 seringa',amount:1350,origin:'Particular'},
  {key:'maos',kind:'procedure',match:['mao','mãos'],label:'Rejuvenescimento mãos - Bioestimulador + Ácido H.',amount:2900,origin:'Particular'},
  {key:'anti-vegf-individual',kind:'procedure',match:['anti vegf'],label:'Anti-VEGF / Intravítreo - individual',amount:4500,origin:'Particular'},
  {key:'anti-vegf-3',kind:'procedure',match:['anti vegf'],label:'Anti-VEGF / Intravítreo - pacote com 3',amount:12500,origin:'Particular'},
  {key:'triancinolona',kind:'procedure',match:['triancinolona'],label:'Triancinolona',amount:1200,origin:'Particular'},
  {key:'ozurdex',kind:'procedure',match:['ozurdex'],label:'Ozurdex - cada aplicação',amount:4800,origin:'Particular'},
  {key:'avastin-1',kind:'procedure',match:['avastin'],label:'Avastin - individual',amount:1900,origin:'Particular'},
  {key:'avastin-3',kind:'procedure',match:['avastin'],label:'Avastin - 3 aplicações',amount:4500,origin:'Particular'},
  {key:'vabysmo-1',kind:'procedure',match:['vabysmo'],label:'Vabysmo - 1 aplicação',amount:5250,origin:'Particular'},
  {key:'vabysmo-4',kind:'procedure',match:['vabysmo'],label:'Vabysmo - 4 aplicações + OCT',amount:19900,origin:'Particular'}
];

const presetAppliesToName = (preset,name) => {
  const n=normalizeCatalogLabel(name);
  const matches=(preset.match||[]).some(term=>n.includes(normalizeCatalogLabel(term)));
  const excluded=(preset.exclude||[]).some(term=>n.includes(normalizeCatalogLabel(term)));
  return matches && !excluded;
};

const normalizeCatalogLabel = value => String(value||'')
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g,'')
  .toLowerCase()
  .replace(/[^a-z0-9]+/g,' ')
  .trim();

const knownCatalogKind = value => {
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

const mapProcedureRows = (rows, surgeryById=new Map(), surgeryByGroup=new Map()) => rows.filter(r=>r.record_kind!=='surgery_legacy').map(r=>{
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
const groupProcedureRows = rows => {
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


function FinancialLineChart({title,labels,series}){
  const width=860,height=290;
  const pad={l:82,r:24,t:34,b:48};
  const values=series.flatMap(item=>item.values||[]).map(numericValue);
  const max=Math.max(...values,1);
  const plotW=width-pad.l-pad.r,plotH=height-pad.t-pad.b;
  const x=i=>labels.length<=1?pad.l+plotW/2:pad.l+(i/(labels.length-1))*plotW;
  const y=v=>pad.t+plotH-(numericValue(v)/max)*plotH;
  const grid=[0,.25,.5,.75,1];
  return <div className="card finance-line-card">
    <div className="finance-chart-head"><div><span>EVOLUÇÃO FINANCEIRA</span><h3>{title}</h3><p>Produção, recebimentos e potencial do período.</p></div>
      <div className="finance-line-legend">{series.map(s=><span key={s.name}><i style={{background:s.color}}/>{s.name}</span>)}</div>
    </div>
    <svg className="finance-line-svg" viewBox={`0 0 ${width} ${height}`} role="img" aria-label={title}>
      {grid.map(g=>{const value=max*g;const gy=y(value);return <g key={g}>
        <line x1={pad.l} x2={width-pad.r} y1={gy} y2={gy} stroke="#e7edef" strokeWidth="1"/>
        <text x={pad.l-10} y={gy+4} textAnchor="end" className="finance-axis">{moneyBR(value).replace(',00','')}</text>
      </g>})}
      {labels.map((label,i)=><text key={`${label}-${i}`} x={x(i)} y={height-15} textAnchor="middle" className="finance-axis">{label}</text>)}
      {series.map(s=>{
        const points=(s.values||[]).map((v,i)=>`${x(i)},${y(v)}`).join(' ');
        return <g key={s.name}>
          <polyline points={points} fill="none" stroke={s.color} strokeWidth="3.2" strokeLinejoin="round" strokeLinecap="round"/>
          {(s.values||[]).map((v,i)=><circle key={i} cx={x(i)} cy={y(v)} r="4.5" fill="#fff" stroke={s.color} strokeWidth="2.5"/>)}
        </g>;
      })}
    </svg>
  </div>;
}

function FinancialBars({title,subtitle,data}){
  const entries=Object.entries(data||{}).filter(([,v])=>numericValue(v)>0).sort((a,b)=>numericValue(b[1])-numericValue(a[1])).slice(0,10);
  const max=Math.max(...entries.map(([,v])=>numericValue(v)),1);
  return <div className="card finance-ranking-card">
    <div className="finance-chart-head"><div><span>ANÁLISE FINANCEIRA</span><h3>{title}</h3>{subtitle&&<p>{subtitle}</p>}</div></div>
    {entries.length?<div className="finance-ranking-list">{entries.map(([label,value])=><div key={label} className="finance-ranking-row">
      <div className="finance-ranking-label" title={label}>{label}</div>
      <div className="finance-ranking-track"><i style={{width:`${Math.max(3,(numericValue(value)/max)*100)}%`}}/></div>
      <b>{moneyBR(value)}</b>
    </div>)}</div>:<div className="empty">Sem valores financeiros para os filtros atuais.</div>}
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


const automaticPriceForItem = ({kind,name,origin='Particular',insuranceName='',eye='',detail='',associatedWithSurgery=false}) => {
  if(!name) return null;
  const candidates=STANDARD_PRICE_PRESETS.filter(p=>p.kind===kind&&presetAppliesToName(p,name));
  if(!candidates.length) return null;

  const originWanted=origin==='SUS'?'SUS':origin==='Convênio'?'Convênio':'Particular';
  const ins=normalizeCatalogLabel(insuranceName);
  const eyeNorm=normalizeCatalogLabel(eye);
  const detailNorm=normalizeCatalogLabel(detail);
  const serviceNorm=normalizeCatalogLabel(name);

  // Condições especiais da tabela de exames.
  if(kind==='exam'&&serviceNorm.includes('consulta')){
    if(ins.includes('saude vida')) return {amount:100,key:'exam-consulta-saude-vida',label:'Consulta - Saúde Vida + Organização',origin:'Convênio'};
    if(ins.includes('vale saude')) return {amount:200,key:'exam-consulta-vale-saude',label:'Consulta - Vale Saúde',origin:'Convênio'};
  }

  // Valores de procedimentos da tabela são gerais. Quando não existe preço
  // específico por convênio/SUS, o mesmo valor da tabela continua válido.
  let pool=candidates.filter(p=>p.origin===originWanted);
  if(!pool.length && kind==='procedure') pool=[...candidates];

  // Cirurgias sem diferenciação explícita por convênio usam o valor geral da tabela.
  // Catarata e Pterígio continuam respeitando os grupos específicos.
  if(!pool.length && kind==='surgery'){
    const isCataract=serviceNorm.includes('catarata')||serviceNorm.includes('faco')||serviceNorm==='trec';
    const isPterygium=serviceNorm.includes('pterig');
    if(!isCataract&&!isPterygium) pool=[...candidates];
    if(isPterygium && !ins.includes('iamspe')){
      const generic=candidates.filter(p=>normalizeCatalogLabel(p.label).includes('pterigio particular convenio'));
      if(generic.length) pool=generic;
    }
  }

  if(!pool.length) return null;

  const score=p=>{
    const label=normalizeCatalogLabel(p.label);
    let s=0;
    if(label.includes(serviceNorm)) s+=5;

    if(kind==='surgery'){
      if((ins.includes('unimed')||ins.includes('oeste'))&&label.includes('unimed oeste')) s+=30;
      if((ins.includes('cassi')||ins.includes('iamspe')||ins.includes('sulamerica'))&&(label.includes('cassi')||label.includes('iamspe')||label.includes('sulamerica'))) s+=30;
      if(ins.includes('iamspe')&&label.includes('iamspe')) s+=35;
      if(serviceNorm.includes('pterig')&&!ins.includes('iamspe')&&label.includes('particular convenio')) s+=25;
    }

    if(eyeNorm.includes('ambos')){
      if(label.includes('2 olhos')||label.includes(' ao ')) s+=18;
      if(label.includes('1 olho')||label.includes('individual')) s-=10;
    }else if(eyeNorm.includes('direito')||eyeNorm.includes('esquerdo')){
      if(label.includes('1 olho')||label.includes('individual')||label.includes('cada olho')) s+=18;
      if(label.includes('2 olhos')||label.includes(' ao ')) s-=10;
    }

    if(detailNorm){
      for(const token of detailNorm.split(' ').filter(x=>x.length>2)) if(label.includes(token)) s+=7;
    }

    if(associatedWithSurgery){
      if(label.includes('na cirurgia')||label.includes('pacote na cirurgia')) s+=30;
      else if(kind==='procedure'&&(serviceNorm.includes('botox')||serviceNorm.includes('toxina')||serviceNorm.includes('laser')||serviceNorm.includes('bioestimul'))) s-=6;
    }else if(kind==='procedure'&&(label.includes('na cirurgia')||label.includes('pacote na cirurgia'))){
      s-=20;
    }
    return s;
  };

  const ranked=pool.map(p=>({p,s:score(p)})).sort((a,b)=>b.s-a.s);
  if(!ranked.length) return null;
  const best=ranked[0];
  const ties=ranked.filter(x=>x.s===best.s);
  const amounts=[...new Set(ties.map(x=>numericValue(x.p.amount)))];
  if(amounts.length>1) return null;
  return {amount:numericValue(best.p.amount),key:best.p.key,label:best.p.label,origin:best.p.origin};
};

const automaticPriceTotal = ({kind,names,origin,insuranceName,eye,detail}) => {
  const clean=(Array.isArray(names)?names:[names]).filter(Boolean);
  if(!clean.length) return null;
  const picked=clean.map(name=>automaticPriceForItem({kind,name,origin,insuranceName,eye,detail}));
  if(picked.some(x=>!x)) return null;
  return {amount:picked.reduce((sum,x)=>sum+x.amount,0),keys:picked.map(x=>x.key),labels:picked.map(x=>x.label)};
};

function PaymentTrafficLight({record,onClick}){
  const tone=paymentTone(record);
  const label=paymentToneLabel(record);
  return <button type="button" className={`payment-light ${tone}`} onClick={onClick} title={`${label}. Clique para abrir o lançamento financeiro.`} aria-label={label}>
    <span className="payment-light-dot"/>
    <small>{label}</small>
  </button>;
}

function PaymentBreakdownEditor({form,setForm,insurances}){
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

function RepassesEditor({form,setForm}){
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

function FinanceQuickSummary({form,setForm}){
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


const V36PaymentSummaryStyle = () => <style>{`
  .finance-summary-payment-select{
    display:flex;
    align-items:center;
    gap:8px;
    margin-top:4px;
  }
  .finance-summary-payment-select>i{
    width:10px;
    height:10px;
    border-radius:999px;
    background:currentColor;
    flex:0 0 auto;
  }
  .finance-summary-payment-select select{
    width:100%;
    min-width:132px;
    border:1px solid rgba(49,95,114,.18);
    border-radius:10px;
    background:#fff;
    color:#20313b;
    font:inherit;
    font-weight:700;
    padding:7px 30px 7px 10px;
    cursor:pointer;
    outline:none;
  }
  .finance-summary-status.paid .finance-summary-payment-select{color:#4f8a6b}
  .finance-summary-status.partial .finance-summary-payment-select{color:#b9802b}
  .finance-summary-status.unpaid .finance-summary-payment-select{color:#b85c5c}
  .finance-summary-payment-select select:focus{
    border-color:#4d8191;
    box-shadow:0 0 0 3px rgba(77,129,145,.12);
  }
  .auto-price-field{display:grid;gap:6px}.auto-price-field small{color:#667985;font-size:10px;line-height:1.45}.auto-price-field input{font-weight:800}.package-auto-price{display:block;margin-top:4px;color:#667985;font-size:9px}.brand-mark.brand-logo-ready{padding:0;background:transparent;box-shadow:none;overflow:visible}.brand-mark.brand-logo-ready img{display:block;width:100%;height:100%;object-fit:contain}.brand-mark span{display:grid;place-items:center;width:100%;height:100%}
`}</style>;


const V40LoginPriceStyle = () => <style>{`
  .user-box small{display:block;margin-top:5px;color:#83929a;font-size:9px;line-height:1.35}
`}</style>;
const V42AlignmentStyle = () => <style>{`
  .filter-check-section{display:flex;flex-direction:column;gap:7px}.filter-check-section>label{font-size:12px;font-weight:800;color:#52636b;text-transform:uppercase;letter-spacing:.04em}.filter-check-grid{display:flex;gap:8px;flex-wrap:wrap}.filter-check{display:inline-flex;align-items:center;gap:8px;border:1px solid #d6e0e2;background:#fff;color:#40525b;border-radius:10px;padding:8px 11px;font:inherit;font-weight:700;cursor:pointer;transition:transform .18s ease,box-shadow .18s ease,background .18s ease,border-color .18s ease;position:relative;overflow:hidden}.filter-check::after{content:'';position:absolute;inset:0;background:linear-gradient(110deg,transparent 35%,rgba(255,255,255,.65) 50%,transparent 65%);transform:translateX(-140%);transition:transform .45s ease}.filter-check:hover{transform:translateY(-2px) scale(1.015)}.filter-check:active{transform:scale(.97)}.filter-check.selected{animation:filterPop .28s ease}.filter-check.selected::after{transform:translateX(140%)}@keyframes filterPop{0%{transform:scale(.94)}60%{transform:scale(1.04)}100%{transform:scale(1)}}.filter-check span{display:grid;place-items:center;width:18px;height:18px;border:1px solid #bdcacc;border-radius:5px;background:#fff;color:#fff;font-size:12px}.filter-check:hover{border-color:#8eaaa0;background:#f7faf9}.filter-check.selected{border-color:#75a48e;background:#eaf5ef;color:#315f4c;box-shadow:0 0 0 2px rgba(79,138,107,.08)}.filter-check.selected span{background:#4f8a6b;border-color:#4f8a6b}.daily-report-filters{align-items:flex-end}.daily-report-filters .full{grid-column:1/-1}.promissory-status{display:inline-flex;padding:5px 9px;border-radius:999px;font-size:12px;font-weight:800}.promissory-status.paid{background:#e8f4ed;color:#397156}.promissory-status.partial{background:#fff4dc;color:#9a671d}.promissory-status.open{background:#edf1f3;color:#55646b}.promissory-status.overdue{background:#fbeaea;color:#a94a4a}.finance-kpi.danger{border-top-color:#b85c5c}.promissory-page .finance-table td{vertical-align:middle}

  .period-filter{display:flex;gap:6px;align-items:center;min-width:210px}.period-filter select{max-width:80px}.period-filter input{min-width:132px}
  .form-validation-banner{display:flex;flex-direction:column;gap:4px;background:#fff1f1;border:1px solid #e4a1a1;color:#963f3f;border-radius:12px;padding:12px 14px;animation:validationShake .34s ease}
  .form-validation-banner strong{font-size:14px}.form-validation-banner span{font-size:12px}
  @keyframes validationShake{0%,100%{transform:translateX(0)}25%{transform:translateX(-5px)}50%{transform:translateX(5px)}75%{transform:translateX(-3px)}}
  .field-error label{color:#a74343!important}.field-error input,.field-error select,.field-error textarea,.field-error .exam-multi-select,.field-error .procedure-multi-select,.field-error .finance-quick-summary{border-color:#c75a5a!important;background:#fff8f8!important;box-shadow:0 0 0 3px rgba(199,90,90,.09)!important}
`}</style>;

function PeriodFilter({mode,setMode,month,setMonth,date,setDate}){
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

function PatientCell({name,phone,birthDate}){
  const url = whatsappUrl(phone);
  return <div className="patient-cell">
    <div><b>{name}</b>{birthDate&&<small>Nasc.: {brDate(birthDate)}</small>}{phone&&<small>{phone}</small>}</div>
    {url&&<a className="wa-btn" href={url} target="_blank" rel="noreferrer" title={`Abrir WhatsApp de ${name}`}><WhatsAppIcon/></a>}
  </div>;
}
function ObservationCell({text}) {
  return text ? <div className="obs-cell" title={text}>{text}</div> : <span className="muted">—</span>;
}
function LinkedProceduresCell({names=[]}){
  const items=(Array.isArray(names)?names:[]).filter(Boolean);
  if(!items.length) return <span className="muted">Sem procedimentos vinculados</span>;
  return <div className="linked-procedures-cell">
    <small>Vinculados ao pacote</small>
    <div>{items.map((name,index)=><span className="linked-procedure-chip" key={`${name}-${index}`}>{name}</span>)}</div>
  </div>;
}

function ProcedureListCell({names=[]}){
  const items=(Array.isArray(names)?names:[]).filter(Boolean);
  if(!items.length) return <span className="muted">—</span>;
  return <div className="procedure-launch-cell">
    <small>{items.length} {items.length===1?'procedimento':'procedimentos'} neste lançamento</small>
    <div>{items.map((name,index)=><span className="linked-procedure-chip" key={`${name}-${index}`}>{name}</span>)}</div>
  </div>;
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

function BrandMark(){
  const [failed,setFailed]=useState(false);
  return <div className={`brand-mark ${failed?'':'brand-logo-ready'}`} aria-label="OftalmoCastro">
    {failed
      ? <span>OC</span>
      : <img src="/logo-oftalmocastro.png" alt="OftalmoCastro" onError={()=>setFailed(true)}/>}
  </div>;
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

function FinancialDocumentsField({form,setForm,onOpen,onDownload}){
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


const UNIMED_STEP1_DOCUMENTS = [
  'Autorização do convênio',
  'Termo de consentimento',
  'Folha de sala',
  'Descrição cirúrgica',
  'Folha de anestesia (quando houver)',
  'Entrevista com a enfermagem',
  'Prontuário físico do paciente',
  'Selo de uso conforme o procedimento'
];
const UNIMED_STEP1_CHECKLIST = [
  'Autorização assinada pelo paciente e pelo médico',
  'Autorização, descrição cirúrgica e folha de anestesia carimbadas pelo médico responsável pela cirurgia e pelo auxiliar',
  'Termo de consentimento assinado pelo paciente',
  'Folha de sala preenchida',
  'Horário de início e de término da cirurgia informados',
  'Selo correto colado conforme o procedimento',
  'Execução no prestador/login 714 (clínica Oftalmo Castro), nunca em nome do médico',
  'Entrevista com a enfermagem anexada',
  'Folha de anestesia anexada (quando houver)',
  'Prontuário físico do paciente incluído'
];
const UNIMED_STEP1_DOC_STATUS = ['Pendente','Recebido','Conferido','Devolvido'];
const UNIMED_STEP1_DOSSIER_STATUS = ['Aguardando documentos','Em conferência','Enviado ao Faturamento','Devolvido à Recepção','Devolvido ao Centro Cirúrgico','Concluído'];
const UNIMED_STEP_STATUS = ['Não iniciada','Em andamento','Concluída'];

const FINANCE_RISK_OPTIONS=[['','Risco: definir'],['baixo','Risco baixo'],['medio','Risco médio'],['alto','Risco alto']];
const FINANCE_COMPLEXITY_OPTIONS=[['','Complexidade: definir'],['baixo','Complexidade baixa'],['medio','Complexidade média'],['alto','Complexidade alta']];
const FINANCE_DELEGATION_OPTIONS=[['','Delegação: definir'],['nao','Mantém com o gestor'],['sim','Delegar para…']];
const FINANCE_RECURRENCE=['Todos os dias','Dia que tem procedimento','Quando tem pedido','1x por semana','2x ao mês','1x mês','1x mês até dia 10'];
const financeLevelScore={baixo:0,medio:1,alto:2};
const financeLevelLabel=v=>v==='baixo'?'baixo':v==='medio'?'médio':v==='alto'?'alto':'a definir';
const financeLevelClass=v=>v==='baixo'?'ok':v==='medio'?'attention':v==='alto'?'danger':'neutral';
const financeStepEvaluated=s=>Boolean(s?.risco&&s?.complexidade&&s?.delegar);
const financeDateBR=value=>{if(!value)return '';const d=new Date(value);return Number.isNaN(d.getTime())?'':d.toLocaleDateString('pt-BR');};
const financeUid=()=>crypto?.randomUUID?.()||`${Date.now()}-${Math.random().toString(36).slice(2)}`;
const financeInitialStep={id:'unimed-etapa-1',texto:'Reunir todas as cirurgias realizadas: folhas de sala, autorizações e prontuário físico do paciente',delegar:'sim',responsavel:'Centro Cirúrgico',risco:'alto',complexidade:'medio'};
const financeStep1Pop={
  objetivo:'Reunir, organizar e conferir a documentação obrigatória de cada cirurgia Unimed para que o Faturamento receba o dossiê completo e sem pendências.',
  quando:'Ao final de cada cirurgia, antes do envio ao Faturamento.',
  responsavel:'Centro Cirúrgico',
  sistemas:['Prontuário físico','Autorização do convênio','Folha de sala'],
  passos:[
    {titulo:'Receber a documentação da Recepção',detalhe:'Receber autorização e termo de consentimento já assinados pelo paciente.',alerta:'A autorização deve também estar assinada e carimbada pelo médico responsável.',responsavel:'Centro Cirúrgico'},
    {titulo:'Completar a documentação da sala',detalhe:'Após a cirurgia, preencher a folha de sala, registrar horários de início e término e aplicar o selo correspondente ao procedimento.',alerta:'Aplicação = selo da injeção intravítrea; catarata = selo da lente intraocular; autotransplante conjuntival = selo da cola biológica.',responsavel:'Centro Cirúrgico'},
    {titulo:'Conferir a descrição cirúrgica',detalhe:'O médico preenche a descrição cirúrgica na folha de sala após desfazer a paramentação.',alerta:'Autorização, descrição cirúrgica e folha de anestesia devem estar carimbadas pelo médico e pelo auxiliar.',responsavel:'Médico / Centro Cirúrgico'},
    {titulo:'Conferir o prestador de execução',detalhe:'Antes de liberar o dossiê, confirmar que a execução foi registrada no prestador/login 714 - clínica Oftalmo Castro.',alerta:'Nunca executar a cirurgia Unimed em nome do médico.',responsavel:'Centro Cirúrgico'},
    {titulo:'Montar e enviar o dossiê',detalhe:'Juntar autorização, termo, folha de sala, descrição, folha de anestesia quando houver, entrevista da enfermagem e prontuário físico; grampear, conferir e enviar ao Faturamento.',alerta:'Documento faltando pode gerar glosa total.',responsavel:'Centro Cirúrgico'}
  ],
  atencao:['Execução sempre pelo prestador/login 714 - clínica Oftalmo Castro.','Falta de assinatura/carimbo do médico na autorização: devolver à Recepção.','Falta de preenchimento, carimbo ou assinatura na folha de sala ou descrição cirúrgica: devolver ao Centro Cirúrgico.','Conferir selo, horários, assinaturas e carimbos antes do envio.'],
  duvidas:'Quando o Faturamento devolver por assinatura/carimbo na autorização, a correção é da Recepção. Pendências de folha de sala ou descrição cirúrgica retornam ao Centro Cirúrgico.'
};

function TasksPopsPanel({profile,session}){
  const canWrite=profile?.role==='admin'||profile?.role==='orientadora';
  const canUpload=canWrite;
  const canAdmin=profile?.role==='admin';
  const storageKey=`oc_finance_tasks_tab_${session?.user?.id||'user'}`;
  const [tab,setTabState]=useState(()=>localStorage.getItem(storageKey)||'tarefas');
  const setTab=v=>{setTabState(v);try{localStorage.setItem(storageKey,v)}catch{}};
  const [loading,setLoading]=useState(true);
  const [notice,setNotice]=useState('Carregando as tarefas…');
  const [tasks,setTasks]=useState([]); const [pops,setPops]=useState([]); const [videos,setVideos]=useState([]); const [materials,setMaterials]=useState([]); const [submissions,setSubmissions]=useState([]);
  const [filter,setFilter]=useState('todas'); const [openTasks,setOpenTasks]=useState(()=>new Set());
  const [editor,setEditor]=useState(null); const [popTask,setPopTask]=useState(null); const [popScrollUpload,setPopScrollUpload]=useState(false); const [busy,setBusy]=useState('');
  const [inlineDelete,setInlineDelete]=useState(''); const [materialPicker,setMaterialPicker]=useState(null); const [viewer,setViewer]=useState(null);
  const taskRefs=useRef({}); const responsibleInputRefs=useRef({});

  const flash=(message,ms=6000)=>{setNotice(message);if(ms)setTimeout(()=>setNotice(v=>v===message?'':v),ms)};
  const dbError=(error,fallback)=>{console.error(error);flash(`${fallback} (${error?.code||'erro'}).`,6000)};
  const load=async(silent=false)=>{
    if(!silent){setLoading(true);setNotice('Carregando as tarefas…')}
    const [t,p,v,m,e]=await Promise.all([
      supabase.from('finance_tasks').select('*').eq('cargo','financeiro').order('ordem',{ascending:true}),
      supabase.from('finance_pops').select('*'), supabase.from('finance_videos').select('*').order('enviado_em',{ascending:false}),
      supabase.from('finance_materials').select('*').order('enviado_em',{ascending:false}), supabase.from('finance_submissions').select('*').order('enviado_em',{ascending:false})
    ]);
    const firstError=[t,p,v,m,e].find(x=>x.error)?.error;
    if(firstError){setLoading(false);return dbError(firstError,'Não foi possível carregar os dados')}
    setTasks(t.data||[]);setPops(p.data||[]);setVideos(v.data||[]);setMaterials(m.data||[]);setSubmissions(e.data||[]);setLoading(false);setNotice('');
  };
  useEffect(()=>{load();},[]);
  useEffect(()=>{
    const ch=supabase.channel('finance-tasks-pops-live')
      .on('postgres_changes',{event:'*',schema:'public',table:'finance_tasks'},()=>load(true))
      .on('postgres_changes',{event:'*',schema:'public',table:'finance_pops'},()=>load(true))
      .on('postgres_changes',{event:'*',schema:'public',table:'finance_videos'},()=>load(true))
      .on('postgres_changes',{event:'*',schema:'public',table:'finance_materials'},()=>load(true))
      .on('postgres_changes',{event:'*',schema:'public',table:'finance_submissions'},()=>load(true)).subscribe();
    return()=>{supabase.removeChannel(ch)};
  },[]);

  const taskPop=t=>pops.find(p=>p.id===t.id)||null;
  const taskVideos=t=>videos.filter(v=>v.tarefa_id===t.id);
  const taskSubmissions=t=>submissions.filter(v=>v.tarefa_id===t.id);
  const popState=t=>{
    const vids=taskVideos(t);const env=taskSubmissions(t);const p=taskPop(t);
    if(vids.some(v=>v.status!=='processado')||env.some(x=>x.status==='aguardando'))return{tone:'attention',label:'Vídeo recebido · atualizando'};
    if(p?.status==='validado')return{tone:'ok',label:`POP validado v${p.versao||1}`};
    if(p)return{tone:'neutral',label:`Rascunho v${p.versao||1}`};return{tone:'neutral',label:'Sem POP'};
  };
  const totalSteps=tasks.reduce((s,t)=>s+(t.subtarefas||[]).length,0); const delegated=tasks.reduce((s,t)=>s+(t.subtarefas||[]).filter(x=>x.delegar==='sim').length,0); const popVideoCount=tasks.filter(t=>taskVideos(t).length>0).length;
  const filteredTasks=tasks.filter(t=>{const steps=t.subtarefas||[];if(filter==='sem-video')return taskVideos(t).length===0;if(filter==='avaliacao')return steps.some(s=>!financeStepEvaluated(s));if(filter==='delegadas')return steps.some(s=>s.delegar==='sim');return true;});
  const toggleTask=id=>setOpenTasks(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n});
  const scrollTask=id=>{setTab('tarefas');setFilter('todas');setOpenTasks(prev=>new Set([...prev,id]));setTimeout(()=>taskRefs.current[id]?.scrollIntoView({behavior:'smooth',block:'center'}),80)};
  const updateStep=async(task,index,patch)=>{
    const steps=[...(task.subtarefas||[])];steps[index]={...steps[index],...patch};if(patch.delegar!=='sim'&&'delegar'in patch)steps[index].responsavel='';
    setTasks(prev=>prev.map(t=>t.id===task.id?{...t,subtarefas:steps}:t));
    const {error}=await supabase.from('finance_tasks').update({subtarefas:steps,atualizado_em:new Date().toISOString(),atualizado_por:session.user.id}).eq('id',task.id);
    if(error){dbError(error,'Não foi possível salvar');load(true)}else flash('Salvo',1600);
  };
  const newTask=()=>setEditor({mode:'new',id:null,tarefa:'',recorrencia:'',subtarefas:[{id:financeUid(),texto:'',delegar:'',responsavel:'',risco:'',complexidade:''}]});
  const editTask=t=>setEditor({mode:'edit',id:t.id,tarefa:t.tarefa||'',recorrencia:t.recorrencia||'',subtarefas:(t.subtarefas||[]).map(x=>({...x}))});
  const saveTask=async()=>{
    if(!editor?.tarefa?.trim())return flash('Dê um nome para a tarefa.',4000);
    const steps=(editor.subtarefas||[]).map(s=>({...s,texto:String(s.texto||'').trim()})).filter(s=>s.texto); const now=new Date().toISOString();
    if(editor.mode==='new'){
      const nums=tasks.map(t=>Number(String(t.codigo||'').match(/(\d+)/)?.[1]||0));const next=(Math.max(0,...nums)+1);const maxOrder=Math.max(0,...tasks.map(t=>Number(t.ordem||0)));
      const row={codigo:`FIN-${String(next).padStart(2,'0')}`,ordem:maxOrder+1,cargo:'financeiro',recorrencia:editor.recorrencia.trim(),tarefa:editor.tarefa.trim(),subtarefas:steps,expected_steps:steps.length,criado_em:now,criado_por:session.user.id,atualizado_em:now,atualizado_por:session.user.id};
      const {data,error}=await supabase.from('finance_tasks').insert(row).select('*').single();if(error)return dbError(error,'Não foi possível salvar');setEditor(null);flash('Tarefa criada.',3000);if(data)scrollTask(data.id);
    }else{
      const original=tasks.find(t=>t.id===editor.id);const merged=steps.map(s=>{const old=(original?.subtarefas||[]).find(x=>x.id===s.id);return old?{...old,...s}:s});
      const {error}=await supabase.from('finance_tasks').update({tarefa:editor.tarefa.trim(),recorrencia:editor.recorrencia.trim(),subtarefas:merged,atualizado_em:now,atualizado_por:session.user.id}).eq('id',editor.id);if(error)return dbError(error,'Não foi possível salvar');setEditor(null);flash('Alterações salvas.',3000);
    }
  };
  const duplicateTask=async t=>{setBusy(`dup:${t.id}`);const nums=tasks.map(x=>Number(String(x.codigo||'').match(/(\d+)/)?.[1]||0));const next=Math.max(0,...nums)+1;const sorted=[...tasks].sort((a,b)=>Number(a.ordem||9999)-Number(b.ordem||9999));const idx=sorted.findIndex(x=>x.id===t.id);const following=sorted[idx+1];const order=following?((Number(t.ordem||0)+Number(following.ordem||0))/2):Math.max(0,...tasks.map(x=>Number(x.ordem||0)))+1;const now=new Date().toISOString();
    const row={codigo:`FIN-${String(next).padStart(2,'0')}`,ordem:order,cargo:t.cargo||'financeiro',recorrencia:t.recorrencia||'',tarefa:`Cópia de ${t.tarefa||''}`,subtarefas:(t.subtarefas||[]).map(s=>({...s,id:financeUid()})),expected_steps:t.expected_steps||((t.subtarefas||[]).length),duplicada_de:t.id,criado_em:now,criado_por:session.user.id,atualizado_em:now,atualizado_por:session.user.id};
    const {data,error}=await supabase.from('finance_tasks').insert(row).select('*').single();setBusy('');if(error)return dbError(error,'Não foi possível duplicar');setFilter('todas');if(data){setEditor({mode:'edit',id:data.id,tarefa:data.tarefa,recorrencia:data.recorrencia,subtarefas:data.subtarefas||[]});scrollTask(data.id)}};
  const deleteTask=async t=>{const {error}=await supabase.from('finance_tasks').delete().eq('id',t.id);if(error)return dbError(error,'Não foi possível excluir');setInlineDelete('');flash('Tarefa excluída.',2500)};

  const openPop=(t,toUpload=false)=>{setPopTask(t);setPopScrollUpload(toUpload)};
  const uploadVideo=async(task,step,file)=>{
    if(!file)return;const ext=String(file.name||'').split('.').pop().toLowerCase();const allowed=['mp4','m4v','mov','qt','webm'];if(!allowed.includes(ext))return flash(`Formato .${ext||'?'} não aceito. Envie em MP4 ou WebM.`,6000);const mb=file.size/1024/1024;if(mb>20)return flash(`O vídeo tem ${mb.toFixed(1)} MB; o limite é 20 MB. Grave a etapa em partes menores.`,6000);
    const head=new Uint8Array(await file.slice(0,12).arrayBuffer());const isWebm=head[0]===0x1a&&head[1]===0x45&&head[2]===0xdf&&head[3]===0xa3;const text=String.fromCharCode(...head.slice(4,8));const isMp4=text==='ftyp';if(!isWebm&&!isMp4)return flash('Este arquivo não é um vídeo de verdade. Abra o vídeo no computador para conferir e exporte de novo em MP4.',6000);
    setBusy(`video:${step?.id||task.id}`);const path=`finance-pops/videos/${task.id}/${financeUid()}-${file.name.replace(/[^a-zA-Z0-9._-]+/g,'-')}`;const {error:upErr}=await supabase.storage.from('portal-files').upload(path,file,{contentType:isWebm?'video/webm':'video/mp4'});if(upErr){setBusy('');return dbError(upErr,'O envio falhou')};const {data:urlData}=supabase.storage.from('portal-files').getPublicUrl(path);const row={tarefa_id:task.id,codigo:task.codigo,tarefa:task.tarefa,etapa_id:step?.id||null,etapa_num:step?(task.subtarefas||[]).findIndex(s=>s.id===step.id)+1:null,etapa_texto:step?.texto||null,arquivo_id:path,url:urlData?.publicUrl||'',nome_arquivo:file.name,tamanho_mb:Number(mb.toFixed(2)),enviado_em:new Date().toISOString(),enviado_por:session.user.id,status:'aguardando'};const {error}=await supabase.from('finance_videos').insert(row);setBusy('');if(error)return dbError(error,'O envio falhou');flash('Recebido. O POP será atualizado na próxima rodada.',8000)};
  const registerDrive=async(task,step,link,obs='')=>{if(!/^https:\/\/(drive|docs)\.google\.com\//i.test(link||''))return flash('O link precisa ser do Google Drive (drive.google.com).',5000);const id=(String(link).match(/\/d\/([A-Za-z0-9_-]{20,})/)||String(link).match(/[?&]id=([A-Za-z0-9_-]{20,})/))?.[1];if(!id)return flash('Cole o link do arquivo de vídeo (botão Compartilhar → Copiar link no Drive).',6000);const {error}=await supabase.from('finance_submissions').insert({tarefa_id:task.id,codigo:task.codigo,tarefa:task.tarefa,etapa_id:step?.id||null,etapa_num:step?(task.subtarefas||[]).findIndex(s=>s.id===step.id)+1:null,etapa_texto:step?.texto||null,drive_id:id,drive_link:link,enviado_em:new Date().toISOString(),enviado_por:session.user.id,status:'aguardando',mensagem:obs||null});if(error)return dbError(error,'Não foi possível registrar o link');flash('Link registrado. O vídeo entra no portal na próxima rodada automática.',8000)};
  const importMaterial=async(task,step,type,file)=>{if(!file)return;const head=new Uint8Array(await file.slice(0,5).arrayBuffer());const signature=String.fromCharCode(...head);if(!file.name.toLowerCase().endsWith('.pdf')||signature!=='%PDF-')return flash('Esse arquivo não é um PDF. No Word ou PowerPoint, use Arquivo › Salvar como › PDF e importe de novo.',7000);const mb=file.size/1024/1024;if(mb>20)return flash(`O PDF tem ${mb.toFixed(1)} MB; o limite é 20 MB.`,6000);setBusy(`material:${step.id}:${type}`);const previous=materials.filter(m=>m.tarefa_id===task.id&&m.etapa_id===step.id&&m.tipo===type);for(const old of previous){if(old.arquivo_id)await supabase.storage.from('portal-files').remove([old.arquivo_id]);await supabase.from('finance_materials').delete().eq('id',old.id)}const path=`finance-pops/materials/${task.id}/${step.id}/${type}-${financeUid()}.pdf`;const {error:upErr}=await supabase.storage.from('portal-files').upload(path,file,{contentType:'application/pdf'});if(upErr){setBusy('');return dbError(upErr,'A importação falhou')};const {data:urlData}=supabase.storage.from('portal-files').getPublicUrl(path);const {error}=await supabase.from('finance_materials').insert({tarefa_id:task.id,codigo:task.codigo,tarefa:task.tarefa,etapa_id:step.id,etapa_num:(task.subtarefas||[]).findIndex(s=>s.id===step.id)+1,etapa_texto:step.texto,tipo:type,nome_arquivo:file.name,tamanho_mb:Number(mb.toFixed(2)),enviado_em:new Date().toISOString(),enviado_por:session.user.id,arquivo_id:path,url:urlData?.publicUrl||'',status:'pdf'});setBusy('');setMaterialPicker(null);if(error)return dbError(error,'A importação falhou');flash(`${type==='pop'?'POP':'Cartão de execução'} importado.`,6000)};
  const openMaterial=async m=>{const {data,error}=await supabase.storage.from('portal-files').createSignedUrl(m.arquivo_id,600);if(error)return dbError(error,'Não foi possível abrir');setViewer({...m,signedUrl:data.signedUrl})};

  const mapCells=()=>{const rows=['alto','medio','baixo'],cols=['baixo','medio','alto'];return rows.map(r=>cols.map(c=>({r,c,steps:tasks.flatMap(t=>(t.subtarefas||[]).map((s,i)=>({task:t,step:s,index:i}))).filter(x=>x.step.risco===r&&x.step.complexidade===c)})))};
  const responsibleGroups=useMemo(()=>{const map=new Map();tasks.forEach(t=>(t.subtarefas||[]).forEach((s,i)=>{if(s.delegar==='sim'){const key=(s.responsavel||'Delegadas sem responsável').trim()||'Delegadas sem responsável';const normalized=key.toLocaleLowerCase('pt-BR');if(!map.has(normalized))map.set(normalized,{name:key,items:[]});map.get(normalized).items.push({task:t,step:s,index:i})}}));return [...map.values()].sort((a,b)=>a.name.localeCompare(b.name,'pt-BR'));},[tasks]);
  const managerCount=tasks.reduce((n,t)=>n+(t.subtarefas||[]).filter(s=>s.delegar==='nao').length,0);const undecided=tasks.reduce((n,t)=>n+(t.subtarefas||[]).filter(s=>!s.delegar).length,0);

  return <section className="finance-task-page">
    <div className="topbar"><div><div className="eyebrow">PORTAL DA EQUIPE • FINANCEIRO</div><h1>Gerente de Administração Financeira</h1><div className="subtitle">Faturamento e fechamento dos convênios, conferência de guias e folhas de sala, repasse médico, caixa diário e solicitações de cirurgia.</div></div><div className="finance-task-indicators"><div><b>{loading?'–':tasks.length}</b><span>tarefas</span></div><div><b>{loading?'–':popVideoCount}</b><span>POPs com vídeo</span></div><div><b>{loading?'–':delegated}</b><span>etapas delegadas</span></div></div></div>
    {notice&&<div className={`finance-task-notice ${notice.includes('Não foi possível')||notice.includes('falhou')?'error':''}`}>{notice}</div>}
    <div className="finance-task-tabs">{[['tarefas','Tarefas'],['mapa','Mapa de delegação'],['responsavel','Por responsável'],['pops','POPs']].map(([id,label])=><button key={id} className={tab===id?'active':''} onClick={()=>setTab(id)}>{label}</button>)}</div>

    {tab==='tarefas'&&<div className="finance-task-tab-body">
      <div className="finance-task-toolbar"><div className="finance-filter-pills">{[['todas','Todas',tasks.length],['sem-video','Sem vídeo',tasks.filter(t=>taskVideos(t).length===0).length],['avaliacao','Avaliação a definir',tasks.filter(t=>(t.subtarefas||[]).some(s=>!financeStepEvaluated(s))).length],['delegadas','Com etapas delegadas',tasks.filter(t=>(t.subtarefas||[]).some(s=>s.delegar==='sim')).length]].map(([id,label,count])=><button key={id} className={filter===id?'active':''} onClick={()=>setFilter(id)}>{label} <b>{count}</b></button>)}</div>{canWrite&&<button className="btn btn-primary" onClick={newTask}>+ Nova tarefa</button>}</div>
      {!loading&&!tasks.length?<div className="card finance-empty"><h3>Nenhuma tarefa cadastrada ainda</h3><p>{canWrite?'Use “Nova tarefa” para incluir a primeira: recorrência, descrição e as etapas para a entrega acontecer.':'As tarefas do cargo aparecem aqui assim que forem cadastradas.'}</p></div>:!filteredTasks.length?<div className="card finance-empty">Nenhuma tarefa neste filtro.</div>:<div className="finance-task-list">{filteredTasks.map(t=>{const steps=t.subtarefas||[];const evalCount=steps.filter(financeStepEvaluated).length;const risks=steps.map(s=>financeLevelScore[s.risco]).filter(v=>v!==undefined);const maxRisk=risks.length?['baixo','medio','alto'][Math.max(...risks)]:'';const decisions=steps.filter(s=>s.delegar).length;const del=steps.filter(s=>s.delegar==='sim').length;const ps=popState(t);const open=openTasks.has(t.id);return <article className={`card finance-task-card ${open?'open':''}`} key={t.id} ref={el=>taskRefs.current[t.id]=el}>
        <button className="finance-task-summary" onClick={()=>toggleTask(t.id)}><span className="finance-code">{t.codigo}</span><span className="finance-task-main"><small>{t.recorrencia||'Sem recorrência'}</small><b>{t.tarefa||'(sem título)'}</b><em>{steps.length}{t.expected_steps&&t.expected_steps>steps.length?` de ${t.expected_steps}`:''} etapa(s){t.atualizado_em?` · atualizada em ${financeDateBR(t.atualizado_em)}`:''}</em></span><span className="finance-task-tags"><i className={steps.length&&evalCount===steps.length?'ok':'neutral'}>{evalCount} de {steps.length} etapas avaliadas</i>{maxRisk&&<i className={financeLevelClass(maxRisk)}>Maior risco: {financeLevelLabel(maxRisk)}</i>}<i className={del>0?'ok':decisions===steps.length&&steps.length?'danger':'neutral'}>{!decisions?'Delegação a definir':`${del} de ${steps.length} etapas delegadas`}</i><i className={ps.tone}>{ps.label}</i></span><span className="finance-chevron">⌄</span></button>
        {open&&<div className="finance-task-content"><div className="finance-task-content-head"><div><h3>Etapas · risco, complexidade, delegação e vídeo</h3><p>{steps.length?'Avalie cada etapa individualmente.':'Nenhuma etapa descrita ainda.'}</p></div>{canWrite&&<button className="btn btn-light" disabled={busy===`dup:${t.id}`} onClick={()=>duplicateTask(t)}>{busy===`dup:${t.id}`?'Duplicando…':'Duplicar'}</button>}</div>
          <div className="finance-step-list">{steps.map((s,i)=><div className="finance-step" id={`finance-step-${s.id}`} key={s.id}><div className="finance-step-number">{i+1}</div><div className="finance-step-body"><b>{s.texto}</b>{canWrite?<><div className="finance-step-controls"><select className={financeLevelClass(s.risco)} value={s.risco||''} onChange={e=>updateStep(t,i,{risco:e.target.value})}>{FINANCE_RISK_OPTIONS.map(([v,l])=><option key={v} value={v}>{l}</option>)}</select><select className={financeLevelClass(s.complexidade)} value={s.complexidade||''} onChange={e=>updateStep(t,i,{complexidade:e.target.value})}>{FINANCE_COMPLEXITY_OPTIONS.map(([v,l])=><option key={v} value={v}>{l}</option>)}</select><select className={s.delegar==='sim'?'ok':'neutral'} value={s.delegar||''} onChange={e=>{const v=e.target.value;updateStep(t,i,{delegar:v});if(v==='sim'&&!s.responsavel)setTimeout(()=>responsibleInputRefs.current[s.id]?.focus(),30)}}>{FINANCE_DELEGATION_OPTIONS.map(([v,l])=><option key={v} value={v}>{l}</option>)}</select>{s.delegar==='sim'&&<input ref={el=>responsibleInputRefs.current[s.id]=el} list="finance-responsibles" placeholder="Cargo ou nome da colaboradora" defaultValue={s.responsavel||''} onKeyDown={e=>{if(e.key==='Enter'){e.currentTarget.blur()}}} onBlur={e=>{if(e.target.value!==s.responsavel)updateStep(t,i,{responsavel:e.target.value.trim()})}}/>}</div></>:<p>Risco {financeLevelLabel(s.risco)} · complexidade {financeLevelLabel(s.complexidade)} · {s.delegar==='sim'?`delegada para ${s.responsavel||'responsável a definir'}`:s.delegar==='nao'?'mantém com o gestor':'delegação a definir'}</p>}
            <div className="finance-step-assets"><div><span>Vídeo</span>{canUpload?<label className="btn btn-light mini">{busy===`video:${s.id}`?'Enviando…':taskVideos(t).some(v=>v.etapa_id===s.id)?'Enviar outro vídeo':'Enviar vídeo desta etapa'}<input type="file" accept="video/mp4,video/webm,.mov,.m4v,.qt" hidden disabled={busy===`video:${s.id}`} onChange={e=>uploadVideo(t,s,e.target.files?.[0])}/></label>:<small>{taskVideos(t).some(v=>v.etapa_id===s.id)?'Vídeo disponível':'Sem vídeo'}</small>}<button className="btn btn-light mini" onClick={()=>openPop(t,true)}>Enviar pelo Google Drive</button></div><div><span>Materiais</span>{materials.filter(m=>m.tarefa_id===t.id&&m.etapa_id===s.id).map(m=><button className="material-pill" key={m.id} onClick={()=>openMaterial(m)}>{m.tipo==='pop'?'POP':'Cartão'}</button>)}{canUpload&&<button className="btn btn-light mini" onClick={()=>setMaterialPicker({task:t,step:s})}>Importar</button>}</div></div>
          </div></div>)}</div>
          <div className="finance-task-actions"><button className="btn btn-primary" onClick={()=>openPop(t)}>Abrir POP</button>{canWrite&&<button className="btn btn-light" onClick={()=>editTask(t)}>Editar tarefa e etapas</button>}{canWrite&&(inlineDelete===t.id?<span className="inline-confirm">Excluir esta tarefa? <button className="btn btn-danger-soft" onClick={()=>deleteTask(t)}>Excluir</button><button className="btn btn-light" onClick={()=>setInlineDelete('')}>Cancelar</button></span>:<button className="btn btn-danger-soft" onClick={()=>setInlineDelete(t.id)}>Excluir</button>)}</div>
        </div>}</article>})}</div>}
      <datalist id="finance-responsibles">{[...new Set(tasks.flatMap(t=>(t.subtarefas||[]).map(s=>s.responsavel).filter(Boolean)))].map(x=><option key={x} value={x}/>)}</datalist>
    </div>}

    {tab==='mapa'&&<div className="finance-task-tab-body"><div className="delegation-layout"><div className="delegation-map"><div className="delegation-axis blank"/><div className="delegation-axis">Complexidade baixa</div><div className="delegation-axis">Complexidade média</div><div className="delegation-axis">Complexidade alta</div>{mapCells().flatMap((row,ri)=>[<div key={`r-${ri}`} className="delegation-axis row">Risco {['alto','médio','baixo'][ri]}</div>,...row.map(cell=>{const score=financeLevelScore[cell.r]+financeLevelScore[cell.c];const zone=score<=1?['Delegar agora','ok']:score===2?['Delegar com treino','attention']:['Manter com o gestor','danger'];return <div className={`delegation-cell ${zone[1]}`} key={`${cell.r}-${cell.c}`}><b>{zone[0]}{cell.steps.length?` · ${cell.steps.length}`:''}</b><div>{cell.steps.map(x=><button key={`${x.task.id}-${x.step.id}`} className={x.step.delegar==='sim'?'delegated':''} title={`${x.task.codigo} · Etapa ${x.index+1} · ${x.step.texto}${x.step.delegar==='sim'&&x.step.responsavel?` → ${x.step.responsavel}`:''}`} onClick={()=>scrollTask(x.task.id)}>{x.task.codigo}·{x.index+1} {x.step.texto.length>48?`${x.step.texto.slice(0,48)}…`:x.step.texto}</button>)}</div></div>})])}</div><aside className="card delegation-side"><h3>Etapas ainda sem avaliação</h3><p>Cada etapa entra no mapa quando recebe risco e complexidade. Etapas já delegadas aparecem com faixa verde.</p>{tasks.filter(t=>(t.subtarefas||[]).some(s=>!s.risco||!s.complexidade)).map(t=><button key={t.id} onClick={()=>scrollTask(t.id)}>{t.codigo} {t.tarefa} · {(t.subtarefas||[]).filter(s=>!s.risco||!s.complexidade).length} etapa(s)</button>)}{!tasks.some(t=>(t.subtarefas||[]).some(s=>!s.risco||!s.complexidade))&&<div className="empty">Todas as etapas já foram avaliadas.</div>}<div className="delegation-legend"><span className="ok">Delegar agora, com POP e conferência no início.</span><span className="attention">Delegar com treino e revisão periódica.</span><span className="danger">Manter com o gestor por enquanto.</span></div></aside></div></div>}

    {tab==='responsavel'&&<div className="finance-task-tab-body"><div className="card responsible-note">Etapas agrupadas por quem faz. É a lista para entregar a cada colaboradora quando a delegação começar.</div><div className="responsible-groups">{responsibleGroups.map(g=><section className="card" key={g.name}><div className="responsible-head"><h3>{g.name}</h3><span>{g.items.length} etapa(s)</span></div>{g.items.map(x=><div className="responsible-item" key={`${x.task.id}-${x.step.id}`}><button onClick={()=>scrollTask(x.task.id)}>{x.task.codigo} · {x.index+1}</button><div><b>{x.step.texto}</b><small>{x.task.tarefa} · {x.task.recorrencia||'Sem recorrência'}</small></div></div>)}</section>)}<section className="card subtle"><div className="responsible-head"><h3>Ainda com o gestor</h3><span>{managerCount} etapas</span></div><p>{undecided?`${undecided} etapas ainda sem decisão de delegação. Defina na aba Tarefas.`:'Todas as etapas já têm decisão de delegação.'}</p></section></div></div>}

    {tab==='pops'&&<div className="finance-task-tab-body"><div className="card pops-note">Dados de pacientes que aparecem nos vídeos são borrados nas telas do POP e nunca entram no texto. Cada tarefa tem um POP. Os que ainda não têm vídeo são rascunhos montados a partir das etapas; o vídeo é o que transforma o rascunho em procedimento validado.</div><div className="pops-grid">{tasks.map(t=>{const p=taskPop(t);const ps=popState(t);return <article className="card pop-card" key={t.id}><div><span className="finance-code">{t.codigo}</span><i className={ps.tone}>{ps.label}</i></div><h3>{t.tarefa}</h3><p>{p?.objetivo||'Ainda sem POP. Envie um vídeo para gerar o procedimento.'}</p><div><button className="btn btn-primary" onClick={()=>openPop(t)}>Abrir POP</button>{canUpload&&<button className="btn btn-light" onClick={()=>openPop(t,true)}>Enviar vídeo</button>}</div></article>})}{!tasks.length&&<div className="card empty">Os POPs aparecem aqui para cada tarefa cadastrada.</div>}</div></div>}

    {editor&&<Modal open={true} onClose={()=>setEditor(null)} title={editor.mode==='new'?'Nova tarefa':'Editar tarefa'} subtitle="Gerente de Administração Financeira" onSave={saveTask} saveText={editor.mode==='new'?'Criar tarefa':'Salvar alterações'}><div className="form-grid"><Field label={editor.mode==='edit'&&String(editor.tarefa).startsWith('Cópia de ')?'Tarefa (cópia: troque o nome e ajuste as etapas)':'Tarefa'} full><input autoFocus value={editor.tarefa} onChange={e=>setEditor({...editor,tarefa:e.target.value})} placeholder="Ex.: Fechamento Unimed"/></Field><Field label="Recorrência" full><input list="finance-recurrence" value={editor.recorrencia} onChange={e=>setEditor({...editor,recorrencia:e.target.value})} placeholder="Ex.: 1x mês até dia 10"/><datalist id="finance-recurrence">{FINANCE_RECURRENCE.map(x=><option key={x} value={x}/>)}</datalist></Field><div className="full task-step-editor"><label>Etapas (na ordem em que são feitas). A delegação de cada etapa é definida depois de salvar.</label>{editor.subtarefas.map((s,i)=><div key={s.id} className="task-step-edit-row"><span>{i+1}</span><textarea rows="2" value={s.texto} onChange={e=>setEditor({...editor,subtarefas:editor.subtarefas.map((x,j)=>j===i?{...x,texto:e.target.value}:x)})}/><div><button type="button" disabled={i===0} onClick={()=>{const a=[...editor.subtarefas];[a[i-1],a[i]]=[a[i],a[i-1]];setEditor({...editor,subtarefas:a})}}>↑</button><button type="button" disabled={i===editor.subtarefas.length-1} onClick={()=>{const a=[...editor.subtarefas];[a[i+1],a[i]]=[a[i],a[i+1]];setEditor({...editor,subtarefas:a})}}>↓</button><button type="button" onClick={()=>{let a=editor.subtarefas.filter((_,j)=>j!==i);if(!a.length)a=[{id:financeUid(),texto:'',delegar:'',responsavel:'',risco:'',complexidade:''}];setEditor({...editor,subtarefas:a})}}>×</button></div></div>)}<button type="button" className="btn btn-light" onClick={()=>setEditor({...editor,subtarefas:[...editor.subtarefas,{id:financeUid(),texto:'',delegar:'',responsavel:'',risco:'',complexidade:''}]})}>+ Adicionar etapa</button></div></div></Modal>}

    {popTask&&<FinancePopModal task={popTask} pop={taskPop(popTask)} videos={taskVideos(popTask)} submissions={taskSubmissions(popTask)} materials={materials.filter(m=>m.tarefa_id===popTask.id)} canUpload={canUpload} canAdmin={canAdmin} onClose={()=>{setPopTask(null);setPopScrollUpload(false)}} scrollUpload={popScrollUpload} onUploadVideo={uploadVideo} onRegisterDrive={registerDrive} onOpenMaterial={openMaterial}/>} 
    {materialPicker&&<div className="modal" onMouseDown={e=>{if(e.target===e.currentTarget)setMaterialPicker(null)}}><div className="card modal-card small-material-modal"><div className="modal-head"><div><h2>Importar material</h2><p>{materialPicker.task.codigo} · Etapa {(materialPicker.task.subtarefas||[]).findIndex(s=>s.id===materialPicker.step.id)+1}</p></div><button className="icon-close" onClick={()=>setMaterialPicker(null)}>×</button></div><p>O que você vai importar (sempre em PDF)?</p><div className="material-type-buttons">{['pop','cartao'].map(type=><label className="btn btn-primary" key={type}>{busy===`material:${materialPicker.step.id}:${type}`?'Importando…':type==='pop'?'POP':'Cartão de execução'}<input type="file" hidden accept="application/pdf,.pdf" disabled={busy===`material:${materialPicker.step.id}:${type}`} onChange={e=>importMaterial(materialPicker.task,materialPicker.step,type,e.target.files?.[0])}/></label>)}</div></div></div>}
    {viewer&&<div className="modal material-viewer" onMouseDown={e=>{if(e.target===e.currentTarget)setViewer(null)}}><div className="card material-viewer-card"><div className="modal-head"><div><h2>{viewer.codigo} · Etapa {viewer.etapa_num} · {viewer.tipo==='pop'?'POP':'Cartão de execução'}</h2><p>{viewer.etapa_texto}</p></div><button className="icon-close" onClick={()=>setViewer(null)}>×</button></div><iframe title="Documento" src={viewer.signedUrl}/><div className="form-actions"><a className="btn btn-primary" href={viewer.signedUrl} download={viewer.nome_arquivo}>Imprimir / baixar PDF</a><button className="btn btn-light" onClick={()=>setViewer(null)}>Fechar</button></div></div></div>}
  </section>;
}

function FinancePopModal({task,pop,videos,submissions,materials,canUpload,canAdmin,onClose,scrollUpload,onUploadVideo,onRegisterDrive,onOpenMaterial}){
  const uploadRef=useRef(null);const [driveLink,setDriveLink]=useState('');const [obs,setObs]=useState('');const [stepId,setStepId]=useState('');
  useEffect(()=>{if(scrollUpload)setTimeout(()=>uploadRef.current?.scrollIntoView({behavior:'smooth',block:'center'}),100)},[scrollUpload]);
  const p=pop||((task.codigo==='FIN-01')?{status:'rascunho',versao:1,atualizado_em:task.atualizado_em,fonte:'Etapa 1 cadastrada',...financeStep1Pop}:null);const ps=(()=>{if(videos.some(v=>v.status!=='processado')||submissions.some(x=>x.status==='aguardando'))return{tone:'attention',label:'Vídeo recebido · atualizando'};if(p?.status==='validado')return{tone:'ok',label:`POP validado v${p.versao||1}`};if(p)return{tone:'neutral',label:`Rascunho v${p.versao||1}`};return{tone:'neutral',label:'Sem POP'}})();const chosenStep=(task.subtarefas||[]).find(s=>s.id===stepId)||null;
  return <div className="modal pop-modal" onMouseDown={e=>{if(e.target===e.currentTarget)onClose()}}><div className="card modal-card pop-modal-card"><div className="modal-head"><div><div className="eyebrow">POP · {task.codigo}</div><h2>{task.tarefa}</h2></div><button className="icon-close" onClick={onClose}>×</button></div><div className="pop-status-line"><i className={ps.tone}>{ps.label}</i>{p?<span>Versão {p.versao||1} · {financeDateBR(p.atualizado_em||task.atualizado_em)} · {p.fonte||'cadastro interno'}</span>:<span>Este POP ainda não foi escrito.</span>}</div>{(videos.some(v=>v.status!=='processado')||submissions.some(x=>x.status==='aguardando'))&&<div className="finance-task-notice">Há vídeo novo aguardando. O POP é reescrito na próxima rodada automática (de hora em hora, das 7h52 às 17h52 nos dias úteis) e esta tela se atualiza sozinha.</div>}
    {p?<><section className="pop-section"><h3>Identificação</h3><div className="pop-ident-grid"><div><span>Objetivo</span><p>{p.objetivo}</p></div><div><span>Quando fazer</span><p>{p.quando||task.recorrencia}</p></div><div><span>Responsável</span><p>{p.responsavel||'A definir'}</p></div><div><span>Sistemas e materiais</span><p>{(p.sistemas||[]).join(' · ')||'A definir'}</p></div></div></section><section className="pop-section"><h3>Passo a passo</h3><div className="pop-steps">{(p.passos||[]).map((x,i)=><div key={i}><span>{i+1}</span><div><b>{x.titulo}</b><p>{x.detalhe}</p>{x.alerta&&<strong>Atenção: {x.alerta}</strong>}<small>{x.responsavel}</small></div></div>)}</div></section><section className="pop-section"><h3>Pontos de atenção</h3><ul>{(p.atencao||[]).map((x,i)=><li key={i}>{x}</li>)}</ul><h3>Em caso de dúvida</h3><p>{p.duvidas||'—'}</p></section></>:<section className="pop-section"><h3>Etapas cadastradas</h3>{(task.subtarefas||[]).length?<ol>{task.subtarefas.map(s=><li key={s.id}>{s.texto}</li>)}</ol>:<p>Nenhuma etapa descrita.</p>}</section>}
    {(task.subtarefas||[]).some(s=>s.delegar)&&<section className="pop-section"><h3>Quem faz cada etapa</h3>{(task.subtarefas||[]).map((s,i)=><div className="pop-owner-row" key={s.id}><b>Etapa {i+1}</b><span>{s.delegar==='sim'?(s.responsavel||'Delegada, responsável a definir'):s.delegar==='nao'?'Gestor':'A definir'}</span></div>)}</section>}
    {canAdmin&&<section className="pop-section"><h3>Vídeos originais</h3><p>Visíveis só para quem administra o portal, porque podem mostrar dados de pacientes. A equipe vê o POP com as telas já borradas.</p>{videos.map(v=><div className="video-row" key={v.id}><div><b>{v.etapa_num?`Etapa ${v.etapa_num} · ${v.etapa_texto}`:'Tarefa inteira'}</b><small>{v.nome_arquivo} · {v.tamanho_mb||0} MB · {financeDateBR(v.enviado_em)}</small></div><i className={v.status==='processado'?'ok':'attention'}>{v.status==='processado'?'No POP':'Aguardando POP'}</i></div>)}</section>}
    <section className="pop-section"><h3>POPs e cartões importados por etapa</h3>{materials.length?<div className="material-list">{materials.map(m=><div key={m.id}><i className="ok">{m.tipo==='pop'?'POP':'Cartão de execução'}</i><span>{m.nome_arquivo} · {financeDateBR(m.enviado_em)}</span><button className="btn btn-light mini" onClick={()=>onOpenMaterial(m)}>{m.tipo==='pop'?'Ver POP':'Ver cartão'}</button></div>)}</div>:<p>Nenhum material importado.</p>}</section>
    {p?.historico?.length?<section className="pop-section"><h3>Histórico de versões</h3>{p.historico.map((h,i)=><div key={i}>v{h.versao} · {financeDateBR(h.data)} · {h.nota}</div>)}</section>:null}
    <section className="pop-section pop-upload" ref={uploadRef}><h3>Enviar vídeo</h3><p>Dica: na aba Tarefas, cada etapa tem o próprio botão de envio. Sempre que possível, grave com um paciente de teste. Dados de pacientes que aparecerem são borrados nas imagens do POP e não entram no texto.</p>{canUpload?<><Field label="Este vídeo mostra"><select value={stepId} onChange={e=>setStepId(e.target.value)}><option value="">Tarefa inteira</option>{(task.subtarefas||[]).map((s,i)=><option value={s.id} key={s.id}>Etapa {i+1} · {s.texto.slice(0,70)}</option>)}</select></Field><label className="video-drop">Escolher vídeo — MP4 ou WebM até 20 MB<input type="file" hidden accept="video/mp4,video/webm,.mov,.m4v,.qt" onChange={e=>onUploadVideo(task,chosenStep,e.target.files?.[0])}/></label><div className="or">ou</div><Field label="Link do vídeo no Google Drive"><input value={driveLink} onChange={e=>setDriveLink(e.target.value)} placeholder="https://drive.google.com/..."/></Field><Field label="Observações para o POP"><textarea rows="3" value={obs} onChange={e=>setObs(e.target.value)} placeholder="Opcional: o que muda em relação ao que está escrito, exceções, prazos, quem confere."/></Field><button className="btn btn-primary" onClick={()=>{if(!driveLink)return alert('Escolha um vídeo ou cole o link do Google Drive.');onRegisterDrive(task,chosenStep,driveLink,obs);setDriveLink('');setObs('')}}>Registrar vídeo do Drive</button></>:<p>O envio de vídeos fica disponível para quem tem permissão de edição neste portal.</p>}</section>
  </div></div>;
}


function App(){
  const [session,setSession] = useState(null);
  const [profile,setProfile] = useState(null);
  const [loading,setLoading] = useState(true);
  const [page,setPage] = useState('exames');

  const [doctors,setDoctors] = useState([]);
  const [examTypes,setExamTypes] = useState([]);
  const [surgeryTypes,setSurgeryTypes] = useState([]);
  const [procedures,setProcedures] = useState([]);
  const [insurances,setInsurances] = useState([]);
  const [profiles,setProfiles] = useState([]);
  const [records,setRecords] = useState([]);
  const [surgeries,setSurgeries] = useState([]);
  const [procedureRecords,setProcedureRecords] = useState([]);
  const [followups,setFollowups] = useState([]);

  const [examModal,setExamModal] = useState(false);
  const [surgeryModal,setSurgeryModal] = useState(false);
  const [procedureModal,setProcedureModal] = useState(false);
  const [catalogModal,setCatalogModal] = useState(false);
  const [editingExamId,setEditingExamId] = useState(null);
  const [editingSurgeryId,setEditingSurgeryId] = useState(null);
  const [editingProcedureId,setEditingProcedureId] = useState(null);
  const [followupModal,setFollowupModal] = useState(false);
  const [followupTarget,setFollowupTarget] = useState(null);
  const [documentPreview,setDocumentPreview] = useState(null);
  const [formErrors,setFormErrors] = useState({});
  const setValidationErrors=(scope,fields,message)=>{
    setFormErrors(prev=>({...prev,[scope]:{fields,message}}));
    setTimeout(()=>document.querySelector(`[data-validation-scope="${scope}"] .field-error, [data-validation-scope="${scope}"] .form-validation-banner`)?.scrollIntoView({behavior:'smooth',block:'center'}),80);
  };
  const clearValidationErrors=scope=>setFormErrors(prev=>({...prev,[scope]:null}));
  const fieldError=(scope,name)=>Boolean(formErrors?.[scope]?.fields?.includes(name));

  const [filterMonth,setFilterMonth] = useState(monthNow);
  const [filterDate,setFilterDate] = useState('');
  const [filterPeriodMode,setFilterPeriodMode] = useState('month');
  const [filterDoctor,setFilterDoctor] = useState('');
  const [filterExam,setFilterExam] = useState('');
  const [examStatus,setExamStatus] = useState('');
  const [examSearch,setExamSearch] = useState('');
  const [sMonth,setSMonth] = useState(monthNow);
  const [sDate,setSDate] = useState('');
  const [sPeriodMode,setSPeriodMode] = useState('month');
  const [sDoctor,setSDoctor] = useState('');
  const [sSurgeryType,setSSurgeryType] = useState('');
  const [sProcedure,setSProcedure] = useState('');
  const [sStatus,setSStatus] = useState('');
  const [sSearch,setSSearch] = useState('');
  const [pMonth,setPMonth] = useState(monthNow);
  const [pDate,setPDate] = useState('');
  const [pPeriodMode,setPPeriodMode] = useState('month');
  const [pDoctor,setPDoctor] = useState('');
  const [pProcedure,setPProcedure] = useState('');
  const [pStatus,setPStatus] = useState('');
  const [pSearch,setPSearch] = useState('');
  const [dashMonth,setDashMonth] = useState(monthNow);
  const [dashDate,setDashDate] = useState('');
  const [dashPeriodMode,setDashPeriodMode] = useState('month');
  const [dashDoctor,setDashDoctor] = useState('');
  const [dashUser,setDashUser] = useState('');
  const [dashStatus,setDashStatus] = useState('');
  const [dashPatient,setDashPatient] = useState('');
  const [finMonth,setFinMonth] = useState(monthNow);
  const [finDate,setFinDate] = useState('');
  const [finPeriodMode,setFinPeriodMode] = useState('month');
  const [finDoctor,setFinDoctor] = useState('');
  const [finInsurance,setFinInsurance] = useState('');
  const [finOrigin,setFinOrigin] = useState('');
  const [finSource,setFinSource] = useState('');
  const [finOperationalStatus,setFinOperationalStatus] = useState('');
  const [finStatus,setFinStatus] = useState('');
  const [finMethod,setFinMethod] = useState('');
  const [finSearch,setFinSearch] = useState('');
  const [followDay,setFollowDay] = useState('');
  const [followType,setFollowType] = useState('');
  const [followDoctor,setFollowDoctor] = useState('');
  const [followSourceStatus,setFollowSourceStatus] = useState('');
  const [followStatus,setFollowStatus] = useState('');
  const [followSearch,setFollowSearch] = useState('');
  const [dailyDate,setDailyDate] = useState(todayISO());
  const [dailySources,setDailySources] = useState([]);
  const [dailyStages,setDailyStages] = useState([]);
  const [dailySearch,setDailySearch] = useState('');
  const [dailyBirthDate,setDailyBirthDate] = useState('');
  const [invoiceDateFrom,setInvoiceDateFrom] = useState('');
  const [invoiceDateTo,setInvoiceDateTo] = useState('');
  const [invoiceSearch,setInvoiceSearch] = useState('');
  const [invoiceSource,setInvoiceSource] = useState('');
  const [invoiceDocumentStatus,setInvoiceDocumentStatus] = useState('');
  const [promMonth,setPromMonth] = useState(monthNow);
  const [promDate,setPromDate] = useState('');
  const [promStatus,setPromStatus] = useState('');
  const [promSearch,setPromSearch] = useState('');
  const unimedSaved=(()=>{try{return JSON.parse(localStorage.getItem('oc_unimed_step1')||'null')}catch{return null}})();
  const [unimedForm,setUnimedForm] = useState(unimedSaved?.form||{
    patient:'',surgery_date:'',procedure:'',doctor:'',assistant:'',execution_login:'714',start_time:'',end_time:'',seal:'',
    dossier_owner:'',reviewer:'',sent_at:'',dossier_status:'Aguardando documentos',step_status:'Não iniciada',observations:'',return_reason:''
  });
  const [unimedDocs,setUnimedDocs] = useState(unimedSaved?.docs||Object.fromEntries(UNIMED_STEP1_DOCUMENTS.map(name=>[name,'Pendente'])));
  const [unimedChecklist,setUnimedChecklist] = useState(unimedSaved?.checklist||Object.fromEntries(UNIMED_STEP1_CHECKLIST.map(name=>[name,false])));

  const [exportingPdf,setExportingPdf] = useState(false);
  const [exportingPanelPdf,setExportingPanelPdf] = useState(false);
  const resultsPageRef = useRef(null);
  const pdfPage1Ref = useRef(null);
  const pdfPage2Ref = useRef(null);
  const pdfPage3Ref = useRef(null);
  const dailyReportRef = useRef(null);
  useEffect(()=>{
    try{ localStorage.setItem('oc_unimed_step1',JSON.stringify({form:unimedForm,docs:unimedDocs,checklist:unimedChecklist})); }catch{}
  },[unimedForm,unimedDocs,unimedChecklist]);
  const unimedChecklistComplete=UNIMED_STEP1_CHECKLIST.every(item=>Boolean(unimedChecklist[item]));
  const unimedLoginOk=String(unimedForm.execution_login||'').trim()==='714';
  const unimedCanSend=unimedChecklistComplete&&unimedLoginOk;
  const unimedSetDossierStatus=status=>setUnimedForm(v=>({...v,dossier_status:status,step_status:status==='Concluído'?'Concluída':'Em andamento'}));
  const unimedSendToBilling=()=>{
    if(!unimedCanSend) return alert('O envio está bloqueado. Conclua todo o checklist e confirme o prestador/login 714.');
    const now=new Date();
    setUnimedForm(v=>({...v,dossier_status:'Enviado ao Faturamento',step_status:'Em andamento',sent_at:now.toISOString().slice(0,16)}));
  };
  const unimedRegisterReturn=destination=>setUnimedForm(v=>({...v,dossier_status:destination,step_status:'Em andamento'}));
  const resetUnimedStep1=()=>{
    if(!confirm('Limpar os dados preenchidos da Etapa 1 neste navegador?')) return;
    const form={patient:'',surgery_date:'',procedure:'',doctor:'',assistant:'',execution_login:'714',start_time:'',end_time:'',seal:'',dossier_owner:'',reviewer:'',sent_at:'',dossier_status:'Aguardando documentos',step_status:'Não iniciada',observations:'',return_reason:''};
    setUnimedForm(form);
    setUnimedDocs(Object.fromEntries(UNIMED_STEP1_DOCUMENTS.map(name=>[name,'Pendente'])));
    setUnimedChecklist(Object.fromEntries(UNIMED_STEP1_CHECKLIST.map(name=>[name,false])));
  };


  const emptyFinanceForm = () => ({
    payment_origin:'Particular',total_amount:'',discount_amount:'',received_amount:'',payment_method:'',installments:'1',payment_breakdown:[],
    financial_status:'',payment_confirmed:false,price_manual:false,coverage_choice:'',expected_receipt_date:'',received_date:'',billed_amount:'',billing_competence:'',billing_date:'',financial_notes:'',price_preset_key:'',
    invoice_number:'',invoice_issued_date:'',invoice_file_path:'',invoice_file:null,financial_documents:[],pending_documents:[],removed_document_paths:[],repasses:[]
  });
  const [examForm,setExamForm] = useState({
    date:todayISO(),realization_date:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',exam_type_ids:[],eye:'',insurance_id:'',status:'Agendado',launcher:'',obs:'',...emptyFinanceForm()
  });
  const [sForm,setSForm] = useState({
    date:todayISO(),scheduled_on:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',surgery_type_id:'',procedure_ids:[],procedure_finance:{},lio_description:'',eye:'',
    insurance_id:'',status:'Solicitação',arrival:'',time:'',payment:'Não informado',launcher:'',obs:'',...emptyFinanceForm()
  });
  const [pForm,setPForm] = useState({
    date:todayISO(),realization_date:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',procedure_ids:[],eye:'',
    insurance_id:'',status:'Solicitação',time:'',payment:'Não informado',launcher:'',obs:'',...emptyFinanceForm()
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

  const teamResponsibleOptions = [...new Set([...(profiles||[]).map(p=>p.full_name).filter(Boolean),'ADM'])];

  // Catálogos físicos separados no Supabase:
  // surgery_types = somente cirurgias | procedures = somente procedimentos.
  // O SQL V17 desativa o item no catálogo errado sem apagar o histórico.
  const procedureOptions = useMemo(()=>procedures,[procedures]);

  const insuranceNameById=id=>insurances.find(x=>x.id===id)?.name||'';
  const explicitOrigin=form=>form.coverage_choice==='__PARTICULAR__'?'Particular':(form.coverage_choice?(normalizeCatalogLabel(insuranceNameById(form.insurance_id)).includes('sus')?'SUS':'Convênio'):'');
  const examAutoPrice=useMemo(()=>{
    if(!examForm.coverage_choice||!examForm.eye) return null;
    return automaticPriceTotal({kind:'exam',names:(examForm.exam_type_ids||[]).map(id=>examTypes.find(x=>x.id===id)?.name),origin:explicitOrigin(examForm),insuranceName:insuranceNameById(examForm.insurance_id),eye:examForm.eye});
  },[examForm.exam_type_ids,examForm.coverage_choice,examForm.insurance_id,examForm.eye,examTypes,insurances]);
  const surgeryAutoPrice=useMemo(()=>{
    if(!sForm.coverage_choice||!sForm.eye) return null;
    return automaticPriceTotal({kind:'surgery',names:surgeryTypes.find(x=>x.id===sForm.surgery_type_id)?.name||'',origin:explicitOrigin(sForm),insuranceName:insuranceNameById(sForm.insurance_id),eye:sForm.eye,detail:sForm.lio_description});
  },[sForm.surgery_type_id,sForm.coverage_choice,sForm.insurance_id,sForm.eye,sForm.lio_description,surgeryTypes,insurances]);
  const procedureAutoPrice=useMemo(()=>{
    if(!pForm.coverage_choice||!pForm.eye) return null;
    return automaticPriceTotal({kind:'procedure',names:(pForm.procedure_ids||[]).map(id=>procedureOptions.find(x=>x.id===id)?.name),origin:explicitOrigin(pForm),insuranceName:insuranceNameById(pForm.insurance_id),eye:pForm.eye});
  },[pForm.procedure_ids,pForm.coverage_choice,pForm.insurance_id,pForm.eye,procedureOptions,insurances]);

  useEffect(()=>{
    if(examForm.price_manual) return;
    const next=examAutoPrice?String(examAutoPrice.amount):'';
    if(String(examForm.total_amount||'')!==next) setExamForm(f=>({...f,total_amount:next,price_preset_key:examAutoPrice?.keys?.join(',')||'',payment_origin:explicitOrigin(f)}));
  },[examAutoPrice,examForm.price_manual]);
  useEffect(()=>{
    if(sForm.price_manual) return;
    const extras=(sForm.procedure_ids||[]).reduce((sum,id)=>{const cfg=sForm.procedure_finance?.[id];return sum+(cfg?.included===false?numericValue(cfg.amount):0)},0);
    const next=surgeryAutoPrice?String(surgeryAutoPrice.amount+extras):'';
    if(String(sForm.total_amount||'')!==next) setSForm(f=>({...f,total_amount:next,price_preset_key:surgeryAutoPrice?.keys?.join(',')||'',payment_origin:explicitOrigin(f)}));
  },[surgeryAutoPrice,sForm.price_manual,sForm.procedure_ids,sForm.procedure_finance]);
  useEffect(()=>{
    if(pForm.price_manual) return;
    const next=procedureAutoPrice?String(procedureAutoPrice.amount):'';
    if(String(pForm.total_amount||'')!==next) setPForm(f=>({...f,total_amount:next,price_preset_key:procedureAutoPrice?.keys?.join(',')||'',payment_origin:explicitOrigin(f)}));
  },[procedureAutoPrice,pForm.price_manual]);

  const currentLaunchName = profile?.full_name?.trim() || session?.user?.email || (profile?.role==='admin'?'ADM':'Usuário');
  const launcherOptionsAll = ['Juliana','Elizângela','Luiza','ADM'];
  const defaultLaunchName = launcherOptionsAll.includes(currentLaunchName)?currentLaunchName:'ADM';
  const examLaunchName = String(examForm.launcher||defaultLaunchName||'').trim();
  const surgeryLaunchName = String(sForm.launcher||defaultLaunchName||'').trim();
  const procedureLaunchName = String(pForm.launcher||defaultLaunchName||'').trim();
  useEffect(()=>{if(examModal&&!editingExamId&&!examForm.launcher&&defaultLaunchName)setExamForm(f=>({...f,launcher:defaultLaunchName}));},[examModal,editingExamId,defaultLaunchName]);
  useEffect(()=>{if(surgeryModal&&!editingSurgeryId&&!sForm.launcher&&defaultLaunchName)setSForm(f=>({...f,launcher:defaultLaunchName}));},[surgeryModal,editingSurgeryId,defaultLaunchName]);
  useEffect(()=>{if(procedureModal&&!editingProcedureId&&!pForm.launcher&&defaultLaunchName)setPForm(f=>({...f,launcher:defaultLaunchName}));},[procedureModal,editingProcedureId,defaultLaunchName]);

  const loadAll = async ()=>{
    setLoading(true);
    const userId = session.user.id;
    const [
      profileRes, doctorsRes, examTypesRes, surgeryTypesRes, proceduresRes,
      insurancesRes, profilesRes, examsRes, surgeriesRes, procedureRecordsRes, followupsRes
    ] = await Promise.all([
      supabase.from('profiles').select('*').eq('id',userId).single(),
      supabase.from('doctors').select('*').eq('active',true).order('name'),
      supabase.from('exam_types').select('*').eq('active',true).order('name'),
      supabase.from('surgery_types').select('*').eq('active',true).order('name'),
      supabase.from('procedures').select('*').eq('active',true).order('name'),
      supabase.from('insurances').select('*').eq('active',true).order('name'),
      supabase.from('profiles').select('id,full_name,role,active').eq('active',true).order('full_name'),
      supabase.from('exams').select(`
        id,launch_group_id,exam_date,realization_date,eye,status,observation,launched_by,launched_by_name,insurance_id,payment_origin,total_amount,discount_amount,received_amount,payment_method,installments,financial_status,expected_receipt_date,received_date,billed_amount,billing_competence,billing_date,financial_notes,price_preset_key,payment_breakdown,invoice_number,invoice_issued_date,invoice_file_path,financial_documents,repasses,
        patients(id,full_name,whatsapp,birth_date),
        doctors:doctors!exams_doctor_id_fkey(id,name),
        indicated_by:doctors!exams_indicated_by_doctor_id_fkey(id,name),
        exam_types(id,name),
        insurances(id,name)
      `).order('exam_date',{ascending:false}),
      supabase.from('surgeries').select(`
        id,launch_group_id,surgery_date,scheduled_on,eye,lio_description,status,arrival_time,surgery_time,payment_status,observation,launched_by,launched_by_name,surgery_type_id,record_kind,payment_origin,total_amount,discount_amount,received_amount,payment_method,installments,financial_status,expected_receipt_date,received_date,billed_amount,billing_competence,billing_date,financial_notes,price_preset_key,payment_breakdown,invoice_number,invoice_issued_date,invoice_file_path,financial_documents,repasses,
        patients(id,full_name,whatsapp,birth_date),
        doctors:doctors!surgeries_doctor_id_fkey(id,name),
        indicated_by:doctors!surgeries_indicated_by_doctor_id_fkey(id,name),
        surgery_type:surgery_types!surgeries_surgery_type_id_fkey(id,name),
        procedures:procedures!surgeries_procedure_id_fkey(id,name),
        insurances(id,name)
      `).order('surgery_date',{ascending:false}),
      supabase.from('procedure_records').select(`
        id,launch_group_id,procedure_date,realization_date,status,procedure_time,payment_status,observation,launched_by,launched_by_name,surgery_id,surgery_group_id,legacy_surgery_id,record_kind,procedure_id,eye,payment_origin,total_amount,discount_amount,received_amount,payment_method,installments,financial_status,expected_receipt_date,received_date,billed_amount,billing_competence,billing_date,financial_notes,price_preset_key,payment_breakdown,invoice_number,invoice_issued_date,invoice_file_path,financial_documents,repasses,included_in_package,additional_amount,
        patients(id,full_name,whatsapp,birth_date),
        doctors:doctors!procedure_records_doctor_id_fkey(id,name),
        indicated_by:doctors!procedure_records_indicated_by_doctor_id_fkey(id,name),
        procedures:procedures!procedure_records_procedure_id_fkey(id,name),
        insurances:insurances!procedure_records_insurance_id_fkey(id,name)
      `).order('procedure_date',{ascending:false}),
      supabase.from('followups').select('*').order('updated_at',{ascending:false})
    ]);

    const errors=[profileRes,doctorsRes,examTypesRes,surgeryTypesRes,proceduresRes,insurancesRes,profilesRes,examsRes,surgeriesRes,procedureRecordsRes,followupsRes]
      .map(x=>x.error).filter(Boolean);
    if(errors.length) console.error(errors);

    setProfile(profileRes.data || null);
    setDoctors(doctorsRes.data || []);
    setExamTypes(examTypesRes.data || []);
    setSurgeryTypes(surgeryTypesRes.data || []);
    setProcedures(proceduresRes.data || []);
    setInsurances(insurancesRes.data || []);
    setProfiles(profilesRes.data || []);

    const groupedSurgeries=groupLaunchRows(surgeriesRes.data||[],'surgery');
    const surgeryById=new Map();
    const surgeryByGroup=new Map();
    const surgeryGroupCount=new Map();
    groupedSurgeries.forEach(item=>{
      (item.db_ids||[]).forEach(id=>surgeryById.set(id,item));
      surgeryGroupCount.set(item.group_id,(surgeryGroupCount.get(item.group_id)||0)+1);
      if(!surgeryByGroup.has(item.group_id)) surgeryByGroup.set(item.group_id,item);
    });

    // Só usa group_id como fallback quando existe UMA cirurgia no pacote.
    // Ex.: Pterígio + Catarata continuam duas cirurgias, nunca cirurgia + procedimento.
    const singleSurgeryByGroup=new Map(
      [...surgeryByGroup.entries()].filter(([groupId])=>surgeryGroupCount.get(groupId)===1)
    );
    // Regra oficial de classificação: o banco continua preservado, mas a interface
    // nunca trata um nome reconhecido como CIRURGIA como se fosse PROCEDIMENTO.
    // Itens desconhecidos permanecem visíveis para não esconder novos cadastros.
    const mappedProcedures=mapProcedureRows(procedureRecordsRes?.data||[],surgeryById,singleSurgeryByGroup)
      .filter(item=>knownCatalogKind(item.procedure)!=='surgery');

    const linkedBySurgeryId=new Map();
    const linkedByGroup=new Map();
    const linkedByPatientDate=new Map();
    const surgeryByPatientDate=new Map();
    const anchoredLaunchGroupsByPatientDate=new Map();

    // Índice auxiliar de cirurgias por paciente + data. Este fallback só é usado
    // quando existe UMA única cirurgia para aquela paciente naquele dia. Assim,
    // Catarata + Pterígio continuam sendo duas cirurgias independentes.
    groupedSurgeries.forEach(item=>{
      const key=`${item.patient_id||''}|${item.date||''}`;
      if(!surgeryByPatientDate.has(key)) surgeryByPatientDate.set(key,[]);
      surgeryByPatientDate.get(key).push(item);
    });

    // Primeiro cria os vínculos fortes. Eles servem como âncora do pacote.
    mappedProcedures.forEach(item=>{
      const patientDateKey=`${item.patient_id||''}|${item.date||''}`;

      if(item.surgery_id){
        if(!linkedBySurgeryId.has(item.surgery_id)) linkedBySurgeryId.set(item.surgery_id,[]);
        linkedBySurgeryId.get(item.surgery_id).push(item);
      }
      if(item.surgery_group_id){
        if(!linkedByGroup.has(item.surgery_group_id)) linkedByGroup.set(item.surgery_group_id,[]);
        linkedByGroup.get(item.surgery_group_id).push(item);
      }

      const resolvedById=Boolean(item.surgery_id && surgeryById.has(item.surgery_id));
      const resolvedByCurrentGroup=Boolean(item.surgery_group_id && singleSurgeryByGroup.has(item.surgery_group_id));
      if((resolvedById || resolvedByCurrentGroup) && item.group_id){
        if(!anchoredLaunchGroupsByPatientDate.has(patientDateKey)) anchoredLaunchGroupsByPatientDate.set(patientDateKey,new Set());
        anchoredLaunchGroupsByPatientDate.get(patientDateKey).add(item.group_id);
      }
    });

    // Depois recupera registros antigos do MESMO pacote. Um procedimento entra no
    // fallback quando já traz pista cirúrgica antiga OU compartilha launch_group_id
    // com outro procedimento que está corretamente ancorado à cirurgia.
    mappedProcedures.forEach(item=>{
      const key=`${item.patient_id||''}|${item.date||''}`;
      const anchoredLaunchGroups=anchoredLaunchGroupsByPatientDate.get(key)||new Set();
      const hasLegacyLink=Boolean(item.surgery_id || item.surgery_group_id || item.legacy_surgery_id);
      const sharesAnchoredLaunch=Boolean(item.group_id && anchoredLaunchGroups.has(item.group_id));

      if(hasLegacyLink || sharesAnchoredLaunch){
        if(!linkedByPatientDate.has(key)) linkedByPatientDate.set(key,[]);
        linkedByPatientDate.get(key).push(item);
      }
    });

    const enrichedSurgeries=groupedSurgeries.map(item=>{
      const surgeryId=item.db_ids?.[0]||'';
      const direct=linkedBySurgeryId.get(surgeryId)||[];

      // Vínculo pelo grupo atual da cirurgia. Só é usado quando esse grupo contém
      // uma única cirurgia, evitando misturar dois atos cirúrgicos distintos.
      const byGroup=surgeryGroupCount.get(item.group_id)===1
        ? (linkedByGroup.get(item.group_id)||[])
        : [];

      // Compatibilidade histórica: recupera Botox, Bioestimulador, Laser CO2 etc.
      // que pertencem ao pacote, mas ficaram com identificadores antigos/diferentes.
      // A associação por paciente + data só acontece quando existe uma única cirurgia.
      const patientDateKey=`${item.patient_id||''}|${item.date||''}`;
      const sameDaySurgeries=surgeryByPatientDate.get(patientDateKey)||[];
      const byPatientDate=sameDaySurgeries.length===1
        ? (linkedByPatientDate.get(patientDateKey)||[])
        : [];

      const linked=[...direct,...byGroup,...byPatientDate]
        .filter(row=>knownCatalogKind(row.procedure)!=='surgery')
        .filter((row,index,arr)=>arr.findIndex(x=>x.id===row.id)===index);

      if(!linked.length) return item;
      const procedureIds=[...new Set(linked.map(x=>x.procedure_id).filter(Boolean))];
      const procedureNames=[...new Set(linked.map(x=>x.procedure).filter(Boolean))];
      return {
        ...item,
        procedure_ids:procedureIds,
        procedure_names:procedureNames,
        procedure_finance:Object.fromEntries(linked.filter(x=>x.procedure_id).map(x=>[
          x.procedure_id,{included:x.included_in_package!==false,amount:numericValue(x.additional_amount)}
        ])),
        procedure:procedureNames.join(' • ')
      };
    });

    setRecords(groupLaunchRows(examsRes.data||[],'exam'));
    setSurgeries(enrichedSurgeries);
    setProcedureRecords(mappedProcedures);
    setFollowups(followupsRes?.data || []);
    setLoading(false);
  };

  const getOrCreatePatient = async (name,whatsapp,birthDate='')=>{
    const cleanName=name.trim();
    const cleanPhone=whatsapp.trim();

    if(cleanPhone){
      const {data:byPhone}=await supabase.from('patients').select('id').eq('whatsapp',cleanPhone).limit(1).maybeSingle();
      if(byPhone?.id){ if(birthDate) await supabase.from('patients').update({birth_date:birthDate}).eq('id',byPhone.id); return byPhone.id; }
    }
    const {data:byName}=await supabase.from('patients').select('id,whatsapp,birth_date').ilike('full_name',cleanName).limit(1).maybeSingle();
    if(byName?.id){
      const patch={};
      if(cleanPhone && normalizePhone(byName.whatsapp)!==normalizePhone(cleanPhone)) patch.whatsapp=cleanPhone;
      if(birthDate && byName.birth_date!==birthDate) patch.birth_date=birthDate;
      if(Object.keys(patch).length){ const {error:updatePatientError}=await supabase.from('patients').update(patch).eq('id',byName.id); if(updatePatientError) throw updatePatientError; }
      return byName.id;
    }

    const {data,error}=await supabase.from('patients').insert({
      full_name:cleanName,
      whatsapp:cleanPhone||null,
      birth_date:birthDate||null,
      created_by:session.user.id
    }).select('id').single();
    if(error) throw error;
    return data.id;
  };

  const uploadInvoiceFile = async (file,sourceType,groupId)=>{
    if(!file) return '';
    const safeName=String(file.name||'nota-fiscal').replace(/[^a-zA-Z0-9._-]+/g,'-');
    const path=`${sourceType}/${groupId}/${Date.now()}-${safeName}`;
    const {error}=await supabase.storage.from('financial-documents').upload(path,file,{upsert:true});
    if(error) throw error;
    return path;
  };

  // V28: múltiplos anexos permanecem apenas no fluxo da cirurgia.
  // Exames e procedimentos mantêm exatamente o comportamento da V25.
  const uploadFinancialDocument = async (file,sourceType,groupId)=>{
    if(!file) return null;
    const safeName=String(file.name||'documento').replace(/[^a-zA-Z0-9._-]+/g,'-');
    const unique=(crypto?.randomUUID?.()||`${Date.now()}-${Math.random().toString(36).slice(2)}`);
    const path=`${sourceType}/${groupId}/${unique}-${safeName}`;
    const {error}=await supabase.storage.from('financial-documents').upload(path,file,{upsert:false,contentType:file.type||undefined});
    if(error) throw error;
    return {
      id:unique,
      path,
      name:file.name||safeName,
      type:file.type||'',
      size:Number(file.size||0),
      uploaded_at:new Date().toISOString()
    };
  };

  const syncFinancialDocuments = async (form,sourceType,groupId)=>{
    const existing=Array.isArray(form.financial_documents)?form.financial_documents:[];
    const pending=Array.isArray(form.pending_documents)?form.pending_documents:[];
    const removed=[...new Set(form.removed_document_paths||[])];
    if(removed.length){
      const {error:removeError}=await supabase.storage.from('financial-documents').remove(removed);
      if(removeError) console.error('Falha ao remover documentos financeiros',removeError);
    }
    const uploaded=[];
    for(const file of pending){
      const doc=await uploadFinancialDocument(file,sourceType,groupId);
      if(doc) uploaded.push(doc);
    }
    return [...existing.filter(doc=>doc?.path&&!removed.includes(doc.path)),...uploaded];
  };

  const openInvoiceFile = async (path,name='Documento')=>{
    if(!path) return;
    const {data,error}=await supabase.storage.from('financial-documents').createSignedUrl(path,60*10);
    if(error||!data?.signedUrl) return alert('Não foi possível visualizar o documento.');
    const ext=String(name||path).split('.').pop()?.toLowerCase()||'';
    setDocumentPreview({path,name:name||String(path).split('/').pop()||'Documento',url:data.signedUrl,isImage:['png','jpg','jpeg','webp'].includes(ext)});
  };

  const downloadInvoiceFile = async (path,fileName='documento')=>{
    if(!path) return;
    const {data,error}=await supabase.storage.from('financial-documents').download(path);
    if(error||!data) return alert('Não foi possível baixar o documento.');
    const url=URL.createObjectURL(data);
    const a=document.createElement('a');
    a.href=url;
    a.download=fileName||'documento';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),1000);
  };

  const addExam = async ()=>{
    const missing=[];
    if(!examForm.date) missing.push('date');
    if(!examForm.realization_date) missing.push('realization_date');
    if(!examForm.patient.trim()) missing.push('patient');
    if(!examForm.indicated_by_doctor_id) missing.push('indicated_by_doctor_id');
    if(!examForm.doctor_id) missing.push('doctor_id');
    if(!examForm.exam_type_ids?.length) missing.push('exam_type_ids');
    if(!examForm.coverage_choice) missing.push('coverage_choice');
    if(!examForm.eye) missing.push('eye');
    if(!examForm.payment_confirmed) missing.push('payment_confirmed');
    if(!examLaunchName) missing.push('launcher');
    if(missing.length){setValidationErrors('exam',missing,'Alguns campos obrigatórios ainda precisam ser preenchidos.');return;}
    clearValidationErrors('exam');
    try{
      const patient_id=await getOrCreatePatient(examForm.patient,examForm.whatsapp,examForm.patient_birth_date);
      const launch_group_id=crypto.randomUUID();
      const rows=examForm.exam_type_ids.map(examTypeId=>({
        launch_group_id,
        exam_date:examForm.date, realization_date:examForm.realization_date||null, eye:examForm.eye||null, patient_id, doctor_id:examForm.doctor_id,
        indicated_by_doctor_id:examForm.indicated_by_doctor_id||null,
        exam_type_id:examTypeId, status:examForm.status,
        insurance_id:examForm.insurance_id||null,
        payment_origin:examForm.payment_origin||'Particular',
        total_amount:numericValue(examForm.total_amount),discount_amount:numericValue(examForm.discount_amount),received_amount:numericValue(examForm.received_amount),
        payment_method:examForm.payment_method||null,installments:Number(examForm.installments||1),financial_status:examForm.financial_status||'Pendente',
        expected_receipt_date:examForm.expected_receipt_date||null,received_date:examForm.received_date||null,billed_amount:numericValue(examForm.billed_amount),
        billing_competence:examForm.billing_competence||null,billing_date:examForm.billing_date||null,financial_notes:examForm.financial_notes||null,price_preset_key:examForm.price_preset_key||null,payment_breakdown:examForm.payment_breakdown||[],invoice_number:examForm.invoice_number||null,invoice_issued_date:examForm.invoice_issued_date||null,invoice_file_path:examForm.invoice_file_path||null,financial_documents:examForm.financial_documents||[],repasses:examForm.repasses||[],
        observation:examForm.obs||null, launched_by:session.user.id,
        launched_by_name:examLaunchName
      }));
      const {data:examSaved,error}=await supabase.from('exams').insert(rows).select('id,launch_group_id');
      if(error) throw error;
      if(examSaved?.length){ const docs=await syncFinancialDocuments(examForm,'exam',launch_group_id); await supabase.from('exams').update({financial_documents:docs,invoice_file_path:docs[0]?.path||null}).eq('launch_group_id',launch_group_id); }
      setExamModal(false);
      setExamForm({date:todayISO(),realization_date:'',eye:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',exam_type_ids:[],insurance_id:'',status:'Agendado',launcher:defaultLaunchName,obs:'',...emptyFinanceForm()});
      await loadAll();
    }catch(e){ console.error(e); alert('Não foi possível salvar os exames.'); }
  };

  const buildLinkedProcedureRows = ({procedureIds,patient_id,surgery_id,surgery_group_id,sourceForm,launcherName,launchGroupId}) => {
    // Todos os procedimentos escolhidos juntos compartilham o mesmo lançamento.
    // No banco continuam sendo linhas individuais para a contagem correta.
    const launch_group_id=launchGroupId||crypto.randomUUID();
    return (procedureIds||[]).map(procedureId=>({
      launch_group_id,
      procedure_date:sourceForm.date,
      patient_id,
      doctor_id:sourceForm.doctor_id,
      indicated_by_doctor_id:sourceForm.indicated_by_doctor_id||null,
      procedure_id:procedureId,
      surgery_id:surgery_id||null,
      surgery_group_id:surgery_group_id||null,
      eye:sourceForm.eye||'Não se aplica',
      insurance_id:sourceForm.insurance_id||null,
      status:sourceForm.status||'Solicitação',
      procedure_time:sourceForm.time||null,
      payment_status:sourceForm.payment||'Não informado',
      payment_origin:sourceForm.payment_origin||'Particular',
      total_amount:0,
      received_amount:0,
      payment_method:sourceForm.payment_method||null,
      installments:Number(sourceForm.installments||1),
      financial_status:sourceForm.financial_status||'Pendente',
      expected_receipt_date:sourceForm.expected_receipt_date||null,
      received_date:sourceForm.received_date||null,
      billed_amount:0,
      billing_competence:sourceForm.billing_competence||null,
      billing_date:sourceForm.billing_date||null,
      financial_notes:sourceForm.financial_notes||null,
      included_in_package:sourceForm.procedure_finance?.[procedureId]?.included!==false,
      additional_amount:sourceForm.procedure_finance?.[procedureId]?.included===false ? numericValue(sourceForm.procedure_finance?.[procedureId]?.amount) : 0,
      observation:sourceForm.obs||null,
      launched_by:session.user.id,
      launched_by_name:examLaunchName
    }));
  };

  const addSurgery = async ()=>{
    const missing=[];
    if(!sForm.date) missing.push('date');
    if(!sForm.scheduled_on) missing.push('scheduled_on');
    if(!sForm.patient.trim()) missing.push('patient');
    if(!sForm.indicated_by_doctor_id) missing.push('indicated_by_doctor_id');
    if(!sForm.doctor_id) missing.push('doctor_id');
    if(!sForm.surgery_type_id) missing.push('surgery_type_id');
    if(!sForm.coverage_choice) missing.push('coverage_choice');
    if(!sForm.eye) missing.push('eye');
    if(!sForm.payment_confirmed) missing.push('payment_confirmed');
    if(!surgeryLaunchName) missing.push('launcher');
    if(missing.length){setValidationErrors('surgery',missing,'Alguns campos obrigatórios ainda precisam ser preenchidos.');return;}
    clearValidationErrors('surgery');

    try{
      const patient_id=await getOrCreatePatient(sForm.patient,sForm.whatsapp,sForm.patient_birth_date);
      const launch_group_id=crypto.randomUUID();

      // V15: uma única linha em surgeries = uma cirurgia.
      const {data:surgeryRow,error}=await supabase.from('surgeries').insert({
        launch_group_id,
        surgery_date:sForm.date,
        scheduled_on:sForm.scheduled_on||null,
        patient_id,
        doctor_id:sForm.doctor_id,
        indicated_by_doctor_id:sForm.indicated_by_doctor_id||null,
        surgery_type_id:sForm.surgery_type_id,
        procedure_id:null,
        eye:sForm.eye,
        lio_description:sForm.lio_description||null,
        insurance_id:sForm.insurance_id||null,
        status:sForm.status,
        arrival_time:sForm.arrival||null,
        surgery_time:sForm.time||null,
        payment_status:sForm.payment,
        payment_origin:sForm.payment_origin||'Particular',
        total_amount:numericValue(sForm.total_amount),discount_amount:numericValue(sForm.discount_amount),
        received_amount:numericValue(sForm.received_amount),
        payment_method:sForm.payment_method||null,
        installments:Number(sForm.installments||1),
        financial_status:sForm.financial_status||'Pendente',
        expected_receipt_date:sForm.expected_receipt_date||null,
        received_date:sForm.received_date||null,
        billed_amount:numericValue(sForm.billed_amount),
        billing_competence:sForm.billing_competence||null,
        billing_date:sForm.billing_date||null,
        financial_notes:sForm.financial_notes||null,payment_breakdown:sForm.payment_breakdown||[],invoice_number:sForm.invoice_number||null,invoice_issued_date:sForm.invoice_issued_date||null,invoice_file_path:sForm.invoice_file_path||null,financial_documents:sForm.financial_documents||[],repasses:sForm.repasses||[],
        price_preset_key:sForm.price_preset_key||null,
        observation:sForm.obs||null,
        launched_by:session.user.id,
        launched_by_name:surgeryLaunchName
      }).select('id').single();
      if(error) throw error;
      { const docs=await syncFinancialDocuments(sForm,'surgery',launch_group_id); await supabase.from('surgeries').update({financial_documents:docs,invoice_file_path:docs[0]?.path||null}).eq('id',surgeryRow.id); }

      // Procedimentos do pacote são gravados no módulo Procedimentos e apenas vinculados à cirurgia.
      const linkedRows=buildLinkedProcedureRows({
        procedureIds:sForm.procedure_ids,
        patient_id,
        surgery_id:surgeryRow.id,
        surgery_group_id:launch_group_id,
        sourceForm:sForm,
        launcherName:surgeryLaunchName
      });
      if(linkedRows.length){
        const {error:procedureError}=await supabase.from('procedure_records').insert(linkedRows);
        if(procedureError){
          console.error(procedureError);
          alert('A cirurgia foi salva, mas não foi possível vincular os procedimentos. Verifique se o SQL V16 foi aplicado.');
        }
      }

      setSurgeryModal(false);
      setSForm({
        date:todayISO(),scheduled_on:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',surgery_type_id:'',procedure_ids:[],procedure_finance:{},lio_description:'',
        eye:'',insurance_id:'',status:'Solicitação',
        arrival:'',time:'',payment:'Não informado',launcher:defaultLaunchName,obs:'',...emptyFinanceForm()
      });
      await loadAll();
    }catch(e){
      console.error(e);
      alert('Não foi possível salvar a cirurgia.');
    }
  };


  const openEditExam = r => {
    setEditingExamId(r.id);
    setExamForm({
      date:r.date,
      realization_date:r.realization_date||'',
      eye:r.eye||'',
      patient:r.patient,
      patient_birth_date:r.patient_birth_date||'',
      whatsapp:r.whatsapp || '',
      doctor_id:r.doctor_id,
      indicated_by_doctor_id:r.indicatedById||'',
      exam_type_ids:r.exam_type_ids||[],
      launcher:r.launchedByName||defaultLaunchName,
      insurance_id:r.insurance_id||'',status:r.status,
            obs:r.obs || '',payment_origin:r.payment_origin||financeOrigin(r),total_amount:r.total_amount||'',discount_amount:r.discount_amount||'',received_amount:r.received_amount||'',
      payment_method:r.payment_method||'',installments:String(r.installments||1),financial_status:r.financial_status||'Pendente',payment_confirmed:true,price_manual:true,coverage_choice:r.insurance_id||'__PARTICULAR__',
      expected_receipt_date:r.expected_receipt_date||'',received_date:r.received_date||'',billed_amount:r.billed_amount||'',billing_competence:r.billing_competence||'',billing_date:r.billing_date||'',financial_notes:r.financial_notes||'',price_preset_key:r.price_preset_key||'',payment_breakdown:r.payment_breakdown||[],invoice_number:r.invoice_number||'',invoice_issued_date:r.invoice_issued_date||'',invoice_file_path:r.invoice_file_path||'',invoice_file:null,financial_documents:(r.financial_documents?.length?r.financial_documents:(r.invoice_file_path?[{id:r.invoice_file_path,path:r.invoice_file_path,name:'Documento anexado',type:'',size:0}]:[])),pending_documents:[],removed_document_paths:[],repasses:r.repasses||[]
    });
    clearValidationErrors('exam');setExamModal(true);
  };

  const openEditSurgery = r => {
    setEditingSurgeryId(r.id);
    setSForm({
      date:r.date,
      scheduled_on:r.scheduled_on||'',
      patient:r.patient,
      patient_birth_date:r.patient_birth_date||'',
      whatsapp:r.whatsapp || '',
      doctor_id:r.doctor_id,
      indicated_by_doctor_id:r.indicatedById||'',
      surgery_type_id:r.surgery_type_id||'',
      launcher:r.launchedByName||defaultLaunchName,
      procedure_ids:r.procedure_ids||[],
      eye:r.eye || 'Não se aplica',
      insurance_id:r.insurance_id || '',
      status:r.status,
      arrival:r.arrival || '',
      time:r.time || '',
      payment:r.payment || 'Não informado',
      payment_origin:r.payment_origin||financeOrigin(r),
      total_amount:r.total_amount||'',discount_amount:r.discount_amount||'',received_amount:r.received_amount||'',payment_method:r.payment_method||'',installments:String(r.installments||1),
      financial_status:r.financial_status||'Pendente',payment_confirmed:true,price_manual:true,coverage_choice:r.insurance_id||'__PARTICULAR__',expected_receipt_date:r.expected_receipt_date||'',received_date:r.received_date||'',
      billed_amount:r.billed_amount||'',billing_competence:r.billing_competence||'',billing_date:r.billing_date||'',financial_notes:r.financial_notes||'',price_preset_key:r.price_preset_key||'',payment_breakdown:r.payment_breakdown||[],invoice_number:r.invoice_number||'',invoice_issued_date:r.invoice_issued_date||'',invoice_file_path:r.invoice_file_path||'',invoice_file:null,financial_documents:(r.financial_documents?.length?r.financial_documents:(r.invoice_file_path?[{id:r.invoice_file_path,path:r.invoice_file_path,name:'Documento anexado',type:'',size:0}]:[])),pending_documents:[],removed_document_paths:[],repasses:r.repasses||[],
      lio_description:r.lio_description||'',procedure_finance:r.procedure_finance||{},
            obs:r.obs || ''
    });
    clearValidationErrors('surgery');setSurgeryModal(true);
  };

  const openEditProcedure = r => {
    setEditingProcedureId(r.group_key||r.id);
    setPForm({
      date:r.date,
      realization_date:r.realization_date||'',
      eye:r.eye||'',
      patient:r.patient,
      patient_birth_date:r.patient_birth_date||'',
      whatsapp:r.whatsapp||'',
      doctor_id:r.doctor_id,
      indicated_by_doctor_id:r.indicatedById||'',
      procedure_ids:r.procedure_ids?.length?r.procedure_ids:(r.procedure_id?[r.procedure_id]:[]),
      launcher:r.launchedByName||defaultLaunchName,
      eye:r.eye||'Não se aplica',
      insurance_id:r.insurance_id||'',
      status:r.status||'Solicitação',
      time:r.time||'',
      payment:r.payment||'Não informado',
      payment_origin:r.payment_origin||financeOrigin(r),
      total_amount:r.total_amount||'',discount_amount:r.discount_amount||'',received_amount:r.received_amount||'',payment_method:r.payment_method||'',installments:String(r.installments||1),
      financial_status:r.financial_status||'Pendente',payment_confirmed:true,price_manual:true,coverage_choice:r.insurance_id||'__PARTICULAR__',expected_receipt_date:r.expected_receipt_date||'',received_date:r.received_date||'',
      billed_amount:r.billed_amount||'',billing_competence:r.billing_competence||'',billing_date:r.billing_date||'',financial_notes:r.financial_notes||'',price_preset_key:r.price_preset_key||'',payment_breakdown:r.payment_breakdown||[],invoice_number:r.invoice_number||'',invoice_issued_date:r.invoice_issued_date||'',invoice_file_path:r.invoice_file_path||'',invoice_file:null,financial_documents:(r.financial_documents?.length?r.financial_documents:(r.invoice_file_path?[{id:r.invoice_file_path,path:r.invoice_file_path,name:'Documento anexado',type:'',size:0}]:[])),pending_documents:[],removed_document_paths:[],repasses:r.repasses||[],
            obs:r.obs||''
    });
    clearValidationErrors('procedure');setProcedureModal(true);
  };


  const updatePatient = async (patientId,name,whatsapp,birthDate='') => {
    if(!patientId) return;
    const {error} = await supabase.from('patients').update({
      full_name:name.trim(),
      whatsapp:whatsapp.trim() || null,
      birth_date:birthDate||null
    }).eq('id',patientId);
    if(error) throw error;
  };

  const updateExam = async ()=>{
    const original = records.find(r=>r.id===editingExamId);
    if(!original) return;
    if(!examForm.date||!examForm.realization_date||!examForm.patient.trim()||!examForm.indicated_by_doctor_id||!examForm.doctor_id||!examForm.exam_type_ids?.length||!examForm.coverage_choice||!examForm.eye||!examForm.payment_confirmed)
      return alert('Preencha os campos obrigatórios: data do agendamento, data da realização, paciente, médico que indicou, médico que realizou/fará, exame(s), convênio/Particular, olho e pagamento.');

    try{
      await updatePatient(original.patient_id, examForm.patient, examForm.whatsapp, examForm.patient_birth_date);

      const {error:deleteError}=await supabase.from('exams').delete().in('id',original.db_ids);
      if(deleteError) throw deleteError;

      const rows=examForm.exam_type_ids.map(examTypeId=>({
        launch_group_id:original.group_id,
        exam_date:examForm.date,
        realization_date:examForm.realization_date||null,
        eye:examForm.eye||null,
        patient_id:original.patient_id,
        doctor_id:examForm.doctor_id,
        indicated_by_doctor_id:examForm.indicated_by_doctor_id||null,
        exam_type_id:examTypeId,
        status:examForm.status,
        insurance_id:examForm.insurance_id||null,
        payment_origin:examForm.payment_origin||'Particular',total_amount:numericValue(examForm.total_amount),discount_amount:numericValue(examForm.discount_amount),received_amount:numericValue(examForm.received_amount),
        payment_method:examForm.payment_method||null,installments:Number(examForm.installments||1),financial_status:examForm.financial_status||'Pendente',
        expected_receipt_date:examForm.expected_receipt_date||null,received_date:examForm.received_date||null,billed_amount:numericValue(examForm.billed_amount),
        billing_competence:examForm.billing_competence||null,billing_date:examForm.billing_date||null,financial_notes:examForm.financial_notes||null,price_preset_key:examForm.price_preset_key||null,payment_breakdown:examForm.payment_breakdown||[],invoice_number:examForm.invoice_number||null,invoice_issued_date:examForm.invoice_issued_date||null,invoice_file_path:examForm.invoice_file_path||null,financial_documents:examForm.financial_documents||[],repasses:examForm.repasses||[],
        observation:examForm.obs||null,
        launched_by:session.user.id,
        launched_by_name:examLaunchName
      }));

      const {data:examSaved,error}=await supabase.from('exams').insert(rows).select('id,launch_group_id');
      if(error) throw error;
      if(examSaved?.length){ const docs=await syncFinancialDocuments(examForm,'exam',original.group_id); await supabase.from('exams').update({financial_documents:docs,invoice_file_path:docs[0]?.path||null}).eq('launch_group_id',original.group_id); }

      setExamModal(false);
      setEditingExamId(null);
      setExamForm({date:todayISO(),realization_date:'',eye:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',exam_type_ids:[],eye:'',insurance_id:'',status:'Agendado',obs:'',...emptyFinanceForm()});
      await loadAll();
    }catch(e){
      console.error(e);
      alert('Não foi possível atualizar o lançamento de exames.');
    }
  };

  const updateSurgery = async ()=>{
    const original = surgeries.find(r=>r.id===editingSurgeryId);
    if(!original) return;
    if(!sForm.date||!sForm.scheduled_on||!sForm.patient.trim()||!sForm.indicated_by_doctor_id||!sForm.doctor_id||!sForm.surgery_type_id||!sForm.coverage_choice||!sForm.eye||!sForm.payment_confirmed)
      return alert('Preencha os campos obrigatórios: data do agendamento, data da realização, paciente, médico que indicou, cirurgião, cirurgia, convênio/Particular, olho e pagamento.');

    try{
      await updatePatient(original.patient_id, sForm.patient, sForm.whatsapp, sForm.patient_birth_date);

      // Remove somente os procedimentos vinculados a ESTA cirurgia.
      const originalSurgeryId=original.db_ids?.[0]||original.id;
      const {error:deleteProcedureError}=await supabase.from('procedure_records').delete().eq('surgery_id',originalSurgeryId);
      if(deleteProcedureError) throw deleteProcedureError;

      const {error:deleteError}=await supabase.from('surgeries').delete().in('id',original.db_ids);
      if(deleteError) throw deleteError;

      // Recria como uma única cirurgia, mantendo o mesmo group_id para preservar o Follow-up.
      const {data:surgeryRow,error}=await supabase.from('surgeries').insert({
        launch_group_id:original.group_id,
        surgery_date:sForm.date,
        scheduled_on:sForm.scheduled_on||null,
        patient_id:original.patient_id,
        doctor_id:sForm.doctor_id,
        indicated_by_doctor_id:sForm.indicated_by_doctor_id||null,
        surgery_type_id:sForm.surgery_type_id,
        procedure_id:null,
        eye:sForm.eye,
        lio_description:sForm.lio_description||null,
        insurance_id:sForm.insurance_id||null,
        status:sForm.status,
        arrival_time:sForm.arrival||null,
        surgery_time:sForm.time||null,
        payment_status:sForm.payment,
        payment_origin:sForm.payment_origin||'Particular',
        total_amount:numericValue(sForm.total_amount),discount_amount:numericValue(sForm.discount_amount),
        received_amount:numericValue(sForm.received_amount),
        payment_method:sForm.payment_method||null,
        installments:Number(sForm.installments||1),
        financial_status:sForm.financial_status||'Pendente',
        expected_receipt_date:sForm.expected_receipt_date||null,
        received_date:sForm.received_date||null,
        billed_amount:numericValue(sForm.billed_amount),
        billing_competence:sForm.billing_competence||null,
        billing_date:sForm.billing_date||null,
        financial_notes:sForm.financial_notes||null,payment_breakdown:sForm.payment_breakdown||[],invoice_number:sForm.invoice_number||null,invoice_issued_date:sForm.invoice_issued_date||null,invoice_file_path:sForm.invoice_file_path||null,financial_documents:sForm.financial_documents||[],repasses:sForm.repasses||[],
        price_preset_key:sForm.price_preset_key||null,
        observation:sForm.obs||null,
        launched_by:session.user.id,
        launched_by_name:surgeryLaunchName
      }).select('id').single();
      if(error) throw error;
      { const docs=await syncFinancialDocuments(sForm,'surgery',original.group_id); await supabase.from('surgeries').update({financial_documents:docs,invoice_file_path:docs[0]?.path||null}).eq('id',surgeryRow.id); }

      const linkedRows=buildLinkedProcedureRows({
        procedureIds:sForm.procedure_ids,
        patient_id:original.patient_id,
        surgery_id:surgeryRow.id,
        surgery_group_id:original.group_id,
        sourceForm:sForm,
        launcherName:surgeryLaunchName
      });
      if(linkedRows.length){
        const {error:procedureError}=await supabase.from('procedure_records').insert(linkedRows);
        if(procedureError){
          console.error(procedureError);
          alert('A cirurgia foi atualizada, mas não foi possível recriar os procedimentos vinculados. Verifique o SQL V16.');
        }
      }

      setSurgeryModal(false);
      setEditingSurgeryId(null);
      setSForm({date:todayISO(),scheduled_on:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',surgery_type_id:'',procedure_ids:[],procedure_finance:{},lio_description:'',eye:'',insurance_id:'',status:'Solicitação',arrival:'',time:'',payment:'Não informado',obs:'',...emptyFinanceForm()});
      await loadAll();
    }catch(e){
      console.error(e);
      alert('Não foi possível atualizar o lançamento cirúrgico.');
    }
  };

  const addProcedure = async ()=>{
    const missing=[];
    if(!pForm.date) missing.push('date');
    if(!pForm.realization_date) missing.push('realization_date');
    if(!pForm.patient.trim()) missing.push('patient');
    if(!pForm.indicated_by_doctor_id) missing.push('indicated_by_doctor_id');
    if(!pForm.doctor_id) missing.push('doctor_id');
    if(!pForm.procedure_ids?.length) missing.push('procedure_ids');
    if(!pForm.coverage_choice) missing.push('coverage_choice');
    if(!pForm.eye) missing.push('eye');
    if(!pForm.payment_confirmed) missing.push('payment_confirmed');
    if(!procedureLaunchName) missing.push('launcher');
    if(missing.length){setValidationErrors('procedure',missing,'Alguns campos obrigatórios ainda precisam ser preenchidos.');return;}
    clearValidationErrors('procedure');
    try{
      const patient_id=await getOrCreatePatient(pForm.patient,pForm.whatsapp,pForm.patient_birth_date);
      const launch_group_id=crypto.randomUUID();
      const rows=pForm.procedure_ids.map(procedureId=>({
        launch_group_id,
        procedure_date:pForm.date,
        realization_date:pForm.realization_date||null,
        patient_id,
        doctor_id:pForm.doctor_id,
        indicated_by_doctor_id:pForm.indicated_by_doctor_id||null,
        procedure_id:procedureId,
        surgery_id:null,
        surgery_group_id:null,
        eye:pForm.eye,
        insurance_id:pForm.insurance_id||null,
        status:pForm.status,
        procedure_time:pForm.time||null,
        payment_status:pForm.payment,
        payment_origin:pForm.payment_origin||'Particular',
        total_amount:numericValue(pForm.total_amount),
        discount_amount:numericValue(pForm.discount_amount),
        received_amount:numericValue(pForm.received_amount),
        payment_method:pForm.payment_method||null,
        installments:Number(pForm.installments||1),
        financial_status:pForm.financial_status||'Pendente',
        expected_receipt_date:pForm.expected_receipt_date||null,
        received_date:pForm.received_date||null,
        billed_amount:numericValue(pForm.billed_amount),
        billing_competence:pForm.billing_competence||null,
        billing_date:pForm.billing_date||null,
        financial_notes:pForm.financial_notes||null,payment_breakdown:pForm.payment_breakdown||[],invoice_number:pForm.invoice_number||null,invoice_issued_date:pForm.invoice_issued_date||null,invoice_file_path:pForm.invoice_file_path||null,financial_documents:pForm.financial_documents||[],repasses:pForm.repasses||[],
        price_preset_key:pForm.price_preset_key||null,
        included_in_package:false,
        additional_amount:0,
        observation:pForm.obs||null,
        launched_by:session.user.id,
        launched_by_name:procedureLaunchName
      }));
      const {data:procedureSaved,error}=await supabase.from('procedure_records').insert(rows).select('id');
      if(error) throw error;
      if(procedureSaved?.length){ const docs=await syncFinancialDocuments(pForm,'procedure',launch_group_id); await supabase.from('procedure_records').update({financial_documents:docs,invoice_file_path:docs[0]?.path||null}).eq('launch_group_id',launch_group_id); }
      setProcedureModal(false);
      setPForm({date:todayISO(),realization_date:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',procedure_ids:[],eye:'',insurance_id:'',status:'Solicitação',time:'',payment:'Não informado',launcher:defaultLaunchName,obs:'',...emptyFinanceForm()});
      await loadAll();
    }catch(e){
      console.error(e);
      alert('Não foi possível salvar o procedimento.');
    }
  };

  const updateProcedure = async ()=>{
    const original=procedureLaunchRows.find(r=>(r.group_key||r.id)===editingProcedureId);
    if(!original) return;
    if(!pForm.date||!pForm.realization_date||!pForm.patient.trim()||!pForm.indicated_by_doctor_id||!pForm.doctor_id||!pForm.procedure_ids?.length||!pForm.coverage_choice||!pForm.eye||!pForm.payment_confirmed)
      return alert('Preencha os campos obrigatórios: data do agendamento, data da realização, paciente, médico que indicou, médico que realizou/fará, procedimento(s), convênio/Particular, olho e pagamento.');
    try{
      await updatePatient(original.patient_id,pForm.patient,pForm.whatsapp,pForm.patient_birth_date);

      const ids=original.db_ids?.length?original.db_ids:[original.id];
      const {error:deleteError}=await supabase.from('procedure_records').delete().in('id',ids);
      if(deleteError) throw deleteError;

      const launch_group_id=original.group_id||crypto.randomUUID();
      const originalProcedureItems=procedureRecords.filter(item=>ids.includes(item.id));
      const rows=pForm.procedure_ids.map(procedureId=>({
        launch_group_id,
        procedure_date:pForm.date,
        realization_date:pForm.realization_date||null,
        patient_id:original.patient_id,
        doctor_id:pForm.doctor_id,
        indicated_by_doctor_id:pForm.indicated_by_doctor_id||null,
        procedure_id:procedureId,
        surgery_id:original.surgery_id||null,
        surgery_group_id:original.surgery_group_id||null,
        eye:pForm.eye,
        insurance_id:pForm.insurance_id||null,
        status:pForm.status,
        procedure_time:pForm.time||null,
        payment_status:pForm.payment,
        payment_origin:pForm.payment_origin||'Particular',
        total_amount:numericValue(pForm.total_amount),
        discount_amount:numericValue(pForm.discount_amount),
        received_amount:numericValue(pForm.received_amount),
        payment_method:pForm.payment_method||null,
        installments:Number(pForm.installments||1),
        financial_status:pForm.financial_status||'Pendente',
        expected_receipt_date:pForm.expected_receipt_date||null,
        received_date:pForm.received_date||null,
        billed_amount:numericValue(pForm.billed_amount),
        billing_competence:pForm.billing_competence||null,
        billing_date:pForm.billing_date||null,
        financial_notes:pForm.financial_notes||null,payment_breakdown:pForm.payment_breakdown||[],invoice_number:pForm.invoice_number||null,invoice_issued_date:pForm.invoice_issued_date||null,invoice_file_path:pForm.invoice_file_path||null,financial_documents:pForm.financial_documents||[],repasses:pForm.repasses||[],
        price_preset_key:pForm.price_preset_key||null,
        included_in_package:(original.surgery_id||original.surgery_group_id)
          ? (originalProcedureItems.find(item=>item.procedure_id===procedureId)?.included_in_package!==false)
          : false,
        additional_amount:(original.surgery_id||original.surgery_group_id)
          ? numericValue(originalProcedureItems.find(item=>item.procedure_id===procedureId)?.additional_amount)
          : 0,
        observation:pForm.obs||null,
        launched_by:session.user.id,
        launched_by_name:procedureLaunchName
      }));

      const {error}=await supabase.from('procedure_records').insert(rows);
      if(error) throw error;
      { const docs=await syncFinancialDocuments(pForm,'procedure',launch_group_id); await supabase.from('procedure_records').update({financial_documents:docs,invoice_file_path:docs[0]?.path||null}).eq('launch_group_id',launch_group_id); }

      setProcedureModal(false);
      setEditingProcedureId(null);
      setPForm({date:todayISO(),realization_date:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',procedure_ids:[],eye:'',insurance_id:'',status:'Solicitação',time:'',payment:'Não informado',obs:'',...emptyFinanceForm()});
      await loadAll();
    }catch(e){
      console.error(e);
      alert('Não foi possível atualizar o lançamento de procedimentos.');
    }
  };


  const closeExamModal = ()=>{
    setExamModal(false);
    setEditingExamId(null);
    setExamForm({date:todayISO(),realization_date:'',eye:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',exam_type_ids:[],eye:'',insurance_id:'',status:'Agendado',obs:'',...emptyFinanceForm()});
  };

  const closeSurgeryModal = ()=>{
    setSurgeryModal(false);
    setEditingSurgeryId(null);
    setSForm({date:todayISO(),scheduled_on:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',surgery_type_id:'',procedure_ids:[],procedure_finance:{},lio_description:'',eye:'',insurance_id:'',status:'Solicitação',arrival:'',time:'',payment:'Não informado',obs:'',...emptyFinanceForm()});
  };

  const closeProcedureModal = ()=>{
    setProcedureModal(false);
    setEditingProcedureId(null);
    setPForm({date:todayISO(),realization_date:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',procedure_ids:[],eye:'',insurance_id:'',status:'Solicitação',time:'',payment:'Não informado',obs:'',...emptyFinanceForm()});
  };

  const delExam = async id=>{
    if(!confirm('Tem certeza que deseja excluir este lançamento de exames? Essa ação não poderá ser desfeita.')) return;
    const item=records.find(r=>r.id===id);
    const {error}=await supabase.from('exams').delete().in('id',item?.db_ids||[id]);
    if(error) return alert('Não foi possível excluir. Verifique as permissões do seu usuário no Supabase.');
    await loadAll();
  };
  const delSurgery = async id=>{
    if(!confirm('Tem certeza que deseja excluir este lançamento cirúrgico? Os procedimentos vinculados a esta cirurgia também serão excluídos.')) return;
    const item=surgeries.find(r=>r.id===id);
    const surgeryIds=item?.db_ids||[id];
    const {error:procedureError}=await supabase.from('procedure_records').delete().in('surgery_id',surgeryIds);
    if(procedureError) return alert('Não foi possível excluir os procedimentos vinculados.');
    const {error}=await supabase.from('surgeries').delete().in('id',surgeryIds);
    if(error) return alert('Não foi possível excluir. Verifique as permissões do seu usuário no Supabase.');
    await loadAll();
  };
  const delProcedure = async item=>{
    const qty=item?.db_ids?.length||1;
    if(!confirm(`Tem certeza que deseja excluir este lançamento com ${qty} procedimento(s)?`)) return;
    const ids=item?.db_ids?.length?item.db_ids:[item?.id];
    const {error}=await supabase.from('procedure_records').delete().in('id',ids);
    if(error) return alert('Não foi possível excluir o lançamento de procedimentos.');
    await loadAll();
  };

  const addCatalog = async (table,name,extra={})=>{
    const clean=name.trim();
    if(!clean) return false;

    // Protege os dois catálogos contra os erros já encontrados no histórico.
    const knownKind=knownCatalogKind(clean);
    if(table==='surgery_types' && knownKind==='procedure'){
      alert(`“${clean}” é um procedimento. Cadastre em Procedimentos, não em Cirurgias.`);
      return false;
    }
    if(table==='procedures' && knownKind==='surgery'){
      alert(`“${clean}” é uma cirurgia. Cadastre em Cirurgias, não em Procedimentos.`);
      return false;
    }

    // Se o nome já existia e foi apenas desativado durante o alinhamento,
    // reaproveita o mesmo cadastro em vez de gerar erro/duplicidade.
    const {data:existing,error:lookupError}=await supabase.from(table)
      .select('id,active,name')
      .ilike('name',clean)
      .limit(1)
      .maybeSingle();
    if(lookupError){
      console.error(lookupError);
      alert(`Não foi possível verificar o cadastro “${clean}”.`);
      return false;
    }
    if(existing?.id){
      if(existing.active){
        alert(`“${existing.name}” já está cadastrado nesta lista.`);
        return false;
      }
      const {error:reactivateError}=await supabase.from(table).update({active:true}).eq('id',existing.id);
      if(reactivateError){
        console.error(reactivateError);
        alert(`Não foi possível reativar “${clean}”.`);
        return false;
      }
      await loadAll();
      return true;
    }

    const {error}=await supabase.from(table).insert({name:clean,...extra});
    if(error){
      console.error(error);
      if(error.code==='23505') alert(`Já existe um cadastro com o nome “${clean}” nesta lista.`);
      else alert(`Não foi possível cadastrar “${clean}”. ${error.message||''}`.trim());
      return false;
    }
    await loadAll();
    return true;
  };
  const deactivateCatalog = async (table,id)=>{
    if(!confirm('Remover este item da lista? Os lançamentos antigos serão preservados.')) return;
    const {error}=await supabase.from(table).update({active:false}).eq('id',id);
    if(error) return alert(error.message);
    await loadAll();
  };

  const filteredExams = useMemo(()=>{
    const q=examSearch.trim().toLowerCase();
    return records.filter(r=>
      (!filterMonth||r.date.startsWith(filterMonth)) &&
      (!filterDate||r.date===filterDate||r.realization_date===filterDate) &&
      (!filterDoctor||r.doctor_id===filterDoctor) &&
      (!filterExam||r.exam_type_ids?.includes(filterExam)) &&
      (!examStatus||r.status===examStatus) &&
      patientMatchesSearch(r,q)
    );
  },[records,filterMonth,filterDate,filterDoctor,filterExam,examStatus,examSearch]);

  const filteredSurgeries = useMemo(()=>{
    const q=sSearch.trim().toLowerCase();
    return surgeries.filter(r=>
      (!sMonth||r.date.startsWith(sMonth)) &&
      (!sDate||r.date===sDate||r.scheduled_on===sDate) &&
      (!sDoctor||r.doctor_id===sDoctor) &&
      (!sSurgeryType||r.surgery_type_id===sSurgeryType) &&
      (!sProcedure||r.procedure_ids?.includes(sProcedure)) &&
      (!sStatus||r.status===sStatus) &&
      patientMatchesSearch(r,q)
    );
  },[surgeries,sMonth,sDate,sDoctor,sSurgeryType,sProcedure,sStatus,sSearch]);

  const procedureLaunchRows = useMemo(()=>groupProcedureRows(procedureRecords),[procedureRecords]);

  // Dados individuais: usados nos KPIs/Resultados para que 3 procedimentos = 3.
  const filteredProcedures = useMemo(()=>{
    const q=pSearch.trim().toLowerCase();
    return procedureRecords.filter(r=>
      (!pMonth||r.date?.startsWith(pMonth)) &&
      (!pDate||r.date===pDate||r.realization_date===pDate) &&
      (!pDoctor||r.doctor_id===pDoctor) &&
      (!pProcedure||r.procedure_id===pProcedure) &&
      (!pStatus||r.status===pStatus) &&
      patientMatchesSearch(r,q)
    );
  },[procedureRecords,pMonth,pDate,pDoctor,pProcedure,pStatus,pSearch]);

  // Visualização agrupada: mesma paciente/pacote/data aparece em UM lançamento.
  // O filtro de procedimento localiza o pacote, mas exibe todos os procedimentos dele.
  const filteredProcedureLaunches = useMemo(()=>{
    const q=pSearch.trim().toLowerCase();
    return procedureLaunchRows.filter(r=>
      (!pMonth||r.date?.startsWith(pMonth)) &&
      (!pDate||r.date===pDate||r.realization_date===pDate) &&
      (!pDoctor||r.doctor_id===pDoctor) &&
      (!pProcedure||r.procedure_ids?.includes(pProcedure)) &&
      (!pStatus||r.status===pStatus) &&
      patientMatchesSearch(r,q)
    );
  },[procedureLaunchRows,pMonth,pDate,pDoctor,pProcedure,pStatus,pSearch]);

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

    const message = `Você está prestes a excluir ${filteredSurgeries.length} cirurgia(s) exibida(s) pelos filtros atuais. Os procedimentos vinculados a esses pacotes também serão excluídos. Deseja continuar?`;
    if(!confirm(message)) return;

    const ids = [...new Set(filteredSurgeries.flatMap(item=>item.db_ids||[]))];
    if(ids.length){
      const {error:procedureError}=await supabase.from('procedure_records').delete().in('surgery_id',ids);
      if(procedureError){
        console.error(procedureError);
        return alert('Não foi possível excluir os procedimentos vinculados.');
      }
    }

    if(!ids.length) return alert('Não foi possível identificar as cirurgias para exclusão.');

    const {error}=await supabase.from('surgeries').delete().in('id',ids);
    if(error){
      console.error(error);
      return alert('Não foi possível excluir as cirurgias filtradas.');
    }

    await loadAll();
    alert('Cirurgias filtradas excluídas com sucesso.');
  };

  const deleteFilteredProcedures = async ()=>{
    if(!filteredProcedures.length) return alert('Nenhum procedimento encontrado com os filtros atuais.');
    const message=`Você está prestes a excluir ${filteredProcedures.length} procedimento(s) exibido(s). Deseja continuar?`;
    if(!confirm(message)) return;
    const ids=filteredProcedures.map(item=>item.id);
    const {error}=await supabase.from('procedure_records').delete().in('id',ids);
    if(error){
      console.error(error);
      return alert('Não foi possível excluir os procedimentos filtrados.');
    }
    await loadAll();
    alert('Procedimentos filtrados excluídos com sucesso.');
  };


  const financialRows = useMemo(()=>{
    const examRows=records.map(r=>({
      key:`exam:${r.group_id||r.id}`,source:'Exame',source_type:'exam',launch_date:r.date,realization_date:r.realization_date||'',date:isRealizedOperational(r.status)?(r.realization_date||r.date):r.date,patient:r.patient,patient_id:r.patient_id,
      doctor:r.doctor,doctor_id:r.doctor_id,insurance:r.insurance||'Particular',insurance_id:r.insurance_id||'',service:r.exam,
      operational_status:r.status,payment_origin:r.payment_origin||financeOrigin(r),payment_method:r.payment_method||'',financial_status:r.financial_status||'Pendente',
      total:netFinancialAmount(r),gross:numericValue(r.total_amount),discount:numericValue(r.discount_amount),received:numericValue(r.received_amount),billed:numericValue(r.billed_amount),expected_receipt_date:r.expected_receipt_date||'',received_date:r.received_date||'',
      installments:Number(r.installments||1),notes:r.financial_notes||'',payment_breakdown:r.payment_breakdown||[],invoice_number:r.invoice_number||'',invoice_issued_date:r.invoice_issued_date||'',invoice_file_path:r.invoice_file_path||'',financial_documents:r.financial_documents||[],indicatedBy:r.indicatedBy||'',base_record:r
    }));
    const surgeryRows=surgeries.map(r=>({
      key:`surgery:${r.group_id||r.id}`,source:'Cirurgia',source_type:'surgery',launch_date:r.date,realization_date:r.scheduled_on||'',date:isRealizedOperational(r.status)?(r.scheduled_on||r.date):r.date,patient:r.patient,patient_id:r.patient_id,
      doctor:r.doctor,doctor_id:r.doctor_id,insurance:r.insurance||'Particular',insurance_id:r.insurance_id||'',service:r.surgery_type,
      operational_status:r.status,payment_origin:r.payment_origin||financeOrigin(r),payment_method:r.payment_method||'',financial_status:r.financial_status||'Pendente',
      total:netFinancialAmount(r),gross:numericValue(r.total_amount),discount:numericValue(r.discount_amount),received:numericValue(r.received_amount),billed:numericValue(r.billed_amount),expected_receipt_date:r.expected_receipt_date||'',received_date:r.received_date||'',
      installments:Number(r.installments||1),notes:r.financial_notes||'',payment_breakdown:r.payment_breakdown||[],invoice_number:r.invoice_number||'',invoice_issued_date:r.invoice_issued_date||'',invoice_file_path:r.invoice_file_path||'',financial_documents:r.financial_documents||[],indicatedBy:r.indicatedBy||'',base_record:r
    }));

    // Procedimento avulso: um lançamento financeiro por launch_group_id, mesmo que tenha vários procedimentos.
    const standalone=groupProcedureRows(procedureRecords.filter(r=>!r.surgery_id&&!r.surgery_group_id)).map(r=>({
      key:`procedure:${r.group_id||r.id}`,source:'Procedimento',source_type:'procedure',launch_date:r.date,realization_date:r.realization_date||'',date:isRealizedOperational(r.status)?(r.realization_date||r.date):r.date,patient:r.patient,patient_id:r.patient_id,
      doctor:r.doctor,doctor_id:r.doctor_id,insurance:r.insurance||'Particular',insurance_id:r.insurance_id||'',service:r.procedure,
      operational_status:r.status,payment_origin:r.payment_origin||financeOrigin(r),payment_method:r.payment_method||'',financial_status:r.financial_status||'Pendente',
      total:netFinancialAmount(r),gross:numericValue(r.total_amount),discount:numericValue(r.discount_amount),received:numericValue(r.received_amount),billed:numericValue(r.billed_amount),expected_receipt_date:r.expected_receipt_date||'',received_date:r.received_date||'',
      installments:Number(r.installments||1),notes:r.financial_notes||'',payment_breakdown:r.payment_breakdown||[],invoice_number:r.invoice_number||'',invoice_issued_date:r.invoice_issued_date||'',invoice_file_path:r.invoice_file_path||'',financial_documents:r.financial_documents||[],indicatedBy:r.indicatedBy||'',base_record:r
    }));

    // Procedimentos vinculados só entram financeiramente quando marcados como cobrados à parte.
    const extras=procedureRecords.filter(r=>(r.surgery_id||r.surgery_group_id)&&r.included_in_package===false&&numericValue(r.additional_amount)>0).map(r=>({
      key:`extra:${r.id}`,source:'Adicional de pacote',source_type:'extra',launch_date:r.date,realization_date:r.realization_date||'',date:isRealizedOperational(r.status)?(r.realization_date||r.date):r.date,patient:r.patient,patient_id:r.patient_id,
      doctor:r.doctor,doctor_id:r.doctor_id,insurance:r.insurance||'Particular',insurance_id:r.insurance_id||'',service:r.procedure,
      operational_status:r.status,payment_origin:r.payment_origin||financeOrigin(r),payment_method:r.payment_method||'',financial_status:r.financial_status||'Pendente',
      total:numericValue(r.additional_amount),gross:numericValue(r.additional_amount),discount:0,received:0,billed:0,expected_receipt_date:r.expected_receipt_date||'',received_date:'',installments:Number(r.installments||1),notes:r.financial_notes||'',payment_breakdown:r.payment_breakdown||[],invoice_number:r.invoice_number||'',invoice_issued_date:r.invoice_issued_date||'',invoice_file_path:r.invoice_file_path||'',financial_documents:r.financial_documents||[],indicatedBy:r.indicatedBy||'',base_record:r
    }));
    return [...examRows,...surgeryRows,...standalone,...extras];
  },[records,surgeries,procedureRecords]);

  const filteredFinancialRows = useMemo(()=>{
    const q=finSearch.trim().toLowerCase();
    return financialRows.filter(r=>
      (!finMonth||r.date?.startsWith(finMonth)) &&
      (!finDate||r.date===finDate||r.launch_date===finDate||r.realization_date===finDate||r.received_date===finDate) &&
      (!finDoctor||r.doctor_id===finDoctor) &&
      (!finInsurance||r.insurance_id===finInsurance || normalizedBreakdown(r.payment_breakdown).some(x=>x.insurance_id===finInsurance) || (finInsurance==='particular'&&(!r.insurance_id || normalizedBreakdown(r.payment_breakdown).some(x=>x.payer_type==='Particular')))) &&
      (!finOrigin||r.payment_origin===finOrigin || (r.payment_origin==='Misto'&&['Particular','Convênio'].includes(finOrigin))) &&
      (!finSource||r.source_type===finSource) &&
      (!finOperationalStatus||r.operational_status===finOperationalStatus) &&
      (!finStatus||paymentTone(r)===finStatus) &&
      (!finMethod||r.payment_method===finMethod) &&
      (!q||patientMatchesSearch(r.base_record||r,q)||String(r.service||'').toLowerCase().includes(q))
    );
  },[financialRows,finMonth,finDate,finDoctor,finInsurance,finOrigin,finSource,finOperationalStatus,finStatus,finMethod,finSearch]);

  const financeSummary=useMemo(()=>{
    const valid=filteredFinancialRows.filter(r=>!isCanceledOperational(r.operational_status));
    const realized=valid.filter(r=>isRealizedOperational(r.operational_status));
    const production=realized.reduce((s,r)=>s+r.total,0);
    const received=valid.reduce((s,r)=>s+r.received,0);
    const receivable=realized.reduce((s,r)=>s+Math.max(r.total-r.received,0),0);
    const potential=valid.filter(r=>isPotentialOperational(r.operational_status)).reduce((s,r)=>s+r.total,0);
    const particular=valid.reduce((s,r)=>s+paymentAllocation(r,'Particular'),0);
    const insurance=valid.reduce((s,r)=>s+paymentAllocation(r,'Convênio'),0);
    const lost=filteredFinancialRows.filter(r=>isCanceledOperational(r.operational_status)).reduce((s,r)=>s+r.total,0);
    const billed=valid.reduce((s,r)=>s+(r.billed||0),0);
    return {production,received,receivable,potential,particular,insurance,lost,billed};
  },[filteredFinancialRows]);

  const financeAnalytics=useMemo(()=>{
    const base=financialRows.filter(r=>(!finDoctor||r.doctor_id===finDoctor)&&(!finOrigin||r.payment_origin===finOrigin));
    const baseMonth=finMonth||monthNow;
    const months=Array.from({length:6},(_,i)=>shiftMonth(baseMonth,i-5));
    const monthRows=m=>base.filter(r=>r.date?.startsWith(m)&&!isCanceledOperational(r.operational_status));
    const productionTrend=months.map(m=>monthRows(m).filter(r=>isRealizedOperational(r.operational_status)).reduce((s,r)=>s+r.total,0));
    const receivedTrend=months.map(m=>monthRows(m).reduce((s,r)=>s+r.received,0));
    const potentialTrend=months.map(m=>monthRows(m).filter(r=>isPotentialOperational(r.operational_status)).reduce((s,r)=>s+r.total,0));
    const sumBy=key=>filteredFinancialRows.filter(r=>!isCanceledOperational(r.operational_status)).reduce((acc,r)=>{const k=r[key]||'Não informado';acc[k]=(acc[k]||0)+r.total;return acc;},{});
    return {
      labels:months.map(shortMonthLabel),productionTrend,receivedTrend,potentialTrend,
      byInsurance:sumBy('insurance'),byDoctor:sumBy('doctor'),byService:sumBy('service'),byMethod:sumBy('payment_method'),byOrigin:sumBy('payment_origin')
    };
  },[financialRows,filteredFinancialRows,finMonth,finDoctor,finOrigin]);

  const promissoryRows=useMemo(()=>financialRows.flatMap(r=>{
    const breakdown=normalizedBreakdown(r.payment_breakdown);
    const rows=breakdown.filter(x=>normalizeCatalogLabel(x.method).includes('promissoria'));
    if(rows.length){
      return rows.map((x,index)=>{
        const amount=numericValue(x.amount);
        const received=x.status==='Recebido'?amount:0;
        const due=x.expected_date||r.expected_receipt_date||'';
        const paidDate=x.received_date||r.received_date||'';
        const overdue=Boolean(due&&x.status!=='Recebido'&&due<todayISO());
        return {id:`${r.key}:prom:${x.id||index}`,record:r,patient:r.patient,service:r.service,source:r.source,reference_date:r.date,amount,received,balance:Math.max(amount-received,0),installments:Number(x.installments||r.installments||1),due_date:due,received_date:paidDate,status:x.status==='Recebido'?'Pago':overdue?'Vencida':'Em aberto'};
      });
    }
    if(normalizeCatalogLabel(r.payment_method).includes('promissoria')){
      const amount=numericValue(r.total),received=numericValue(r.received),due=r.expected_receipt_date||'';
      const overdue=Boolean(due&&received<amount&&due<todayISO());
      return [{id:`${r.key}:prom:main`,record:r,patient:r.patient,service:r.service,source:r.source,reference_date:r.date,amount,received,balance:Math.max(amount-received,0),installments:Number(r.installments||1),due_date:due,received_date:r.received_date||'',status:received>=amount&&amount>0?'Pago':overdue?'Vencida':received>0?'Parcial':'Em aberto'}];
    }
    return [];
  }),[financialRows]);

  const filteredPromissoryRows=useMemo(()=>{
    const q=promSearch.trim().toLowerCase();
    return promissoryRows.filter(r=>
      (!promMonth||r.reference_date?.startsWith(promMonth)||r.due_date?.startsWith(promMonth)||r.received_date?.startsWith(promMonth)) &&
      (!promDate||r.reference_date===promDate||r.due_date===promDate||r.received_date===promDate) &&
      (!promStatus||r.status===promStatus) &&
      (!q||String(r.patient||'').toLowerCase().includes(q)||String(r.service||'').toLowerCase().includes(q))
    ).sort((a,b)=>String(a.due_date||a.reference_date||'').localeCompare(String(b.due_date||b.reference_date||'')));
  },[promissoryRows,promMonth,promDate,promStatus,promSearch]);

  const promissorySummary=useMemo(()=>({
    total:filteredPromissoryRows.reduce((s,r)=>s+r.amount,0),
    received:filteredPromissoryRows.reduce((s,r)=>s+r.received,0),
    balance:filteredPromissoryRows.reduce((s,r)=>s+r.balance,0),
    overdue:filteredPromissoryRows.filter(r=>r.status==='Vencida').reduce((s,r)=>s+r.balance,0)
  }),[filteredPromissoryRows]);

  const openPromissoryRecord=row=>{const r=row.record;if(r.source_type==='exam'){openEditExam(r.base_record);return;}if(r.source_type==='surgery'){openEditSurgery(r.base_record);return;}if(r.source_type==='procedure'){openEditProcedure(r.base_record);return;}};

  const exportPromissoryWord=()=>{
    const rows=filteredPromissoryRows.map(r=>`<tr><td>${brDate(r.reference_date)}</td><td>${String(r.patient||'')}</td><td>${String(r.service||'')}</td><td>${moneyBR(r.amount)}</td><td>${r.installments}x</td><td>${r.due_date?brDate(r.due_date):'—'}</td><td>${r.status}</td><td>${moneyBR(r.received)}</td><td>${moneyBR(r.balance)}</td></tr>`).join('');
    const html=`<html><head><meta charset="utf-8"><style>body{font-family:Arial,sans-serif;color:#24343d}h1{font-size:22px}table{border-collapse:collapse;width:100%;font-size:11px}th,td{border:1px solid #ccd6da;padding:6px;text-align:left}th{background:#edf4f1}</style></head><body><h1>OftalmoCastro — Controle de Promissórias</h1><p>Período: ${promDate?brDate(promDate):(promMonth?monthLabel(promMonth):'Todos')} • ${filteredPromissoryRows.length} registro(s)</p><p>Total: <b>${moneyBR(promissorySummary.total)}</b> • Recebido: <b>${moneyBR(promissorySummary.received)}</b> • Saldo: <b>${moneyBR(promissorySummary.balance)}</b> • Vencido: <b>${moneyBR(promissorySummary.overdue)}</b></p><table><thead><tr><th>Data</th><th>Paciente</th><th>Serviço</th><th>Valor</th><th>Parcelas</th><th>Vencimento</th><th>Situação</th><th>Recebido</th><th>Saldo</th></tr></thead><tbody>${rows||'<tr><td colspan="9">Sem registros.</td></tr>'}</tbody></table></body></html>`;
    const blob=new Blob(['\ufeff',html],{type:'application/msword'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=`OftalmoCastro_Promissorias_${promDate||promMonth||'todos'}.doc`;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
  };

  const exportPromissoryPdf=()=>{
    const pdf=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});const margin=10;let y=12;pdf.setFont('helvetica','bold');pdf.setFontSize(16);pdf.text('OftalmoCastro - Controle de Promissorias',margin,y);y+=7;pdf.setFont('helvetica','normal');pdf.setFontSize(9);pdf.text(`Periodo: ${promDate?brDate(promDate):(promMonth?monthLabel(promMonth):'Todos')} | Registros: ${filteredPromissoryRows.length}`,margin,y);y+=5;pdf.text(`Total: ${moneyBR(promissorySummary.total)} | Recebido: ${moneyBR(promissorySummary.received)} | Saldo: ${moneyBR(promissorySummary.balance)} | Vencido: ${moneyBR(promissorySummary.overdue)}`,margin,y);y+=7;
    const widths=[18,42,55,24,16,24,22,24,24],headers=['Data','Paciente','Servico','Valor','Parc.','Vencimento','Situacao','Recebido','Saldo'];const drawRow=(vals,bold=false)=>{let x=margin;pdf.setFont('helvetica',bold?'bold':'normal');vals.forEach((v,i)=>{const w=widths[i];pdf.rect(x,y-4,w,7);pdf.text(String(v??'').slice(0,i===2?38:i===1?27:18),x+1.5,y);x+=w;});y+=7;};drawRow(headers,true);filteredPromissoryRows.forEach(r=>{if(y>190){pdf.addPage();y=14;drawRow(headers,true);}drawRow([brDate(r.reference_date),r.patient,r.service,moneyBR(r.amount),`${r.installments}x`,r.due_date?brDate(r.due_date):'—',r.status,moneyBR(r.received),moneyBR(r.balance)]);});pdf.save(`OftalmoCastro_Promissorias_${promDate||promMonth||'todos'}.pdf`);
  };

  const toggleDailyFilter=(current,value,setter)=>setter(current.includes(value)?current.filter(x=>x!==value):[...current,value]);

  const dailyReportRows = useMemo(()=>{
    const q=dailySearch.trim().toLowerCase();
    const examRows=records.map(r=>({...r,report_type:'Exame',report_source:'exam',service_name:r.exam}));
    const surgeryRows=surgeries.map(r=>({...r,report_type:'Cirurgia',report_source:'surgery',service_name:r.surgery_type}));
    const procedureRows=procedureLaunchRows.map(r=>({...r,report_type:'Procedimento',report_source:'procedure',service_name:r.procedure}));
    return [...examRows,...surgeryRows,...procedureRows]
      .filter(r=>!dailySources.length||dailySources.includes(r.report_source))
      .filter(r=>!dailyBirthDate||r.patient_birth_date===dailyBirthDate)
      .filter(r=>!q||patientMatchesSearch(r,q)||String(r.service_name||'').toLowerCase().includes(q))
      .filter(r=>{
        const status=String(r.status||'').toLowerCase();
        const realized=status.includes('realiz');
        const canceled=status.includes('cancel');
        const realizationDate=r.report_source==='surgery'?(r.scheduled_on||r.date):(r.realization_date||r.date);
        const stage=realized?'realizado':canceled?'cancelado':'lancado';
        if(dailyStages.length&&!dailyStages.includes(stage)) return false;
        if(!dailyDate) return true;
        return stage==='realizado'?realizationDate===dailyDate:r.date===dailyDate;
      })
      .sort((a,b)=>String(a.patient||'').localeCompare(String(b.patient||''),'pt-BR'));
  },[records,surgeries,procedureLaunchRows,dailyDate,dailySources,dailyStages,dailySearch,dailyBirthDate]);

  const openDailyReportEdit = row=>{
    if(row.report_source==='exam'){
      openEditExam(row);
      return;
    }
    if(row.report_source==='surgery'){
      openEditSurgery(row);
      return;
    }
    if(row.report_source==='procedure'){
      openEditProcedure(row);
    }
  };

  const invoiceManagementRows = useMemo(()=>{
    const q=invoiceSearch.trim().toLowerCase();
    return financialRows
      .filter(r=>!['extra'].includes(r.source_type))
      .filter(r=>!invoiceSource||r.source_type===invoiceSource)
      .filter(r=>{
        const issueDate=r.invoice_issued_date||'';
        if(invoiceDateFrom && (!issueDate||issueDate<invoiceDateFrom)) return false;
        if(invoiceDateTo && (!issueDate||issueDate>invoiceDateTo)) return false;
        return true;
      })
      .filter(r=>!q||String(r.patient||'').toLowerCase().includes(q)||String(r.service||'').toLowerCase().includes(q)||String(r.invoice_number||'').toLowerCase().includes(q))
      .filter(r=>{
        const hasDoc=Boolean((r.financial_documents||[]).length||r.invoice_file_path||r.invoice_number);
        if(invoiceDocumentStatus==='com') return hasDoc;
        if(invoiceDocumentStatus==='sem') return !hasDoc;
        return true;
      })
      .sort((a,b)=>String(b.invoice_issued_date||b.date||'').localeCompare(String(a.invoice_issued_date||a.date||'')));
  },[financialRows,invoiceDateFrom,invoiceDateTo,invoiceSearch,invoiceSource,invoiceDocumentStatus]);

  const paymentSummaryText = record => {
    const rows=normalizedBreakdown(record?.payment_breakdown);
    if(rows.length){
      const parts=rows.map(row=>{
        const date=row.received_date||row.expected_date||'';
        const status=row.status==='Recebido'?'Pago':'Previsto';
        const method=row.method||row.payer_type||'';
        return `${status}${date?` em ${brDate(date)}`:''} ${moneyBR(row.amount)}${method?` via ${method}`:''}`;
      });
      const net=netFinancialAmount(record);
      const received=numericValue(record?.received_amount);
      const balance=Math.max(net-received,0);
      return `${parts.join(' + ')}${balance>0?` — RESTAM ${moneyBR(balance)}`:` — TOTAL ${moneyBR(net)}`}`;
    }
    const net=netFinancialAmount(record);
    const received=numericValue(record?.received_amount);
    if(received>0) return `${moneyBR(received)} recebido${Math.max(net-received,0)>0?` — RESTAM ${moneyBR(Math.max(net-received,0))}`:''}`;
    return net>0?`Valor: ${moneyBR(net)} — sem recebimento lançado`:'Sem informação financeira';
  };

  const repassesSummaryText = record => {
    const rows=Array.isArray(record?.repasses)?record.repasses:[];
    if(!rows.length) return '—';
    return rows.map(r=>`${r.type||'Repasse'} ${r.professional||''} ${moneyBR(r.amount)}`.trim()).join(' / ');
  };

  const weekdayLabel = date => {
    if(!date) return '';
    const d=new Date(`${date}T12:00:00`);
    return d.toLocaleDateString('pt-BR',{weekday:'long'}).replace(/^./,c=>c.toUpperCase());
  };

  const exportDailyReportPdf = async()=>{
    try{
      const pdf=new jsPDF({orientation:'portrait',unit:'mm',format:'a4',compress:true});
      const pageW=210,pageH=297,margin=14,contentW=pageW-margin*2;
      let y=18;
      const safe=value=>String(value??'—').replace(/•/g,'-').replace(/[“”]/g,'"').replace(/[‘’]/g,"'");
      const addHeader=()=>{
        pdf.setTextColor(49,95,114);pdf.setFont('helvetica','bold');pdf.setFontSize(15);
        pdf.text('OFTALMOCASTRO • RELATÓRIO DIÁRIO',margin,y);
        y+=7;pdf.setTextColor(90,107,116);pdf.setFont('helvetica','normal');pdf.setFontSize(9);
        pdf.text(`${brDate(dailyDate)} • ${weekdayLabel(dailyDate)} • ${dailyReportRows.length} registro(s)`,margin,y);
        y+=9;pdf.setDrawColor(221,230,234);pdf.line(margin,y,pageW-margin,y);y+=7;
      };
      const ensureSpace=h=>{if(y+h>pageH-16){pdf.addPage();y=18;addHeader();}};
      const addLine=(label,value,indent=0)=>{
        const labelW=39;
        pdf.setFontSize(8.5);pdf.setFont('helvetica','bold');pdf.setTextColor(49,95,114);
        pdf.text(safe(label),margin+indent,y);
        pdf.setFont('helvetica','normal');pdf.setTextColor(45,57,63);
        const lines=pdf.splitTextToSize(safe(value||'—'),contentW-labelW-indent);
        pdf.text(lines,margin+labelW+indent,y);
        y+=Math.max(5,lines.length*4.3);
      };
      addHeader();
      if(!dailyReportRows.length){pdf.setFontSize(10);pdf.setTextColor(110,125,132);pdf.text('Nenhum lançamento encontrado para os filtros selecionados.',margin,y);}
      dailyReportRows.forEach((r,index)=>{
        const payments=paymentSummaryText(r);
        const repasses=repassesSummaryText(r);
        const estimatedLines=Math.ceil((payments.length+repasses.length)/70);
        ensureSpace(50+estimatedLines*5);
        pdf.setFillColor(248,251,252);pdf.setDrawColor(218,228,232);
        pdf.roundedRect(margin,y-2,contentW,9,2,2,'FD');
        pdf.setFont('helvetica','bold');pdf.setFontSize(9.5);pdf.setTextColor(49,95,114);
        const heading=`${index+1}. ${String(r.report_type||'').toUpperCase()}: ${r.service_name||'—'}`;
        pdf.text(pdf.splitTextToSize(safe(heading),contentW-32),margin+3,y+3);
        pdf.setFontSize(7.5);pdf.setTextColor(90,107,116);pdf.text(safe(r.status||'—'),pageW-margin-3,y+3,{align:'right'});
        y+=12;
        addLine('Paciente:',`${r.patient||'—'}${r.patient_birth_date?` • Nasc. ${brDate(r.patient_birth_date)}`:''}`);
        addLine('Agendamento:',brDate(r.date));
        const realizationDate=r.report_source==='surgery'?r.scheduled_on:r.realization_date;
        if(realizationDate) addLine('Realização:',brDate(realizationDate));
        if(r.report_source==='surgery') addLine('LIO:',r.lio_description||'—');
        addLine('Convênio:',r.insurance||financeOrigin(r)||'Particular');
        addLine('Pagamentos:',payments);
        addLine('Indicação:',r.indicatedBy||'—');
        addLine(r.report_source==='surgery'?'Cirurgião:':'Realizado por:',r.doctor||'—');
        addLine('Repasses:',repasses);
        const docCount=(r.financial_documents||[]).length||(r.invoice_file_path?1:0);
        if(docCount) addLine('Documentos:',`${docCount} anexo(s)`);
        y+=5;pdf.setDrawColor(232,238,240);pdf.line(margin,y,pageW-margin,y);y+=7;
      });
      const pages=pdf.getNumberOfPages();
      for(let i=1;i<=pages;i++){pdf.setPage(i);pdf.setFontSize(7.5);pdf.setTextColor(135,149,155);pdf.text(`Página ${i} de ${pages}`,pageW-margin,pageH-8,{align:'right'});}
      pdf.save(`OftalmoCastro_Relatorio_Diario_${dailyDate||todayISO()}.pdf`);
    }catch(e){console.error(e);alert('Não foi possível gerar o PDF do relatório diário.');}
  };

  const escapeWordHtml = value => String(value??'')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

  const exportDailyReportWord = ()=>{
    try{
      const title=`Exames / Cirurgias / Procedimentos — ${brDate(dailyDate)} — ${weekdayLabel(dailyDate)}`;
      const rowsHtml=dailyReportRows.length ? dailyReportRows.map((r,index)=>{
        const realizationDate=r.report_source==='surgery'?r.scheduled_on:r.realization_date;
        const realization=realizationDate
          ? `<p><b>DATA DA REALIZAÇÃO:</b> ${escapeWordHtml(brDate(realizationDate))}</p>`:'';
        const lio=r.report_source==='surgery'
          ? `<p><b>LIO:</b> ${escapeWordHtml(r.lio_description||'—')}</p>`:'';
        const docs=(r.financial_documents?.length||r.invoice_file_path)
          ? `<p><b>DOCUMENTOS:</b> ${escapeWordHtml(r.financial_documents?.length||1)} anexo(s)</p>`:'';
        return `<div class="item">
          <h2>${index+1}. ${escapeWordHtml(String(r.report_type||'').toUpperCase())}: ${escapeWordHtml(r.service_name||'—')}</h2>
          <p><b>STATUS:</b> ${escapeWordHtml(r.status||'—')}</p>
          <p><b>DATA DO AGENDAMENTO:</b> ${escapeWordHtml(brDate(r.date))}</p>
          ${realization}${lio}
          <p><b>CONVÊNIO:</b> ${escapeWordHtml(r.insurance||financeOrigin(r)||'Particular')}</p>
          <p><b>NOME DO PACIENTE:</b> <strong>${escapeWordHtml(r.patient||'—')}</strong>${r.patient_birth_date?` — Nasc. ${escapeWordHtml(brDate(r.patient_birth_date))}`:''}</p>
          <p><b>VALOR / PAGAMENTOS:</b> ${escapeWordHtml(paymentSummaryText(r))}</p>
          <p><b>INDICAÇÃO:</b> ${escapeWordHtml(r.indicatedBy||'—')}</p>
          <p><b>${r.report_source==='surgery'?'CIRURGIÃO':'REALIZADO POR'}:</b> ${escapeWordHtml(r.doctor||'—')}</p>
          <p><b>REPASSES:</b> ${escapeWordHtml(repassesSummaryText(r))}</p>
          ${docs}
        </div>`;
      }).join('') : '<p>Nenhum lançamento encontrado para os filtros selecionados.</p>';

      const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><style>
        @page{size:A4;margin:1.6cm}body{font-family:Arial,sans-serif;color:#20313b;font-size:10.5pt;line-height:1.4}
        h1{font-size:16pt;color:#315f72;margin:0 0 4pt} .meta{color:#74838c;margin:0 0 18pt}
        .item{border-top:1px solid #dce5e8;padding:10pt 0 8pt;page-break-inside:avoid}.item h2{font-size:11.5pt;color:#315f72;margin:0 0 7pt}
        .item p{margin:2.5pt 0}.footer{margin-top:18pt;color:#74838c;font-size:8.5pt}
      </style></head><body>
        <h1>${escapeWordHtml(title)}</h1>
        <p class="meta">${dailyReportRows.length} registro(s) • Tipo: ${escapeWordHtml(dailySources.length?dailySources.map(x=>x==='exam'?'Exames':x==='surgery'?'Cirurgias':'Procedimentos').join(', '):'Todos')} • Situação: ${escapeWordHtml(dailyStages.length?dailyStages.map(x=>x==='realizado'?'Realizados':x==='cancelado'?'Cancelados':'Lançados / agendados').join(', '):'Todos')}</p>
        ${rowsHtml}
        <p class="footer">Gerado pelo sistema OftalmoCastro em ${escapeWordHtml(new Date().toLocaleString('pt-BR'))}.</p>
      </body></html>`;
      const blob=new Blob(['\ufeff',html],{type:'application/msword;charset=utf-8'});
      const url=URL.createObjectURL(blob);
      const a=document.createElement('a');
      a.href=url;
      a.download=`OftalmoCastro_Relatorio_Diario_${dailyDate||todayISO()}.doc`;
      document.body.appendChild(a);a.click();a.remove();
      setTimeout(()=>URL.revokeObjectURL(url),1000);
    }catch(e){ console.error(e); alert('Não foi possível gerar o arquivo Word do relatório diário.'); }
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
      patient_birth_date:r.patient_birth_date||'',
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
      patient_birth_date:r.patient_birth_date||'',
      source_date:r.date,
      doctor:r.doctor,
      doctor_id:r.doctor_id,
      item:`${r.surgery_type}${r.procedure&&r.procedure!=='Sem procedimentos vinculados'?` • ${r.procedure}`:''}`,
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
      (!q||patientMatchesSearch(row,q)||String(row.item||'').toLowerCase().includes(q)||String(row.doctor||'').toLowerCase().includes(q))
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

  const dashboardExams = useMemo(()=>{
    const q=dashPatient.trim().toLowerCase();
    return records.filter(r=>
      (!dashMonth||r.date.startsWith(dashMonth)) &&
      (!dashDate||r.date===dashDate||r.scheduled_on===dashDate) &&
      (!dashDate||r.date===dashDate||r.realization_date===dashDate) &&
      (!dashDoctor||r.doctor_id===dashDoctor) &&
      (!dashUser||r.launchedByName===dashUser) &&
      (!dashStatus||r.status===dashStatus) &&
      patientMatchesSearch(r,q)
    );
  },[records,dashMonth,dashDate,dashDoctor,dashUser,dashStatus,dashPatient]);

  const dashboardSurgeries = useMemo(()=>{
    const q=dashPatient.trim().toLowerCase();
    return surgeries.filter(r=>
      (!dashMonth||r.date.startsWith(dashMonth)) &&
      (!dashDoctor||r.doctor_id===dashDoctor) &&
      (!dashUser||r.launchedByName===dashUser) &&
      (!dashStatus||r.status===dashStatus) &&
      patientMatchesSearch(r,q)
    );
  },[surgeries,dashMonth,dashDate,dashDoctor,dashUser,dashStatus,dashPatient]);

  const dashboardProcedures = useMemo(()=>{
    const q=dashPatient.trim().toLowerCase();
    return procedureRecords.filter(r=>
      (!dashMonth||r.date?.startsWith(dashMonth)) &&
      (!dashDate||r.date===dashDate||r.realization_date===dashDate) &&
      (!dashDoctor||r.doctor_id===dashDoctor) &&
      (!dashUser||r.launchedByName===dashUser) &&
      (!dashStatus||r.status===dashStatus) &&
      patientMatchesSearch(r,q)
    );
  },[procedureRecords,dashMonth,dashDate,dashDoctor,dashUser,dashStatus,dashPatient]);

  const launchedName = record => record.launchedByName || (profiles.find(p=>p.id===record.launchedBy)?.role==='admin' ? 'ADM' : 'Orientadora');

  const userRows = ['Juliana','Luiza','Elisangela','ADM'].map(name=>({
    key:name,
    name,
    exams:totalItems(dashboardExams.filter(x=>launchedName(x)===name),'exam_type_ids'),
    surgeries:dashboardSurgeries.filter(x=>launchedName(x)===name).length,
    procedures:dashboardProcedures.filter(x=>launchedName(x)===name).length
  }));

  const currentAccessName = currentLaunchName || session?.user?.email || 'Usuário';

  const resultSummary=useMemo(()=>{
    const totalEx=totalItems(dashboardExams,'exam_type_ids');
    const totalSu=dashboardSurgeries.length; // 1 launch_group_id = 1 cirurgia
    const totalProcedures=dashboardProcedures.length;

    const exAgendado=totalItemsByStatus(dashboardExams,'exam_type_ids','Agendado');
    const exPendente=totalItemsByStatus(dashboardExams,'exam_type_ids','Pendente');
    const realEx=totalItemsByStatus(dashboardExams,'exam_type_ids','Realizado');
    const exCancelado=totalItemsByStatus(dashboardExams,'exam_type_ids','Cancelado');
    const exOpen=exAgendado+exPendente;

    const suSolicitacao=countRowsByStatus(dashboardSurgeries,'Solicitação');
    const suPendente=countRowsByStatus(dashboardSurgeries,'Pendente');
    const suAutorizada=countRowsByStatus(dashboardSurgeries,'Autorizada');
    const suAgendada=countRowsByStatus(dashboardSurgeries,'Agendada');
    const realSu=countRowsByStatus(dashboardSurgeries,'Realizada');
    const suCancelada=countRowsByStatus(dashboardSurgeries,'Cancelada');
    const suOpen=suSolicitacao+suPendente+suAutorizada+suAgendada;

    const prSolicitacao=countRowsByStatus(dashboardProcedures,'Solicitação');
    const prPendente=countRowsByStatus(dashboardProcedures,'Pendente');
    const prAutorizada=countRowsByStatus(dashboardProcedures,'Autorizada');
    const prAgendada=countRowsByStatus(dashboardProcedures,'Agendada');
    const realProcedures=countRowsByStatus(dashboardProcedures,'Realizada');
    const prCancelada=countRowsByStatus(dashboardProcedures,'Cancelada');
    const prOpen=prSolicitacao+prPendente+prAutorizada+prAgendada;

    const paymentCounts=countBy(dashboardSurgeries,'payment');
    const paymentPending=(paymentCounts['Pendente']||0)+(paymentCounts['20% pago']||0)+(paymentCounts['Não informado']||0);
    const paymentClosed=(paymentCounts['Pago']||0)+(paymentCounts['Convênio']||0);

    const examTop=Object.entries(countManyBy(dashboardExams,'exam_names')).sort((a,b)=>b[1]-a[1])[0];
    const surgTop=Object.entries(countBy(dashboardSurgeries,'surgery_type')).sort((a,b)=>b[1]-a[1])[0];
    const procedureTop=Object.entries(countBy(dashboardProcedures,'procedure')).sort((a,b)=>b[1]-a[1])[0];

    const professionalCounts={};
    dashboardExams.forEach(r=>{ professionalCounts[r.doctor||'Não informado']=(professionalCounts[r.doctor||'Não informado']||0)+(r.exam_type_ids?.length||0); });
    dashboardSurgeries.forEach(r=>{ professionalCounts[r.doctor||'Não informado']=(professionalCounts[r.doctor||'Não informado']||0)+1; });
    dashboardProcedures.forEach(r=>{ professionalCounts[r.doctor||'Não informado']=(professionalCounts[r.doctor||'Não informado']||0)+1; });
    const doctorTop=Object.entries(professionalCounts).sort((a,b)=>b[1]-a[1])[0];

    const insuranceTop=Object.entries(
      countBy(dashboardSurgeries.map(x=>({...x,insuranceLabel:x.insurance||'Particular'})),'insuranceLabel')
    ).sort((a,b)=>b[1]-a[1])[0];

    const uniquePatients=new Set(
      [...dashboardExams,...dashboardSurgeries,...dashboardProcedures]
        .map(x=>`${String(x.patient||'').trim().toLowerCase()}|${normalizePhone(x.whatsapp)}`)
        .filter(x=>x!=='|')
    ).size;

    return {
      totalEx,totalSu,totalProcedures,realEx,realSu,realProcedures,
      exAgendado,exPendente,exCancelado,exOpen,
      suSolicitacao,suPendente,suAutorizada,suAgendada,suCancelada,suOpen,
      prSolicitacao,prPendente,prAutorizada,prAgendada,prCancelada,prOpen,
      paymentCounts,paymentPending,paymentClosed,
      examTop,surgTop,procedureTop,doctorTop,insuranceTop,uniquePatients,
      exRate:percentOf(realEx,totalEx),
      suRate:percentOf(realSu,totalSu),
      procedureRate:percentOf(realProcedures,totalProcedures),
      exOpenRate:percentOf(exOpen,totalEx),
      suOpenRate:percentOf(suOpen,totalSu),
      procedureOpenRate:percentOf(prOpen,totalProcedures),
      paymentPendingRate:percentOf(paymentPending,totalSu)
    };
  },[dashboardExams,dashboardSurgeries,dashboardProcedures]);

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
    const surgeryTrend=months.map(month=>byMonth(baseSurgeryRows,month).length);
    const examConversion=months.map(month=>{
      const rows=byMonth(baseExamRows,month);
      const total=totalItems(rows,'exam_type_ids');
      const done=totalItemsByStatus(rows,'exam_type_ids','Realizado');
      return percentOf(done,total);
    });
    const surgeryConversion=months.map(month=>{
      const rows=byMonth(baseSurgeryRows,month);
      const total=rows.length;
      const done=countRowsByStatus(rows,'Realizada');
      return percentOf(done,total);
    });

    const currentExamRows=byMonth(baseExamRows,baseMonth);
    const previousExamRows=byMonth(baseExamRows,previousMonth);
    const currentSurgeryRows=byMonth(baseSurgeryRows,baseMonth);
    const previousSurgeryRows=byMonth(baseSurgeryRows,previousMonth);

    const currentExamTotal=totalItems(currentExamRows,'exam_type_ids');
    const previousExamTotal=totalItems(previousExamRows,'exam_type_ids');
    const currentSurgeryTotal=currentSurgeryRows.length;
    const previousSurgeryTotal=previousSurgeryRows.length;

    const currentExamDone=totalItemsByStatus(currentExamRows,'exam_type_ids','Realizado');
    const previousExamDone=totalItemsByStatus(previousExamRows,'exam_type_ids','Realizado');
    const currentSurgeryDone=countRowsByStatus(currentSurgeryRows,'Realizada');
    const previousSurgeryDone=countRowsByStatus(previousSurgeryRows,'Realizada');

    const examStatuses=['Realizado','Agendado','Pendente','Cancelado'];
    const surgeryStatuses=['Realizada','Agendada','Autorizada','Pendente','Solicitação','Cancelada'];

    const examByStatusType=Object.fromEntries(examStatuses.map(status=>[
      status,
      countManyBy(dashboardExams.filter(r=>r.status===status),'exam_names')
    ]));

    const surgeryByStatusType=Object.fromEntries(surgeryStatuses.map(status=>[
      status,
      countBy(dashboardSurgeries.filter(r=>r.status===status),'surgery_type')
    ]));

    const examStatusTotals=Object.fromEntries(examStatuses.map(status=>[
      status,totalItemsByStatus(dashboardExams,'exam_type_ids',status)
    ]));

    const surgeryStatusTotals=Object.fromEntries(surgeryStatuses.map(status=>[
      status,countRowsByStatus(dashboardSurgeries,status)
    ]));
    const procedureStatusTotals=Object.fromEntries(surgeryStatuses.map(status=>[
      status,dashboardProcedures.filter(r=>r.status===status).length
    ]));
    const procedureByStatusType=Object.fromEntries(surgeryStatuses.map(status=>[
      status,countBy(dashboardProcedures.filter(r=>r.status===status),'procedure')
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
    const surgeryByDoctor=buildDoctorRows(dashboardSurgeries,'surgery_type_names','doctor');
    const procedureByDoctor=buildDoctorRows(dashboardProcedures.map(r=>({...r,procedure_names:[r.procedure]})),'procedure_names','doctor');

    const examByIndication=buildDoctorRows(
      dashboardExams.filter(r=>r.indicatedBy),
      'exam_names',
      'indicatedBy'
    );

    const surgeryByIndication=buildDoctorRows(
      dashboardSurgeries.filter(r=>r.indicatedBy),
      'surgery_type_names',
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
      row.surgeryTotal+=1;
      if(record.status==='Realizada') row.surgeryDone+=1;
      else if(record.status==='Agendada') row.surgeryScheduled+=1;
      else if(record.status==='Autorizada') row.surgeryAuthorized+=1;
      else if(record.status==='Pendente') row.surgeryPending+=1;
      else if(record.status==='Solicitação') row.surgeryRequested+=1;
      else if(record.status==='Cancelada') row.surgeryCanceled+=1;
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
      examByStatusType,surgeryByStatusType,procedureByStatusType,examStatusTotals,surgeryStatusTotals,procedureStatusTotals,
      examStatusPie,surgeryStatusPie,
      examByDoctor,surgeryByDoctor,procedureByDoctor,examByIndication,surgeryByIndication,doctorConversion,
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
  },[records,surgeries,procedureRecords,dashboardExams,dashboardSurgeries,dashboardProcedures,dashMonth,dashDoctor,dashUser]);

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
      ['Cirurgias por tipo',countBy(dashboardSurgeries,'surgery_type'),'A35:F48'],
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
      `e a cirurgia com maior volume foi ${resultSummary.surgTop?.[0]||'—'}.`;
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

    wsSurgeries.mergeCells('A1:N1');
    styleTitle(wsSurgeries.getCell('A1'),18);
    wsSurgeries.getCell('A1').value=`CIRURGIAS • ${periodLabel}`;
    wsSurgeries.mergeCells('A2:N2');
    wsSurgeries.getCell('A2').value=`Filtros: Médico ${selectedDoctor} | Responsável ${selectedUser} | Status ${selectedStatus}`;
    wsSurgeries.getCell('A2').font={size:10,color:{argb:COLORS.muted}};

    const surgeryHeaders=[
      'Data','Paciente','WhatsApp','Cirurgião','Cirurgia','Procedimentos vinculados','Olho','Convênio',
      'Status','Horário chegada','Horário cirurgia','Pagamento','Lançado por','Observação'
    ];
    const surgeryRows=dashboardSurgeries.map(r=>[
      brDate(r.date),r.patient,r.whatsapp,r.doctor,r.surgery_type,r.procedure,r.eye,r.insurance||'Particular',
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

    [13,28,18,24,25,32,14,22,18,18,18,18,18,45].forEach((w,i)=>wsSurgeries.getColumn(i+1).width=w);
    wsSurgeries.getColumn(14).alignment={wrapText:true,vertical:'top'};
    styleStatusCells(wsSurgeries,9,4,3+Math.max(1,surgeryRows.length));
    wsSurgeries.autoFilter={from:'A3',to:'N3'};
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
      pageHeader('ANÁLISE POR MÉDICO','Exames por médico e tipo + cirurgias por cirurgião e tipo de cirurgia.');
      y=49;
      drawDoctorColumn(margin,y,half,231,'Exames por médico e tipo',advancedAnalytics.examByDoctor,'Exame');
      drawDoctorColumn(margin+half+4,y,half,231,'Cirurgias por cirurgião e tipo',advancedAnalytics.surgeryByDoctor,'Cirurgia');

      // =====================================================
      // PÁGINA 3 — DETALHAMENTO POR STATUS
      // =====================================================
      pdf.addPage('a4','portrait');
      pageHeader('DETALHAMENTO POR STATUS','Tipos de exame e tipos de cirurgia dentro de cada status do período.');
      y=49;
      drawStatusColumn(margin,y,half,231,'Exames por status e tipo',
        ['Realizado','Pendente','Agendado','Cancelado'],advancedAnalytics.examByStatusType,advancedAnalytics.examStatusTotals,'Exame');
      drawStatusColumn(margin+half+4,y,half,231,'Cirurgias por status e tipo',
        ['Pendente','Solicitação','Autorizada','Agendada','Realizada','Cancelada'],advancedAnalytics.surgeryByStatusType,advancedAnalytics.surgeryStatusTotals,'Cirurgia');

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
      drawRankingBox(margin,y,boxW,boxH,'Cirurgias por tipo',topEntries(countBy(dashboardSurgeries,'surgery_type')));
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
        {label:'Total operacional',value:resultSummary.totalEx+resultSummary.totalSu,detail:'Exames + cirurgias'}
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
        {label:'Financeiro pendente',value:resultSummary.paymentPending,detail:`${resultSummary.paymentPendingRate}% das cirurgias`,fill:[255,247,229],border:[239,220,179],valueColor:[151,104,32]},
        {label:'Pago / convenio',value:resultSummary.paymentClosed,detail:'Cirurgias com situacao financeira fechada',fill:[237,246,240],border:[211,231,218],valueColor:[62,118,86]}
      ],4);

      const quickText=
        `No periodo foram registrados ${resultSummary.totalEx} exames e ${resultSummary.totalSu} cirurgias para ${resultSummary.uniquePatients} pacientes unicos. `+
        `Nos exames, ${resultSummary.realEx} foram realizados, ${resultSummary.exOpen} permanecem em aberto e ${resultSummary.exCancelado} foram cancelados. `+
        `Nas cirurgias, ${resultSummary.realSu} foram realizadas, ${resultSummary.suOpen} permanecem em fluxo operacional e ${resultSummary.suCancelada} foram canceladas. `+
        `O exame de maior volume foi ${resultSummary.examTop?.[0]||'nao identificado'} e a cirurgia mais frequente foi ${resultSummary.surgTop?.[0]||'nao identificado'}.`;

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

      drawStatusPanel('Status das cirurgias',resultSummary.totalSu,[
        {label:'Realizada',value:resultSummary.realSu},
        {label:'Agendada',value:resultSummary.suAgendada},
        {label:'Autorizada',value:resultSummary.suAutorizada},
        {label:'Pendente',value:resultSummary.suPendente},
        {label:'Solicitacao',value:resultSummary.suSolicitacao},
        {label:'Cancelada',value:resultSummary.suCancelada}
      ]);

      drawStatusPanel('Situacao financeira das cirurgias',resultSummary.totalSu,[
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
        'Cirurgia'
      );

      drawDoctorTypeBreakdown(
        'CIRURGIAS POR CIRURGIAO E TIPO',
        advancedAnalytics.surgeryByDoctor,
        'Cirurgia'
      );

      if(advancedAnalytics.surgeryByIndication.length){
        drawDoctorTypeBreakdown(
          'INDICACOES DE CIRURGIA POR MEDICO',
          advancedAnalytics.surgeryByIndication,
          'Cirurgia indicada'
        );
      }

      drawTable(
        'DETALHAMENTO DAS CIRURGIAS',
        [
          {key:'date',label:'Data',width:16,value:r=>brDate(r.date)},
          {key:'patient',label:'Paciente',width:31,bold:true},
          {key:'doctor',label:'Cirurgiao',width:22},
          {key:'indicatedBy',label:'Indicado por',width:20,value:r=>r.indicatedBy||'—'},
          {key:'surgery_type',label:'Cirurgia',width:26,bold:true},
          {key:'procedure',label:'Procedimentos',width:30},
          {key:'eye',label:'Olho',width:12},
          {key:'insurance',label:'Convenio',width:18,value:r=>r.insurance||'Particular'},
          {key:'status',label:'Status',width:18},
          {key:'payment',label:'Pagamento',width:18},
          {key:'time',label:'Horario',width:14,value:r=>r.time||'—'},
          {key:'responsible',label:'Responsavel',width:18,value:r=>launchedName(r)},
          {key:'obs',label:'Observacao',width:30}
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
        countBy(dashboardSurgeries.map(x=>({...x,insuranceLabel:x.insurance||'Particular'})),'insuranceLabel'),
        8
      );

      drawSimpleRanking('Cirurgias por tipo',countBy(dashboardSurgeries,'surgery_type'),8);

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

  const logout=async()=>{ await supabase.auth.signOut(); setProfile(null); };

  if(loading) return <div className="loading-screen">Carregando...</div>;
  if(!session) return <Login onLogged={setSession}/>;

  return <div className="app">
    <aside className="sidebar">
      <div className="brand"><BrandMark/><div><b>Oftalmocastro</b><span>Gestão de exames e cirurgias</span></div></div>
      <div className="user-box">
        <b>{currentLaunchName}</b>
        <span>{session.user.email}</span>
        <small>{accessLabel(profile)} · lançamentos identificados automaticamente pelo login</small>
      </div>
      <div className="nav">
        <button className={page==='exames'?'active':''} onClick={()=>setPage('exames')}><Stethoscope className="nav-icon" size={18}/><span>Exames</span></button>
        <button className={page==='cirurgias'?'active':''} onClick={()=>setPage('cirurgias')}><Activity className="nav-icon" size={18}/><span>Cirurgias</span></button>
        <button className={page==='procedimentos'?'active':''} onClick={()=>setPage('procedimentos')}><ClipboardList className="nav-icon" size={18}/><span>Procedimentos</span></button>
        <button className={page==='followup'?'active':''} onClick={()=>setPage('followup')}><RefreshCw className="nav-icon" size={18}/><span>Follow-up</span></button>
        <button className={page==='resultados'?'active':''} onClick={()=>setPage('resultados')}><BarChart3 className="nav-icon" size={18}/><span>Resultados</span></button>
        <div className={`nav-finance-group ${['financeiro','relatorio-diario','promissorias','tarefas-pops'].includes(page)?'open active':''}`}>
          <button className={['financeiro','relatorio-diario','promissorias','tarefas-pops'].includes(page)?'active':''} onClick={()=>setPage('financeiro')} aria-haspopup="true">
            <span className="nav-finance-main"><WalletCards className="nav-icon" size={18}/><span>Financeiro</span></span><ChevronRight className="nav-chevron-icon" size={17}/>
          </button>
          <div className="nav-finance-submenu">
            <button className={page==='financeiro'?'active':''} onClick={()=>setPage('financeiro')}><LayoutDashboard size={16}/><span>Visão geral</span></button>
            <button className={page==='relatorio-diario'?'active':''} onClick={()=>setPage('relatorio-diario')}><CalendarDays size={16}/><span>Relatório diário</span></button>
            <button className={page==='promissorias'?'active':''} onClick={()=>setPage('promissorias')}><FileText size={16}/><span>Controle de Promissórias</span></button>
            <button className={page==='tarefas-pops'?'active':''} onClick={()=>setPage('tarefas-pops')}><ListChecks size={16}/><span>Tarefas e POPs</span></button>
          </div>
        </div>
        <button className={page==='cadastros'?'active':''} onClick={()=>setPage('cadastros')}><Settings className="nav-icon" size={18}/><span>Cadastros</span></button>
        <button className={page==='manual'?'active':''} onClick={()=>setPage('manual')}><CircleHelp className="nav-icon" size={18}/><span>Manual</span></button>
      </div>
      <button className="btn btn-light logout-btn" onClick={logout}><LogOut size={16}/> Sair</button>
    </aside>

    <main className="main">
      {page==='exames'&&<section>
        <div className="topbar">
          <div><div className="eyebrow">GESTÃO OPERACIONAL</div><h1>Controle mensal de exames</h1><div className="subtitle">Dados compartilhados entre os usuários da clínica.</div></div>
          <div className="top-actions">
            <button className="btn btn-light" onClick={()=>setPage('cadastros')}><Settings size={16}/> Cadastros</button>
            <button className="btn btn-primary" onClick={()=>{setEditingExamId(null);setExamForm({date:todayISO(),realization_date:'',eye:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',exam_type_ids:[],eye:'',insurance_id:'',status:'Agendado',obs:'',...emptyFinanceForm()});clearValidationErrors('exam');setExamModal(true)}}><Plus size={16}/> Novo exame</button>
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
            <PeriodFilter mode={filterPeriodMode} setMode={setFilterPeriodMode} month={filterMonth} setMonth={setFilterMonth} date={filterDate} setDate={setFilterDate}/>
            <Select value={filterDoctor} setValue={setFilterDoctor} options={doctors} first="Todos os médicos"/>
            <Select value={filterExam} setValue={setFilterExam} options={examTypes} first="Todos os exames"/>
            <Select value={examStatus} setValue={setExamStatus} options={['Agendado','Pendente','Realizado','Cancelado']} first="Todos os status"/>
            <input className="patient-search-input" value={examSearch} onChange={e=>setExamSearch(e.target.value)} placeholder="Buscar nome ou data de nascimento" aria-label="Buscar paciente por nome ou nascimento nos exames"/>
            <button className="btn btn-light" onClick={()=>{setFilterMonth('');setFilterDate('');setFilterPeriodMode('month');setFilterDoctor('');setFilterExam('');setExamStatus('');setExamSearch('')}}>Limpar</button>
            <button
              className="btn btn-danger-soft bulk-delete-btn"
              onClick={deleteFilteredExams}
              disabled={!filteredExams.length}
              title="Excluir todos os lançamentos exibidos pelos filtros atuais"
            >
              <Trash2 size={15}/> Excluir filtrados ({filteredExams.length})
            </button>
          </div>
          <div className="table-wrap"><table><thead><tr><th>Agendado em</th><th>Realização</th><th>Paciente</th><th>Médico</th><th>Indicado por</th><th>Exame</th><th>Status</th><th>Financeiro</th><th>Observação</th><th>Lançado por</th><th>Ações</th></tr></thead>
          <tbody>{filteredExams.length?filteredExams.map(r=><tr key={r.id}>
            <td>{brDate(r.date)}</td><td>{r.realization_date?brDate(r.realization_date):<span className="muted">—</span>}</td><td><PatientCell name={r.patient} phone={r.whatsapp} birthDate={r.patient_birth_date}/></td>
            <td>{r.doctor}</td><td>{r.indicatedBy||<span className="muted">—</span>}</td><td>{r.exam}</td><td><span className={`status ${statusClass(r.status)}`}>{r.status}</span></td>
            <td><PaymentTrafficLight record={r} onClick={()=>openEditExam(r)}/></td>
            <td><ObservationCell text={r.obs}/></td><td>{launchedName(r)}</td>
            <td><div className="row-actions">
              <button className="btn btn-edit-soft" onClick={()=>openEditExam(r)}>Abrir exame</button>
              <button className="btn btn-danger-soft" onClick={()=>delExam(r.id)}>Excluir</button>
            </div></td>
          </tr>):<tr><td colSpan="9" className="empty">Nenhum exame encontrado.</td></tr>}</tbody></table></div>
        </div>
      </section>}

      {page==='cirurgias'&&<section>
        <div className="topbar"><div><div className="eyebrow">CENTRO CIRÚRGICO</div><h1>Controle de cirurgias</h1><div className="subtitle">Programação, autorização, contato e observações.</div></div>
          <div className="top-actions">
            <button className="btn btn-light" onClick={()=>setPage('cadastros')}><Settings size={16}/> Cadastros</button>
            <button className="btn btn-primary" onClick={()=>{setEditingSurgeryId(null);setSForm({date:todayISO(),scheduled_on:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',surgery_type_id:'',procedure_ids:[],procedure_finance:{},lio_description:'',eye:'',insurance_id:'',status:'Solicitação',arrival:'',time:'',payment:'Não informado',obs:'',...emptyFinanceForm()});setSurgeryModal(true)}}><Plus size={16}/> Nova cirurgia</button>
          </div>
        </div>
        <div className="kpis"><Kpi label="Cirurgias no período" value={filteredSurgeries.length}/><Kpi label="Agendadas" value={countRowsByStatus(filteredSurgeries,'Agendada')}/><Kpi label="Pendentes" value={countRowsByStatus(filteredSurgeries,'Pendente')}/><Kpi label="Realizadas" value={countRowsByStatus(filteredSurgeries,'Realizada')}/><Kpi label="Procedimentos vinculados" value={totalItems(filteredSurgeries,'procedure_ids')}/></div>
        <div className="card">
          <div className="filters">
            <PeriodFilter mode={sPeriodMode} setMode={setSPeriodMode} month={sMonth} setMonth={setSMonth} date={sDate} setDate={setSDate}/>
            <Select value={sDoctor} setValue={setSDoctor} options={doctors} first="Todos os cirurgiões"/>
            <Select value={sSurgeryType} setValue={setSSurgeryType} options={surgeryTypes} first="Todas as cirurgias"/>
            <Select value={sProcedure} setValue={setSProcedure} options={procedureOptions} first="Todos os procedimentos vinculados"/>
            <Select value={sStatus} setValue={setSStatus} options={['Solicitação','Pendente','Autorizada','Agendada','Realizada','Cancelada']} first="Todos os status"/>
            <input className="patient-search-input" value={sSearch} onChange={e=>setSSearch(e.target.value)} placeholder="Buscar nome ou data de nascimento" aria-label="Buscar paciente por nome ou nascimento nas cirurgias"/>
            <button className="btn btn-light" onClick={()=>{setSMonth('');setSDate('');setSPeriodMode('month');setSDoctor('');setSSurgeryType('');setSProcedure('');setSStatus('');setSSearch('')}}>Limpar</button>
            <button
              className="btn btn-danger-soft bulk-delete-btn"
              onClick={deleteFilteredSurgeries}
              disabled={!filteredSurgeries.length}
              title="Excluir todos os lançamentos exibidos pelos filtros atuais"
            >
              <Trash2 size={15}/> Excluir filtrados ({filteredSurgeries.length})
            </button>
          </div>
          <div className="table-wrap"><table><thead><tr><th>Agendado em</th><th>Realização</th><th>Paciente</th><th>Cirurgião</th><th>Indicado por</th><th>Cirurgia</th><th>Procedimentos</th><th>Olho</th><th>Convênio</th><th>Status</th><th>Financeiro</th><th>Observação</th><th>Lançado por</th><th>Ações</th></tr></thead>
          <tbody>{filteredSurgeries.length?filteredSurgeries.map(r=><tr key={r.id}>
            <td>{brDate(r.date)}</td><td>{r.scheduled_on?brDate(r.scheduled_on):<span className="muted">—</span>}</td><td><PatientCell name={r.patient} phone={r.whatsapp} birthDate={r.patient_birth_date}/></td>
            <td>{r.doctor}</td><td>{r.indicatedBy||<span className="muted">—</span>}</td><td><b>{r.surgery_type}</b>{r.classification_review&&<small className="classification-review">Revisar classificação</small>}</td><td><LinkedProceduresCell names={r.procedure_names}/></td><td>{r.eye}</td><td>{r.insurance}</td>
            <td><span className={`status ${statusClass(r.status)}`}>{r.status}</span></td>
            <td><PaymentTrafficLight record={r} onClick={()=>openEditSurgery(r)}/></td>
            <td><ObservationCell text={r.obs}/></td><td>{launchedName(r)}</td>
            <td><div className="row-actions">
              <button className="btn btn-edit-soft" onClick={()=>openEditSurgery(r)}>Abrir cirurgia</button>
              <button className="btn btn-danger-soft" onClick={()=>delSurgery(r.id)}>Excluir</button>
            </div></td>
          </tr>):<tr><td colSpan="14" className="empty">Nenhuma cirurgia encontrada.</td></tr>}</tbody></table></div>
        </div>
      </section>}

      {page==='procedimentos'&&<section>
        <div className="topbar">
          <div>
            <div className="eyebrow">PROCEDIMENTOS</div>
            <h1>Controle de procedimentos</h1>
            <div className="subtitle">Área exclusiva para procedimentos. Itens vinculados a uma cirurgia aparecem aqui sem aumentar a quantidade de cirurgias.</div>
          </div>
          <div className="top-actions">
            <button className="btn btn-light" onClick={()=>setPage('cadastros')}><Settings size={16}/> Cadastros</button>
            <button className="btn btn-primary" onClick={()=>{
              setEditingProcedureId(null);
              setPForm({date:todayISO(),realization_date:'',patient:'',patient_birth_date:'',whatsapp:'',doctor_id:'',indicated_by_doctor_id:'',procedure_ids:[],eye:'',insurance_id:'',status:'Solicitação',time:'',payment:'Não informado',obs:'',...emptyFinanceForm()});
              setProcedureModal(true);
            }}><Plus size={16}/> Novo procedimento</button>
          </div>
        </div>

        <div className="kpis">
          <Kpi label="Procedimentos no período" value={filteredProcedures.length}/>
          <Kpi label="Agendados" value={countRowsByStatus(filteredProcedures,'Agendada')}/>
          <Kpi label="Pendentes" value={countRowsByStatus(filteredProcedures,'Pendente')}/>
          <Kpi label="Realizados" value={countRowsByStatus(filteredProcedures,'Realizada')}/>
          <Kpi label="Vinculados a cirurgias" value={filteredProcedures.filter(r=>r.surgery_group_id).length}/>
        </div>

        <div className="card">
          <div className="filters">
            <PeriodFilter mode={pPeriodMode} setMode={setPPeriodMode} month={pMonth} setMonth={setPMonth} date={pDate} setDate={setPDate}/>
            <Select value={pDoctor} setValue={setPDoctor} options={doctors} first="Todos os médicos"/>
            <Select value={pProcedure} setValue={setPProcedure} options={procedureOptions} first="Todos os procedimentos"/>
            <Select value={pStatus} setValue={setPStatus} options={['Solicitação','Pendente','Autorizada','Agendada','Realizada','Cancelada']} first="Todos os status"/>
            <input className="patient-search-input" value={pSearch} onChange={e=>setPSearch(e.target.value)} placeholder="Buscar nome ou data de nascimento" aria-label="Buscar paciente por nome ou nascimento nos procedimentos"/>
            <button className="btn btn-light" onClick={()=>{setPMonth('');setPDate('');setPPeriodMode('month');setPDoctor('');setPProcedure('');setPStatus('');setPSearch('')}}>Limpar</button>
            <button className="btn btn-danger-soft bulk-delete-btn" onClick={deleteFilteredProcedures} disabled={!filteredProcedures.length}>
              <Trash2 size={15}/> Excluir filtrados ({filteredProcedures.length})
            </button>
          </div>

          <div className="table-wrap"><table><thead><tr>
            <th>Agendado em</th><th>Realização</th><th>Paciente</th><th>Médico</th><th>Indicado por</th><th>Procedimento</th><th>Origem</th><th>Olho</th><th>Convênio</th><th>Status</th><th>Financeiro</th><th>Observação</th><th>Lançado por</th><th>Ações</th>
          </tr></thead>
          <tbody>{filteredProcedureLaunches.length?filteredProcedureLaunches.map(r=><tr key={r.group_key||r.id}>
            <td>{brDate(r.date)}{r.time&&<><br/><small>{r.time}</small></>}</td>
            <td>{r.realization_date?brDate(r.realization_date):<span className="muted">—</span>}</td>
            <td><PatientCell name={r.patient} phone={r.whatsapp} birthDate={r.patient_birth_date}/></td>
            <td>{r.doctor}</td>
            <td>{r.indicatedBy||<span className="muted">—</span>}</td>
            <td><ProcedureListCell names={r.procedure_names}/></td>
            <td>{r.surgery_group_id?<span className="procedure-origin linked">Pacote: {r.linked_surgery||'Cirurgia'}</span>:<span className="procedure-origin standalone">Avulso</span>}</td>
            <td>{r.eye}</td>
            <td>{r.insurance||'Particular'}</td>
            <td><span className={`status ${statusClass(r.status)}`}>{r.status}</span></td>
            <td><PaymentTrafficLight record={r} onClick={()=>openEditProcedure(r)}/></td>
            <td><ObservationCell text={r.obs}/></td>
            <td>{launchedName(r)}</td>
            <td><div className="row-actions">
              <button className="btn btn-edit-soft" onClick={()=>openEditProcedure(r)}>Abrir procedimento</button>
              <button className="btn btn-danger-soft" onClick={()=>delProcedure(r)}>Excluir</button>
            </div></td>
          </tr>):<tr><td colSpan="14" className="empty">Nenhum procedimento encontrado.</td></tr>}</tbody></table></div>
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
                <button className="btn btn-light" onClick={()=>openOriginalFromFollowup(row)}>Ir para lançamento de origem</button>
              </div>
            </article>;
          }) : <div className="card empty followup-empty">Nenhum paciente encontrado com os filtros selecionados.</div>}
        </div>
      </section>}

      {page==='relatorio-diario'&&<section className="daily-report-page">
        <div className="topbar">
          <div><div className="eyebrow">RELATÓRIO OPERACIONAL</div><h1>Relatório diário</h1><div className="subtitle">Conferência diária de exames, cirurgias e procedimentos, com situação operacional e financeira.</div></div>
          <div className="top-actions daily-export-actions">
            <button className="btn btn-light" onClick={()=>setPage('financeiro')}>← Financeiro</button>
            <button className="btn btn-light" onClick={exportDailyReportWord}><FileText size={16}/> Exportar Word</button>
            <button className="btn btn-primary" onClick={exportDailyReportPdf}><FileDown size={16}/> Gerar PDF</button>
          </div>
        </div>
        <div className="card daily-report-toolbar">
          <div className="filters daily-report-filters">
            <Field label="Data"><input type="date" value={dailyDate} onChange={e=>setDailyDate(e.target.value)}/></Field>
            <Field label="Nascimento"><input type="date" value={dailyBirthDate} onChange={e=>setDailyBirthDate(e.target.value)}/></Field>
            <Field label="Paciente"><input value={dailySearch} onChange={e=>setDailySearch(e.target.value)} placeholder="Nome ou nascimento"/></Field>
            <div className="full filter-check-section"><label>Tipo</label><div className="filter-check-grid">
              {[['exam','Exames'],['surgery','Cirurgias'],['procedure','Procedimentos']].map(([id,label])=><button type="button" key={id} className={`filter-check ${dailySources.includes(id)?'selected':''}`} onClick={()=>toggleDailyFilter(dailySources,id,setDailySources)}><span>{dailySources.includes(id)?'✓':''}</span>{label}</button>)}
            </div></div>
            <div className="full filter-check-section"><label>Situação</label><div className="filter-check-grid">
              {[['lancado','Lançados / agendados'],['realizado','Realizados'],['cancelado','Cancelados']].map(([id,label])=><button type="button" key={id} className={`filter-check ${dailyStages.includes(id)?'selected':''}`} onClick={()=>toggleDailyFilter(dailyStages,id,setDailyStages)}><span>{dailyStages.includes(id)?'✓':''}</span>{label}</button>)}
            </div></div>
            <button className="btn btn-light" onClick={()=>{setDailyDate(todayISO());setDailySources([]);setDailyStages([]);setDailySearch('');setDailyBirthDate('')}}>Hoje / limpar filtros</button>
          </div>
        </div>
        <div className="card daily-report-sheet" ref={dailyReportRef}>
          <div className="daily-report-title"><b>Exames / Cirurgias / Procedimentos — {brDate(dailyDate)} — {weekdayLabel(dailyDate)}</b><small>{dailyReportRows.length} registro(s) encontrado(s)</small></div>
          {dailyReportRows.length?dailyReportRows.map((r,index)=><article className="daily-report-item" key={`${r.report_source}:${r.group_id||r.id}:${index}`}>
            <div className="daily-report-item-head">
              <strong>{index+1}. {r.report_type.toUpperCase()}: {r.service_name}</strong>
              <div className="daily-report-item-actions no-pdf">
                <span className={`status ${statusClass(r.status)}`}>{r.status}</span>
                <button type="button" className="btn btn-edit-soft daily-edit-btn" onClick={()=>openDailyReportEdit(r)}>Editar</button>
              </div>
            </div>
            <div><b>DATA DO AGENDAMENTO:</b> {brDate(r.date)} &nbsp; <b>DATA DA REALIZAÇÃO:</b> {brDate(r.report_source==='surgery'?(r.scheduled_on||''):(r.realization_date||''))||'—'}</div>
            {r.report_source==='surgery'&&<div><b>LIO:</b> {r.lio_description||'—'}</div>}
            <div><b>CONVÊNIO:</b> {r.insurance||financeOrigin(r)||'Particular'}</div>
            <div><b>NOME DO PACIENTE:</b> <strong>{r.patient}</strong>{r.patient_birth_date?` — Nasc. ${brDate(r.patient_birth_date)}`:''}</div>
            <div><b>PAGAMENTO:</b> {paymentToneLabel(r)}{r.payment_method?` — ${r.payment_method}`:''}{Number(r.installments||1)>1?` — ${r.installments}x`:''}</div>
            <div><b>VALOR / PAGAMENTOS:</b> {paymentSummaryText(r)}</div>
            <div><b>INDICAÇÃO:</b> {r.indicatedBy||'—'}</div>
            <div><b>{r.report_source==='surgery'?'CIRURGIÃO':'REALIZADO POR'}:</b> {r.doctor||'—'}</div>
            <div><b>REPASSES:</b> {repassesSummaryText(r)}</div>
            {(r.financial_documents?.length||r.invoice_file_path)&&<div><b>DOCUMENTOS:</b> {r.financial_documents?.length||1} anexo(s)</div>}
          </article>):<div className="empty">Nenhum lançamento encontrado para os filtros selecionados.</div>}
        </div>
      </section>}

      {page==='resultados'&&<section className="results-export-root" ref={resultsPageRef}>
        <div className="pdf-page-block" ref={pdfPage1Ref}>
          <div className="topbar results-header">
            <div>
              <div className="eyebrow">INTELIGÊNCIA OPERACIONAL</div>
              <h1>Resultados e análise mensal</h1>
              <div className="subtitle">Visão consolidada e separada de exames, cirurgias, procedimentos, pacientes e responsáveis pelos lançamentos.</div>
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
              <PeriodFilter mode={dashPeriodMode} setMode={setDashPeriodMode} month={dashMonth} setMonth={setDashMonth} date={dashDate} setDate={setDashDate}/>
              <Select value={dashDoctor} setValue={setDashDoctor} options={doctors} first="Todos os médicos"/>
              <Select value={dashUser} setValue={setDashUser} options={['Juliana','Luiza','Elisangela','ADM']} first="Todos os responsáveis"/>
              <Select value={dashStatus} setValue={setDashStatus} options={['Agendado','Realizado','Cancelado','Solicitação','Pendente','Autorizada','Agendada','Realizada','Cancelada']} first="Todos os status"/>
              <input className="patient-search-input" value={dashPatient} onChange={e=>setDashPatient(e.target.value)} placeholder="Buscar nome ou data de nascimento" aria-label="Buscar paciente por nome ou nascimento nos resultados"/>
            </div>
            <div className="active-period">Período analisado: <b>{dashDate?brDate(dashDate):monthLabel(dashMonth)}</b></div>
          </div>

          <div className="analysis-grid results-kpis">
            <div className="analysis-card highlight"><span>Pacientes únicos</span><strong>{resultSummary.uniquePatients}</strong><small>considerando os três módulos</small></div>
            <div className="analysis-card"><span>Total de exames</span><strong>{resultSummary.totalEx}</strong><small>{resultSummary.realEx} realizados • {resultSummary.exRate}% de realização</small></div>
            <div className="analysis-card"><span>Total de cirurgias</span><strong>{resultSummary.totalSu}</strong><small>{resultSummary.realSu} realizadas • {resultSummary.suRate}% de realização</small></div>
            <div className="analysis-card procedure-kpi"><span>Total de procedimentos</span><strong>{resultSummary.totalProcedures}</strong><small>{resultSummary.realProcedures} realizados • {resultSummary.procedureRate}% de realização</small></div>
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

            <div className="card status-summary-card procedure-status-card">
              <div className="status-summary-head">
                <div>
                  <span>PROCEDIMENTOS</span>
                  <h3>Situação dos procedimentos</h3>
                  <p>Procedimentos são analisados separadamente e não aumentam a quantidade de cirurgias.</p>
                </div>
                <div className="status-rate success"><b>{resultSummary.procedureRate}%</b><small>realização</small></div>
              </div>
              <div className="status-breakdown-grid surgery">
                <div className="status-breakdown-item success"><span>Realizados</span><b>{resultSummary.realProcedures}</b><small>concluídos</small></div>
                <div className="status-breakdown-item info"><span>Agendados</span><b>{resultSummary.prAgendada}</b><small>com data</small></div>
                <div className="status-breakdown-item info"><span>Autorizados</span><b>{resultSummary.prAutorizada}</b><small>liberados</small></div>
                <div className="status-breakdown-item warning"><span>Pendentes</span><b>{resultSummary.prPendente}</b><small>precisam atenção</small></div>
                <div className="status-breakdown-item neutral"><span>Solicitações</span><b>{resultSummary.prSolicitacao}</b><small>em análise</small></div>
                <div className="status-breakdown-item danger"><span>Cancelados</span><b>{resultSummary.prCancelada}</b><small>fora do fluxo</small></div>
              </div>
              <div className="status-summary-foot"><b>{resultSummary.prOpen}</b> procedimento(s) ainda em fluxo • {resultSummary.procedureOpenRate}% do total</div>
            </div>
          </div>

          <div className="results-operational-strip">
            <div><span>Total operacional</span><b>{resultSummary.totalEx+resultSummary.totalSu+resultSummary.totalProcedures}</b><small>exames + cirurgias + procedimentos, sem misturar as categorias</small></div>
            <div><span>Financeiro em aberto</span><b>{resultSummary.paymentPending}</b><small>cirurgias: pendente, 20% pago ou não informado</small></div>
            <div><span>Exame de maior volume</span><b className="text-value">{resultSummary.examTop?.[0]||'Sem dados'}</b><small>{resultSummary.examTop?.[1]||0} exame(s)</small></div>
            <div><span>Cirurgia de maior volume</span><b className="text-value">{resultSummary.surgTop?.[0]||'Sem dados'}</b><small>{resultSummary.surgTop?.[1]||0} cirurgia(s)</small></div>
            <div><span>Procedimento de maior volume</span><b className="text-value">{resultSummary.procedureTop?.[0]||'Sem dados'}</b><small>{resultSummary.procedureTop?.[1]||0} procedimento(s)</small></div>
          </div>

          <div className="results-highlights">
            <div className="card insight-card"><span>Exame com maior volume</span><b>{resultSummary.examTop?.[0]||'Sem dados'}</b><small>{resultSummary.examTop?.[1]||0} exame(s)</small></div>
            <div className="card insight-card"><span>Cirurgia com maior volume</span><b>{resultSummary.surgTop?.[0]||'Sem dados'}</b><small>{resultSummary.surgTop?.[1]||0} cirurgia(s)</small></div>
            <div className="card insight-card"><span>Procedimento com maior volume</span><b>{resultSummary.procedureTop?.[0]||'Sem dados'}</b><small>{resultSummary.procedureTop?.[1]||0} procedimento(s)</small></div>
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
                <p>Veja exatamente quais exames e cirurgias estão associados a cada médico.</p>
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
                title="Cirurgias por cirurgião e tipo"
                subtitle="Cada caso cirúrgico conta uma vez, independentemente dos procedimentos associados."
                rows={advancedAnalytics.surgeryByDoctor}
                itemLabel="Cirurgia"
              />

              <DoctorTypeMatrix
                title="Procedimentos por médico e tipo"
                subtitle="Procedimentos avulsos e procedimentos de pacote são contados aqui, separados das cirurgias."
                rows={advancedAnalytics.procedureByDoctor}
                itemLabel="Procedimento"
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
                subtitle="Quantidade e tipo de cirurgia indicada por cada médico."
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
            title="Cirurgias por status e por tipo"
            subtitle="Mostra quantas cirurgias de cada tipo estão em solicitação, pendentes, autorizadas, agendadas, realizadas ou canceladas. Cada caso conta uma única vez."
            statuses={['Pendente','Solicitação','Autorizada','Agendada','Realizada','Cancelada']}
            data={advancedAnalytics.surgeryByStatusType}
            totalByStatus={advancedAnalytics.surgeryStatusTotals}
            labelName="Cirurgia"
          />

          <StatusTypeMatrix
            title="Procedimentos por status e por tipo"
            subtitle="Mostra Botox, Laser CO₂, Bioestimulador, anti-VEGF e demais procedimentos em uma análise própria, sem somá-los às cirurgias."
            statuses={['Pendente','Solicitação','Autorizada','Agendada','Realizada','Cancelada']}
            data={advancedAnalytics.procedureByStatusType}
            totalByStatus={advancedAnalytics.procedureStatusTotals}
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
            <div className="results-section-title"><div><h3>Análise de cirurgias</h3><p>Produção, tipos de cirurgia, status e convênios.</p></div><div className="pill">{resultSummary.totalSu} cirurgias</div></div>
            <div className="grid">
              <div className="card chart-card"><h2>Cirurgias por tipo</h2><Bars data={countBy(dashboardSurgeries,'surgery_type')}/></div>
              <div className="card chart-card"><h2>Cirurgias por cirurgião</h2><Bars data={countBy(dashboardSurgeries,'doctor')}/></div>
            </div>
          </div>

          <div className="results-section procedure-results-section">
            <div className="results-section-title"><div><h3>Análise de procedimentos</h3><p>Procedimentos são analisados em uma seção própria e permanecem vinculados ao pacote quando houver cirurgia.</p></div><div className="pill">{resultSummary.totalProcedures} procedimentos</div></div>
            <div className="grid">
              <div className="card chart-card"><h2>Procedimentos por tipo</h2><Bars data={countBy(dashboardProcedures,'procedure')}/></div>
              <div className="card chart-card"><h2>Procedimentos por médico</h2><Bars data={countBy(dashboardProcedures,'doctor')}/></div>
            </div>
          </div>
        </div>

        <div className="pdf-page-block" ref={pdfPage3Ref}>
          <div className="grid pdf-continuation-grid">
            <div className="card chart-card"><h2>Status das cirurgias</h2><Bars data={countBy(dashboardSurgeries,'status')}/></div>
            <div className="card chart-card"><h2>Cirurgias por convênio</h2><Bars data={countBy(dashboardSurgeries.map(x=>({...x,insuranceLabel:x.insurance||'Particular'})),'insuranceLabel')}/></div>
          </div>

          <div className="grid pdf-continuation-grid procedure-results-grid">
            <div className="card chart-card"><h2>Status dos procedimentos</h2><Bars data={countBy(dashboardProcedures,'status')}/></div>
            <div className="card chart-card"><h2>Origem dos procedimentos</h2><Bars data={countBy(dashboardProcedures.map(x=>({...x,originLabel:x.surgery_group_id?'Vinculado a cirurgia':'Avulso'})),'originLabel')}/></div>
          </div>

          <div className="card results-users-card">
            <div className="section-title">
              <div><h2>Lançamentos por responsável</h2><p className="section-description">Acompanhe separadamente exames, cirurgias e procedimentos registrados por cada usuário.</p></div>
              <div className="pill">{resultSummary.totalEx+resultSummary.totalSu+resultSummary.totalProcedures} registros</div>
            </div>
            <div className="table-wrap"><table><thead><tr><th>Responsável</th><th>Exames</th><th>Cirurgias</th><th>Procedimentos</th><th>Total</th><th>Participação</th></tr></thead>
            <tbody>{userRows.map(r=>{
              const total=r.exams+r.surgeries+r.procedures;
              const all=resultSummary.totalEx+resultSummary.totalSu+resultSummary.totalProcedures;
              return <tr key={r.key}><td><b>{r.name}</b></td><td>{r.exams}</td><td>{r.surgeries}</td><td>{r.procedures}</td><td><b>{total}</b></td><td>{all?Math.round(total/all*100):0}%</td></tr>
            })}</tbody></table></div>
          </div>
        </div>
      </section>}

      {page==='financeiro'&&<section className="finance-page">
        <div className="topbar finance-topbar">
          <div><div className="eyebrow">GESTÃO FINANCEIRA</div><h1>Financeiro</h1><div className="subtitle">Produção, recebimentos, contas a receber e potencial financeiro sem duplicar valores de pacotes cirúrgicos.</div></div>
        </div>

        <div className="card finance-filter-card">
          <div className="filter-title">Filtros financeiros</div>
          <div className="finance-filters">
            <PeriodFilter mode={finPeriodMode} setMode={setFinPeriodMode} month={finMonth} setMonth={setFinMonth} date={finDate} setDate={setFinDate}/>
            <Select value={finDoctor} setValue={setFinDoctor} options={doctors} first="Todos os médicos"/>
            <select value={finInsurance} onChange={e=>setFinInsurance(e.target.value)}><option value="">Todos os convênios</option><option value="particular">Particular</option>{insurances.map(i=><option key={i.id} value={i.id}>{i.name}</option>)}</select>
            <Select value={finOrigin} setValue={setFinOrigin} options={['Particular','Convênio','Misto']} first="Todas as origens"/>
            <select value={finSource} onChange={e=>setFinSource(e.target.value)}><option value="">Exames + Cirurgias + Procedimentos</option><option value="exam">Exames</option><option value="surgery">Cirurgias</option><option value="procedure">Procedimentos</option><option value="extra">Adicionais de pacote</option></select>
            <Select value={finOperationalStatus} setValue={setFinOperationalStatus} options={['Solicitação','Pendente','Autorizada','Agendada','Realizada','Cancelada','Realizado','Agendado','Cancelado']} first="Todos os status operacionais"/>
            <Select value={finStatus} setValue={setFinStatus} options={[{id:'paid',name:'Pago'},{id:'partial',name:'Pago parcial'},{id:'unpaid',name:'Não pago'}]} first="Todos os pagamentos"/>
            <Select value={finMethod} setValue={setFinMethod} options={['PIX','Dinheiro','Débito','Crédito','Crédito parcelado','Boleto','Transferência','Convênio','Múltiplos']} first="Todas as formas de pagamento"/>
            <input value={finSearch} onChange={e=>setFinSearch(e.target.value)} placeholder="Buscar paciente, nascimento ou serviço"/>
            <button className="btn btn-light" onClick={()=>{setFinMonth('');setFinDate('');setFinPeriodMode('month');setFinDoctor('');setFinInsurance('');setFinOrigin('');setFinSource('');setFinOperationalStatus('');setFinStatus('');setFinMethod('');setFinSearch('')}}>Limpar</button>
          </div>
        </div>

        <div className="finance-kpis">
          <div className="finance-kpi primary"><span>Produção realizada</span><strong>{moneyBR(financeSummary.production)}</strong><small>Serviços efetivamente realizados</small></div>
          <div className="finance-kpi success"><span>Recebido</span><strong>{moneyBR(financeSummary.received)}</strong><small>Valores já recebidos</small></div>
          <div className="finance-kpi warning"><span>A receber</span><strong>{moneyBR(financeSummary.receivable)}</strong><small>Realizados ainda não recebidos</small></div>
          <div className="finance-kpi info"><span>Potencial em aberto</span><strong>{moneyBR(financeSummary.potential)}</strong><small>Agendados, autorizados e pendentes</small></div>
          <div className="finance-kpi"><span>Particular</span><strong>{moneyBR(financeSummary.particular)}</strong><small>Valor total particular</small></div>
          <div className="finance-kpi"><span>Convênios</span><strong>{moneyBR(financeSummary.insurance)}</strong><small>Valor total por convênios</small></div>
        </div>

        <div className="finance-main-grid">
          <FinancialLineChart title="Previsto x realizado x recebido" labels={financeAnalytics.labels} series={[
            {name:'Produção realizada',color:'#315f72',values:financeAnalytics.productionTrend},
            {name:'Recebido',color:'#4f8a6b',values:financeAnalytics.receivedTrend},
            {name:'Potencial',color:'#b9802b',values:financeAnalytics.potentialTrend}
          ]}/>
          <div className="card finance-side-summary">
            <div className="finance-chart-head"><div><span>LEITURA RÁPIDA</span><h3>Posição financeira</h3><p>Resumo dos valores filtrados.</p></div></div>
            <div className="finance-summary-list">
              <div><span>Valor faturado a convênios</span><b>{moneyBR(financeSummary.billed)}</b></div>
              <div><span>Cancelado / perdido</span><b>{moneyBR(financeSummary.lost)}</b></div>
              <div><span>Conversão em recebimento</span><b>{financeSummary.production?Math.round((financeSummary.received/financeSummary.production)*100):0}%</b></div>
              <div><span>Lançamentos financeiros</span><b>{filteredFinancialRows.length}</b></div>
            </div>
          </div>
        </div>

        <div className="finance-analysis-grid">
          <FinancialBars title="Faturamento por convênio" subtitle="Particular também aparece para comparação." data={financeAnalytics.byInsurance}/>
          <FinancialBars title="Faturamento por médico" subtitle="Valor associado aos atendimentos de cada profissional." data={financeAnalytics.byDoctor}/>
          <FinancialBars title="Por cirurgia / procedimento" subtitle="Serviços com maior impacto financeiro." data={financeAnalytics.byService}/>
          <FinancialBars title="Por forma de pagamento" subtitle="Distribuição dos valores por meio de recebimento." data={financeAnalytics.byMethod}/>
        </div>

        <div className="card finance-detail-card">
          <div className="section-title"><div><h2>Detalhamento financeiro</h2><p className="finance-detail-subtitle">Produção e recebimento por lançamento. Pacotes aparecem uma única vez.</p></div></div>
          <div className="table-wrap"><table className="finance-table"><thead><tr><th>Data referência</th><th>Paciente</th><th>Pagamento</th><th>Origem</th><th>Serviço</th><th>Indicado por</th><th>Realizado por</th><th>Convênio</th><th>Status</th><th>Valor bruto</th><th>Desconto</th><th>Valor líquido</th><th>Recebido</th><th>Saldo</th><th>Forma</th><th>NF</th></tr></thead>
          <tbody>{filteredFinancialRows.length?filteredFinancialRows.map(r=><tr key={r.key}>
            <td>{brDate(r.date)}</td><td><b>{r.patient}</b></td><td><PaymentTrafficLight record={r} onClick={()=>{if(r.source_type==='surgery'){openEditSurgery(r.base_record);setPage('cirurgias')}else if(r.source_type==='exam'){openEditExam(r.base_record);setPage('exames')}else if(r.source_type==='procedure'){openEditProcedure(r.base_record);setPage('procedimentos')}}}/></td>
            <td>{r.payment_origin||'—'}</td><td><b>{r.service}</b></td><td>{r.indicatedBy||'—'}</td><td>{r.doctor||'—'}</td><td>{r.insurance||'Particular'}</td><td><span className={`status ${statusClass(r.operational_status)}`}>{r.operational_status}</span></td>
            <td>{moneyBR(r.gross??r.total)}</td><td>{moneyBR(r.discount||0)}</td><td><b>{moneyBR(r.total)}</b></td><td>{moneyBR(r.received)}</td><td>{moneyBR(Math.max(r.total-r.received,0))}</td><td>{r.payment_method||normalizedBreakdown(r.payment_breakdown).map(x=>x.method).filter(Boolean).join(' + ')||'—'}</td>
            <td>{(r.financial_documents?.length||r.invoice_file_path)?<div className="table-document-actions">{(r.financial_documents?.length?r.financial_documents:[{path:r.invoice_file_path,name:r.invoice_number||'Nota fiscal'}]).map((doc,i)=><span key={doc.path||i} className="table-document-item"><button type="button" className="btn btn-light invoice-open-btn" onClick={()=>openInvoiceFile(doc.path,doc.name)}>Visualizar</button><button type="button" className="btn btn-light invoice-open-btn" onClick={()=>downloadInvoiceFile(doc.path,doc.name||`documento-${i+1}`)}>Baixar</button></span>)}</div>:r.invoice_number||'—'}</td>
          </tr>):<tr><td colSpan="16" className="empty">Nenhum valor financeiro encontrado.</td></tr>}</tbody></table></div>
        </div>

        <div className="card invoice-management-card">
          <div className="section-title invoice-management-title">
            <div><h2>Notas fiscais e comprovantes</h2><p className="finance-detail-subtitle">Gestão de documentos financeiros por data de emissão, paciente e tipo de atendimento.</p></div>
            <div className="pill">{invoiceManagementRows.length} registro(s)</div>
          </div>
          <div className="invoice-management-filters">
            <Field label="Emitida de"><input type="date" value={invoiceDateFrom} onChange={e=>setInvoiceDateFrom(e.target.value)}/></Field>
            <Field label="Até"><input type="date" value={invoiceDateTo} onChange={e=>setInvoiceDateTo(e.target.value)}/></Field>
            <Field label="Paciente / NF"><input value={invoiceSearch} onChange={e=>setInvoiceSearch(e.target.value)} placeholder="Nome do paciente ou nº da NF"/></Field>
            <Field label="Tipo"><Select value={invoiceSource} setValue={setInvoiceSource} options={[{id:'exam',name:'Exames'},{id:'surgery',name:'Cirurgias'},{id:'procedure',name:'Procedimentos'}]} first="Todos"/></Field>
            <Field label="Documento"><Select value={invoiceDocumentStatus} setValue={setInvoiceDocumentStatus} options={[{id:'com',name:'Com NF / comprovante'},{id:'sem',name:'Sem NF / comprovante'}]} first="Todos"/></Field>
            <button className="btn btn-light invoice-clear-btn" onClick={()=>{setInvoiceDateFrom('');setInvoiceDateTo('');setInvoiceSearch('');setInvoiceSource('');setInvoiceDocumentStatus('')}}>Limpar</button>
          </div>
          <div className="table-wrap invoice-management-table"><table><thead><tr><th>Emissão</th><th>Paciente</th><th>Tipo</th><th>Serviço</th><th>Nº NF / referência</th><th>Valor</th><th>Status financeiro</th><th>Documentos</th></tr></thead>
          <tbody>{invoiceManagementRows.length?invoiceManagementRows.map(r=>{
            const docs=r.financial_documents?.length?r.financial_documents:(r.invoice_file_path?[{path:r.invoice_file_path,name:r.invoice_number||'Nota fiscal'}]:[]);
            return <tr key={`invoice:${r.key}`}>
              <td>{r.invoice_issued_date?brDate(r.invoice_issued_date):<span className="muted">—</span>}</td>
              <td><b>{r.patient}</b><small className="invoice-patient-date">Lançamento: {brDate(r.date)}</small></td>
              <td>{r.source}</td><td>{r.service}</td><td>{r.invoice_number||<span className="muted">—</span>}</td><td><b>{moneyBR(r.total)}</b></td><td><span className={`payment-label ${paymentTone(r)}`}>{paymentToneLabel(r)}</span></td>
              <td>{docs.length?<div className="invoice-document-list">{docs.map((doc,i)=><div key={doc.path||i}><span>{doc.name||`Documento ${i+1}`}</span><button type="button" className="btn btn-light" onClick={()=>openInvoiceFile(doc.path,doc.name)}>Visualizar</button><button type="button" className="btn btn-light" onClick={()=>downloadInvoiceFile(doc.path,doc.name||`documento-${i+1}`)}>Baixar</button></div>)}</div>:<span className="muted">Sem documento</span>}</td>
            </tr>;
          }):<tr><td colSpan="8" className="empty">Nenhuma nota fiscal ou lançamento encontrado com estes filtros.</td></tr>}</tbody></table></div>
        </div>
      </section>}


      {page==='promissorias'&&<section className="promissory-page finance-page">
        <div className="topbar finance-topbar"><div><div className="eyebrow">GESTÃO FINANCEIRA</div><h1>Controle de Promissórias</h1><div className="subtitle">Acompanhe promissórias em aberto, pagas e vencidas a partir dos pagamentos registrados em Exames, Cirurgias e Procedimentos.</div></div><div className="top-actions"><button className="btn btn-light" onClick={exportPromissoryWord}><FileText size={16}/> Exportar Word</button><button className="btn btn-primary" onClick={exportPromissoryPdf}><FileDown size={16}/> Gerar PDF</button></div></div>
        <div className="card finance-filter-card"><div className="filter-title">Filtros das promissórias</div><div className="finance-filters promissory-filters"><input type="month" value={promMonth} onChange={e=>setPromMonth(e.target.value)}/><input type="date" value={promDate} onChange={e=>{setPromDate(e.target.value);if(e.target.value)setPromMonth(e.target.value.slice(0,7))}} title="Filtrar por dia, vencimento ou pagamento"/><input value={promSearch} onChange={e=>setPromSearch(e.target.value)} placeholder="Buscar paciente ou serviço"/><Select value={promStatus} setValue={setPromStatus} options={['Em aberto','Parcial','Vencida','Pago']} first="Todas as situações"/><button className="btn btn-light" onClick={()=>{setPromMonth('');setPromDate('');setPromStatus('');setPromSearch('')}}>Limpar</button></div></div>
        <div className="finance-kpis promissory-kpis"><div className="finance-kpi primary"><span>Total em promissórias</span><strong>{moneyBR(promissorySummary.total)}</strong><small>{filteredPromissoryRows.length} registro(s)</small></div><div className="finance-kpi success"><span>Recebido</span><strong>{moneyBR(promissorySummary.received)}</strong><small>Valores quitados/recebidos</small></div><div className="finance-kpi warning"><span>Saldo em aberto</span><strong>{moneyBR(promissorySummary.balance)}</strong><small>Valor ainda a receber</small></div><div className="finance-kpi danger"><span>Vencido</span><strong>{moneyBR(promissorySummary.overdue)}</strong><small>Promissórias vencidas ainda abertas</small></div></div>
        <div className="card finance-detail-card"><div className="section-title"><div><h2>Promissórias</h2><p className="finance-detail-subtitle">Os dados vêm diretamente das formas de pagamento dos lançamentos. Para corrigir valor, vencimento ou situação, abra o lançamento original.</p></div></div><div className="table-wrap"><table className="finance-table"><thead><tr><th>Data</th><th>Paciente</th><th>Origem</th><th>Serviço</th><th>Valor</th><th>Parcelas</th><th>Vencimento</th><th>Situação</th><th>Recebido</th><th>Saldo</th><th>Ação</th></tr></thead><tbody>{filteredPromissoryRows.length?filteredPromissoryRows.map(r=><tr key={r.id}><td>{brDate(r.reference_date)}</td><td><b>{r.patient}</b></td><td>{r.source}</td><td>{r.service}</td><td><b>{moneyBR(r.amount)}</b></td><td>{r.installments}x</td><td>{r.due_date?brDate(r.due_date):'—'}</td><td><span className={`promissory-status ${r.status==='Pago'?'paid':r.status==='Vencida'?'overdue':r.status==='Parcial'?'partial':'open'}`}>{r.status}</span></td><td>{moneyBR(r.received)}</td><td><b>{moneyBR(r.balance)}</b></td><td><button className="btn btn-edit-soft" onClick={()=>openPromissoryRecord(r)}>Abrir lançamento</button></td></tr>):<tr><td colSpan="11" className="empty">Nenhuma promissória encontrada com os filtros atuais.</td></tr>}</tbody></table></div></div>
      </section>}

      {page==='tarefas-pops'&&<TasksPopsPanel profile={profile} session={session}/>}


      {page==='cadastros' &&<section className="catalog-page">
        <div className="topbar">
          <div>
            <div className="eyebrow">CONFIGURAÇÕES DO SISTEMA</div>
            <h1>Cadastros da clínica</h1>
            <div className="subtitle">ADM e Orientadora podem cadastrar médicos, exames, cirurgias e procedimentos em catálogos separados. Procedimento nunca é contabilizado como cirurgia.</div>
          </div>
        </div>

        <div className="catalog-info card">
          <div>
            <b>Cadastros dinâmicos</b>
            <p>Cirurgias e procedimentos ficam em cadastros independentes: cirurgia é o evento cirúrgico; procedimento é um item avulso ou vinculado ao pacote de uma cirurgia.</p>
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
            title="Cirurgias"
            description="Cadastre somente cirurgias. Este catálogo é separado dos procedimentos e cada lançamento conta como 1 cirurgia."
            placeholder="Ex.: Blefaroplastia"
            items={surgeryTypes}
            onAdd={v=>addCatalog('surgery_types',v)}
            onRemove={id=>deactivateCatalog('surgery_types',id)}
          />
          <Catalog
            title="Procedimentos"
            description="Cadastre somente procedimentos, como Botox, Laser CO₂, Bioestimulador e Profhilo. Eles nunca entram como cirurgia."
            placeholder="Ex.: Botox"
            items={procedureOptions}
            onAdd={v=>addCatalog('procedures',v)}
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
            <p>O sistema centraliza os lançamentos de exames, cirurgias e procedimentos da clínica. As informações ficam salvas no Supabase e podem ser acompanhadas nos Resultados e exportadas para Excel.</p>
          </div>
          <div className="manual-flow">
            <span>1. Entrar</span><b>→</b><span>2. Selecionar responsável</span><b>→</b><span>3. Lançar</span><b>→</b><span>4. Acompanhar</span>
          </div>
        </div>

        <div className="manual-grid">
          <article className="card manual-card">
            <div className="manual-number">01</div>
            <h3>Acesso ao sistema</h3>
            <p>Entre com o e-mail e a senha cadastrados. O acesso pode ser de <b>ADM</b> ou <b>Orientadora</b>.</p>
            <div className="manual-note">Os lançamentos são identificados automaticamente pelo usuário que está logado. Não é necessário escolher manualmente quem está lançando.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">02</div>
            <h3>Identificação do lançamento</h3>
            <p>O sistema registra automaticamente o <b>nome do usuário logado</b> em cada novo lançamento. O nome aparece no menu lateral junto ao e-mail da conta.</p>
            <div className="manual-note">No acesso ADM, o responsável é preenchido automaticamente como ADM.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">03</div>
            <h3>Lançar um exame</h3>
            <p>Acesse <b>Exames → Novo exame</b>. Informe a <b>data do agendamento</b> e, quando definida, a <b>data da realização</b>. Preencha paciente, médico, tipo de exame, status e observação. No Financeiro, registre cada pagamento como <b>Pago</b> ou <b>Não pago</b>; o sistema calcula saldo e situação automaticamente. Depois clique em <b>Salvar exame</b>.</p>
            <div className="manual-note">Use o campo Observação para registrar informações importantes que a equipe precise consultar depois.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">04</div>
            <h3>Lançar uma cirurgia</h3>
            <p>Acesse <b>Cirurgias → Nova cirurgia</b>. Informe a <b>data do agendamento</b> e a <b>data da realização</b>, paciente, cirurgião e cirurgia. Se fizer parte do pacote, vincule os procedimentos adicionais. No Financeiro, cada recebimento deve ser marcado como <b>Pago</b> ou <b>Não pago</b>; o sistema calcula recebido, saldo e situação automaticamente.</p>
            <div className="manual-note">Mantenha o status atualizado para que os indicadores mensais representem a situação real da cirurgia.</div><div className="manual-note">Nos lançamentos de Exames, Cirurgias e Procedimentos, data de realização, médico que indicou, profissional que realizou/fará, convênio ou Particular e situação do pagamento são campos obrigatórios. O valor é sugerido automaticamente pela tabela quando houver correspondência única e pode ser alterado manualmente.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">05</div>
            <h3>Lançar um procedimento</h3>
            <p>Acesse <b>Procedimentos → Novo procedimento</b>. Informe a <b>data do agendamento</b> e, quando definida, a <b>data da realização</b>, paciente, profissional, procedimento, convênio e status. No Financeiro, use <b>Pago</b> ou <b>Não pago</b> em cada pagamento. Procedimentos vinculados a uma cirurgia continuam identificados como parte do pacote.</p>
            <div className="manual-note">A data da realização é independente da data em que o procedimento entrou na agenda.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">06</div>
            <h3>Editar informações</h3>
            <p>Na tabela de Exames, Cirurgias ou Procedimentos, ou diretamente no <b>Relatório diário</b>, clique em <b>Editar</b>. O mesmo formulário completo do lançamento será aberto. É possível corrigir atendimento, convênio, valores, <b>Pago / Pago parcial / Não pago</b>, forma de pagamento, parcelas, documentos e repasses. Ao salvar, a alteração também atualiza a <b>Visão geral do Financeiro</b>.</p>
            <h3>Filtro de período</h3>
            <p>Exames, Cirurgias, Procedimentos, Resultados e Financeiro usam um único controle de período. Escolha <b>Mês</b> ou <b>Data</b> no próprio campo e depois informe o período desejado.</p>
            <h3>Validação dos lançamentos</h3>
            <p>Nos novos lançamentos de Exames, Cirurgias e Procedimentos, <b>Quem está lançando?</b> voltou a aparecer dentro do formulário. Se alguém tentar salvar sem preencher um campo obrigatório, o sistema exibe um aviso visual e destaca em vermelho os campos que precisam de atenção.</p>
            <h3>Relatório diário</h3>
            <p>No Relatório diário, <b>Tipo</b> e <b>Situação</b> aparecem em caixinhas selecionáveis. É possível marcar uma ou várias opções; as selecionadas ficam destacadas em verde. Também existe filtro específico pela <b>data de nascimento</b> do paciente.</p>
            <h3>Controle de Promissórias</h3>
            <p>No submenu <b>Financeiro → Controle de Promissórias</b>, o sistema reúne automaticamente os pagamentos registrados como Promissória. É possível filtrar por mês, dia, paciente, serviço e situação, acompanhar total, recebido, saldo e vencidos, abrir o lançamento original e exportar o relatório em Word ou PDF.</p>
          </article>

          <article className="card manual-card">
            <div className="manual-number">07</div>
            <h3>Excluir um lançamento</h3>
            <p>Clique em <b>Excluir</b> na linha desejada. O sistema solicitará uma confirmação antes de remover o registro.</p>
            <div className="manual-warning">A exclusão é definitiva. Confira o paciente e o procedimento antes de confirmar.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">08</div>
            <h3>Filtros</h3>
            <p>Use os filtros no topo das páginas para localizar informações por <b>mês, médico, exame ou procedimento</b>. Em Resultados também é possível filtrar por responsável e status.</p>
          </article>

          <article className="card manual-card">
            <div className="manual-number">09</div>
            <h3>Resultados</h3>
            <p>A página Resultados mostra pacientes únicos, volumes, taxas de realização, exames e cirurgias mais registrados, produção por médico e tipo, indicações por médico quando informadas, status, convênio e lançamentos por responsável.</p>
            <div className="manual-note">Os indicadores acompanham os filtros selecionados. Assim você pode analisar um mês, profissional, responsável ou status específico.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">10</div>
            <h3>Exportar resultados</h3>
            <p>Na página Resultados, aplique os filtros desejados e clique em <b>PDF completo</b>. O relatório mostra resumo executivo, indicadores, situação dos exames e cirurgias, pendências, cancelamentos, situação financeira, comparação com o mês anterior, gráficos de linha, gráficos de pizza e o detalhamento de cada tipo de exame/procedimento dentro de cada status.</p>
            <p>O PDF é paginado automaticamente para que todos os exames e cirurgias filtrados apareçam com <b>paciente, médico/cirurgião, procedimento, status, responsável e observações</b>. A exportação em Excel continua disponível separadamente para conferência dos dados.</p>
          </article>

          <article className="card manual-card">
            <div className="manual-number">11</div>
            <h3>Cadastros da clínica</h3>
            <p>Tanto no acesso <b>ADM</b> quanto no acesso <b>Orientadora</b>, abra <b>Cadastros</b> no menu lateral. Nessa página é possível adicionar ou remover <b>médicos, exames, cirurgias e procedimentos</b>. Os novos itens passam a aparecer automaticamente nos formulários de lançamento.</p><div className="manual-note">Itens removidos deixam de aparecer para novos lançamentos, mas os registros antigos permanecem preservados.</div>
          </article>

          <article className="card manual-card">
            <div className="manual-number">12</div>
            <h3>WhatsApp do paciente</h3>
            <p>O WhatsApp é opcional ao lançar ou editar um exame/cirurgia. Quando informado, fica salvo no cadastro do paciente e pode ser usado no Follow-up. No Follow-up, o WhatsApp fica oculto na visão principal para manter a tela limpa. Ao clicar em <b>Registrar follow-up</b> ou <b>Atualizar follow-up</b>, o contato aparece junto com o botão <b>Enviar WhatsApp</b>, que abre a conversa com uma mensagem pronta alterando automaticamente o nome do paciente.</p>
          </article>

          <article className="card manual-card">
            <div className="manual-number">13</div>
            <h3>Boas práticas</h3>
            <p>Confira os dados antes de salvar, mantenha os status atualizados, registre observações importantes e evite criar lançamentos duplicados.</p>
            <div className="manual-note">Ao finalizar o uso em um computador compartilhado, clique sempre em <b>Sair</b>.</div>
          </article>
        </div>

        <div className="card manual-help">
          <h3>Fluxo recomendado no dia a dia</h3>
          <div className="manual-checklist">
            <span>✓ Selecione seu nome antes de começar</span>
            <span>✓ Lance exames, cirurgias e procedimentos conforme forem agendados</span>
            <span>✓ Informe a data da realização quando o atendimento acontecer</span>
            <span>✓ Marque os pagamentos como Pago ou Não pago e revise os saldos</span>
            <span>✓ Exporte o Excel para conferência e análise</span>
          </div>
        </div>
      </section>}


      {documentPreview&&<div className="document-preview-overlay" onMouseDown={e=>{if(e.target===e.currentTarget)setDocumentPreview(null)}}>
        <div className="document-preview-card">
          <div className="document-preview-head"><div><b>{documentPreview.name}</b><small>Visualização do documento financeiro</small></div><div className="document-preview-actions"><button className="btn btn-light" onClick={()=>downloadInvoiceFile(documentPreview.path,documentPreview.name)}>Baixar</button><button className="icon-close" onClick={()=>setDocumentPreview(null)}>×</button></div></div>
          <div className="document-preview-body">{documentPreview.isImage?<img src={documentPreview.url} alt={documentPreview.name}/>:<iframe src={documentPreview.url} title={documentPreview.name}/>}</div>
        </div>
      </div>}
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
        grid-template-columns:repeat(5,minmax(0,1fr));
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

    <style>{`
      .patient-search-input{min-width:190px}
      .linked-procedures-cell{min-width:180px}
      .linked-procedures-cell>small{display:block;margin-bottom:6px;color:#74838c;font-size:8px;font-weight:800;letter-spacing:.05em;text-transform:uppercase}
      .linked-procedures-cell>div{display:flex;flex-wrap:wrap;gap:5px}
      .linked-procedure-chip{display:inline-flex;align-items:center;max-width:230px;padding:5px 8px;border:1px solid #d8e7eb;border-radius:999px;background:#f1f7f9;color:#315f72;font-size:9px;font-weight:750;line-height:1.2}
      .procedure-kpi{border-color:#dce9df!important;background:linear-gradient(180deg,#fff,#f5faf7)!important}
      .procedure-status-card{grid-column:1/-1}
      .results-highlights{grid-template-columns:repeat(4,minmax(0,1fr))!important}
      .results-operational-strip{grid-template-columns:repeat(5,minmax(0,1fr))}
      @media(max-width:1180px){.results-highlights{grid-template-columns:repeat(2,minmax(0,1fr))!important}.results-operational-strip{grid-template-columns:repeat(2,minmax(0,1fr))}.patient-search-input{min-width:150px}}
      @media(max-width:700px){.results-highlights,.results-operational-strip{grid-template-columns:1fr!important}.patient-search-input{width:100%;min-width:0}.linked-procedures-cell{min-width:150px}}
    `}</style>

    <Modal open={examModal} onClose={closeExamModal} title={editingExamId?'Editar exame':'Novo exame'} subtitle="Lançamento simplificado: atendimento, financeiro e repasses." onSave={editingExamId?updateExam:addExam} saveText={editingExamId?'Salvar alterações':'Salvar exame'}>
      <div className="form-grid simplified-form-grid" data-validation-scope="exam">
        {formErrors.exam?.message&&<div className="form-validation-banner full"><strong>Faltam informações obrigatórias</strong><span>{formErrors.exam.message} Os campos destacados em vermelho precisam ser preenchidos.</span></div>}
        <div className="form-section full">
          <div className="form-section-title"><span>1</span><div><b>DADOS DO ATENDIMENTO</b><small>Informações clínicas e operacionais principais.</small></div></div>
          <div className="form-section-grid">
            <div className={fieldError('exam','launcher')?'field-error':''}><Field label="Quem está lançando? *"><Select value={examForm.launcher||currentLaunchName||''} setValue={v=>setExamForm({...examForm,launcher:v})} options={launcherOptionsAll} first="Selecione"/></Field></div>
            <div className={fieldError('exam','date')?'field-error':''}><Field label="Data do agendamento *"><input required type="date" value={examForm.date} onChange={e=>setExamForm({...examForm,date:e.target.value})}/></Field></div>
            <div className={fieldError('exam','realization_date')?'field-error':''}><Field label="Data da realização *"><input required type="date" value={examForm.realization_date||''} onChange={e=>setExamForm({...examForm,realization_date:e.target.value})}/></Field></div>
            <div className={fieldError('exam','patient')?'field-error':''}><Field label="Paciente *"><input required value={examForm.patient} onChange={e=>setExamForm({...examForm,patient:e.target.value})}/></Field></div>
            <Field label="Data de nascimento"><input type="date" value={examForm.patient_birth_date||''} onChange={e=>setExamForm({...examForm,patient_birth_date:e.target.value})}/></Field>
            <div className={fieldError('exam','coverage_choice')?'field-error':''}><Field label="Convênio / Particular *"><Select value={examForm.coverage_choice||''} setValue={v=>setExamForm({...examForm,coverage_choice:v,insurance_id:v==='__PARTICULAR__'?'':v,payment_origin:v==='__PARTICULAR__'?'Particular':'Convênio',price_manual:false})} options={[{id:'__PARTICULAR__',name:'Particular'},...insurances]} first="Selecione"/></Field></div>
            <div className={fieldError('exam','eye')?'field-error':''}><Field label="Olho *"><Select value={examForm.eye||''} setValue={v=>setExamForm({...examForm,eye:v,price_manual:false})} options={['Não se aplica','Direito (OD)','Esquerdo (OE)','Ambos (AO)']} first="Selecione"/></Field></div>
            <div className={fieldError('exam','indicated_by_doctor_id')?'field-error':''}><Field label="Médico que indicou *"><Select value={examForm.indicated_by_doctor_id||''} setValue={v=>setExamForm({...examForm,indicated_by_doctor_id:v})} options={doctors} first="Selecione"/></Field></div>
            <div className={fieldError('exam','doctor_id')?'field-error':''}><Field label="Médico que realizou / realizará *"><Select value={examForm.doctor_id} setValue={v=>setExamForm({...examForm,doctor_id:v})} options={doctors} first="Selecione"/></Field></div>
            <Field label="Status"><Select value={examForm.status} setValue={v=>setExamForm({...examForm,status:v})} options={['Agendado','Pendente','Realizado','Cancelado']}/></Field>
            <div className={fieldError('exam','exam_type_ids')?'field-error full':'full'}><Field label={editingExamId ? "Exame *" : "Exames *"} full>
              <div className="exam-multi-select"><div className="exam-multi-head"><span>Selecione um ou vários exames</span><b>{examForm.exam_type_ids?.length||0} selecionado(s)</b></div><div className="exam-options">
                {examTypes.map(exam=>{const checked=examForm.exam_type_ids?.includes(exam.id);return <label key={exam.id} className={`exam-option ${checked?'selected':''}`}><input type="checkbox" checked={checked} onChange={()=>{const ids=examForm.exam_type_ids||[];setExamForm({...examForm,exam_type_ids:checked?ids.filter(id=>id!==exam.id):[...ids,exam.id],price_manual:false});}}/><span>{exam.name}</span></label>;})}
              </div></div>
            </Field></div>
          </div>
        </div>

        <div className="form-section finance-section full">
          <div className="form-section-title finance"><span>2</span><div><b>FINANCEIRO</b><small>Valor sugerido automaticamente pela tabela, com edição manual quando necessário. Pagamento é obrigatório.</small></div></div>
          <div className="form-section-grid finance-simple-grid">
            <Field label="Valor total (R$)" full><div className="auto-price-field"><input type="number" min="0" step="0.01" value={examForm.total_amount} onChange={e=>setExamForm({...examForm,total_amount:e.target.value,price_manual:true})}/><small>{examForm.price_manual?'Valor ajustado manualmente.':examAutoPrice?`Prévia automática pela tabela: ${examAutoPrice.labels.join(' + ')}.`:'Selecione o atendimento. Antes de escolher convênio/Particular, a prévia usa o valor Particular e recalcula automaticamente depois.'} Você pode editar este valor.</small></div></Field>
            <Field label="Desconto (R$)"><input type="number" min="0" step="0.01" value={examForm.discount_amount||''} onChange={e=>setExamForm({...examForm,discount_amount:e.target.value})}/></Field>
            <div className={fieldError('exam','payment_confirmed')?'field-error full':'full'}><FinanceQuickSummary form={examForm} setForm={setExamForm}/></div>
            <PaymentBreakdownEditor form={examForm} setForm={setExamForm} insurances={insurances}/>
            <Field label="Nota fiscal / comprovante nº"><input value={examForm.invoice_number||''} onChange={e=>setExamForm({...examForm,invoice_number:e.target.value})} placeholder="Número ou referência"/></Field>
            <Field label="Data de emissão"><input type="date" value={examForm.invoice_issued_date||''} onChange={e=>setExamForm({...examForm,invoice_issued_date:e.target.value})}/></Field>
            <FinancialDocumentsField form={examForm} setForm={setExamForm} onOpen={openInvoiceFile} onDownload={downloadInvoiceFile}/>
            <Field label="Observação financeira / promissória" full><textarea rows="2" value={examForm.financial_notes} onChange={e=>setExamForm({...examForm,financial_notes:e.target.value})}/></Field>
          </div>
        </div>

        <div className="form-section repasses-section full"><div className="form-section-title repasses"><span>3</span><div><b>REPASSES</b><small>Honorários e custos de profissionais vinculados a este atendimento.</small></div></div><RepassesEditor form={examForm} setForm={setExamForm}/></div>
        <Field label="Observação geral" full><textarea value={examForm.obs} onChange={e=>setExamForm({...examForm,obs:e.target.value})} rows="3"/></Field>
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


    <style>{`
      /* V28 — interações sutis, sem alterar a identidade visual */
      @keyframes ocFadeUp{
        from{opacity:0;transform:translateY(8px)}
        to{opacity:1;transform:translateY(0)}
      }
      @keyframes ocModalIn{
        from{opacity:0;transform:translateY(10px) scale(.992)}
        to{opacity:1;transform:translateY(0) scale(1)}
      }
      @keyframes ocSoftPulse{
        0%,100%{box-shadow:0 0 0 0 rgba(49,95,114,0)}
        50%{box-shadow:0 0 0 5px rgba(49,95,114,.06)}
      }
      .main > section{
        animation:ocFadeUp .28s ease both;
      }
      .modal-card{
        animation:ocModalIn .22s ease-out both;
      }
      .card,.kpi,.analysis-card,.form-section{
        transition:transform .18s ease,box-shadow .18s ease,border-color .18s ease;
      }
      .card:hover,.kpi:hover,.analysis-card:hover{
        transform:translateY(-1px);
      }
      .btn{
        transition:transform .15s ease,box-shadow .15s ease,filter .15s ease,background-color .15s ease;
      }
      .btn:hover{
        transform:translateY(-1px);
      }
      .btn:active{
        transform:translateY(0) scale(.985);
      }
      tbody tr{
        transition:background-color .16s ease,transform .16s ease;
      }
      tbody tr:hover{
        transform:translateX(1px);
      }
      .payment-light:hover .payment-light-dot{
        animation:ocSoftPulse 1.2s ease-in-out infinite;
      }
      .document-chip,.payment-breakdown-row,.repasse-row{
        transition:transform .16s ease,border-color .16s ease,box-shadow .16s ease;
      }
      .document-chip:hover,.payment-breakdown-row:hover,.repasse-row:hover{
        transform:translateY(-1px);
        box-shadow:0 5px 14px rgba(36,55,70,.055);
      }
      .form-section{
        scroll-margin-top:18px;
      }
      @media (prefers-reduced-motion:reduce){
        *,*::before,*::after{
          animation-duration:.001ms!important;
          animation-iteration-count:1!important;
          transition-duration:.001ms!important;
          scroll-behavior:auto!important;
        }
      }
    `}</style>


    <style>{`
      /* V31 — relatórios com respiro e gestão de notas fiscais */
      .daily-report-filters{align-items:end;grid-template-columns:repeat(4,minmax(150px,1fr)) auto}
      .daily-report-sheet{padding:28px!important}
      .daily-report-title{padding-bottom:18px;margin-bottom:6px;border-bottom:1px solid #dce5e8}
      .daily-report-title b{display:block;font-size:17px;color:#315f72;line-height:1.45}
      .daily-report-title small{display:block;margin-top:5px;color:#74838c}
      .daily-report-item{padding:21px 4px 22px!important;margin:0!important;border-bottom:1px solid #dce5e8!important;line-height:1.65!important;page-break-inside:avoid}
      .daily-report-item+.daily-report-item{margin-top:5px!important}
      .daily-report-item-head{margin-bottom:12px!important;gap:18px}
      .daily-report-item-actions{display:flex;align-items:center;gap:9px;flex-wrap:wrap}
      .daily-edit-btn{padding:7px 11px!important;min-height:32px}
      .daily-report-item>div:not(.daily-report-item-head){margin:5px 0!important}
      .daily-report-item b{margin-right:4px}
      .invoice-management-card{margin-top:22px}
      .invoice-management-title{align-items:center}
      .invoice-management-filters{display:grid;grid-template-columns:150px 150px minmax(220px,1.5fr) minmax(160px,1fr) minmax(180px,1fr) auto;gap:10px;align-items:end;margin:16px 0 18px;padding:14px;border:1px solid #e1e9ec;border-radius:14px;background:#f9fbfc}
      .invoice-management-filters label{display:block;margin-bottom:6px;color:#647781;font-size:10px;font-weight:800}
      .invoice-clear-btn{height:40px}
      .invoice-management-table td{vertical-align:top;padding-top:13px;padding-bottom:13px}
      .invoice-patient-date{display:block;margin-top:4px;color:#87959a;font-size:9px}
      .invoice-document-list{display:grid;gap:7px;min-width:250px}
      .invoice-document-list>div{display:flex;align-items:center;gap:6px;flex-wrap:wrap;padding:7px 8px;border:1px solid #e4eaed;border-radius:9px;background:#fbfcfd}
      .invoice-document-list span{flex:1;min-width:120px;font-size:9.5px;color:#53636e;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
      .invoice-document-list .btn{padding:5px 7px;font-size:9px}
      @media(max-width:1150px){.invoice-management-filters{grid-template-columns:repeat(3,minmax(0,1fr))}.daily-report-filters{grid-template-columns:repeat(2,minmax(0,1fr))}}
      @media(max-width:700px){.invoice-management-filters,.daily-report-filters{grid-template-columns:1fr}.daily-report-sheet{padding:18px!important}.invoice-document-list{min-width:200px}}
      @media print{.daily-report-item{padding:18px 0!important;page-break-inside:avoid}.daily-report-sheet{box-shadow:none!important;border:0!important}}
    `}</style>

    <Modal open={surgeryModal} onClose={closeSurgeryModal} title={editingSurgeryId?'Editar cirurgia':'Nova cirurgia'} subtitle="Agendamento e data de realização separados, com financeiro e repasses." onSave={editingSurgeryId?updateSurgery:addSurgery} saveText={editingSurgeryId?'Salvar alterações':'Salvar cirurgia'}>
      <div className="form-grid simplified-form-grid" data-validation-scope="surgery">
        {formErrors.surgery?.message&&<div className="form-validation-banner full"><strong>Faltam informações obrigatórias</strong><span>{formErrors.surgery.message} Os campos destacados em vermelho precisam ser preenchidos.</span></div>}
        <div className="form-section full">
          <div className="form-section-title"><span>1</span><div><b>DADOS DA CIRURGIA</b><small>Informações principais para identificação e acompanhamento.</small></div></div>
          <div className="form-section-grid">
            <div className={fieldError('surgery','launcher')?'field-error':''}><Field label="Quem está lançando? *"><Select value={sForm.launcher||currentLaunchName||''} setValue={v=>setSForm({...sForm,launcher:v})} options={launcherOptionsAll} first="Selecione"/></Field></div>
            <div className={fieldError('surgery','date')?'field-error':''}><Field label="Data do agendamento *"><input required type="date" value={sForm.date} onChange={e=>setSForm({...sForm,date:e.target.value})}/></Field></div>
            <div className={fieldError('surgery','scheduled_on')?'field-error':''}><Field label="Data da realização *"><input required type="date" value={sForm.scheduled_on||''} onChange={e=>setSForm({...sForm,scheduled_on:e.target.value})}/></Field></div>
            <div className={fieldError('surgery','patient')?'field-error':''}><Field label="Paciente *"><input required value={sForm.patient} onChange={e=>setSForm({...sForm,patient:e.target.value})}/></Field></div>
            <Field label="Data de nascimento"><input type="date" value={sForm.patient_birth_date||''} onChange={e=>setSForm({...sForm,patient_birth_date:e.target.value})}/></Field>
            <div className={fieldError('surgery','coverage_choice')?'field-error':''}><Field label="Convênio / Particular *"><Select value={sForm.coverage_choice||''} setValue={v=>setSForm({...sForm,coverage_choice:v,insurance_id:v==='__PARTICULAR__'?'':v,payment_origin:v==='__PARTICULAR__'?'Particular':'Convênio',price_manual:false})} options={[{id:'__PARTICULAR__',name:'Particular'},...insurances]} first="Selecione"/></Field></div>
            <Field label="Cirurgia *"><Select value={sForm.surgery_type_id||''} setValue={v=>setSForm({...sForm,surgery_type_id:v,price_manual:false})} options={surgeryTypes} first="Selecione a cirurgia"/></Field>
            <Field label="LIO / lente (quando aplicável)"><input value={sForm.lio_description||''} onChange={e=>setSForm({...sForm,lio_description:e.target.value,price_manual:false})} placeholder="Ex.: Asférica / Trifocal / Mono Tórica"/></Field>
            <div className={fieldError('surgery','eye')?'field-error':''}><Field label="Olho *"><Select value={sForm.eye||''} setValue={v=>setSForm({...sForm,eye:v,price_manual:false})} options={['Não se aplica','Direito (OD)','Esquerdo (OE)','Ambos (AO)']} first="Selecione"/></Field></div>
            <Field label="Status"><Select value={sForm.status} setValue={v=>setSForm({...sForm,status:v})} options={['Solicitação','Pendente','Autorizada','Agendada','Realizada','Cancelada']}/></Field>
            <div className={fieldError('surgery','indicated_by_doctor_id')?'field-error':''}><Field label="Médico que indicou *"><Select value={sForm.indicated_by_doctor_id||''} setValue={v=>setSForm({...sForm,indicated_by_doctor_id:v})} options={doctors} first="Selecione"/></Field></div>
            <Field label="Cirurgião que realizou / realizará *"><Select value={sForm.doctor_id} setValue={v=>setSForm({...sForm,doctor_id:v})} options={doctors} first="Selecione"/></Field>
            <Field label="Procedimentos vinculados (opcional)" full>
              <div className="procedure-multi-select"><div className="procedure-multi-head"><span>Selecione apenas os procedimentos que pertencem a este pacote</span><b>{sForm.procedure_ids?.length||0} selecionado(s)</b></div><div className="procedure-options">
                {procedureOptions.map(procedure=>{const checked=sForm.procedure_ids?.includes(procedure.id);return <label key={procedure.id} className={`procedure-option ${checked?'selected':''}`}><input type="checkbox" checked={checked} onChange={()=>{const ids=sForm.procedure_ids||[];setSForm({...sForm,procedure_ids:checked?ids.filter(id=>id!==procedure.id):[...ids,procedure.id],price_manual:false});}}/><span>{procedure.name}</span></label>;})}
              </div></div>
            </Field>
          </div>
        </div>

        <div className="form-section finance-section full">
          <div className="form-section-title finance"><span>2</span><div><b>FINANCEIRO</b><small>Valor do pacote sugerido automaticamente pela tabela, incluindo adicionais dos procedimentos quando houver. O valor continua editável.</small></div></div>
          <div className="form-section-grid finance-simple-grid">
            <Field label="Valor total (R$)" full><div className="auto-price-field"><input type="number" min="0" step="0.01" value={sForm.total_amount} onChange={e=>setSForm({...sForm,total_amount:e.target.value,price_manual:true})} placeholder="0,00"/><small>{sForm.price_manual?'Valor ajustado manualmente.':surgeryAutoPrice?`Prévia automática pela tabela: ${surgeryAutoPrice.labels.join(' + ')}${(sForm.procedure_ids||[]).some(id=>sForm.procedure_finance?.[id]?.included===false)?' + adicionais dos procedimentos':''}.`:'Selecione a cirurgia. A prévia começa pela tabela Particular e recalcula ao escolher convênio, olho e lente quando aplicável.'} Você pode editar este valor.</small></div></Field>
            <Field label="Desconto (R$)"><input type="number" min="0" step="0.01" value={sForm.discount_amount||''} onChange={e=>setSForm({...sForm,discount_amount:e.target.value})}/></Field>
            <div className={fieldError('surgery','payment_confirmed')?'field-error full':'full'}><FinanceQuickSummary form={sForm} setForm={setSForm}/></div>
            <PaymentBreakdownEditor form={sForm} setForm={setSForm} insurances={insurances}/>
            {sForm.procedure_ids?.length>0&&<div className="package-finance-box full"><div className="package-finance-head"><b>Procedimentos do pacote</b><span>Por padrão ficam incluídos no valor total. Informe adicional somente quando houver cobrança separada.</span></div>{sForm.procedure_ids.map(id=>{const proc=procedureOptions.find(p=>p.id===id);const cfg=sForm.procedure_finance?.[id]||{included:true,amount:''};const autoProc=sForm.coverage_choice?automaticPriceForItem({kind:'procedure',name:proc?.name||'',origin:explicitOrigin(sForm),insuranceName:insuranceNameById(sForm.insurance_id),eye:sForm.eye,associatedWithSurgery:true}):null;return <div className="package-finance-row" key={id}><span>{proc?.name||'Procedimento'}</span><label className="package-check"><input type="checkbox" checked={cfg.included!==false} onChange={e=>setSForm({...sForm,price_manual:false,procedure_finance:{...(sForm.procedure_finance||{}),[id]:{...cfg,included:e.target.checked,amount:e.target.checked?(cfg.amount||''):((cfg.amount||autoProc?.amount)?String(cfg.amount||autoProc.amount):'')}}})}/><b>Incluído no pacote</b></label>{cfg.included===false?<div><input className="package-extra-input" type="number" min="0" step="0.01" placeholder="Valor adicional" value={cfg.amount||''} onChange={e=>setSForm({...sForm,price_manual:false,procedure_finance:{...(sForm.procedure_finance||{}),[id]:{...cfg,amount:e.target.value}}})}/>{autoProc&&<small className="package-auto-price">Tabela: {moneyBR(autoProc.amount)}</small>}</div>:<span className="package-included-label">Incluído no valor do pacote</span>}</div>})}</div>}
            <Field label="Nota fiscal / comprovante nº"><input value={sForm.invoice_number||''} onChange={e=>setSForm({...sForm,invoice_number:e.target.value})} placeholder="Número ou referência"/></Field>
            <Field label="Data de emissão"><input type="date" value={sForm.invoice_issued_date||''} onChange={e=>setSForm({...sForm,invoice_issued_date:e.target.value})}/></Field>
            <FinancialDocumentsField form={sForm} setForm={setSForm} onOpen={openInvoiceFile} onDownload={downloadInvoiceFile}/>
            <Field label="Observação financeira / promissória" full><textarea rows="2" value={sForm.financial_notes} onChange={e=>setSForm({...sForm,financial_notes:e.target.value})}/></Field>
          </div>
        </div>

        <div className="form-section repasses-section full"><div className="form-section-title repasses"><span>3</span><div><b>REPASSES</b><small>Honorários médicos, anestesista, auxiliar e outros.</small></div></div><RepassesEditor form={sForm} setForm={setSForm}/></div>
        <Field label="Observação geral" full><textarea value={sForm.obs} onChange={e=>setSForm({...sForm,obs:e.target.value})} rows="3"/></Field>
      </div>
    </Modal>

    <style>{`
      .classification-review{display:block;margin-top:4px;color:#9a6b22;font-size:8px;font-weight:800;text-transform:uppercase;letter-spacing:.04em}
      .procedure-origin{display:inline-flex;align-items:center;padding:5px 8px;border-radius:999px;font-size:9px;font-weight:800;white-space:nowrap}
      .procedure-origin.linked{background:#edf6f0;color:#417257;border:1px solid #d5e8db}
      .procedure-origin.standalone{background:#f2f5f6;color:#61727a;border:1px solid #e2e8ea}
    `}</style>

    <Modal open={procedureModal} onClose={closeProcedureModal} title={editingProcedureId?'Editar procedimento':'Novo procedimento'} subtitle="Lançamento simplificado: atendimento, financeiro e repasses." onSave={editingProcedureId?updateProcedure:addProcedure} saveText={editingProcedureId?'Salvar alterações':'Salvar procedimento'}>
      <div className="form-grid simplified-form-grid" data-validation-scope="procedure">
        {formErrors.procedure?.message&&<div className="form-validation-banner full"><strong>Faltam informações obrigatórias</strong><span>{formErrors.procedure.message} Os campos destacados em vermelho precisam ser preenchidos.</span></div>}
        <div className="form-section full">
          <div className="form-section-title"><span>1</span><div><b>DADOS DO PROCEDIMENTO</b><small>Informações principais para identificação e acompanhamento.</small></div></div>
          <div className="form-section-grid">
            <div className={fieldError('procedure','launcher')?'field-error':''}><Field label="Quem está lançando? *"><Select value={pForm.launcher||currentLaunchName||''} setValue={v=>setPForm({...pForm,launcher:v})} options={launcherOptionsAll} first="Selecione"/></Field></div>
            <div className={fieldError('procedure','date')?'field-error':''}><Field label="Data do agendamento *"><input required type="date" value={pForm.date} onChange={e=>setPForm({...pForm,date:e.target.value})}/></Field></div>
            <div className={fieldError('procedure','realization_date')?'field-error':''}><Field label="Data da realização *"><input required type="date" value={pForm.realization_date||''} onChange={e=>setPForm({...pForm,realization_date:e.target.value})}/></Field></div>
            <div className={fieldError('procedure','patient')?'field-error':''}><Field label="Paciente *"><input required value={pForm.patient} onChange={e=>setPForm({...pForm,patient:e.target.value})}/></Field></div>
            <Field label="Data de nascimento"><input type="date" value={pForm.patient_birth_date||''} onChange={e=>setPForm({...pForm,patient_birth_date:e.target.value})}/></Field>
            <div className={fieldError('procedure','coverage_choice')?'field-error':''}><Field label="Convênio / Particular *"><Select value={pForm.coverage_choice||''} setValue={v=>setPForm({...pForm,coverage_choice:v,insurance_id:v==='__PARTICULAR__'?'':v,payment_origin:v==='__PARTICULAR__'?'Particular':'Convênio',price_manual:false})} options={[{id:'__PARTICULAR__',name:'Particular'},...insurances]} first="Selecione"/></Field></div>
            <div className={fieldError('procedure','indicated_by_doctor_id')?'field-error':''}><Field label="Médico que indicou *"><Select value={pForm.indicated_by_doctor_id||''} setValue={v=>setPForm({...pForm,indicated_by_doctor_id:v})} options={doctors} first="Selecione"/></Field></div>
            <div className={fieldError('procedure','doctor_id')?'field-error':''}><Field label="Médico que realizou / realizará *"><Select value={pForm.doctor_id} setValue={v=>setPForm({...pForm,doctor_id:v})} options={doctors} first="Selecione"/></Field></div>
            <div className={fieldError('procedure','eye')?'field-error':''}><Field label="Olho *"><Select value={pForm.eye||''} setValue={v=>setPForm({...pForm,eye:v,price_manual:false})} options={['Não se aplica','Direito (OD)','Esquerdo (OE)','Ambos (AO)']} first="Selecione"/></Field></div>
            <Field label="Status"><Select value={pForm.status} setValue={v=>setPForm({...pForm,status:v})} options={['Solicitação','Pendente','Autorizada','Agendada','Realizada','Cancelada']}/></Field>
            <div className={fieldError('procedure','procedure_ids')?'field-error full':'full'}><Field label="Procedimentos *" full><div className="procedure-multi-select"><div className="procedure-multi-head"><span>Um ou vários procedimentos podem permanecer no mesmo lançamento</span><b>{pForm.procedure_ids?.length||0} selecionado(s)</b></div><div className="procedure-options">{procedureOptions.map(procedure=>{const checked=pForm.procedure_ids?.includes(procedure.id);return <label key={procedure.id} className={`procedure-option ${checked?'selected':''}`}><input type="checkbox" checked={checked} onChange={()=>{const ids=pForm.procedure_ids||[];setPForm({...pForm,procedure_ids:checked?ids.filter(id=>id!==procedure.id):[...ids,procedure.id],price_manual:false});}}/><span>{procedure.name}</span></label>;})}</div></div></Field></div>
          </div>
        </div>

        <div className="form-section finance-section full">
          <div className="form-section-title finance"><span>2</span><div><b>FINANCEIRO</b><small>Valor sugerido automaticamente pela tabela. Você pode ajustar manualmente quando necessário.</small></div></div>
          <div className="form-section-grid finance-simple-grid">
            <Field label="Valor total (R$)" full><div className="auto-price-field"><input type="number" min="0" step="0.01" value={pForm.total_amount} onChange={e=>setPForm({...pForm,total_amount:e.target.value,price_manual:true})}/><small>{pForm.price_manual?'Valor ajustado manualmente.':procedureAutoPrice?`Prévia automática pela tabela: ${procedureAutoPrice.labels.join(' + ')}.`:'Selecione o procedimento. A prévia aparece imediatamente e é recalculada ao escolher convênio/Particular quando houver diferença de tabela.'} Você pode editar este valor.</small></div></Field>
            <Field label="Desconto (R$)"><input type="number" min="0" step="0.01" value={pForm.discount_amount||''} onChange={e=>setPForm({...pForm,discount_amount:e.target.value})}/></Field>
            <div className={fieldError('procedure','payment_confirmed')?'field-error full':'full'}><FinanceQuickSummary form={pForm} setForm={setPForm}/></div>
            <PaymentBreakdownEditor form={pForm} setForm={setPForm} insurances={insurances}/>
            <Field label="Nota fiscal / comprovante nº"><input value={pForm.invoice_number||''} onChange={e=>setPForm({...pForm,invoice_number:e.target.value})} placeholder="Número ou referência"/></Field>
            <Field label="Data de emissão"><input type="date" value={pForm.invoice_issued_date||''} onChange={e=>setPForm({...pForm,invoice_issued_date:e.target.value})}/></Field>
            <FinancialDocumentsField form={pForm} setForm={setPForm} onOpen={openInvoiceFile} onDownload={downloadInvoiceFile}/>
            <Field label="Observação financeira / promissória" full><textarea rows="2" value={pForm.financial_notes} onChange={e=>setPForm({...pForm,financial_notes:e.target.value})}/></Field>
          </div>
        </div>

        <div className="form-section repasses-section full"><div className="form-section-title repasses"><span>3</span><div><b>REPASSES</b><small>Honorários e profissionais vinculados ao procedimento.</small></div></div><RepassesEditor form={pForm} setForm={setPForm}/></div>
        <Field label="Observação geral" full><textarea value={pForm.obs} onChange={e=>setPForm({...pForm,obs:e.target.value})} rows="3"/></Field>
      </div>
    </Modal>

    <style>{`
      .simplified-form-grid{display:block!important}
      .form-section{margin-bottom:18px;padding:18px;border:1px solid var(--border);border-radius:16px;background:#fff}
      .form-section.finance-section{border-color:#d7e7ec;background:linear-gradient(180deg,#fbfdfe,#f5f9fa)}
      .form-section.repasses-section{border-color:#e5e5dc;background:linear-gradient(180deg,#fff,#fbfbf7)}
      .form-section-title{display:flex;align-items:center;gap:11px;margin-bottom:16px;padding-bottom:12px;border-bottom:1px solid #edf1f3}
      .form-section-title>span{width:30px;height:30px;display:grid;place-items:center;flex:0 0 30px;border-radius:9px;background:var(--primary-soft);color:var(--primary);font-size:11px;font-weight:900}
      .form-section-title.finance>span{background:#eaf5ee;color:#3f7557}.form-section-title.repasses>span{background:#f4f0e8;color:#80663d}
      .form-section-title b{display:block;color:var(--text);font-size:13px;letter-spacing:.035em}.form-section-title small{display:block;margin-top:3px;color:var(--muted);font-size:10.5px}
      .form-section-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:13px}.form-section-grid>.full{grid-column:1/-1}
      .finance-simple-grid{grid-template-columns:repeat(4,minmax(0,1fr))}
      .finance-quick-summary{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:9px;margin:2px 0 4px}.finance-quick-summary>div{padding:12px;border:1px solid var(--border);border-radius:11px;background:#fff}.finance-quick-summary span{display:block;color:var(--muted);font-size:9px;font-weight:850;text-transform:uppercase}.finance-quick-summary b{display:block;margin-top:5px;color:var(--primary);font-size:16px}.finance-summary-status b{display:flex;align-items:center;gap:6px}.finance-summary-status i{width:9px;height:9px;border-radius:50%}.finance-summary-status.paid i{background:var(--green)}.finance-summary-status.partial i{background:var(--yellow)}.finance-summary-status.unpaid i{background:var(--red)}
      .payment-breakdown-row.simplified{grid-template-columns:repeat(4,minmax(0,1fr))}.payment-breakdown-row.simplified .payment-row-remove{align-self:end}
      .repasses-box{padding:0;border:0;background:transparent}.repasses-head{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin-bottom:11px}.repasses-head span{display:block;color:#80663d;font-size:9px;font-weight:900;letter-spacing:.08em}.repasses-head b{display:block;margin-top:3px;color:var(--text);font-size:13px}.repasses-head small{display:block;margin-top:3px;color:var(--muted);font-size:10px}.repasses-list{display:grid;gap:9px}.repasse-row{display:grid;grid-template-columns:2fr 1.3fr 1fr auto;gap:9px;align-items:end;padding:11px;border:1px solid var(--border);border-radius:11px;background:#fff}.repasse-total{margin-top:10px;text-align:right;color:var(--muted);font-size:11px}.repasse-total b{color:var(--primary);font-size:13px}
      .financial-documents{display:grid;gap:10px}.documents-upload-zone{display:flex;align-items:center;justify-content:space-between;gap:14px;padding:14px;border:1px dashed #cddde3;border-radius:13px;background:#fff;transition:border-color .18s ease,box-shadow .18s ease,transform .18s ease}.documents-upload-zone:hover{border-color:#8eb1bd;box-shadow:0 6px 18px rgba(49,95,114,.06)}.documents-upload-zone b{display:block;color:var(--text);font-size:12px}.documents-upload-zone small{display:block;margin-top:4px;color:var(--muted);font-size:10.5px}.document-picker{position:relative;overflow:hidden;white-space:nowrap}.document-picker input{position:absolute;inset:0;opacity:0;cursor:pointer}.documents-list{display:grid;gap:8px}.document-chip{display:grid;grid-template-columns:auto minmax(0,1fr) auto;align-items:center;gap:10px;padding:10px 11px;border:1px solid var(--border);border-radius:11px;background:#fff;animation:ocDocIn .18s ease-out}.document-chip.pending{border-style:dashed;background:#fbfdfe}.document-icon{width:30px;height:30px;display:grid;place-items:center;border-radius:9px;background:var(--primary-soft);color:var(--primary);font-weight:900}.document-meta{min-width:0}.document-meta b{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--text);font-size:11.5px}.document-meta small{display:block;margin-top:3px;color:var(--muted);font-size:9.5px}.document-actions{display:flex;gap:6px}.document-actions .btn,.document-picker{display:inline-flex;align-items:center;justify-content:center;gap:6px}.documents-empty{padding:12px;border:1px dashed var(--border);border-radius:10px;color:var(--muted);font-size:10.5px;text-align:center;background:#fff}.table-document-actions{display:flex;align-items:center;gap:7px;flex-wrap:wrap}.table-document-item{display:inline-flex;gap:5px}.table-document-actions small{color:var(--muted);font-size:9px}.form-section{transition:box-shadow .2s ease,border-color .2s ease,transform .2s ease;animation:ocSectionIn .22s ease-out both}.form-section:hover{box-shadow:0 7px 20px rgba(36,55,70,.045)}.financial-status-dot,.finance-status-interactive{transition:transform .16s ease,box-shadow .16s ease}.finance-status-interactive:hover{transform:translateY(-1px)}@keyframes ocSectionIn{from{opacity:.65;transform:translateY(4px)}to{opacity:1;transform:none}}@keyframes ocDocIn{from{opacity:0;transform:translateY(-3px)}to{opacity:1;transform:none}}@media(prefers-reduced-motion:reduce){.form-section,.document-chip{animation:none!important}.form-section,.documents-upload-zone,.financial-status-dot,.finance-status-interactive{transition:none!important}}
      @media(max-width:1050px){.form-section-grid,.finance-simple-grid,.payment-breakdown-row.simplified{grid-template-columns:repeat(2,minmax(0,1fr))}.finance-quick-summary{grid-template-columns:repeat(2,minmax(0,1fr))}.repasse-row{grid-template-columns:1fr 1fr}}
      @media(max-width:650px){.form-section-grid,.finance-simple-grid,.payment-breakdown-row.simplified,.finance-quick-summary,.repasse-row{grid-template-columns:1fr}.form-section{padding:14px}.repasses-head,.documents-upload-zone{flex-direction:column;align-items:stretch}.document-chip{grid-template-columns:auto minmax(0,1fr)}.document-actions{grid-column:1/-1}.document-actions .btn{flex:1}}
    `}</style>

    <style>{`
      .standard-price-field{grid-column:span 2;padding:12px;border:1px solid #dce8ec;border-radius:12px;background:#f8fbfc}
      .standard-price-field label{color:var(--primary);font-weight:900}
      .standard-price-field small{display:block;margin-top:7px;color:var(--muted);font-size:10.5px;line-height:1.4}
      .finance-page{max-width:1550px}
      .finance-filter-card{margin-bottom:18px}
      .finance-filters{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px}
      .finance-kpis{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;margin-bottom:18px}
      .finance-kpi{min-height:126px;padding:20px;border:1px solid var(--border);border-radius:17px;background:#fff;box-shadow:var(--shadow-sm);position:relative;overflow:hidden}
      .finance-kpi::after{content:"";position:absolute;width:90px;height:90px;right:-30px;top:-28px;border-radius:50%;background:var(--primary-soft);opacity:.7}
      .finance-kpi span,.finance-kpi small,.finance-kpi strong{position:relative;z-index:1;display:block}
      .finance-kpi span{color:var(--muted);font-size:10px;font-weight:900;text-transform:uppercase;letter-spacing:.055em}
      .finance-kpi strong{margin:10px 0 7px;color:var(--primary);font-size:28px;letter-spacing:-.04em}
      .finance-kpi small{color:var(--muted);font-size:11px}
      .finance-kpi.success strong{color:var(--green)} .finance-kpi.success::after{background:var(--green-soft)}
      .finance-kpi.warning strong{color:var(--yellow)} .finance-kpi.warning::after{background:var(--yellow-soft)}
      .finance-main-grid{display:grid;grid-template-columns:minmax(0,2.1fr) minmax(280px,.9fr);gap:15px;margin-bottom:15px}
      .finance-line-card{padding:20px;min-height:390px}.finance-line-svg{width:100%;height:auto;margin-top:8px}.finance-axis{fill:#74838c;font-size:10px}
      .finance-chart-head{display:flex;align-items:flex-start;justify-content:space-between;gap:18px}.finance-chart-head span{display:block;color:var(--primary-2);font-size:9px;font-weight:900;letter-spacing:.1em}.finance-chart-head h3{margin:5px 0 3px;font-size:17px;color:var(--text)}.finance-chart-head p{margin:0;color:var(--muted);font-size:11px}
      .finance-line-legend{display:flex;gap:12px;flex-wrap:wrap;justify-content:flex-end}.finance-line-legend span{display:flex;align-items:center;gap:6px;color:#667880;font-size:10px;letter-spacing:0}.finance-line-legend i{width:18px;height:4px;border-radius:999px}
      .finance-side-summary{padding:20px}.finance-summary-list{display:grid;gap:10px;margin-top:18px}.finance-summary-list>div{padding:14px;border:1px solid var(--border);border-radius:12px;background:#fafcfd}.finance-summary-list span{display:block;color:var(--muted);font-size:10px}.finance-summary-list b{display:block;margin-top:5px;color:var(--primary);font-size:20px}
      .finance-analysis-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:15px;margin-bottom:18px}.finance-ranking-card{min-height:330px;padding:20px}.finance-ranking-list{display:grid;gap:14px;margin-top:20px}.finance-ranking-row{display:grid;grid-template-columns:minmax(125px,180px) 1fr minmax(95px,125px);gap:10px;align-items:center}.finance-ranking-label{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#52646f;font-size:11px}.finance-ranking-track{height:9px;background:#edf1f3;border-radius:999px;overflow:hidden}.finance-ranking-track i{display:block;height:100%;border-radius:999px;background:linear-gradient(90deg,var(--primary),#78a5b1)}.finance-ranking-row b{color:var(--primary);font-size:11px;text-align:right}
      .finance-detail-card{margin-top:18px}.finance-detail-subtitle{margin:4px 0 0;color:var(--muted);font-size:12px}.finance-table{min-width:1450px}.money-cell{font-weight:800;color:var(--primary)}.money-cell.received{color:var(--green)}.money-cell.pending{color:var(--yellow)}
      .finance-status{display:inline-flex;padding:6px 9px;border-radius:999px;font-size:9px;font-weight:850;background:#f1f4f5;color:#65757d;white-space:nowrap}.finance-status.finance-recebido{background:var(--green-soft);color:var(--green)}.finance-status.finance-parcial,.finance-status.finance-pendente,.finance-status.finance-aguardando-convênio{background:var(--yellow-soft);color:var(--yellow)}.finance-status.finance-glosado{background:var(--red-soft);color:var(--red)}.finance-status.finance-faturado-ao-convênio{background:var(--primary-soft);color:var(--primary)}
      .finance-form-section{padding:18px;margin-top:4px;border:1px solid #dbe7eb;border-radius:14px;background:linear-gradient(135deg,#fbfdfe,#f3f8f9)}.finance-form-title{margin-bottom:14px}.finance-form-title span{display:block;color:var(--primary-2);font-size:9px;font-weight:900;letter-spacing:.12em}.finance-form-title b{display:block;margin-top:4px;color:var(--text);font-size:16px}.finance-form-title small{display:block;margin-top:4px;color:var(--muted);font-size:11px}.finance-form-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}.finance-form-grid>.full{grid-column:1/-1}
      .package-finance-box{margin-top:16px;padding-top:15px;border-top:1px solid #dfeaec}.package-finance-head{display:flex;justify-content:space-between;gap:12px;margin-bottom:9px}.package-finance-head b{font-size:12px;color:var(--text)}.package-finance-head span{font-size:10px;color:var(--muted)}.package-finance-row{display:grid;grid-template-columns:minmax(180px,1fr) 180px 180px;align-items:center;gap:10px;padding:10px 0;border-bottom:1px solid #e8eef0}.package-finance-row:last-child{border-bottom:0}.package-finance-row>span:first-child{font-size:11px;font-weight:750;color:#455963}.package-check{display:flex;align-items:center;gap:7px;margin:0;color:#526b5b}.package-check input{width:16px;min-height:16px}.package-check b{font-size:10px}.package-extra-input{min-height:38px}.package-included-label{font-size:10px!important;color:var(--green)!important}

      .payment-light{display:inline-flex;align-items:center;gap:7px;padding:6px 9px;border:1px solid var(--border);border-radius:999px;background:#fff;cursor:pointer;white-space:nowrap}.payment-light:hover{transform:translateY(-1px);box-shadow:var(--shadow-sm)}.payment-light-dot{width:11px;height:11px;border-radius:50%;box-shadow:0 0 0 4px rgba(0,0,0,.035)}.payment-light small{font-size:9px;font-weight:850}.payment-light.paid{color:var(--green);background:var(--green-soft);border-color:#d4e8da}.payment-light.paid .payment-light-dot{background:var(--green)}.payment-light.partial{color:var(--yellow);background:var(--yellow-soft);border-color:#f0dfb7}.payment-light.partial .payment-light-dot{background:var(--yellow)}.payment-light.unpaid{color:var(--red);background:var(--red-soft);border-color:#efdcdc}.payment-light.unpaid .payment-light-dot{background:var(--red)}
      .payment-breakdown{padding:14px;margin-top:4px;border:1px dashed #cddde3;border-radius:13px;background:#fff}.payment-breakdown-head{display:flex;align-items:flex-start;justify-content:space-between;gap:15px;margin-bottom:12px}.payment-breakdown-head span{display:block;color:var(--primary-2);font-size:9px;font-weight:900;letter-spacing:.1em}.payment-breakdown-head b{display:block;margin-top:3px;color:var(--text);font-size:14px}.payment-breakdown-head small{display:block;margin-top:4px;color:var(--muted);font-size:10.5px}.payment-breakdown-list{display:grid;gap:10px}.payment-breakdown-row{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:9px;padding:12px;border:1px solid var(--border);border-radius:12px;background:#f9fbfc}.payment-row-remove{align-self:end}.payment-breakdown-empty{padding:15px;color:var(--muted);font-size:11px;text-align:center;border:1px dashed var(--border);border-radius:10px}.invoice-existing-btn{margin-top:8px}.invoice-open-btn{min-height:30px;padding:6px 9px;font-size:10px}.finance-table{min-width:1750px}
      @media(max-width:1150px){.payment-breakdown-row{grid-template-columns:repeat(3,minmax(0,1fr))}}@media(max-width:800px){.payment-breakdown-head{flex-direction:column}.payment-breakdown-row{grid-template-columns:1fr}}
      @media(max-width:1150px){.finance-filters{grid-template-columns:repeat(3,minmax(0,1fr))}.finance-kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.finance-main-grid{grid-template-columns:1fr}.finance-form-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
      @media(max-width:800px){.finance-filters,.finance-kpis,.finance-analysis-grid,.finance-form-grid{grid-template-columns:1fr}.package-finance-row{grid-template-columns:1fr}.finance-ranking-row{grid-template-columns:minmax(100px,150px) 1fr}.finance-ranking-row b{grid-column:2;text-align:left}.finance-line-legend{justify-content:flex-start}}
    `}</style>

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

    <style>{`
      .nav-finance-group{position:relative;width:100%}
      .nav-finance-group>button{width:100%;display:flex;align-items:center;justify-content:space-between}
      .nav-finance-main{display:flex;align-items:center;gap:10px;min-width:0}.nav-icon{flex:0 0 auto;stroke-width:1.9}.nav button{display:flex;align-items:center;gap:10px}.nav-finance-submenu>button svg{flex:0 0 auto;stroke-width:1.9}.logout-btn{display:flex;align-items:center;justify-content:center;gap:7px}
      .nav-chevron-icon{opacity:.65;transition:transform .18s ease,opacity .18s ease;flex:0 0 auto}
      .nav-finance-submenu{display:grid;grid-template-rows:0fr;opacity:0;transform:translateY(-3px);transition:grid-template-rows .2s ease,opacity .18s ease,transform .18s ease;overflow:hidden}
      .nav-finance-submenu>button{min-height:0;overflow:hidden;margin:0 0 0 18px;padding-top:0!important;padding-bottom:0!important;border-left:2px solid #d8e5e9!important;border-radius:0 9px 9px 0!important;background:transparent!important;color:#62757e!important;font-size:12px!important;transition:min-height .2s ease,padding .2s ease,background .15s ease,color .15s ease,transform .15s ease}
      .nav-finance-group:hover .nav-finance-submenu,.nav-finance-group:focus-within .nav-finance-submenu,.nav-finance-group.open .nav-finance-submenu{grid-template-rows:1fr;opacity:1;transform:none}
      .nav-finance-group:hover .nav-finance-submenu>button,.nav-finance-group:focus-within .nav-finance-submenu>button,.nav-finance-group.open .nav-finance-submenu>button{min-height:36px;padding-top:8px!important;padding-bottom:8px!important}
      .nav-finance-group:hover .nav-chevron-icon,.nav-finance-group:focus-within .nav-chevron-icon,.nav-finance-group.open .nav-chevron-icon{transform:rotate(90deg);opacity:1}
      .nav-finance-submenu>button:hover{background:#eef5f7!important;color:#315f72!important;transform:translateX(2px)}
      .nav-finance-submenu>button.active{background:#eaf2f5!important;color:#315f72!important;border-left-color:#4d8191!important;font-weight:800}
      .daily-export-actions{flex-wrap:wrap}
      @media(max-width:760px){.nav-finance-submenu{grid-template-rows:1fr;opacity:1;transform:none}.nav-finance-submenu>button{min-height:36px;padding-top:8px!important;padding-bottom:8px!important}.nav-chevron-icon{transform:rotate(90deg)}.daily-export-actions .btn{flex:1 1 140px}}
      @media(prefers-reduced-motion:reduce){.nav-chevron,.nav-finance-submenu,.nav-finance-submenu>button{transition:none!important}}
    `}</style>

    <style>{`
      .procedure-launch-cell{min-width:180px}
      .procedure-launch-cell>small{display:block;margin-bottom:6px;color:#74838c;font-size:8px;font-weight:800;text-transform:uppercase;letter-spacing:.035em}
      .procedure-launch-cell>div{display:flex;flex-wrap:wrap;gap:5px}
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
        <Field label="Responsável"><Select value={followupForm.responsible} setValue={v=>setFollowupForm({...followupForm,responsible:v})} options={teamResponsibleOptions} first="Selecione"/></Field>
        <Field label="Observação do follow-up" full><textarea rows="5" value={followupForm.note} onChange={e=>setFollowupForm({...followupForm,note:e.target.value})} placeholder="Ex.: paciente pediu retorno na sexta, aguarda autorização, não respondeu WhatsApp..."/></Field>
      </div>
    </Modal>

    <style>{`
      .catalog-item-actions{display:flex;align-items:center;gap:6px}
      .catalog-move-btn{width:auto!important;min-width:30px;padding:0 8px!important;border-radius:8px!important;color:#315f72!important;background:#eef5f7!important;border:1px solid #d9e7eb!important;font-size:14px!important}
      .catalog-move-btn:hover{background:#e2eff3!important}
    `}</style>


    <style>{`
      .nav-symbol{display:inline-grid;place-items:center;min-width:23px;height:23px;padding:0 4px;border-radius:7px;background:rgba(49,95,114,.10);color:#315f72;font-size:11px;font-weight:900;line-height:1;letter-spacing:-.02em}
      .nav button.active .nav-symbol,.nav-finance-group.active>.active .nav-symbol{background:rgba(255,255,255,.18);color:inherit}
      .unimed-page{animation:pageFadeIn .24s ease both}
      .unimed-topbar{align-items:flex-start}
      .unimed-progress{display:flex;justify-content:space-between;gap:24px;align-items:center;padding:22px 24px;margin-bottom:14px;border-left:4px solid #315f72}
      .unimed-progress>div:first-child{display:grid;gap:5px}.unimed-progress span{font-size:11px;font-weight:850;letter-spacing:.09em;color:#74838c}.unimed-progress b{font-size:18px;color:#20313b}.unimed-progress small{color:#74838c}
      .unimed-step-status{min-width:185px;padding:12px 14px;border-radius:12px;background:#f4f7f8;display:grid;gap:4px}.unimed-step-status b{font-size:14px}.unimed-step-status.concluida{background:#edf6f0;color:#4f8a6b}.unimed-step-status.em-andamento{background:#fff5df;color:#9b6b20}
      .unimed-alert{display:flex;gap:13px;align-items:flex-start;border-radius:13px;padding:14px 17px;margin:0 0 16px}.unimed-alert.critical{background:#fbefef;border:1px solid #edcccc;color:#7c3939}.unimed-alert strong{white-space:nowrap;text-transform:uppercase;font-size:11px;letter-spacing:.07em}.unimed-alert span{line-height:1.45}
      .unimed-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;margin-bottom:14px}.unimed-grid.two{grid-template-columns:repeat(2,minmax(0,1fr))}.unimed-block{padding:20px!important;margin-bottom:14px}.unimed-block.full{width:100%}
      .unimed-block-title{display:flex;align-items:center;gap:11px;margin-bottom:15px}.unimed-block-title>span{width:34px;height:34px;border-radius:10px;display:grid;place-items:center;background:#eaf2f5;color:#315f72;font-weight:900;font-size:12px}.unimed-block-title>div{display:grid;gap:2px}.unimed-block-title b{font-size:15px;color:#20313b}.unimed-block-title small{color:#74838c}
      .unimed-info-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}.unimed-info-grid>div{padding:10px;border-radius:10px;background:#f7fafb;display:grid;gap:4px}.unimed-info-grid span,.unimed-timeline span,.unimed-role-list span{color:#74838c;font-size:12px;line-height:1.45}.unimed-info-grid b{font-size:12px}
      .unimed-timeline,.unimed-decision,.unimed-role-list{display:grid;gap:10px}.unimed-timeline>div,.unimed-decision>div,.unimed-role-list>div{padding:11px 12px;background:#f7fafb;border-radius:10px;display:grid;gap:4px}.unimed-decision>div{border-left:3px solid #b9802b}
      .unimed-doc-list{display:grid;gap:8px}.unimed-doc-row{display:grid;grid-template-columns:32px minmax(0,1fr) 180px;gap:10px;align-items:center;padding:9px 10px;border:1px solid #e4ebee;border-radius:10px;transition:transform .15s ease,box-shadow .15s ease,background .15s ease}.unimed-doc-row:hover{transform:translateY(-1px);box-shadow:0 5px 14px rgba(32,49,59,.06);background:#fbfdfe}.unimed-index{width:26px;height:26px;border-radius:8px;background:#eef5f7;color:#315f72;display:grid;place-items:center;font-weight:850;font-size:11px}.unimed-doc-row b{font-size:12px}
      .unimed-form-grid{grid-template-columns:repeat(3,minmax(0,1fr))}.input-ok{border-color:#8bc0a3!important;background:#f6fbf8!important}.input-alert{border-color:#d79b9b!important;background:#fffafa!important}.field-ok{color:#4f8a6b!important;font-weight:700}.field-alert{color:#b85c5c!important;font-weight:800}
      .unimed-steps{margin:0;padding-left:22px;display:grid;gap:10px;color:#53636e}.unimed-steps li{padding-left:4px;line-height:1.48}.unimed-check-list{display:grid;gap:7px}.unimed-check-list label{display:flex;gap:9px;align-items:flex-start;padding:9px 10px;border-radius:9px;background:#f8fafb;border:1px solid #e7edef;cursor:pointer;transition:background .15s ease,border .15s ease}.unimed-check-list label.done{background:#edf6f0;border-color:#cce4d5}.unimed-check-list input{margin-top:2px}.unimed-check-list span{font-size:12px;line-height:1.42}
      .unimed-inline-actions,.unimed-functions{display:flex;flex-wrap:wrap;gap:8px;margin-top:12px}.critical-text{color:#a24a4a;font-weight:800}.unimed-alert-list{display:grid;gap:9px;margin:0;padding-left:19px}.unimed-alert-list li{line-height:1.45}
      .unimed-send-rule{margin-top:12px;padding:10px 12px;border-radius:9px;font-weight:750;font-size:12px}.unimed-send-rule.ok{background:#edf6f0;color:#4f8a6b}.unimed-send-rule.blocked{background:#fff5df;color:#94671f}.unimed-functions .btn:disabled{opacity:.45;cursor:not-allowed}
      .unimed-status-columns{display:grid;grid-template-columns:1fr 1fr;gap:12px}.unimed-status-columns>div{display:grid;gap:6px}.unimed-status-columns>div>b{margin-bottom:3px}.unimed-status-columns span{padding:7px 9px;background:#f7fafb;border-radius:8px;font-size:12px;color:#53636e}.unimed-pdf-card{display:flex;justify-content:space-between;align-items:center;gap:18px;padding:17px 20px!important}.unimed-pdf-card>div:first-child{display:grid;gap:4px}.unimed-pdf-card span{color:#74838c;font-size:12px}
      @media(max-width:1100px){.unimed-grid,.unimed-grid.two{grid-template-columns:1fr}.unimed-form-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
      @media(max-width:700px){.unimed-progress{align-items:flex-start;flex-direction:column}.unimed-step-status{width:100%}.unimed-info-grid,.unimed-form-grid,.unimed-status-columns{grid-template-columns:1fr}.unimed-doc-row{grid-template-columns:30px 1fr}.unimed-doc-row select{grid-column:2}.unimed-pdf-card{align-items:flex-start;flex-direction:column}}
      @media(prefers-reduced-motion:reduce){.unimed-page,.unimed-doc-row{animation:none!important;transition:none!important}}
.finance-task-page{animation:pageFadeIn .24s ease both}.finance-task-indicators{display:flex;gap:9px;align-items:stretch}.finance-task-indicators>div{min-width:112px;padding:12px 14px;border:1px solid #dde6ea;border-radius:12px;background:#fff;display:grid;gap:2px}.finance-task-indicators b{font-size:21px;color:#315f72}.finance-task-indicators span{font-size:11px;color:#74838c;font-weight:750}.finance-task-notice{margin:0 0 14px;padding:11px 14px;border-radius:10px;background:#fff5df;color:#8a641f;border:1px solid #eed59b;font-size:12px;font-weight:700}.finance-task-notice.error{background:#fbeeee;color:#9d4949;border-color:#eccaca}.finance-task-tabs{display:flex;gap:6px;margin-bottom:14px;padding:5px;background:#edf2f4;border-radius:12px;width:max-content;max-width:100%;overflow:auto}.finance-task-tabs button{border:0;background:transparent;padding:9px 15px;border-radius:9px;color:#61717b;font-weight:800;white-space:nowrap;cursor:pointer}.finance-task-tabs button.active{background:#fff;color:#315f72;box-shadow:0 2px 9px rgba(32,49,59,.08)}.finance-task-toolbar{display:flex;justify-content:space-between;gap:14px;align-items:center;margin-bottom:13px}.finance-filter-pills{display:flex;gap:7px;flex-wrap:wrap}.finance-filter-pills button{border:1px solid #dce5e9;background:#fff;border-radius:999px;padding:8px 11px;color:#667680;font-weight:750;cursor:pointer}.finance-filter-pills button b{margin-left:5px;color:#315f72}.finance-filter-pills button.active{background:#315f72;color:#fff;border-color:#315f72}.finance-filter-pills button.active b{color:#fff}.finance-task-list{display:grid;gap:11px}.finance-task-card{padding:0!important;overflow:hidden;transition:transform .18s ease,box-shadow .18s ease}.finance-task-card:hover{transform:translateY(-1px);box-shadow:0 9px 28px rgba(32,49,59,.08)}.finance-task-summary{width:100%;display:grid;grid-template-columns:auto minmax(220px,1fr) minmax(300px,auto) 24px;gap:13px;align-items:center;text-align:left;border:0;background:#fff;padding:16px 18px;cursor:pointer}.finance-code{display:inline-grid;place-items:center;padding:6px 9px;border-radius:8px;background:#eaf2f5;color:#315f72;font-weight:900;font-size:11px;letter-spacing:.04em}.finance-task-main{display:grid;gap:2px}.finance-task-main small,.finance-task-main em{font-style:normal;color:#7b8991;font-size:11px}.finance-task-main b{font-size:14px;color:#20313b}.finance-task-tags{display:flex;gap:5px;flex-wrap:wrap;justify-content:flex-end}.finance-task-tags i,.pop-card i,.pop-status-line i,.video-row i{font-style:normal;font-size:10px;font-weight:850;padding:5px 7px;border-radius:999px;background:#f1f4f5;color:#6e7c84}.finance-task-tags i.ok,.pop-card i.ok,.pop-status-line i.ok,.video-row i.ok{background:#edf6f0;color:#4f8a6b}.finance-task-tags i.attention,.pop-card i.attention,.pop-status-line i.attention,.video-row i.attention{background:#fff5df;color:#9b6b20}.finance-task-tags i.danger{background:#fbeeee;color:#b85c5c}.finance-chevron{font-size:20px;color:#8b989f;transition:transform .2s ease}.finance-task-card.open .finance-chevron{transform:rotate(180deg)}.finance-task-content{border-top:1px solid #e6edef;padding:18px;background:#fbfcfd}.finance-task-content-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:13px}.finance-task-content-head h3{margin:0;font-size:14px}.finance-task-content-head p{margin:3px 0 0;color:#74838c;font-size:12px}.finance-step-list{display:grid;gap:9px}.finance-step{display:grid;grid-template-columns:31px 1fr;gap:10px;padding:12px;border:1px solid #e4ebee;border-radius:11px;background:#fff}.finance-step-number{width:29px;height:29px;display:grid;place-items:center;border-radius:8px;background:#eef5f7;color:#315f72;font-weight:900;font-size:11px}.finance-step-body{display:grid;gap:9px}.finance-step-body>b{font-size:12px;line-height:1.45}.finance-step-body>p{margin:0;color:#74838c;font-size:12px}.finance-step-controls{display:grid;grid-template-columns:repeat(3,minmax(145px,1fr)) minmax(180px,1.2fr);gap:7px}.finance-step-controls select,.finance-step-controls input{min-height:36px}.finance-step-controls .ok{border-color:#a9cfb8;background:#f4faf6}.finance-step-controls .attention{border-color:#e2c47b;background:#fffaf0}.finance-step-controls .danger{border-color:#dcabab;background:#fff7f7}.finance-step-assets{display:grid;gap:7px;padding-top:8px;border-top:1px dashed #e0e8eb}.finance-step-assets>div{display:flex;align-items:center;gap:6px;flex-wrap:wrap}.finance-step-assets>div>span{width:60px;font-size:10px;font-weight:900;color:#7d8991;text-transform:uppercase;letter-spacing:.05em}.btn.mini{padding:6px 8px;font-size:10px}.material-pill{border:0;background:#edf6f0;color:#4f8a6b;font-weight:850;border-radius:999px;padding:5px 8px;cursor:pointer}.finance-task-actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:14px}.inline-confirm{display:flex;gap:6px;align-items:center;font-size:11px;color:#b85c5c;font-weight:800}.finance-empty{padding:30px;text-align:center;color:#74838c}.delegation-layout{display:grid;grid-template-columns:minmax(0,1fr) 300px;gap:14px}.delegation-map{display:grid;grid-template-columns:110px repeat(3,minmax(180px,1fr));grid-auto-rows:minmax(145px,auto);gap:7px}.delegation-axis{display:grid;place-items:center;text-align:center;font-size:11px;font-weight:850;color:#6d7c84;background:#f5f8f9;border-radius:9px}.delegation-axis.row{padding:10px}.delegation-axis.blank{background:transparent}.delegation-cell{border:1px solid #e0e7e9;border-radius:10px;padding:9px;overflow:auto;max-height:220px}.delegation-cell>b{font-size:10px;text-transform:uppercase;letter-spacing:.04em}.delegation-cell.ok{background:#f4faf6}.delegation-cell.attention{background:#fffaf0}.delegation-cell.danger{background:#fff7f7}.delegation-cell>div{display:grid;gap:5px;margin-top:8px}.delegation-cell button{border:1px solid #dfe7ea;background:#fff;text-align:left;border-radius:8px;padding:7px;font-size:10px;color:#53636e;cursor:pointer}.delegation-cell button.delegated{border-left:4px solid #4f8a6b}.delegation-side{padding:16px!important;display:grid;gap:9px;align-content:start}.delegation-side h3{margin:0}.delegation-side p{font-size:12px;color:#74838c}.delegation-side>button{border:1px solid #dde6e9;background:#fff;text-align:left;padding:9px;border-radius:8px;cursor:pointer}.delegation-legend{display:grid;gap:5px;margin-top:7px}.delegation-legend span{padding:7px;border-radius:8px;font-size:10px}.delegation-legend .ok{background:#edf6f0;color:#4f8a6b}.delegation-legend .attention{background:#fff5df;color:#9b6b20}.delegation-legend .danger{background:#fbeeee;color:#b85c5c}.responsible-note,.pops-note{padding:13px 15px!important;margin-bottom:12px;color:#60707a;font-size:12px}.responsible-groups,.pops-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:11px}.responsible-groups>section{padding:16px!important}.responsible-groups .subtle{background:#f8fafb}.responsible-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}.responsible-head h3{margin:0;font-size:14px}.responsible-head span{font-size:11px;color:#74838c}.responsible-item{display:grid;grid-template-columns:auto 1fr;gap:9px;padding:9px 0;border-top:1px solid #edf1f3}.responsible-item>button{border:0;background:#eaf2f5;color:#315f72;border-radius:7px;padding:6px 8px;font-weight:850;cursor:pointer;height:max-content}.responsible-item>div{display:grid;gap:3px}.responsible-item b{font-size:11px}.responsible-item small{color:#74838c}.pop-card{padding:17px!important;display:grid;gap:10px}.pop-card>div:first-child{display:flex;justify-content:space-between;align-items:center}.pop-card h3{margin:0}.pop-card p{margin:0;color:#74838c;font-size:12px;line-height:1.5}.pop-card>div:last-child{display:flex;gap:7px}.task-step-editor{display:grid;gap:8px}.task-step-edit-row{display:grid;grid-template-columns:28px minmax(0,1fr) auto;gap:7px;align-items:start}.task-step-edit-row>span{width:27px;height:27px;display:grid;place-items:center;background:#eef5f7;border-radius:8px;color:#315f72;font-weight:900}.task-step-edit-row>div{display:flex;gap:3px}.task-step-edit-row button{border:1px solid #dce5e8;background:#fff;border-radius:7px;min-width:30px;height:30px;cursor:pointer}.pop-modal-card{max-width:980px!important;max-height:92vh;overflow:auto}.pop-status-line{display:flex;align-items:center;gap:9px;padding:10px 0 14px;border-bottom:1px solid #e8eef0}.pop-status-line span{font-size:11px;color:#74838c}.pop-section{padding:16px 0;border-bottom:1px solid #e8eef0}.pop-section h3{margin:0 0 10px;font-size:14px}.pop-ident-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}.pop-ident-grid>div{padding:10px;border-radius:9px;background:#f7fafb}.pop-ident-grid span{font-size:10px;font-weight:850;color:#7a8991;text-transform:uppercase}.pop-ident-grid p{margin:4px 0 0;font-size:12px;line-height:1.5}.pop-steps{display:grid;gap:8px}.pop-steps>div{display:grid;grid-template-columns:28px 1fr;gap:9px}.pop-steps>div>span{width:27px;height:27px;display:grid;place-items:center;border-radius:8px;background:#eaf2f5;color:#315f72;font-weight:900}.pop-steps>div>div{display:grid;gap:4px;padding:10px;background:#f8fafb;border-radius:9px}.pop-steps p{margin:0;font-size:12px}.pop-steps strong{font-size:11px;color:#a14a4a}.pop-steps small{color:#74838c}.pop-owner-row,.video-row,.material-list>div{display:flex;justify-content:space-between;gap:10px;align-items:center;padding:8px 0;border-top:1px solid #eef2f3}.video-row>div{display:grid;gap:2px}.video-row small{color:#74838c}.material-list{display:grid}.video-drop{display:grid;place-items:center;min-height:95px;border:2px dashed #cbd8dd;border-radius:11px;background:#f8fbfc;color:#53636e;font-weight:800;cursor:pointer}.or{text-align:center;color:#9aa4aa;font-size:11px;margin:8px}.small-material-modal{max-width:480px!important}.material-type-buttons{display:flex;gap:8px}.material-viewer-card{max-width:980px!important;width:min(96vw,980px);height:90vh;display:flex;flex-direction:column}.material-viewer-card iframe{width:100%;flex:1;border:1px solid #dfe7ea;border-radius:10px}.material-viewer-card .form-actions a{text-decoration:none}.finance-task-page .neutral{color:#6f7d85}.finance-task-page .ok{color:#4f8a6b}.finance-task-page .attention{color:#9b6b20}.finance-task-page .danger{color:#b85c5c}@media(max-width:1050px){.finance-task-summary{grid-template-columns:auto 1fr 24px}.finance-task-tags{grid-column:2/4;justify-content:flex-start}.delegation-layout{grid-template-columns:1fr}.finance-step-controls{grid-template-columns:1fr 1fr}.responsible-groups,.pops-grid{grid-template-columns:1fr}}@media(max-width:720px){.finance-task-indicators{width:100%;overflow:auto}.finance-task-toolbar{align-items:flex-start;flex-direction:column}.finance-task-summary{grid-template-columns:auto 1fr 20px}.finance-step-controls{grid-template-columns:1fr}.delegation-map{grid-template-columns:82px repeat(3,minmax(150px,1fr));overflow:auto}.pop-ident-grid{grid-template-columns:1fr}}
    `}
</style>

    <Modal open={catalogModal} onClose={()=>setCatalogModal(false)} title="Cadastros da clínica" subtitle="Disponível para ADM e Orientadora." onSave={()=>setCatalogModal(false)} saveText="Concluir">
      <Catalog title="Médicos / profissionais" items={doctors} onAdd={v=>addCatalog('doctors',v)} onRemove={id=>deactivateCatalog('doctors',id)} />
      <Catalog title="Exames" items={examTypes} onAdd={v=>addCatalog('exam_types',v)} onRemove={id=>deactivateCatalog('exam_types',id)} />
      <Catalog title="Cirurgias" items={surgeryTypes} onAdd={v=>addCatalog('surgery_types',v)} onRemove={id=>deactivateCatalog('surgery_types',id)} />
      <Catalog title="Procedimentos" items={procedureOptions} onAdd={v=>addCatalog('procedures',v)} onRemove={id=>deactivateCatalog('procedures',id)} />
    </Modal>
  </div>;
}

function Catalog({title,description,placeholder='Adicionar novo',items,onAdd,onRemove,onMove,moveLabel}){
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

createRoot(document.getElementById('root')).render(<><V36PaymentSummaryStyle/><V40LoginPriceStyle/><V42AlignmentStyle/><App/></>);
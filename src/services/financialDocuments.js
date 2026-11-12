import { supabase } from '../supabase';

export const uploadFinancialDocument = async (file,sourceType,groupId)=>{
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

export const syncFinancialDocuments = async (form,sourceType,groupId)=>{
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

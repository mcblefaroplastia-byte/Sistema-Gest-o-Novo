import React from 'react';

export const V36PaymentSummaryStyle = () => <style>{`
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


export const V40LoginPriceStyle = () => <style>{`
  .user-box small{display:block;margin-top:5px;color:#83929a;font-size:9px;line-height:1.35}
`}</style>;
export const V42AlignmentStyle = () => <style>{`
  .filter-check-section{display:flex;flex-direction:column;gap:7px}.filter-check-section>label{font-size:12px;font-weight:800;color:#52636b;text-transform:uppercase;letter-spacing:.04em}.filter-check-grid{display:flex;gap:8px;flex-wrap:wrap}.filter-check{display:inline-flex;align-items:center;gap:8px;border:1px solid #d6e0e2;background:#fff;color:#40525b;border-radius:10px;padding:8px 11px;font:inherit;font-weight:700;cursor:pointer;transition:transform .18s ease,box-shadow .18s ease,background .18s ease,border-color .18s ease;position:relative;overflow:hidden}.filter-check::after{content:'';position:absolute;inset:0;background:linear-gradient(110deg,transparent 35%,rgba(255,255,255,.65) 50%,transparent 65%);transform:translateX(-140%);transition:transform .45s ease}.filter-check:hover{transform:translateY(-2px) scale(1.015)}.filter-check:active{transform:scale(.97)}.filter-check.selected{animation:filterPop .28s ease}.filter-check.selected::after{transform:translateX(140%)}@keyframes filterPop{0%{transform:scale(.94)}60%{transform:scale(1.04)}100%{transform:scale(1)}}.filter-check span{display:grid;place-items:center;width:18px;height:18px;border:1px solid #bdcacc;border-radius:5px;background:#fff;color:#fff;font-size:12px}.filter-check:hover{border-color:#8eaaa0;background:#f7faf9}.filter-check.selected{border-color:#75a48e;background:#eaf5ef;color:#315f4c;box-shadow:0 0 0 2px rgba(79,138,107,.08)}.filter-check.selected span{background:#4f8a6b;border-color:#4f8a6b}.daily-report-filters{align-items:flex-end}.daily-report-filters .full{grid-column:1/-1}.promissory-status{display:inline-flex;padding:5px 9px;border-radius:999px;font-size:12px;font-weight:800}.promissory-status.paid{background:#e8f4ed;color:#397156}.promissory-status.partial{background:#fff4dc;color:#9a671d}.promissory-status.open{background:#edf1f3;color:#55646b}.promissory-status.overdue{background:#fbeaea;color:#a94a4a}.finance-kpi.danger{border-top-color:#b85c5c}.promissory-page .finance-table td{vertical-align:middle}

  .period-filter{display:flex;gap:6px;align-items:center;min-width:210px}.period-filter select{max-width:80px}.period-filter input{min-width:132px}
  .form-validation-banner{display:flex;flex-direction:column;gap:4px;background:#fff1f1;border:1px solid #e4a1a1;color:#963f3f;border-radius:12px;padding:12px 14px;animation:validationShake .34s ease}
  .form-validation-banner strong{font-size:14px}.form-validation-banner span{font-size:12px}
  @keyframes validationShake{0%,100%{transform:translateX(0)}25%{transform:translateX(-5px)}50%{transform:translateX(5px)}75%{transform:translateX(-3px)}}
  .field-error label{color:#a74343!important}.field-error input,.field-error select,.field-error textarea,.field-error .exam-multi-select,.field-error .procedure-multi-select,.field-error .finance-quick-summary{border-color:#c75a5a!important;background:#fff8f8!important;box-shadow:0 0 0 3px rgba(199,90,90,.09)!important}
`}</style>;

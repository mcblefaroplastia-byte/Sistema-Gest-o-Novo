import React from 'react';
import { moneyBR, numericValue, percentOf, statusClass } from '../utils/helpers';

export function Bars({data}){
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


export function DonutChart({title,data,centerLabel='Total'}){
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

export function TrendLineChart({title,labels,series,percent=false}){
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


export function FinancialLineChart({title,labels,series}){
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

export function FinancialBars({title,subtitle,data}){
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

export function StatusTypeMatrix({title,subtitle,statuses,data,totalByStatus,labelName='Tipo'}){
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


export function DoctorTypeMatrix({title,subtitle,rows,itemLabel='Procedimento'}){
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

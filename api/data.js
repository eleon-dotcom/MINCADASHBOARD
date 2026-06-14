const XLSX = require('xlsx');
const https = require('https');

// Lee el Excel directo desde GitHub (raw) — sin credenciales ni costos
// Cada vez que subas el archivo al repo, el dashboard se actualiza automáticamente
const GITHUB_RAW_URL =
  'https://raw.githubusercontent.com/eleon-dotcom/MINCADASHBOARD/master/' +
  'Entrgable%20Samana%2015(Recuperado%20autom%C3%A1ticamente).xlsx';

let cache = { data: null, ts: 0 };
const CACHE_MS = 60 * 60 * 1000; // 1 hora

const TALLERES = ['CHICO','JAVERIANA','CEDRITOS','AV CHILE','LAS VEGAS'];
const MES_N = {1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',
               7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'};

function fetchBuffer(url) {
  return new Promise((resolve, reject) => {
    https.get(url, res => {
      if (res.statusCode === 302 || res.statusCode === 301) {
        return fetchBuffer(res.headers.location).then(resolve).catch(reject);
      }
      if (res.statusCode !== 200) {
        return reject(new Error(`HTTP ${res.statusCode} fetching Excel`));
      }
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

function processData(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const ws = wb.Sheets['Ventas Talleres'];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

  const rows = [];
  for (let i = 1; i < raw.length; i++) {
    const r = raw[i];
    if (!r || !r[0]) continue;
    const total = parseFloat(String(r[7]||'0').replace(/[^0-9.\-]/g,''))||0;
    const cant  = parseFloat(String(r[6]||'0').replace(/[^0-9.\-]/g,''))||0;
    rows.push({
      ano: parseInt(r[1])||0, sem: parseInt(r[2])||0, mes: parseInt(r[3])||0,
      prod: String(r[4]||''), tipo: String(r[5]||''), cant, total,
      orden: r[9]!=null ? String(r[9]) : null,
      tec:    String(r[10]||'SIN ASIGNAR').trim()||'SIN ASIGNAR',
      taller: String(r[11]||'SIN TALLER').trim()||'SIN TALLER'
    });
  }

  const ymSet = new Set();
  rows.forEach(r => { if(r.ano&&r.mes) ymSet.add(`${r.ano}-${r.mes}`); });
  const all_ym = Array.from(ymSet)
    .map(k => { const [y,m]=k.split('-').map(Number); return [y,m]; })
    .sort((a,b) => a[0]!==b[0] ? a[0]-b[0] : a[1]-b[1]);

  const VT_LAST_YR = Math.max(...rows.map(r=>r.ano));
  const mesSems = {};
  rows.filter(r=>r.ano===VT_LAST_YR).forEach(r=>{
    if(!mesSems[r.mes]) mesSems[r.mes]=new Set();
    mesSems[r.mes].add(r.sem);
  });
  const complete = Object.entries(mesSems).filter(([,s])=>s.size>=4).map(([m])=>parseInt(m));
  const VT_CURR_MES = complete.length ? Math.max(...complete) : Math.max(...Object.keys(mesSems).map(Number));
  const VT_PREV_MES = VT_CURR_MES>1 ? VT_CURR_MES-1 : 12;
  const VT_PREV_YR  = VT_CURR_MES>1 ? VT_LAST_YR : VT_LAST_YR-1;

  const sumV  = arr => arr.reduce((s,r)=>s+r.total, 0);
  const uniqO = arr => new Set(arr.map(r=>r.orden).filter(Boolean)).size;
  const filt  = (yr,mes,tal) => rows.filter(r=>r.ano===yr&&r.mes===mes&&(!tal||r.taller===tal));

  // vm
  const vm = {};
  all_ym.forEach(([yr,mes]) => {
    const k=`${yr}-${mes}`, all=filt(yr,mes);
    vm[k] = { total:Math.round(sumV(all)), ordenes:uniqO(all), talleres:{} };
    TALLERES.forEach(t => {
      const d=filt(yr,mes,t); const v=Math.round(sumV(d)); const o=uniqO(d); const l=d.length;
      vm[k].talleres[t]={v,o,l,tk:o?Math.round(v/o):0,lpo:o?+(l/o).toFixed(2):0};
    });
  });

  // vs (últimas 12 semanas)
  const last12 = [...new Set(rows.map(r=>r.sem))].sort((a,b)=>a-b).slice(-12);
  const vs = {};
  last12.forEach(s => {
    const d=rows.filter(r=>r.sem===s);
    vs[String(s)] = { tot:Math.round(sumV(d)), t:{} };
    TALLERES.forEach(t => { vs[String(s)].t[t]=Math.round(sumV(d.filter(r=>r.taller===t))); });
  });

  // tec_month_v para delta
  const tec_mv = {};
  all_ym.forEach(([yr,mes]) => {
    const k=`${yr}-${mes}`;
    const d=rows.filter(r=>r.ano===yr&&r.mes===mes&&r.tec!=='SIN ASIGNAR');
    const byT={};
    d.forEach(r=>{byT[r.tec]=(byT[r.tec]||0)+r.total;});
    tec_mv[k]=byT;
  });

  // tc
  const tc = {};
  all_ym.forEach(([yr,mes]) => {
    const k=`${yr}-${mes}`;
    const pm=mes>1?mes-1:12, py=mes>1?yr:yr-1, pk=`${py}-${pm}`;
    const prevV=tec_mv[pk]||{};
    const d=rows.filter(r=>r.ano===yr&&r.mes===mes&&r.tec!=='SIN ASIGNAR');
    if(!d.length){tc[k]=[];return;}
    const byTec={};
    d.forEach(r=>{
      if(!byTec[r.tec]) byTec[r.tec]={v:0,l:0,ords:new Set(),talleres:{}};
      byTec[r.tec].v+=r.total; byTec[r.tec].l++;
      if(r.orden) byTec[r.tec].ords.add(r.orden);
      byTec[r.tec].talleres[r.taller]=(byTec[r.tec].talleres[r.taller]||0)+r.total;
    });
    const avgT={};
    Object.values(byTec).forEach(t=>{
      const mainT=Object.entries(t.talleres).sort((a,b)=>b[1]-a[1])[0][0];
      if(!avgT[mainT]) avgT[mainT]=[];
      avgT[mainT].push(t.v);
    });
    Object.keys(avgT).forEach(t=>{avgT[t]=avgT[t].reduce((a,b)=>a+b,0)/avgT[t].length;});
    tc[k]=Object.entries(byTec).map(([name,t])=>{
      const sorted=Object.entries(t.talleres).sort((a,b)=>b[1]-a[1]);
      const mainT=sorted[0][0]; const o=t.ords.size; const v=Math.round(t.v); const l=t.l;
      const vp=Math.round(prevV[name]||0);
      return {
        TECNICO:name, TALLER:mainT, v, o, l,
        tk: o?Math.round(v/o):0,
        lpo: o?+(l/o).toFixed(2):0,
        vs: +((v-(avgT[mainT]||0))/Math.max(avgT[mainT]||1,1)*100).toFixed(1),
        nt: Object.keys(t.talleres).length,
        tl: sorted.map(([t])=>t),
        vp, dp: vp>0?+((v-vp)/vp*100).toFixed(1):null
      };
    }).sort((a,b)=>b.v-a.v);
  });

  // tp
  const tp = {};
  all_ym.forEach(([yr,mes]) => {
    const k=`${yr}-${mes}`; tp[k]={};
    ['ALL',...TALLERES].forEach(scope => {
      const d=rows.filter(r=>r.ano===yr&&r.mes===mes&&(scope==='ALL'||r.taller===scope));
      if(!d.length){tp[k][scope]=[];return;}
      const byP={};
      d.forEach(r=>{
        const pk=r.prod+'||'+r.tipo;
        if(!byP[pk]) byP[pk]={p:r.prod,ti:r.tipo,v:0,c:0};
        byP[pk].v+=r.total; byP[pk].c+=r.cant;
      });
      const tot=sumV(d);
      tp[k][scope]=Object.values(byP).sort((a,b)=>b.v-a.v).slice(0,15).map(x=>({
        p:x.p, ti:x.ti, v:Math.round(x.v), c:Math.round(x.c),
        pct: tot?+(x.v/tot*100).toFixed(1):0
      }));
    });
  });

  // mx
  const mx = {};
  all_ym.forEach(([yr,mes]) => {
    const k=`${yr}-${mes}`; mx[k]={};
    TALLERES.forEach(t => {
      const d=rows.filter(r=>r.ano===yr&&r.mes===mes&&r.taller===t);
      const sv=Math.round(sumV(d.filter(r=>r.tipo==='SERVICIO')));
      const pv=Math.round(sumV(d.filter(r=>r.tipo==='PRODUCTO')));
      const tot=sv+pv;
      mx[k][t]={sv,pv,tot,ps:tot?+(sv/tot*100).toFixed(1):0,pp:tot?+(pv/tot*100).toFixed(1):0};
    });
  });

  const dc=filt(VT_LAST_YR,VT_CURR_MES), dp2=filt(VT_PREV_YR,VT_PREV_MES);
  const vc=Math.round(sumV(dc)), vp2=Math.round(sumV(dp2));
  const vyr=Math.round(sumV(rows.filter(r=>r.ano===VT_LAST_YR&&r.mes<=VT_CURR_MES)));

  return {
    meta:{
      lm:VT_CURR_MES, ly:VT_LAST_YR, pm:VT_PREV_MES, py:VT_PREV_YR,
      yrs:[...new Set(rows.map(r=>r.ano))].sort(),
      T:TALLERES, M:MES_N, all_ym,
      updated: new Date().toISOString(),
      source: 'GitHub'
    },
    kpis:{
      vc, vp:vp2, d:vp2?+((vc-vp2)/vp2*100).toFixed(1):0,
      vyr, oc:uniqO(dc),
      mes_lbl:`${MES_N[VT_CURR_MES]} ${VT_LAST_YR}`,
      prev_lbl:`${MES_N[VT_PREV_MES]} ${VT_PREV_YR}`
    },
    vm, vs, tc, tp, mx
  };
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 's-maxage=3600, stale-while-revalidate');
  try {
    const now = Date.now();
    if (!cache.data || now - cache.ts > CACHE_MS) {
      console.log('Descargando Excel desde GitHub...');
      const buffer = await fetchBuffer(GITHUB_RAW_URL);
      cache.data = processData(buffer);
      cache.ts = now;
      console.log(`OK — ${cache.data.meta.mes_lbl}`);
    }
    res.status(200).json(cache.data);
  } catch (err) {
    console.error('Error:', err.message);
    res.status(500).json({ error: err.message });
  }
};

/* Tour Engine — shared logic for tour player */

const SPEED={foot:1.25,transit:5.5,car:8.5};
const WAIT ={foot:0,  transit:180,  car:60};
const MCOL ={foot:'#5bb8f5',transit:'#b07ef5',car:'#f5a623'};
const MICO ={foot:'🚶',transit:'🚌',car:'🚗'};
const TICO={'turn-left':'←','turn-right':'→','turn-sharp-left':'↰','turn-sharp-right':'↱',
  'turn-slight-left':'↖','turn-slight-right':'↗','uturn':'↩','continue':'↑',
  'depart':'🚦','arrive':'🏁','roundabout':'⟳','rotary':'⟳','end of road':'↑',
  'new name':'↑','fork':'⑂','merge':'⤵'};
const TES={'turn-left':'Gira a la izquierda','turn-right':'Gira a la derecha',
  'turn-sharp-left':'Giro cerrado izquierda','turn-sharp-right':'Giro cerrado derecha',
  'turn-slight-left':'Dobla levemente izquierda','turn-slight-right':'Dobla levemente derecha',
  'uturn':'Da la vuelta','continue':'Continúa recto','depart':'Sal de','arrive':'Has llegado',
  'roundabout':'Toma la rotonda','rotary':'Toma la rotonda','end of road':'Al final gira',
  'new name':'Continúa por','fork':'En el cruce','merge':'Incorpórate'};

const fmtD=m=>m<1000?Math.round(m)+'m':(m/1000).toFixed(1)+' km';
const fmtT=s=>{const m=Math.ceil(s/60);return m<60?m+' min':Math.floor(m/60)+'h '+(m%60)+'m'};
function hav(a,b,c,d){
  const R=6371000,dL=(c-a)*Math.PI/180,dl=(d-b)*Math.PI/180,
    x=Math.sin(dL/2)**2+Math.cos(a*Math.PI/180)*Math.cos(c*Math.PI/180)*Math.sin(dl/2)**2;
  return R*2*Math.atan2(Math.sqrt(x),Math.sqrt(1-x));
}

let toastTimer=null;
function toast(m){
  const t=document.getElementById('toast');if(!t)return;
  t.textContent=m;t.classList.add('on');
  if(toastTimer)clearTimeout(toastTimer);
  toastTimer=setTimeout(()=>t.classList.remove('on'),2800);
}

/* localStorage helpers */
const LS={
  get:(k,def=null)=>{try{const v=localStorage.getItem(k);return v?JSON.parse(v):def}catch{return def}},
  set:(k,v)=>{try{localStorage.setItem(k,JSON.stringify(v))}catch{}},
  del:(k)=>{try{localStorage.removeItem(k)}catch{}}
};

/* Tour index management */
function getTourIndex(){return LS.get('tours_index',{});}
function saveTourToIndex(meta){
  const idx=getTourIndex();
  idx[meta.id]={id:meta.id,city:meta.city,country:meta.country,emoji:meta.emoji,
    stops:meta.stops?.length||0,duration:meta.duration,custom:true,
    updatedAt:Date.now()};
  LS.set('tours_index',idx);
}
function deleteTourFromIndex(id){const idx=getTourIndex();delete idx[id];LS.set('tours_index',idx);}

/* Recalculate stop times from first stop */
function recalcTimes(stops){
  if(!stops.length)return stops;
  let[h,m]=stops[0].time.split(':').map(Number);
  stops.forEach((s,i)=>{
    if(i>0){s.time=`${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;}
    m+=s.stayMin+10;h+=Math.floor(m/60);m%=60;
  });
  return stops;
}

/* OSRM routing */
async function getTT(fLa,fLo,tLa,tLo,mode,cache){
  const k=`${fLa.toFixed(3)},${fLo.toFixed(3)}-${tLa.toFixed(3)},${tLo.toFixed(3)}-${mode}`;
  if(cache&&cache.has(k))return cache.get(k);
  const dist=hav(fLa,fLo,tLa,tLo);
  try{
    const p=mode==='car'?'car':'foot';
    const r=await fetch(
      `https://router.project-osrm.org/route/v1/${p}/${fLo},${fLa};${tLo},${tLa}?overview=false&alternatives=false`,
      {signal:AbortSignal.timeout(6000)}
    );
    const d=await r.json();
    if(d.code==='Ok'){
      const rd=d.routes[0].distance;
      let dur=d.routes[0].duration;
      if(mode==='transit')dur=rd<400?rd/SPEED.foot:rd/SPEED.transit+WAIT.transit;
      else if(mode==='car')dur+=WAIT.car;
      const res={dist:rd,dur:Math.round(dur)};
      if(cache)cache.set(k,res);
      return res;
    }
  }catch(e){}
  return{dist,dur:Math.round(dist/SPEED[mode]+(mode==='transit'&&dist>400?WAIT.transit:0))};
}

async function fetchRoute(fromLng,fromLat,toLng,toLat,mode){
  const p=mode==='car'?'car':'foot';
  const url=`https://router.project-osrm.org/route/v1/${p}/${fromLng},${fromLat};${toLng},${toLat}`+
    `?steps=true&overview=full&geometries=geojson&alternatives=false&continue_straight=false`;
  const res=await fetch(url,{signal:AbortSignal.timeout(9000)});
  if(!res.ok)throw new Error('HTTP '+res.status);
  const data=await res.json();
  if(data.code!=='Ok')throw new Error(data.code);
  return data.routes[0];
}

async function fetchTourRoute(stops,mode='foot'){
  if(stops.length<2)return null;
  const p=mode==='car'?'car':'foot';
  const coords=stops.map(s=>`${s.lng},${s.lat}`).join(';');
  try{
    const r=await fetch(
      `https://router.project-osrm.org/trip/v1/${p}/${coords}`+
      `?overview=full&geometries=geojson&source=first&destination=last&roundtrip=false`,
      {signal:AbortSignal.timeout(10000)}
    );
    const d=await r.json();
    if(d.code==='Ok'&&d.trips?.length)return d.trips[0].geometry.coordinates.map(c=>[c[1],c[0]]);
  }catch(e){}
  return stops.map(s=>[s.lat,s.lng]);
}

/* Wikipedia photo */
async function fetchWikiPhoto(wikiUrl){
  if(!wikiUrl)return null;
  const title=decodeURIComponent(wikiUrl.split('/wiki/')[1]||'');
  if(!title)return null;
  try{
    const r=await fetch(
      `https://es.wikipedia.org/w/api.php?action=query&titles=${encodeURIComponent(title)}&prop=pageimages&pithumbsize=400&format=json&origin=*`,
      {signal:AbortSignal.timeout(5000)}
    );
    const d=await r.json();
    return Object.values(d.query?.pages||{})[0]?.thumbnail?.source||null;
  }catch{return null;}
}

/* Compass / GPS helpers */
function buildGPSIcon(){
  return L.divIcon({className:'',
    html:`<div class="gw"><div class="gr"></div><div class="gb"></div><div class="gc hid" id="gps-cone"></div></div>`,
    iconSize:[44,44],iconAnchor:[22,22]});
}

/* Export for module use if needed */
if(typeof module!=='undefined')module.exports={SPEED,WAIT,MCOL,MICO,TICO,TES,fmtD,fmtT,hav,toast,LS,getTourIndex,saveTourToIndex,deleteTourFromIndex,recalcTimes,getTT,fetchRoute,fetchTourRoute,fetchWikiPhoto};

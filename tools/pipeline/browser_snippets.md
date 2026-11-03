# Browser snippets (run with the built-in browser's javascript tool)

Boutiqaat blocks server requests, so product data is collected in the user's browser.
Large results are saved by the tool to a file; parse that file and save the list as JSON for `bq_json` / `off_json`.

## A. Boutiqaat brand (open the brand page, e.g. https://www.boutiqaat.com/ar-sa/women/<brand>/br/)
```js
const base=location.pathname, prods={};
const nd=h=>JSON.parse(h.match(/<script id="__NEXT_DATA__" type="application\/json">(.*?)<\/script>/s)[1]);
for(let pg=1;pg<=40;pg++){const r=nd(await fetch(base+'?page='+pg).then(r=>r.text())).props.pageProps.response; const ps=(r.products||[[]])[0]||[]; let n=0;
 ps.forEach(p=>{if(!prods[p.entity_id])n++; prods[p.entity_id]={id:p.entity_id,sku:p.sku,name:p.name.trim(),price:parseFloat(String(p.final_price).replace(',','')),avail:p.is_saleable=='1',url:p.product_url}}); if(!ps.length||!n)break;}
const find=(o,id)=>{if(!o||typeof o!=='object')return null; if(o.entity_id==id&&o.description!==undefined)return o; for(const k in o){const r=find(o[k],id); if(r)return r;} return null;};
const list=Object.values(prods); let i=0;
await Promise.all([...Array(6)].map(async()=>{while(i<list.length){const p=list[i++]; try{const o=find(nd(await fetch(p.url).then(r=>r.text())),p.id)||{};
 let g=o.gallery_images||[]; if(typeof g==='string'){try{g=JSON.parse(g)}catch(e){g=[]}}
 let imgs=g.filter(x=>x&&!x.isVideo).map(x=>x.image); if(o.image_url) imgs=[o.image_url,...imgs.filter(x=>x!==o.image_url)];
 p.desc=(o.description||'').replace(/<[^>]+>/g,' ').replace(/&nbsp;/g,' ').replace(/\s+/g,' ').trim(); p.imgs=imgs;
 if(o.final_price_with_tax)p.price=parseFloat(String(o.final_price_with_tax).replace(',','')); if(o.is_saleable!==undefined)p.avail=o.is_saleable=='1';}catch(e){p.err=String(e)}}}));
JSON.stringify(list)
```

## B. Official Salla store (open the store home page)
```js
let all=[], r=await salla.api.request('products',{params:{per_page:100}},'get'); all=all.concat(r.data);
while(r.cursor&&r.cursor.next){const u=new URL(r.cursor.next); r=await salla.api.request('products',{params:Object.fromEntries(u.searchParams)},'get'); all=all.concat(r.data);}
const list=all.map(p=>({id:String(p.id),name:p.name.trim(),price:p.price,avail:!!p.is_available,url:p.url,cur:p.currency}));
let i=0; await Promise.all([...Array(6)].map(async()=>{while(i<list.length){const p=list[i++]; if(!p.avail)continue; try{
 const h=await fetch(p.url).then(r=>r.text()); const doc=new DOMParser().parseFromString(h,'text/html');
 p.imgs=[...new Set([...doc.querySelectorAll('salla-slider img, .swiper-slide img, [class*=gallery] img, [class*=product] img')].map(x=>x.getAttribute('data-src')||x.getAttribute('src')).filter(u=>u&&u.includes('cdn.salla.sa')&&!u.includes('cdn-cgi')))];
 const d=await salla.api.request('products/'+p.id+'/details',{},'get'); p.desc=(d.data.description||'').replace(/<[^>]+>/g,' ').replace(/&nbsp;/g,' ').replace(/\s+/g,' ').trim().slice(0,900);}catch(e){p.err=String(e)}}}));
JSON.stringify(list)
```
If the official site is not a Salla store, skip it (or collect equivalent fields: id, name, price, avail, desc, imgs).

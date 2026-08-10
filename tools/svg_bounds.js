const { chromium } = require('playwright'); const fs=require('fs');
const body = fs.readFileSync('/home/user/claude_playground/docs/asset-tokenization-graphic-record.html','utf8');
const doc = `<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>*,*::before,*::after{box-sizing:border-box}body{margin:0}img,svg,video{max-width:100%}</style></head><body>${body}</body></html>`;
(async()=>{
 const b=await chromium.launch();
 const ctx=await b.newContext({viewport:{width:1280,height:900}});
 const p=await ctx.newPage(); await p.setContent(doc,{waitUntil:'load'});
 await p.evaluate(()=>document.querySelectorAll('.rv').forEach(e=>e.classList.add('in')));
 await p.waitForTimeout(400);
 const res = await p.evaluate(()=>{
   const out=[];
   document.querySelectorAll('svg[viewBox]').forEach((svg,si)=>{
     const vb=svg.viewBox.baseVal;
     svg.querySelectorAll('text').forEach(t=>{
       let bb; try{ bb=t.getBBox(); }catch(e){ return; }
       const overR = bb.x+bb.width  - (vb.x+vb.width);
       const overB = bb.y+bb.height - (vb.y+vb.height);
       const overL = vb.x - bb.x, overT = vb.y - bb.y;
       const worst = Math.max(overR,overB,overL,overT);
       if(worst > 0.5) out.push({svg:si, text:t.textContent.slice(0,34),
         x:+bb.x.toFixed(1), w:+bb.width.toFixed(1), right:+(bb.x+bb.width).toFixed(1),
         vbW:vb.width, vbH:vb.height, bottom:+(bb.y+bb.height).toFixed(1), over:+worst.toFixed(1)});
     });
   });
   // also: does any <text> collide past its own sibling rect? (shape containment for boxed labels)
   return out;
 });
 // element-level clipping: is any SVG's rendered content taller/wider than its box?
 const clip = await p.evaluate(()=>{
   const o=[];
   document.querySelectorAll('svg[viewBox]').forEach((svg,si)=>{
     const r=svg.getBoundingClientRect();
     const parent=svg.parentElement.getBoundingClientRect();
     if(r.width>parent.width+1 && getComputedStyle(svg.parentElement).overflowX==='visible')
       o.push({svg:si, w:r.width, parentW:parent.width});
   });
   return o;
 });
 console.log('SVG <text> overflowing its viewBox:', res.length===0 ? 'none ✔' : '');
 res.forEach(r=>console.log('  ', JSON.stringify(r)));
 console.log('SVGs wider than their (non-scrolling) container:', clip.length===0 ? 'none ✔' : JSON.stringify(clip));
 await b.close();
 process.exit(res.length||clip.length ? 1 : 0);
})();

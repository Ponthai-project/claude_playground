// Deterministic runtime verification of the artifact in headless Chromium.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const SRC = '/home/user/claude_playground/docs/asset-tokenization-graphic-record.html';
const OUT = '/tmp/claude-0/-home-user-claude-playground/1e08d9d7-3372-5de1-b011-ea1450a7f21e/scratchpad/shots';
fs.mkdirSync(OUT, { recursive: true });

const body = fs.readFileSync(SRC, 'utf8');
// Reproduce the artifact host wrapper: doctype + head + minimal reset + body.
const doc = `<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>*,*::before,*::after{box-sizing:border-box}body{margin:0}img,svg,video{max-width:100%}</style>
</head><body>${body}</body></html>`;

const results = [];
const P = (m) => results.push(['PASS', m]);
const F = (m) => results.push(['FAIL', m]);
const N = (m) => results.push(['NOTE', m]);

// ── WCAG contrast, computed in-page ──────────────────────────
const CONTRAST_FN = `
  function parse(c){ const m=c.match(/[\\d.]+/g); if(!m) return null;
    return {r:+m[0],g:+m[1],b:+m[2],a:m.length>3?+m[3]:1}; }
  function lum({r,g,b}){ const f=v=>{v/=255; return v<=0.03928?v/12.92:Math.pow((v+0.055)/1.055,2.4)};
    return 0.2126*f(r)+0.7152*f(g)+0.0722*f(b); }
  function over(fg,bg){ const a=fg.a; return {r:fg.r*a+bg.r*(1-a), g:fg.g*a+bg.g*(1-a), b:fg.b*a+bg.b*(1-a), a:1}; }
  function effBg(el){
    let cur=el, stack=[];
    while(cur){ const c=parse(getComputedStyle(cur).backgroundColor);
      if(c && c.a>0){ stack.push(c); if(c.a===1) break; }
      cur=cur.parentElement; }
    if(!stack.length) return {r:255,g:255,b:255,a:1};
    let base=stack[stack.length-1];
    for(let i=stack.length-2;i>=0;i--) base=over(stack[i],base);
    return base;
  }
  function ratio(el){
    const bg=effBg(el); let fg=parse(getComputedStyle(el).color); if(!fg) return null;
    if(fg.a<1) fg=over(fg,bg);
    const L1=lum(fg), L2=lum(bg);
    return (Math.max(L1,L2)+0.05)/(Math.min(L1,L2)+0.05);
  }
`;

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium/chrome-linux/chrome' })
    .catch(() => chromium.launch());

  const themes = [
    { key: 'system-light', colorScheme: 'light', stamp: null },
    { key: 'system-dark',  colorScheme: 'dark',  stamp: null },
    { key: 'stamp-light',  colorScheme: 'dark',  stamp: 'light' },  // explicit light must beat dark OS
    { key: 'stamp-dark',   colorScheme: 'light', stamp: 'dark'  },  // explicit dark must beat light OS
  ];
  const viewports = [
    { key: 'desktop', width: 1280, height: 1000 },
    { key: 'mobile',  width: 375,  height: 780  },
  ];

  const bgByTheme = {};

  for (const t of themes) {
    for (const v of viewports) {
      const ctx = await browser.newContext({
        viewport: { width: v.width, height: v.height },
        colorScheme: t.colorScheme,
        deviceScaleFactor: 1,
      });
      const page = await ctx.newPage();
      const consoleErrs = [], pageErrs = [], requests = [];
      page.on('console', m => { if (m.type() === 'error') consoleErrs.push(m.text()); });
      page.on('pageerror', e => pageErrs.push(String(e)));
      page.on('request', r => { if (!r.url().startsWith('data:') && !r.url().startsWith('about:')) requests.push(r.url()); });

      await page.setContent(doc, { waitUntil: 'load' });
      if (t.stamp) await page.evaluate(s => document.documentElement.setAttribute('data-theme', s), t.stamp);
      // scroll through so IntersectionObserver reveals everything, then settle
      await page.evaluate(async () => {
        const h = document.documentElement.scrollHeight;
        for (let y = 0; y < h; y += 600) { window.scrollTo(0, y); await new Promise(r => setTimeout(r, 12)); }
        window.scrollTo(0, 0);
      });
      await page.waitForTimeout(700);

      const tag = `${t.key}-${v.key}`;

      // ── errors ─────────────────────────────────────────
      if (consoleErrs.length === 0 && pageErrs.length === 0) P(`${tag}: no console errors, no uncaught exceptions`);
      else F(`${tag}: errors → console=${JSON.stringify(consoleErrs)} page=${JSON.stringify(pageErrs)}`);

      const netExt = requests.filter(u => !u.startsWith('file:') && u !== 'about:blank');
      if (netExt.length === 0) P(`${tag}: zero network requests (fully self-contained)`);
      else F(`${tag}: network requests attempted → ${JSON.stringify(netExt.slice(0, 5))}`);

      // ── horizontal overflow ────────────────────────────
      const of = await page.evaluate(() => {
        const de = document.documentElement;
        const offenders = [];
        const vw = de.clientWidth;
        document.querySelectorAll('*').forEach(el => {
          const r = el.getBoundingClientRect();
          if (r.width > 0 && (r.right > vw + 1.5 || r.left < -1.5)) {
            // ignore nodes inside an overflow-x:auto scroller
            let p = el.parentElement, inScroller = false;
            while (p) { const ov = getComputedStyle(p).overflowX;
              if (ov === 'auto' || ov === 'scroll') { inScroller = true; break; } p = p.parentElement; }
            if (!inScroller) offenders.push(el.tagName + '.' + ((el.className && el.className.baseVal) || el.className || '').toString().split(' ')[0]
              + ` [${Math.round(r.left)}..${Math.round(r.right)} vs ${vw}]`);
          }
        });
        return { scrollW: de.scrollWidth, clientW: de.clientWidth, bodyScrollW: document.body.scrollWidth,
                 offenders: [...new Set(offenders)].slice(0, 8) };
      });
      if (of.scrollW <= of.clientW + 1 && of.offenders.length === 0)
        P(`${tag}: no horizontal page scroll (scrollWidth ${of.scrollW} <= clientWidth ${of.clientW})`);
      else F(`${tag}: horizontal overflow — scrollWidth ${of.scrollW} vs ${of.clientW}; offenders ${JSON.stringify(of.offenders)}`);

      // ── theme resolution ───────────────────────────────
      const th = await page.evaluate(() => {
        const cs = getComputedStyle(document.body);
        return { bg: cs.backgroundColor, color: cs.color,
                 paper: getComputedStyle(document.documentElement).getPropertyValue('--paper').trim(),
                 ink: getComputedStyle(document.documentElement).getPropertyValue('--ink').trim() };
      });
      bgByTheme[tag] = th;
      if (th.bg !== 'rgba(0, 0, 0, 0)' && th.bg !== 'transparent')
        P(`${tag}: body background opaque → ${th.bg} (--paper ${th.paper}, --ink ${th.ink})`);
      else F(`${tag}: body background is transparent — would borrow host ground`);

      // ── contrast ───────────────────────────────────────
      const cr = await page.evaluate(`(() => { ${CONTRAST_FN}
        const sels = ['.head h1','.head .lede','.sec-title h2','.sec-title .q','h3.blk','.card p','.note p',
                      'ul.pt li','ol.pt li','.tl-desc','.tl-when','.tl-what','tbody td','thead th','tbody th',
                      'figcaption','.callout p','.verdict p','.sources li','.sources a','.agenda a','.tag','.chip','.head-meta'];
        const out=[];
        for(const s of sels){
          document.querySelectorAll(s).forEach(el=>{
            if(!el.textContent.trim()) return;
            const cs=getComputedStyle(el);
            const px=parseFloat(cs.fontSize);
            const bold=parseInt(cs.fontWeight,10)>=700;
            const large = px>=24 || (px>=18.66 && bold);
            const r=ratio(el); if(r==null) return;
            out.push({sel:s, r:+r.toFixed(2), px:+px.toFixed(1), large, need: large?3:4.5});
          });
        }
        // worst per selector
        const worst={}; out.forEach(o=>{ if(!worst[o.sel]||o.r<worst[o.sel].r) worst[o.sel]=o; });
        return Object.values(worst).sort((a,b)=>a.r-b.r);
      })()`);
      const bad = cr.filter(c => c.r < c.need);
      if (bad.length === 0) P(`${tag}: WCAG AA contrast met on all ${cr.length} sampled text roles (worst ${cr[0].sel} = ${cr[0].r}:1, needs ${cr[0].need})`);
      else F(`${tag}: contrast below AA → ${bad.map(b => `${b.sel} ${b.r}:1 (needs ${b.need}, ${b.px}px)`).join('; ')}`);

      // desktop-only structural checks
      if (v.key === 'desktop' && t.key === 'system-light') {
        const struct = await page.evaluate(() => {
          const revealed = document.querySelectorAll('.rv.in').length, total = document.querySelectorAll('.rv').length;
          const roughApplied = getComputedStyle(document.querySelector('.card.rough'), '::before').filter;
          const secs = [...document.querySelectorAll('section.sec')].map(s => s.id);
          const tables = [...document.querySelectorAll('table')].map(t => {
            const w = t.parentElement; return { scrollable: getComputedStyle(w).overflowX, fits: t.scrollWidth <= w.clientWidth };
          });
          // text measure: running prose line length
          const p = document.querySelector('.card.rough p');
          const chW = (() => { const s = document.createElement('span'); s.textContent = '0'.repeat(100);
            s.style.cssText = 'position:absolute;visibility:hidden;font:' + getComputedStyle(p).font;
            document.body.appendChild(s); const w = s.getBoundingClientRect().width / 100; s.remove(); return w; })();
          return { revealed, total, roughApplied, secs,
                   tables, proseCh: Math.round(p.getBoundingClientRect().width / chW),
                   h1: getComputedStyle(document.querySelector('.head h1')).fontSize,
                   svgCount: document.querySelectorAll('svg').length,
                   filterDefs: document.querySelectorAll('filter').length };
        });
        if (struct.revealed === struct.total) P(`structure: reveal script ran — ${struct.revealed}/${struct.total} .rv elements visible`);
        else F(`structure: only ${struct.revealed}/${struct.total} .rv elements revealed`);
        if (/url\(.*#rough.*\)/.test(struct.roughApplied)) P(`structure: hand-drawn border filter applied → ::before filter = ${struct.roughApplied}`);
        else F(`structure: rough filter NOT applied (::before filter = ${struct.roughApplied})`);
        const want = ['s1','s2','s3','s4','s5'];
        if (JSON.stringify(struct.secs) === JSON.stringify(want)) P(`structure: 5 sections present in order ${struct.secs.join(', ')}`);
        else F(`structure: sections = ${JSON.stringify(struct.secs)}`);
        P(`structure: ${struct.svgCount} inline SVGs, ${struct.filterDefs} filter defs, h1 ${struct.h1}`);
        N(`structure: prose measure ≈ ${struct.proseCh} characters per line`);
        struct.tables.forEach((t, i) => {
          if (t.scrollable === 'auto' || t.scrollable === 'scroll') P(`structure: table ${i + 1} wrapper overflow-x=${t.scrollable}${t.fits ? ' (fits, no scroll needed)' : ' (scrolls internally)'}`);
          else F(`structure: table ${i + 1} wrapper overflow-x=${t.scrollable}`);
        });

        // anchor navigation
        const nav = await page.evaluate(async () => {
          const a = document.querySelector('.agenda a[href="#s4"]'); a.click();
          await new Promise(r => setTimeout(r, 400));
          const el = document.getElementById('s4');
          return { top: Math.round(el.getBoundingClientRect().top) };
        });
        if (Math.abs(nav.top) < 120) P(`interaction: in-page anchor #s4 scrolls target into view (top offset ${nav.top}px)`);
        else F(`interaction: anchor #s4 landed at ${nav.top}px`);
        await page.evaluate(() => window.scrollTo(0, 0));
        await page.waitForTimeout(200);

        // focus visibility
        const focus = await page.evaluate(() => {
          const a = document.querySelector('.agenda a'); a.focus();
          const cs = getComputedStyle(a);
          return { outline: cs.outlineStyle + ' ' + cs.outlineWidth + ' ' + cs.outlineColor, active: document.activeElement === a };
        });
        if (focus.active && focus.outline && !/none/.test(focus.outline)) P(`a11y: focused link shows outline → ${focus.outline}`);
        else N(`a11y: focus outline on .agenda a = ${focus.outline} (browser default applies without :focus-visible heuristic)`);
      }

      // screenshots (desktop full page for the two system states)
      if (v.key === 'desktop' && (t.key === 'system-light' || t.key === 'system-dark')) {
        await page.screenshot({ path: path.join(OUT, `${tag}-full.png`), fullPage: true });
        await page.screenshot({ path: path.join(OUT, `${tag}-above-fold.png`) });
      }
      if (v.key === 'mobile' && t.key === 'system-light') {
        await page.screenshot({ path: path.join(OUT, `${tag}-above-fold.png`) });
      }
      await ctx.close();
    }
  }

  // ── cross-theme: light and dark must actually differ, and stamps must win ──
  const sl = bgByTheme['system-light-desktop'].bg, sd = bgByTheme['system-dark-desktop'].bg;
  const tl = bgByTheme['stamp-light-desktop'].bg, td = bgByTheme['stamp-dark-desktop'].bg;
  if (sl !== sd) P(`theming: light and dark render different grounds (${sl} vs ${sd})`);
  else F(`theming: light and dark identical (${sl}) — dark tokens not applying`);
  if (tl === sl) P(`theming: data-theme="light" overrides a dark OS (${tl})`);
  else F(`theming: data-theme="light" did not win — got ${tl}, expected ${sl}`);
  if (td === sd) P(`theming: data-theme="dark" overrides a light OS (${td})`);
  else F(`theming: data-theme="dark" did not win — got ${td}, expected ${sd}`);

  await browser.close();

  console.log('='.repeat(78));
  console.log('RUNTIME VERIFICATION (headless Chromium)');
  console.log('='.repeat(78));
  const pass = results.filter(r => r[0] === 'PASS').length;
  const fail = results.filter(r => r[0] === 'FAIL');
  for (const [k, m] of results) console.log(`  ${k}  ${m}`);
  console.log('-'.repeat(78));
  console.log(`${pass} passed, ${results.filter(r => r[0] === 'NOTE').length} notes, ${fail.length} failed`);
  process.exit(fail.length ? 1 : 0);
})();

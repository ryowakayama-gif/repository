const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const fs = require('fs'), path = require('path');
(async () => {
  const dir = process.argv[2];
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome', args: ['--no-sandbox'] });
  const p = await b.newPage();
  for (const f of fs.readdirSync(dir).filter(x => x.endsWith('.svg')).sort()) {
    await p.goto('file://' + path.resolve(dir, f));
    const res = await p.evaluate(() => {
      const svg = document.querySelector('svg'); const vb = svg.viewBox.baseVal; const out = [];
      for (const t of svg.querySelectorAll('text')) {
        const bb = t.getBBox();
        if (bb.x + bb.width > vb.width - 6) {
          const s = t.textContent.trim();
          const avail = vb.width - 20 - bb.x;
          out.push({ s, len: s.length, w: Math.round(bb.width), avail: Math.round(avail),
                     keep: Math.floor(s.length * avail / bb.width) });
        }
      }
      return out;
    });
    res.forEach(o => console.log(`${f}\t現${o.len}字/${o.w}px\t枠${o.avail}px\t上限${o.keep}字\t${o.s.slice(0,30)}`));
  }
  await b.close();
})();

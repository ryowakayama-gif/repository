const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const fs = require('fs'), path = require('path');
(async () => {
  const dir = process.argv[2];
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome', args: ['--no-sandbox'] });
  const p = await b.newPage();
  for (const f of fs.readdirSync(dir).filter(x => x.endsWith('.svg')).sort()) {
    await p.goto('file://' + path.resolve(dir, f));
    const res = await p.evaluate(() => {
      const svg = document.querySelector('svg'), vb = svg.viewBox.baseVal;
      const rects = [...svg.querySelectorAll('rect')].map(r => r.getBBox())
        .filter(r => r.width > 40 && r.width < vb.width - 20);
      const out = [];
      for (const t of svg.querySelectorAll('text')) {
        const bb = t.getBBox();
        if (bb.width < 1) continue;
        const cx = bb.x + 2, cy = bb.y + bb.height / 2;
        const holders = rects.filter(r => r.x <= cx && cx <= r.x + r.width &&
                                          r.y <= cy && cy <= r.y + r.height);
        if (!holders.length) continue;
        const h = holders.reduce((a, c) => (c.width < a.width ? c : a));
        const over = Math.round(bb.x + bb.width - (h.x + h.width - 5));
        if (over > 0) out.push({ s: t.textContent.trim().slice(0, 22), over,
                                 cap: Math.floor(t.textContent.trim().length * (h.width - 12) / bb.width) });
      }
      return out;
    });
    console.log(`${f}\tはみ出し ${res.length}件`);
    res.slice(0, 8).forEach(o => console.log(`    +${o.over}px 上限${o.cap}字  ${o.s}`));
  }
  await b.close();
})();

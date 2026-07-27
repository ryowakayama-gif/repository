const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const fs = require('fs'), path = require('path');
(async () => {
  const dir = process.argv[2];
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
                                    args: ['--no-sandbox'] });
  const p = await b.newPage({ deviceScaleFactor: 2 });
  for (const f of fs.readdirSync(dir).filter(x => x.endsWith('.svg')).sort()) {
    await p.setViewportSize({ width: 1300, height: 900 });
    await p.goto('file://' + path.resolve(dir, f));
    const res = await p.evaluate(() => {
      const svg = document.querySelector('svg');
      const vb = svg.viewBox.baseVal;
      const out = [];
      for (const t of svg.querySelectorAll('text')) {
        const bb = t.getBBox();
        if (bb.x + bb.width > vb.width - 6 || bb.y + bb.height > vb.height - 2 || bb.x < 2) {
          out.push({ s: t.textContent.trim().slice(0, 28),
                     r: Math.round(bb.x + bb.width), b: Math.round(bb.y + bb.height) });
        }
      }
      return { w: vb.width, h: vb.height, over: out };
    });
    if (res.over.length) {
      console.log(`## ${f} (${res.w}x${res.h}) はみ出し ${res.over.length}件`);
      res.over.slice(0, 6).forEach(o => console.log(`   右端${o.r} 下端${o.b}  ${o.s}`));
    } else {
      console.log(`OK ${f}`);
    }
  }
  await b.close();
})();

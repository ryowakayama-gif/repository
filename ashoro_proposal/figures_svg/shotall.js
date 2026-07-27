const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const fs = require('fs'), path = require('path');
(async () => {
  const dir = process.argv[2], out = process.argv[3];
  fs.mkdirSync(out, { recursive: true });
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome', args: ['--no-sandbox'] });
  const p = await b.newPage({ deviceScaleFactor: 2 });
  await p.setViewportSize({ width: 1300, height: 900 });
  for (const f of fs.readdirSync(dir).filter(x => x.endsWith('.svg')).sort()) {
    await p.goto('file://' + path.resolve(dir, f));
    const el = await p.$('svg');
    await el.screenshot({ path: path.join(out, f.replace('.svg', '.png')) });
  }
  await b.close();
})();

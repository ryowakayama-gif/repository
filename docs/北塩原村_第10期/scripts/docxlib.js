// docx ライブラリの置き場所を1か所で決める。
//
// これまで各組立てスクリプトが '/tmp/node_modules/docx' をじか書きしていた。
// /tmp は作業環境が作り直されると消えるため、復元したリポジトリでは
// 3本の組立てがいずれも止まる。本文そのものは版管理にあるので失われないが、
// 「組み直せない」状態になる。
//
// 探す順番
//   1 このリポジトリの node_modules（package.json に版を固定している）
//   2 /tmp/node_modules（いまの作業環境に置かれているもの）
// どちらにもなければ、入れ方を示して止まる。
const fs = require('fs');
const path = require('path');

const CAND = [
  path.join(__dirname, '..', 'node_modules', 'docx'),
  path.join(__dirname, '..', '..', '..', 'node_modules', 'docx'),
  '/tmp/node_modules/docx',
];

function load() {
  for (const p of CAND) {
    if (fs.existsSync(p)) return require(p);
  }
  console.error(
    'docx ライブラリが見つかりません。探した場所:\n  ' + CAND.join('\n  ') +
    '\n\n次のいずれかで入れてください。\n' +
    '  npm install --prefix docs/北塩原村_第10期\n' +
    '  npm install --prefix /tmp docx@9.7.1\n' +
    '版は docs/北塩原村_第10期/package.json で固定しています。');
  process.exit(1);
}

module.exports = load();

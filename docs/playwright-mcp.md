# Playwright MCP

このリポジトリのセッションでブラウザ操作ができるように、[@playwright/mcp](https://github.com/microsoft/playwright-mcp)
を MCP サーバーとして登録しています。`mcp__playwright__browser_*` ツール（ナビゲート、
スナップショット、クリック、フォーム入力、スクリーンショット、`evaluate` など）が使えます。

## 構成ファイル

| ファイル | 役割 |
| --- | --- |
| `.mcp.json` | `playwright` サーバーを stdio で登録 |
| `.claude/mcp/playwright-mcp.sh` | サーバー起動ラッパー（環境ごとの引数を解決） |
| `.claude/hooks/session-start.sh` | クラウドセッション起動時の前準備 |
| `.claude/settings.json` | 上記フックの登録と、`.mcp.json` の承認スキップ |

`.playwright-mcp/` に自動生成のスナップショット・コンソールログ・名前なしスクリーンショット
が出力されます（`.gitignore` 済み）。ファイル名を明示したスクリーンショットはリポジトリ
ルートに出るので、不要なら消してください。

## クラウドセッション固有の注意点

起動ラッパーとフックは、Claude Code on the web のコンテナ特有の事情を吸収しています。

- **`--headless --no-sandbox`**: ディスプレイがなく、root で動くため Chromium の
  サンドボックスが起動できない。ローカル実行時（root 以外）は付けない。
- **`--executable-path /opt/pw-browsers/chromium`**: `@playwright/mcp` が同梱する
  playwright-core が要求する Chromium ビルドは、このコンテナではダウンロードできない
  （`PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1`）。イメージ同梱の Chromium を使い回す。
  この変数が無い環境では Playwright 同梱のブラウザにフォールバックする。
- **CA の NSS 登録**: Chromium は `/etc/ssl/certs` ではなく独自の NSS ストアを見るため、
  セッションの HTTPS 傍受 CA を `~/.pki/nssdb` に入れないと、すべての `https://`
  ナビゲートが `ERR_CERT_AUTHORITY_INVALID` で失敗する。フックが `certutil` で登録する。
- **外向き通信は egress ポリシー配下**: 許可されていないホストは
  `ERR_TUNNEL_CONNECTION_FAILED` になる。TLS 設定の問題ではないので、
  `curl -sS "$HTTPS_PROXY/__agentproxy/status"` の `recentRelayFailures` で確認する。

## バージョン

`@playwright/mcp` は `0.0.83` 固定です。変更する場合は
`.claude/mcp/playwright-mcp.sh` と `.claude/hooks/session-start.sh` の
`PLAYWRIGHT_MCP_VERSION` 既定値を揃えてください。環境変数で上書きもできます。

```bash
PLAYWRIGHT_MCP_VERSION=0.0.82 .claude/mcp/playwright-mcp.sh --help
```

`vision`（座標ベース操作）や `pdf` ツールが必要な場合は、ラッパーの `args` に
`--caps vision,pdf` を足してください。

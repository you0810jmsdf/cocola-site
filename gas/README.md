# gas/ — Google Apps Script ソース管理ルール

このディレクトリは `clasp` 経由で GAS プロジェクト
（scriptId: `1zOZBorvfkKZAmCtJB9KDkQLXMxsTfGhUXWOdbxJ4lFLyA92Ts_gM0t7B`）
と双方向同期する作業ツリーです。

## 厳守事項（記録漏れ防止）

GAS 側でデプロイされている全ファイルは **必ずローカルに `clasp pull` し、リポジトリに版管理** すること。
GAS エディタ上だけで編集して終わらせない。changelog にも記録すること。

### 管理対象ファイル（最低限）

ローカルに存在しなければならない GAS ファイル一覧：

- `appsscript.json`
- `コード.js`
- `dao-points.js`
- `dao-proposals.html`
- `event-topics-normalized-submit.gs.js`
- `event-topics-row-repair.gs.js`
- `image-scan.gs.js`
- `image-upload.html`
- **`threads-daily-post.gs.js`** ← 毎日の Threads 自動投稿（千葉イベント／自分をつくる）。GAS 側に存在するが過去にローカル取得が漏れていた。要 `clasp pull`。
- `calendar-trigger.gs.js` ← changelog v176 記載。要 `clasp pull` 確認。

新規ファイルを GAS 側に追加した場合は、同時にこの README の一覧へ追記すること。

## 標準ワークフロー

```bash
# 認証（初回のみ／トークン失効時）
npm run gas:login

# GAS → ローカル取得（編集前に必ず実行）
npm run gas:pull

# ローカル → GAS デプロイ
npm run gas:push

# 状態確認
npm run gas:status
```

## キーワード設定の正規記録

毎日の Threads 自動投稿（自分をつくる）で参照する情報収集キーワードは、
`gas/threads-daily-post-keywords.json` を **正規記録（source of truth）** とする。

GAS 側 `threads-daily-post.gs.js` の `SELF_POST_KEYWORDS` 配列は、
このJSONと **常に一致** していなければならない。

キーワード追加・削除時の手順:

1. `gas/threads-daily-post-keywords.json` を更新（PR で版管理）
2. `npm run gas:pull` で最新の `threads-daily-post.gs.js` を取得
3. `SELF_POST_KEYWORDS` を JSON と同一内容に修正
4. `npm run gas:push` でデプロイ
5. GAS エディタで新バージョンを公開（デプロイ）
6. `changelog/index.html` に追記
7. 同一PR内でコミット

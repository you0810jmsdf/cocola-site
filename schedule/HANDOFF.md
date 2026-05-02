# COCoLa schedule Web App 引き継ぎ文書

Claude Code が本プロジェクトの実装を引き継ぐための前提共有ドキュメント。
本書を最初に読んだうえで、`schedule/` 配下のドラフトファイル群を起点に作業を進めること。

---

## 1. 概要

COCoLa の Google カレンダーに対して、メンバーがブラウザから予定登録・閲覧できる軽量 Web アプリ。
Google Apps Script (GAS) を Web App として公開し、COCoLa サイト (GitHub Pages) に iframe で埋め込む構成。

```
[ブラウザ: form.html / schedule.html]
        │  (POST / GET)
        ▼
[GAS Web App: cocola_calendar.gs]
        │  CalendarApp API
        ▼
[Google カレンダー: cocola.project@gmail.com]
        │  iframe (任意)
        ▼
[COCoLa LP: index.html / 任意ページに埋め込み]
```

---

## 2. 確定仕様

| 項目 | 内容 |
|---|---|
| ランタイム | Google Apps Script (V8) |
| 公開形態 | Web App (`doGet` / `doPost`) |
| 認証 | 「全員」アクセス可。書き込みは合言葉 (`ADMIN_PASS`) で軽くゲート |
| カレンダー ID | `cocola.project@gmail.com` |
| データ保管 | Google カレンダー本体 (シートには持たない) |
| タイムゾーン | `Asia/Tokyo` |
| 言語 | 日本語 UI |
| 埋め込み先 | COCoLa LP の任意ページ (iframe) |

### COCoLa サイト配色リファレンス

`styles.css` のトークンに合わせて UI を統一する。

| 用途 | 変数 | 値 |
|---|---|---|
| アクセント (主) | `--green` | `#1d9e75` |
| アクセント (濃) | `--green-dark` | `#0f6e56` |
| サブ (情報) | `--blue` | `#1fa7c9` |
| 強調 (CTA) | `--orange` | `#ef7f4d` |
| 注意 | `--yellow` | `#f2b134` |
| 背景 | `--bg` | `#f7faf8` |
| テキスト | `--text` | `#183128` |
| 補助文字 | `--muted` | `#607068` |
| 角丸 | `--radius-md` / `--radius-lg` | `18px` / `24px` |

---

## 3. ファイル構成

`schedule/` 配下に以下を配置 (Claude Code はこの構成を前提に編集する)。

```
schedule/
├── HANDOFF.md              ← 本書
├── cocola_calendar.gs      ← GAS 本体 (doGet / doPost / Calendar 操作)
├── form.html               ← 予定登録フォーム
├── schedule.html           ← 予定一覧 / カレンダー表示
└── セットアップ手順.txt    ← clasp / Web App デプロイ手順
```

GAS プロジェクト側 (clasp) では拡張子は `.gs` / `.html` で扱う。
リポジトリ上ではドラフトとして同名で管理する。

---

## 4. 設定値

`cocola_calendar.gs` 冒頭にスクリプトプロパティを参照する形で定義する。

| キー | 値 | 備考 |
|---|---|---|
| `CALENDAR_ID` | `cocola.project@gmail.com` | 既存 `gas/コード.js:10` と同一 |
| `ADMIN_PASS` | (スクリプトプロパティ) | 平文埋め込み禁止 |
| `DEFAULT_DURATION_MIN` | `60` | 終了時刻未指定時の既定 |
| `LIST_RANGE_DAYS` | `30` | `getUpcomingEvents` の取得範囲 |
| `TIMEZONE` | `Asia/Tokyo` | `Session.getScriptTimeZone()` ではなく明示 |

デプロイ:

- 実行ユーザー: 自分
- アクセス可能ユーザー: 全員
- 新バージョンはデプロイ ID を維持して「新しいバージョン」で更新する (URL 不変)

---

## 5. 主要関数仕様

### `doGet(e)`

| パラメータ | 値 | 戻り値 |
|---|---|---|
| `mode=list` | なし | `{ ok, events: [{id,title,start,end,location,description,url}] }` JSON |
| `mode=ical` | なし | text/calendar (将来拡張用、初版は省略可) |
| (なし) | なし | `schedule.html` をレンダリング (`HtmlService.createTemplateFromFile`) |

### `doPost(e)` → `createEvent_(payload)`

入力 (form-encoded または JSON):

```
{
  pass: string,         // ADMIN_PASS と一致
  title: string,        // 必須
  start: ISO8601,       // 必須 (例: "2026-05-10T19:00:00+09:00")
  end:   ISO8601,       // 任意 (未指定は start + DEFAULT_DURATION_MIN)
  location: string,     // 任意
  description: string,  // 任意
  url: string           // 任意 (description 末尾に追記)
}
```

戻り値: `{ ok: true, id: <eventId> }` / `{ ok: false, error: <msg> }`

### `getUpcomingEvents_()`

`CalendarApp.getCalendarById(CALENDAR_ID).getEvents(now, now + LIST_RANGE_DAYS)` を返却用にマップ。
`start` / `end` は ISO 文字列で返す。

---

## 6. UI 設計指針

- ベース背景は `--bg`、カードは白 + `--shadow`、角丸は `--radius-lg`
- ボタンは `--green` 塗り / 押下時 `--green-dark`
- 入力欄ボーダーは `--border`、フォーカスで `--green`
- フォントは既存サイトと同じシステムフォントスタック (`styles.css` 参照)
- iframe 埋め込み前提のため、`html/body` のマージン 0、最大幅 720px の中央寄せ
- スマホファースト。`@media (min-width: 720px)` で 2 カラム化

---

## 7. 拡張候補

短期 (MVP 後すぐ):

1. 予定の編集・削除 (`mode=update` / `mode=delete`)
2. カテゴリタグ (`mono` / `machi` / `tsunagari` / `shoku` / `jibun` / `tomoni`) の付与とフィルタ
3. iCal フィード (`mode=ical`) で外部カレンダー購読

中期:

4. COCoLa メンバー認証 (`member-app/` の仕組みと連携)
5. TOPICS (`gas/コード.js` の TOPICS シート) との双方向連携
6. 印西イベント (`INZAI_SHEET`) の自動取り込み結果をカレンダー表示

---

## 8. 制約・注意点

- GAS の 1 日あたり Calendar 書き込み上限 (実質 5,000 events/day) を超える運用はしない
- `doPost` は CSRF 対策が弱い → 合言葉 + Origin チェックで最低限ガード
- iframe 埋め込み時は `X-Frame-Options` を GAS 側で制御できない (GAS の制約) ため、
  `IFRAME` モードで `HtmlService.XFrameOptionsMode.ALLOWALL` を明示
- 個人情報を `description` に入れない運用ルールをフォーム文言で明示
- スクリプトプロパティに `ADMIN_PASS` を保存し、コード中にハードコードしない

---

## 9. 動作確認チェックリスト

- [ ] `clasp push` 成功
- [ ] Web App URL に GET アクセス → `schedule.html` が表示される
- [ ] `?mode=list` で JSON が返る
- [ ] `form.html` から予定登録 → Google カレンダーに反映される
- [ ] 合言葉なし / 誤りで POST → `{ ok:false }` が返る
- [ ] スマホ幅 (375px) でレイアウト崩れなし
- [ ] iframe 埋め込み (LP 内) で表示・操作できる
- [ ] タイムゾーンが JST で記録される

---

## 10. 関連リソース

- COCoLa LP: `index.html` / `styles.css`
- 既存 GAS (TOPICS): `gas/コード.js` (CALENDAR_ID 等の定数を流用)
- メンバーアプリ: `member-app/コード.js` (将来の認証連携先)
- DAO 投票: `dao/index.html` (UI トーンの参考)
- イベント一覧: `events/index.html` (一覧表示の参考)

---

## 11. やりとり履歴 (設計時の確定事項)

1. データストアはシートではなく Google カレンダー本体に一本化する
2. CALENDAR_ID は既存 TOPICS と同じ `cocola.project@gmail.com` を共用
3. 認証は「合言葉方式」で開始し、メンバー認証は中期課題に回す
4. 公開は GAS Web App + iframe 埋め込み (Vercel 等は使わない)
5. UI は COCoLa LP の配色トークンを流用し、新規パレットは作らない
6. タイムゾーンは `Asia/Tokyo` 固定 (海外利用は想定しない)
7. 一覧の既定取得範囲は「今日から 30 日先まで」
8. ファイル名・関数名は日本語を避け、コメント・UI 文言は日本語

---

> Claude Code への指示テンプレート
>
> 「`schedule/HANDOFF.md` を読んでから、`cocola_calendar.gs` / `form.html` / `schedule.html` /
> `セットアップ手順.txt` を作成してください。第 9 章のチェックリストを満たすこと。」

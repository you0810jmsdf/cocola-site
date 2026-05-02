# COCoLa サイト運営補助 — Claude Code 向けガイド

このリポジトリで作業する Claude Code が、最初に読む前提共有文書。
ユーザー向けの説明は `README.md`、移行設計は `GITHUB_PAGES移行設計.md`、
個別プロジェクトの仕様は各ディレクトリの `HANDOFF.md` を参照する。

---

## 1. プロジェクト概要

COCoLa (印西を拠点とした地域プロジェクト) の公式サイト。
Google Sites から GitHub Pages 静的サイトへ移行中で、
裏側に複数の Google Apps Script (GAS) Web App が稼働している。

公開 URL: `https://you0810jmsdf.github.io/cocola-site/`

---

## 2. リポジトリ構成

```
cocola-site/
├── index.html                  ← LP 本体 (トップページ)
├── styles.css                  ← LP 共通スタイル (配色トークン定義)
├── categories/<6種>/index.html ← 「○○をつくる」カテゴリページ
│     mono / machi / tsunagari / shoku / jibun / tomoni
├── apps/index.html             ← Apps 入口
├── events/index.html           ← イベント一覧 (GAS 連携)
├── dao/index.html              ← DAO 投票
├── members/                    ← メンバー一覧 (data.json のみ、表示は他ページ)
├── member-app/                 ← メンバー向けアプリ (GAS Web App + html)
├── changelog/index.html        ← 更新履歴
├── calendar/index.html         ← イベントカレンダー (月グリッド表示)
├── schedule/                   ← スケジュール Web App (GAS、HANDOFF.md 参照)
│   └── data.json               ← schedule GAS のミラー (Actions が更新)
├── gas/                        ← TOPICS 系 GAS のソース (clasp 連携先)
│   ├── コード.js                ← TOPICS Web App 本体
│   ├── dao-points.js
│   ├── image-scan.gs.js
│   └── ...
├── docs/                       ← 利用者向けマニュアル
├── scripts/                    ← ビルド・運用スクリプト
└── .github/workflows/          ← GitHub Actions
```

---

## 3. アーキテクチャ

```
┌────────────── ブラウザ ──────────────┐
│  index.html / categories/* / events/* │
│  fetch('./events/data.json')          │
└──────────────────┬───────────────────┘
                   │ 静的配信
┌──────────────────▼───────────────────┐
│        GitHub Pages (本リポジトリ)     │
│    data.json は Actions で自動更新     │
└──────────────────┬───────────────────┘
                   │ 5分ごと cron
┌──────────────────▼───────────────────┐
│       GitHub Actions (.github/)       │
│   curl GAS → data.json をコミット      │
└──────────────────┬───────────────────┘
                   │ HTTPS
┌──────────────────▼───────────────────┐
│        GAS Web Apps (cocola.project)  │
│   gas/コード.js  : TOPICS / events     │
│   member-app/    : メンバー管理        │
│   schedule/      : カレンダー (新規)   │
└──────────────────┬───────────────────┘
                   │ Calendar / Sheets API
              Google ドライブ
```

ブラウザは GAS を直接呼ばない (CORS と速度のため)。
GAS の出力は GitHub Actions が定期的に取得して `*/data.json` に書き出す。

---

## 4. コーディング規則

### 4.1 配色トークン (`styles.css` 冒頭で定義)

新規 UI を作るときは必ずこのトークンを使う。新色を勝手に追加しない。

| 用途 | 変数 | 値 |
|---|---|---|
| アクセント (主) | `--green` | `#1d9e75` |
| アクセント (濃) | `--green-dark` | `#0f6e56` |
| サブ (情報) | `--blue` | `#1fa7c9` |
| 強調 (CTA) | `--orange` | `#ef7f4d` |
| 注意 | `--yellow` | `#f2b134` |
| 補助 | `--violet` / `--pink` | `#6d79d8` / `#e56b8f` |
| 背景 | `--bg` | `#f7faf8` |
| 背景 (柔) | `--bg-soft` | `#eef8f3` |
| テキスト | `--text` | `#183128` |
| 補助文字 | `--muted` | `#607068` |
| 角丸 | `--radius-md` / `--lg` / `--xl` | `18` / `24` / `32` px |

### 4.2 カテゴリ ID

URL・data.json・GAS 側で揃える。日本語表示は GAS 側 `CATEGORY_MAP` 参照。

```
mono       → ものをつくる
machi      → まちをつくる
tsunagari  → つながりをつくる
shoku      → 食をつくる
jibun      → 自分をつくる
tomoni     → ともにつくる
```

### 4.3 HTML / CSS / JS

- フレームワーク非導入。素の HTML/CSS/JS で書く。
- `index.html` のスタイルはインライン `<style>` と `styles.css` が混在。新規追加は `styles.css` を優先。
- スマホファースト。`@media (min-width: 720px)` でデスクトップ拡張。
- 文字コード UTF-8、改行 LF。日本語 UI、コメントも日本語可。
- フォントは OS 標準のシステムフォントスタックを使う (Web フォントは追加しない)。

### 4.4 ファイル命名

- HTML: ディレクトリ + `index.html` (URL がきれいになる)
- データ: `<dir>/data.json` (Actions が更新する規約)
- GAS のソース: `gas/` 配下に `.js` または `.gs.js`、HTML は `.html`

---

## 5. GAS 開発フロー (clasp)

ルートの `.clasp.json` は **TOPICS 系 (`gas/`)** に紐付いている。
他の GAS プロジェクト (`member-app/`、`schedule/`) は別 scriptId で独立管理する。

```bash
# TOPICS 系
npm run gas:status   # 差分確認
npm run gas:pull     # GAS → ローカル
npm run gas:push     # ローカル → GAS
npm run gas:open     # GAS エディタを開く
```

GAS 側のシークレット (`ADMIN_PASS` など) はコードに埋め込まず、
スクリプトプロパティで管理する。コードで参照するときは `PropertiesService`。

---

## 6. GitHub Actions の役割

| ワークフロー | 頻度 | 役割 |
|---|---|---|
| `update-events-data.yml` | 5 分 | GAS から印西イベントを取得して `events/data.json` 更新 |
| `update-schedule-data.yml` | 5 分 | schedule GAS から予定を取得して `schedule/data.json` 更新 (`GAS_SCHEDULE_URL` 未設定時はスキップ) |
| `update-dao-data.yml` | 5 分 | DAO データ更新 |
| `update-members-data.yml` | 5 分 | メンバーデータ更新 |
| `sync-event-sources.yml` | 毎日 JST 9:00 | 外部イベントソースの巡回トリガー |
| `health-check.yml` | **毎時** | サイト自動巡回・異常時 Issue 通知 (詳細は §7) |

`data.json` は **Actions による自動コミット** (`[skip ci]` 付き)。
**手で編集しない** — 次回の Actions 実行で上書きされる。

---

## 7. サイト自動巡回と管理者通知

`health-check.yml` がサイト全体を毎時巡回し、異常を検知したら GitHub Issue で通知する。
Issue は `auto-monitor` ラベルで購読者にメール通知され、これが「管理者通知」のチャネル。

### 7.1 巡回項目

| 区分 | 観点 | 異常条件 |
|---|---|---|
| GAS 疎通 | TOPICS / DAO / Members の Web App | エンドポイントが想定 JSON を返さない |
| Actions 健全性 | update-* ワークフローの実行履歴 | 直近 30 分以内に成功実行がない |
| データ鮮度 | `*/data.json` の `_generatedMs` | 30 分以上更新されていない |
| HTML 到達性 | `/`、6 カテゴリ、`/apps/`、`/events/`、`/dao/`、`/changelog/` | HTTP 200 以外 |

### 7.2 通知の挙動

- **異常発生時**: `auto-monitor` ラベル付き Issue を新規作成。既に open Issue があれば追記コメント
- **正常復旧時**: open Issue を自動クローズし「✅ 全項目正常に復旧しました」を残す
- **頻度抑制**: 既存 Issue へのコメントで通知が重複するため、リポジトリ Watch を「Issues only」推奨

### 7.3 巡回項目を追加・変更するとき

`health-check.yml` の `Aggregate and judge` ステップ内 Python に観点を追記する。
新しい区分を足したら本書 §7.1 の表も更新する (実装と説明の乖離を防ぐ)。

例: `schedule/` Web App をデプロイしたら以下を追加:
- `env` に `GAS_SCHEDULE` を追加
- `mode=list` を curl して JSON 妥当性を判定
- 結果を `results['GAS[schedule]']` に格納

### 7.4 通知を一時停止したいとき

`workflow_dispatch` 中の保守時間帯は、`health-check.yml` を一時的に無効化する代わりに
リポジトリ Settings → Actions → Disable workflow を使う。
コードを書き換えてコミットを汚さない。

---

## 8. やってはいけないこと / 注意点

- `events/data.json` `dao/data.json` `members/data.json` を手動編集しない (Actions が上書き)
- GAS のスクリプト ID や Web App URL を README に書かない (`.github/workflows/` 内に集約)
- 新しい色・フォント・フレームワークを断りなく導入しない
- カテゴリ ID (`mono` 等) を変えない (GAS / data.json / URL の整合が壊れる)
- `index.html` のインラインスクリプトに API キーや合言葉を書かない
- GAS 側で個人情報を平文ログに出さない
- Google Sites 側の本番運用に影響を与える作業 (リンク差し替え等) は確認してから

---

## 9. 作業前のチェックリスト

新規実装・改修に入る前に確認する。

- [ ] 同じ機能が既に `gas/` または既存ページに実装されていないか grep する
- [ ] 触るデータが Actions 自動更新の対象でないか確認する
- [ ] 配色トークン・カテゴリ ID を再利用できるか確認する
- [ ] GAS 側の変更があるなら clasp の対象プロジェクトを意識する
- [ ] 関連 `HANDOFF.md` (例: `schedule/HANDOFF.md`) を読む

---

## 10. 関連ドキュメント

- `README.md` — リポジトリの公開向け説明
- `GITHUB_PAGES移行設計.md` — Google Sites からの移行設計
- `schedule/HANDOFF.md` — schedule Web App の仕様
- `member-app/member.md` — member-app の仕様
- `docs/member-actions-video-guide.md` — メンバー操作の動画ガイド
- `gas/コード.js` 冒頭 — 共通定数 (`CALENDAR_ID`、`SHEET_ID`、カテゴリマップ等)

---

## 11. コミット・PR の作法

- ブランチ名は `claude/<目的>-<短いID>` 形式 (例: `claude/create-handoff-document-27Ysx`)
- 自動更新コミット (`Auto-update * data [skip ci]`) には触らない
- 人手のコミットメッセージは「動詞 + 対象」を 1 行目に。本文に背景を 1〜3 行
- PR は明示的に依頼されたときだけ作る
- `data.json` の差分を含むコミットは避け、コードと分けて push する

# COCoLa メンバー紹介ページ 運用方針

最終更新: 2026-04-06

---

## ファイル保存場所

| ファイル | パス |
|---|---|
| この運用方針 | C:\Users\nsfactory\OneDrive\COCoLa\サイト運営\member\member.md |
| GAS（メイン） | C:\Users\nsfactory\OneDrive\COCoLa\サイト運営\member\cocola_form_update.gs |
| HTML（CSV版・予備） | C:\Users\nsfactory\OneDrive\COCoLa\サイト運営\member\cocola_members.html |
| clasp設定 | C:\Users\nsfactory\OneDrive\COCoLa\サイト運営\member\.clasp.json |

---

## 関連リンク

| 名称 | URL |
|---|---|
| メンバーページ（Googleサイト） | https://sites.google.com/view/cocola-project/member |
| 入会・修正フォーム | https://docs.google.com/forms/d/e/1FAIpQLSfZYSm7QglMdTIQjfRFzbTuRWxnMlXJfpby3uHEAGTQ-6lc3A/viewform |
| スプレッドシート（フォーム回答） | https://docs.google.com/spreadsheets/d/13FwQyLeZK_Kgvl6LqY4UhzMkxJJUkLLi01vhA2jwUcc |
| GAS ウェブアプリURL | https://script.google.com/macros/s/AKfycbxQ7gBHJQFVZEXcyisJdiL7xa64RcKzYjzBZLEGlkelSvxXCfPPHTW5eAVHB09aue1p/exec |
| GAS スクリプトID | 1eexKs-ZcEd01ywddBs9vBYSiQ4QvcnSX2SLuamoC4vvkgyvHomouw6Q5 |
| 写真保存フォルダ（Drive） | https://drive.google.com/drive/folders/1MoPb8B3kaVocHWJ2l7Fp7b-K_KodJLfv |

---

## システム構成

```
Googleフォーム（入会 / 情報修正 兼用）
    ↓ 回答
スプレッドシート（メンバー登録情報）
    ↓ onFormSubmit トリガーが自動上書き処理
    ↓ 写真を指定フォルダへ移動＋共有設定を自動公開
GAS doGet（スプレッドシートを直接読んでHTML生成）
    ↓ iframeで埋め込み
Googleサイト メンバーページ
```

---

## スプレッドシート列マッピング（確定版）

| 列 | 番号（0始まり） | 内容 |
|---|---|---|
| A | 0 | タイムスタンプ |
| B | 1 | メールアドレス（フォーム送信者・非公開） |
| C | 2 | このフォームの目的（新規入会 / 情報の修正・更新） |
| D | 3 | 希望するニックネーム |
| E | 4 | 本名（公開可能な方のみ） |
| F | 5 | 職業（公開可能な方のみ） |
| G | 6 | 公開顔写真（GoogleドライブURL） |
| H | 7 | 公開可能なメールアドレス |
| I | 8 | SNS① |
| J | 9 | SNS② |
| K | 10 | 関連ホームページ① |
| L | 11 | 関連ホームページ② |
| M | 12 | 保有資格 |
| N | 13 | 好きなこと・得意なこと |
| O | 14 | 今後やりたいこと |
| P | 15 | コメント（自己PR・COCoLaに期待すること等） |
| Q | 16 | カテゴリ（チェックボックス・複数選択可）★追加列 |
| R | 17 | 印西市民アカデミーの期別 |
| S | 18 | COCoLa以外で所属している市民活動団体名 |
| T | 19 | 所属団体1のホームページ・SNS等のURL |
| U | 20 | COCoLa以外で所属している市民活動団体名2 |
| V | 21 | 所属団体2のホームページ・SNS等のURL |
| W | 22 | 団体内役職 |

---

## Apps Script CONFIG 設定値

```javascript
const CONFIG = {
  sheetName     : 'フォームの回答 1',
  colPurpose    : 3,                   // C列（1始まり）
  colNickname   : 4,                   // D列（1始まり）
  updateLabel   : '情報の修正・更新',
  skipEmpty     : true,
  adminEmail    : 'you0810jmsdf@gmail.com',
  formUrl       : 'https://docs.google.com/forms/d/e/1FAIpQLSfZYSm7QglMdTIQjfRFzbTuRWxnMlXJfpby3uHEAGTQ-6lc3A/viewform',
  photoFolderId : '1MoPb8B3kaVocHWJ2l7Fp7b-K_KodJLfv',  // 写真保存先フォルダ
};
```

---

## WEBアプリ機能一覧（v2）

| 機能 | 内容 |
|---|---|
| レスポンシブ | スマホ:1列 / PC:画面幅に応じて自動 |
| カテゴリフィルター | Q列の値で6カテゴリに絞り込み |
| 全文検索 | 名前・スキル・コメント等をリアルタイム検索 |
| 詳細モーダル | カードクリックで全情報をポップアップ表示 |
| プロフィール写真 | drive.google.com/thumbnail 形式に自動変換 |
| 写真自動移動 | フォーム送信時に指定フォルダへ移動＋共有設定を自動公開 |
| デフォルト画像 | 写真未登録は人影シルエット（グレー）を表示 |
| SNSラベル自動判定 | URLからInstagram/Facebook/X/LINE等を自動認識 |
| スキル見出し | カード上に「好きなこと・得意なこと」ラベルを表示 |
| メーリングリストコピー | パスワード認証後にH列＋B列の全メールを収集してコピー |

---

## カテゴリ（Q列）の選択肢

フォームにチェックボックス形式で追加。以下の6つ（＋その他）：

- ものをつくる
- まちをつくる
- つながりをつくる
- 食をつくる
- 自分をつくる
- ともにつくる
- その他（管理者がQ列を手動修正）

フィルターボタンのラベルと完全一致させること。

---

## メーリングリストコピーのパスワード

- パスワード：`cocola2026`
- H列（公開）＋B列（非公開）の両方を収集して重複排除
- ※クライアントサイド認証のため、ソースを見れば確認可能。誤操作防止用途

---

## clasp（コード更新ツール）設定

Node.js + clasp をインストール済み。

```
# ログイン（初回のみ）
clasp login

# フォルダに移動
cd "C:\Users\nsfactory\OneDrive\COCoLa\サイト運営\member"

# GASに反映
clasp push
```

その後 Apps Scriptエディタで再デプロイ：
`デプロイ → 既存のデプロイを管理 → ✏️ → 新しいバージョン → デプロイ`

---

## GAS 初回デプロイ手順

1. スプレッドシート → 拡張機能 → Apps Script
2. `cocola_form_update.gs` の内容を全選択して貼り付け → 保存（Ctrl+S）
3. 「デプロイ」→「新しいデプロイ」
4. 種類：ウェブアプリ / 実行ユーザー：自分 / アクセス：全員
5. デプロイ → 認証許可 → URLをコピー

### onFormSubmit トリガー設定

| 項目 | 設定値 |
|---|---|
| 実行する関数 | onFormSubmit |
| 実行するデプロイ | Head |
| イベントのソース | スプレッドシートから |
| イベントの種類 | フォーム送信時 |

---

## Googleサイトへの埋め込み

1. Googleサイト編集画面 → 挿入 → 埋め込む → URLを埋め込む
2. GAS ウェブアプリURLを貼り付け
3. ※Googleサイト側に表題があるため、GAS側のヘッダーは削除済み

---

## メンバーが情報を修正したいとき

1. 入会フォームを開く
2. 「このフォームの目的」で **「情報の修正・更新」** を選択
3. ニックネームは**登録済みと完全に同じ表記**で入力
4. 変更したい項目だけ入力（変更しない項目は空欄でOK）
5. 送信 → Apps Scriptが自動で既存データを上書き
6. 管理者（ゆうさん）に通知メールが届く

### 注意事項
- ニックネームの表記が1文字でも違うと上書きされず新規行として追加される
- 空欄で送信した項目は元の値が維持される（削除はできない）
- 削除したい場合はスプレッドシートを直接編集する

---

## 今後の改善候補

- [ ] clasp login 完了 → clasp push → 再デプロイ
- [ ] 動作テスト：フォーム送信 → 写真自動移動・公開設定確認
- [ ] 動作テスト：メーリングリストコピー（パスワード認証）確認
- [ ] カード上の「好きなこと・得意なこと」ラベル表示確認（再デプロイ後）
- [ ] 管理者通知メールの動作確認

# freefree マイページ：キャッシュクリア（ログアウト）ボタンの追加

freefree（千葉ニュータウン事業主応援掲示板）のマイページで、一度ログインすると別のメールアドレスで入れなくなる問題を解決するため、`mypage.html` に2箇所修正を加える。

## 原因

`mypage.html` の以下のコードにより、ログイン情報が `sessionStorage` に保存され、別アカウントへの切替ができなくなっている。

- 184行目: `let myEmail = sessionStorage.getItem('myEmail') || '';`
- 273行目: `sessionStorage.setItem('myEmail', email);`
- 244-247行目: ページ読み込み時に `myEmail` があれば自動ログイン

## 修正手順

GASエディタで「千葉ニュータウン事業主応援掲示板」プロジェクトを開き、`mypage.html` を以下の2箇所修正する。

---

## 修正① 15行目あたり（HTMLヘッダー部）

### 探す場所（修正前）

```html
<div><a href="<?= webAppUrl ?>?page=index" target="_top">← 掲示板へ戻る</a></div>
```

### 置き換える内容（修正後）

```html
<div>
  <a href="<?= webAppUrl ?>?page=index" target="_top">← 掲示板へ戻る</a>
  <button onclick="logoutMypage()" style="margin-left:12px;padding:4px 10px;font-size:12px;background:#fff;border:1px solid #ccc;border-radius:4px;cursor:pointer">🔄 ログアウト（別アカウントへ）</button>
</div>
```

---

## 修正② 184行目の下（JavaScript部）

### 探す場所

```js
let myEmail = sessionStorage.getItem('myEmail') || '';
```

### この行のすぐ下に、以下の関数を追加

```js
function logoutMypage() {
  if (!confirm('ログアウトして別のメールアドレスでログインし直しますか？')) return;
  sessionStorage.removeItem('myEmail');
  location.reload();
}
```

---

## デプロイ

1. GASエディタで Ctrl+S（Mac は Cmd+S）で保存
2. 右上の「デプロイ」→「デプロイを管理」
3. 既存のデプロイの「✏️編集」アイコンをクリック
4. バージョンを「新バージョン」に変更
5. 「デプロイ」をクリック

## 動作確認

1. メールAでマイページにログイン
2. 「🔄 ログアウト（別アカウントへ）」ボタンを押す
3. 確認ダイアログで「OK」
4. ページがリロードされ、ログイン画面が表示される
5. メールBを入力 → 別アカウントでログインできる

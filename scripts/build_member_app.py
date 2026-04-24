#!/usr/bin/env python3
"""
GAS版 html.html → GitHub Pages版 member-app/html.html 変換スクリプト

使い方:
  python scripts/build_member_app.py

cocola-gas/html.html を Single Source of Truth として、
GASテンプレートタグを静的な値に置換してGitHub Pages版を生成する。
"""

import sys
from pathlib import Path

# ─── パス設定 ────────────────────────────────────────────────
ROOT      = Path(__file__).parent.parent.parent  # C:/AI/
GAS_SRC   = ROOT / "cocola-gas" / "html.html"
SITE_DEST = ROOT / "cocola-site" / "member-app" / "html.html"

# ─── テンプレートタグ → 静的値 の置換マップ ───────────────────
APP_URL  = "https://script.google.com/macros/s/AKfycbxQ7gBHJQFVZEXcyisJdiL7xa64RcKzYjzBZLEGlkelSvxXCfPPHTW5eAVHB09aue1p/exec"
FORM_URL = "https://docs.google.com/forms/d/1V_0080ht4b7iBG4kpckyWhqYmwNVtQkbSStas7fAeGs/viewform"
HOME_URL = "https://you0810jmsdf.github.io/cocola-site/index.html"

REPLACEMENTS = [
    # HTML属性内のテンプレートタグ
    ('href="<?= HOME_URL ?>"',  f'href="{HOME_URL}"'),
    ('href="<?= FORM_URL ?>"',  f'href="{FORM_URL}"'),
    # JS内のJSON.stringify形式（クォート付き文字列になる）
    (f'<?!= JSON.stringify(APP_URL) ?>',  f"'{APP_URL}'"),
    (f'<?!= JSON.stringify(FORM_URL) ?>', f"'{FORM_URL}'"),
    (f'<?!= JSON.stringify(HOME_URL) ?>', f"'{HOME_URL}'"),
]

def build():
    if not GAS_SRC.exists():
        print(f"ERROR: ソースが見つかりません: {GAS_SRC}", file=sys.stderr)
        sys.exit(1)

    text = GAS_SRC.read_text(encoding="utf-8")

    for old, new in REPLACEMENTS:
        if old in text:
            text = text.replace(old, new)
        else:
            print(f"WARN: 置換パターンが見つかりません: {old!r}")

    # 残存するテンプレートタグをチェック
    remaining = [line for line in text.splitlines() if "<?=" in line or "<?!=" in line]
    if remaining:
        print("WARN: 未処理のテンプレートタグが残っています:")
        for line in remaining:
            print(f"  {line.strip()}")

    SITE_DEST.write_text(text, encoding="utf-8")
    print(f"OK: {GAS_SRC} → {SITE_DEST}")
    print(f"    {len(text.splitlines())} 行")

if __name__ == "__main__":
    build()

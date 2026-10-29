# ことなら（kotonara.tech）ブランド PPTX — カスタマイズ手引き

公式サイト https://kotonara.tech/ の配色・トーンに合わせた PowerPoint テーマ一式です。

## ことテク2026 ハッカソン チーム発表テンプレート

成果発表会（発表8分＋デモ2分＋質疑3分）用の9枚構成テンプレートは `hackathon-2026/` にあります。
Claude Code で `/create-hackathon-slide <チームの成果物リポジトリのパスまたはURL>` を実行すれば、
成果物を読んで Markdown を埋め、このテーマの PPTX まで生成します。詳細は `hackathon-2026/README.md` を参照。

## テンプレート3種

| ファイル | 版 | メインの緑（accent1） | 「ことテク」ロゴ | フッター |
| --- | --- | --- | --- | --- |
| `kotonara-brand-template.pptx` | メイン | `#4D9622` | あり | 空欄 |
| `kotonara-brand-template-nologo-1.pptx` | ロゴなし版1 | `#4D9622` | なし | 空欄 |
| `kotonara-brand-template-nologo-2.pptx` | ロゴなし版2（旧メイン） | `#1A832A` | なし | 「ことテク2026 2daysハッカソン」 |

ふだんはメインを使います。イベント資料・営業資料のどちらにも使えるよう、フッターは空欄にしてあります
（［挿入］＞［ヘッダーとフッター］で入力）。メインとロゴなし版1は1枚ずつ5レイアウトの見本が入っています。

メインのロゴの位置:

| レイアウト | ロゴ |
| --- | --- |
| 表紙（タイトル） | 下部の帯の左に白いロゴ。キャッチコピー『ひらめきを、カタチに。』（13pt）はロゴの右で、下端をロゴにそろえる |
| タイトルとコンテンツ／表スライド／セクション見出し | 右上に緑のロゴ |
| クロージング（反転） | なし（背景が緑のため） |

ロゴはスライドではなくレイアウトに置いてあります。［ホーム］＞［新しいスライド］でレイアウトを選べばロゴも付きます。
PNG（4000×1412）と SVG の両方を埋め込んであり、PowerPoint 2016 以降は SVG で描画します。

`hackathon-2026/` のチーム発表テンプレートは、ロゴなし版2と同じ旧テーマのままです。

### メイン・ロゴなし版1を md から作る

リポジトリルートで実行します。初回は `pip install python-pptx` が必要です。

```
python scripts/build_kotonara_theme_pptx.py 原稿.md -o 出力.pptx --footer "フッターの文言"
python scripts/build_kotonara_theme_pptx.py 原稿.md -o 出力.pptx --no-logo   # ロゴなし版1のテーマ
```

- `--footer` を省くとフッターは空欄
- `{{...}}` が1つでも残っているとビルドは失敗する。社内確認用に残したまま作るときは `--draft` を付ける
- md の書式は `hackathon-2026/README.md` の「md を編集するときの約束ごと」と同じ。違いは次の2点
  - 1枚目は表紙、最後はクロージング、表があれば表スライド、それ以外はタイトルとコンテンツになる。
    セクション見出しにしたい枚には `<!-- layout: section -->` を書く（`cover` `content` `table` `closing` も指定できる）
  - 表スライドの枚数と位置に決まりはない
- 自動検証: テーマの色、ロゴがレイアウトにあること、スライドに色やロゴを直接置いていないこと、`{{...}}` の残存

カレンダー表: ヘッダーを `| 日 | 月 | 火 | 水 | 木 | 金 | 土 | この週のこと |` にした表は、月のカレンダーとして描きます。

- 日付セルは `19` または `19 ラベル`。ラベル付きはベージュ、ラベルに「当日」を含むと緑で塗る
- 空セルは月の外（灰色）
- 「この週のこと」が縦に同じ文言で続くと、セルを結合する
- 見本は `examples/kotonara-calendar-sample.md`

### 雛形とロゴを作り直す

```
python scripts/build_kotonara_theme_pptx.py --build-templates     # メインとロゴなし版1
python scripts/make_kotonara_logo.py assets/kotonara-logo-source.jpg assets
python scripts/make_kotonara_logo.py assets/kotonara-logo-source.jpg assets --color FFFFFF --name kotonara-logo-white
python -m pytest -q scripts/test_build_kotonara_theme_pptx.py
```

- 雛形の元は `kotonara-brand-template.md`
- ロゴの作り直しには `pip install potracer cairosvg numpy pillow` が必要
- `build_kotonara_theme_pptx.py` は `hackathon-2026/scripts/build_hackathon_slide_pptx.py` のテーマ生成を読み込んで使う
  （書き込みはしない）。そちらを変えるとメインとロゴなし版1にも反映されるので、両方のテストを流す
- `#4D9622` は `#1A832A` より明るいため、緑の上の小さい白文字は読みにくくなる。本番の投影環境で一度確認する

## ファイル

| ファイル | 内容 |
| --- | --- |
| `kotonara-brand-template.pptx` | メインの雛形（ロゴあり） |
| `kotonara-brand-template-nologo-1.pptx` | ロゴなし版1の雛形 |
| `kotonara-brand-template-nologo-2.pptx` | ロゴなし版2の雛形（旧メイン、表紙1枚） |
| `kotonara-brand-template.md` | メインとロゴなし版1の元になる md |
| `scripts/build_kotonara_theme_pptx.py` | メインとロゴなし版1のテーマで md から PPTX を作る |
| `scripts/make_kotonara_logo.py` | 元画像からロゴの SVG・PNG を作る |
| `scripts/test_build_kotonara_theme_pptx.py` | 上記の受け入れテスト |
| `assets/kotonara-logo.svg` / `.png` | 緑（`#4D9622`）のロゴ |
| `assets/kotonara-logo-white.svg` / `.png` | 白のロゴ（表紙の帯用） |
| `assets/kotonara-logo-source.jpg` | ロゴの元画像 |
| `examples/kotonara-calendar-sample.md` | カレンダー表の見本 |
| `kototech-aug-hackathon-kickoff-kotonara.pptx` | 8月キックオフ本編19枚（旧テーマ、スピーカーノート付き） |
| `scripts/build_kotonara_pptx.py` | 旧テーマのビルドスクリプト（キックオフ本編とロゴなし版2を生成） |
| `kototech-aug-hackathon-kickoff-slides.md` | キックオフ本編の入力 Markdown |

以下は旧テーマ（ロゴなし版2）の説明です。メインとロゴなし版1は、accent1 が `#4D9622` であることと
ロゴ・フッターの違いを除いて同じ構成です。

## テーマの中身

### カラースキーム（theme1.xml）

| スロット | 色 | 用途 |
| --- | --- | --- |
| dk1 / tx1 | `#222222` | 見出しの黒 |
| lt1 / bg1 | `#FFFFFF` | 背景の白 |
| dk2 / tx2 | `#333333` | 本文テキスト |
| lt2 / bg2 | `#EADECC` | ベージュ（カード背景） |
| accent1 | `#1A832A` | ブランドグリーン（メインアクセント） |
| accent2 | `#4169E1` | CTAブルー |
| accent3 | `#337AB7` | サブブルー |
| accent4 | `#CCCC33` | マスタード |
| accent5 | `#EADECC` | ベージュ |
| accent6 | `#666666` | 補助テキスト・罫線 |
| hlink / folHlink | `#4169E1` / `#337AB7` | リンク |

### フォントスキーム（theme1.xml）

- 見出し（majorFont）: 欧文 **Bahnschrift** ／ 和文 **游ゴシック**
- 本文（minorFont）: 欧文 **Segoe UI** ／ 和文 **游ゴシック**

スライド・レイアウト・マスターの色はすべて `<a:schemeClr>` 参照、フォントはすべて
`+mj-lt` / `+mj-ea` / `+mn-lt` / `+mn-ea` 参照で書かれています（本文スライドの
srgbClr 直書きはゼロ）。そのためテーマを差し替えるだけで全ページの見た目が入れ替わります。

### スライドマスターとレイアウト（5種）

| # | レイアウト名 | 種別 | プレースホルダー |
| --- | --- | --- | --- |
| 1 | 表紙（タイトル） | title | ctrTitle / subTitle(idx1) / body(idx2)。下部にブランドグリーンの帯＋『ひらめきを、カタチに。』 |
| 2 | タイトルとコンテンツ | obj | title / body(idx1) / footer(idx11) / slideNumber(idx12)。タイトル下に accent1 のアクセントバー |
| 3 | 表スライド | tbl | title / footer / slideNumber。表はネイティブ表オブジェクトを配置 |
| 4 | セクション見出し | secHead | title（accent1・大）/ body / footer / slideNumber。左に縦のグリーンバー |
| 5 | クロージング（反転） | obj | 背景 accent1、白文字の title / body / footer / slideNumber |

フッターは全レイアウト共通で「ことテク2026 2daysハッカソン」＋右にスライド番号。
本編19枚はすべて上記レイアウトのプレースホルダーに流し込んでいます
（マスター継承が効くので、マスターを直せば全ページに反映されます）。

---

## カスタマイズ方法 A: スクリプトの CONFIG を書き換えて再生成する

`scripts/build_kotonara_pptx.py` の冒頭にある `CONFIG` だけを編集し、再実行します。

```python
CONFIG = {
    "theme_name": "Kotonara",
    "colors": {                 # ← テーマ12色。RRGGBB を # なしで
        "dk1": "222222", "lt1": "FFFFFF", "dk2": "333333", "lt2": "EADECC",
        "accent1": "1A832A",    # ブランドグリーン
        "accent2": "4169E1", "accent3": "337AB7", "accent4": "CCCC33",
        "accent5": "EADECC", "accent6": "666666",
        "hlink": "4169E1", "folHlink": "337AB7",
    },
    "fonts": {                  # ← 和文/欧文 × 見出し/本文 の4種
        "major_latin": "Bahnschrift", "major_ea": "游ゴシック",
        "minor_latin": "Segoe UI",    "minor_ea": "游ゴシック",
    },
    "sizes": { "title": 30, "body": 18, "table_body": 14, ... },  # ← pt
    "footer_text": "ことテク2026 2daysハッカソン",
    "catch_copy": "ひらめきを、カタチに。",
}
```

実行:

```powershell
cd "C:\Users\Ryuto\OneDrive\ドキュメント\karabiner\kototech\scripts"
& "C:\Users\Ryuto\AppData\Local\Programs\Python\Python312\python.exe" build_kotonara_pptx.py
```

- `accent1` を変えると、アクセントバー・箇条書きの点・表のヘッダー行・表紙の帯・
  クロージングの背景がまとめて変わります。
- フォント4種を変えると、見出し／本文／表／ノートすべてに一括で反映されます。
- 入出力パスも冒頭の `MD_PATH` / `DECK_PATH` / `TEMPLATE_PATH` で変更できます。
- 実行すると最後に自動検証（スライド数19・ノート有無・表3枚・タイトル一致・
  theme1.xml の配色/フォント・zip 健全性・参照切れ）が走り、結果が表示されます。

## カスタマイズ方法 B: PowerPoint 上で配色／フォントを差し替える

theme1.xml に正式なカラースキーム／フォントスキームとして定義してあるので、
PowerPoint の GUI からそのまま差し替えられます。

**配色を変える**

1. リボンの **［デザイン］** タブを開く
2. 右側 **［バリエーション］** グループの **▼（その他）** をクリック
3. **［配色］** にポインタを合わせ、一覧から選ぶ
4. 自分の色に変えるなら **［色のカスタマイズ…］** →
   「アクセント1」に `#1A832A` 相当を入れ替えるなど12色を編集 →
   名前を付けて **［保存］**（保存した配色は他の資料でも再利用できます）

**フォントを変える**

1. **［デザイン］** → **［バリエーション］** の **▼** → **［フォント］**
2. **［フォントのカスタマイズ…］** を開く
3. 「英数字用の見出しのフォント／本文のフォント」と
   「日本語文字用の見出しのフォント／本文のフォント」をそれぞれ指定 → **［保存］**

※ このデッキは色もフォントもテーマ参照で書かれているため、上記の操作だけで
全19枚（表・フッター・スピーカーノートを含む）が一括で切り替わります。
個別スライドで色を直接指定してしまうと、その箇所だけテーマ差し替えが効かなくなるので注意してください。

**新しい資料を作るとき**

`kotonara-brand-template.pptx`（メイン）を開いて［名前を付けて保存］するか、
`.potx` として保存してテンプレート化してください。ロゴを入れない資料はロゴなし版1を使います。
［ホーム］→［新しいスライド］から上記5レイアウトを選べます。

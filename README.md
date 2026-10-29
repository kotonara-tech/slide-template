# ことなら（kotonara.tech）ブランド PPTX — カスタマイズ手引き

公式サイト https://kotonara.tech/ の配色・トーンに合わせた PowerPoint テーマ一式です。

## ことテク2026 ハッカソン チーム発表テンプレート

成果発表会（発表8分＋デモ2分＋質疑3分）用の9枚構成テンプレートは `hackathon-2026/` にあります。
Claude Code で `/create-hackathon-slide <チームの成果物リポジトリのパスまたはURL>` を実行すれば、
成果物を読んで Markdown を埋め、このテーマの PPTX まで生成します。詳細は `hackathon-2026/README.md` を参照。

## ことテク 汎用スライドテーマ

イベント資料・営業資料に使う汎用テーマ（メインの緑 `#4D9622`、「ことテク」ロゴ入り、フッター空欄）は
`kototech-theme/` にあります。雛形は `kototech-theme/kototech-brand-template.pptx`。
md から作るときは `kototech-theme/scripts/build_kototech_pptx.py` を使います。詳細は `kototech-theme/README.md` を参照。

## ファイル

| ファイル | 内容 |
| --- | --- |
| `kototech-aug-hackathon-kickoff-kotonara.pptx` | 本編19枚（スピーカーノート付き） |
| `kotonara-brand-template.pptx` | 同じテーマ／スライドマスターだけを持つ雛形（表紙1枚） |
| `scripts/build_kotonara_pptx.py` | 上記2ファイルを生成するビルドスクリプト |
| `kototech-aug-hackathon-kickoff-slides.md` | 入力となる確定済み Markdown |

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

`kotonara-brand-template.pptx` を開いて［名前を付けて保存］するか、
`.potx` として保存してテンプレート化してください。
［ホーム］→［新しいスライド］から上記5レイアウトを選べます。

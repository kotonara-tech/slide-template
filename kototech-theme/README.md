# ことテク 汎用スライドテーマ

ことテク（ことならテックプロジェクト）のイベント資料・営業資料に使う PowerPoint テーマです。
特定のイベントに依存しない汎用テンプレートで、フッターの文言は空欄にしてあります。

リポジトリ直下の「ことなら（kotonara.tech）ブランド PPTX」と同じテーマ・5レイアウトを元に、
次の点を変えています。

| 項目 | 内容 |
| --- | --- |
| メインの緑（accent1） | `#4D9622`（元テンプレートは `#1A832A`） |
| ロゴ | 「ことテク」ロゴをスライドレイアウトに配置 |
| 表紙の帯 | 左に白いロゴ。右にキャッチコピー『ひらめきを、カタチに。』（13pt、下端をロゴにそろえる） |
| フッター | 空欄（［挿入］＞［ヘッダーとフッター］で入力） |
| カレンダー表 | md からカレンダー形式の表を作れる |

## ファイル

| ファイル | 内容 |
| --- | --- |
| `kototech-brand-template.pptx` | 雛形。5レイアウトの見本が1枚ずつ入っている |
| `template.md` | 雛形の元になる md |
| `examples/calendar-sample.md` | カレンダー表の見本 |
| `scripts/build_kototech_pptx.py` | md から PPTX を作るビルドスクリプト（自動検証つき） |
| `scripts/make_logo.py` | 元画像からロゴの SVG・PNG を作るスクリプト |
| `scripts/test_build_kototech_pptx.py` | 受け入れテスト |
| `assets/kototech-logo.svg` / `.png` | 緑（`#4D9622`）のロゴ。PNG は 4000×1412・背景透過 |
| `assets/kototech-logo-white.svg` / `.png` | 白のロゴ（表紙の帯用） |
| `assets/kototech-logo-source.jpg` | ロゴの元画像（黒背景・白抜き） |

## レイアウトとロゴ

| # | レイアウト | ロゴ |
| --- | --- | --- |
| 1 | 表紙（タイトル） | 下部の帯の左に白いロゴ |
| 2 | タイトルとコンテンツ | 右上に緑のロゴ |
| 3 | 表スライド | 右上に緑のロゴ |
| 4 | セクション見出し | 右上に緑のロゴ |
| 5 | クロージング（反転） | なし（背景が緑のため） |

ロゴ・配色・フォントはスライドマスター、レイアウト、テーマに定義してあります。
スライドに色やロゴを直接置いていないので、［ホーム］＞［新しいスライド］でレイアウトを選べば
そのまま付きます。ロゴは PNG と SVG の両方を入れてあり、PowerPoint 2016 以降は SVG で描画します。

## 使い方A: 雛形から作る

`kototech-brand-template.pptx` を開いて［名前を付けて保存］します。見本のスライドは消して使います。

## 使い方B: md から作る

リポジトリルートで実行します。初回は `pip install python-pptx` が必要です。

```
python kototech-theme/scripts/build_kototech_pptx.py 原稿.md -o 出力.pptx --footer "フッターの文言"
```

- `--footer` を省くとフッターは空欄
- `{{...}}` が1つでも残っているとビルドは失敗する。社内確認用に残したまま作るときは `--draft` を付ける
- md の書式は `hackathon-2026/README.md` の「md を編集するときの約束ごと」と同じ。違いは次の2点
  - 1枚目は表紙、最後はクロージング、表があれば表スライド、それ以外はタイトルとコンテンツになる。
    セクション見出しにしたい枚には `<!-- layout: section -->` を書く（`cover` `content` `table` `closing` も指定できる）
  - 表スライドの枚数と位置に決まりはない

### カレンダー表

ヘッダーを `| 日 | 月 | 火 | 水 | 木 | 金 | 土 | この週のこと |` にした表は、月のカレンダーとして描きます。

- 日付セルは `19` または `19 ラベル`。ラベル付きはベージュ、ラベルに「当日」を含むと緑で塗る
- 空セルは月の外（灰色）
- 「この週のこと」が縦に同じ文言で続くと、セルを結合する

見本は `examples/calendar-sample.md` です。

## 雛形とロゴを作り直す

```
python kototech-theme/scripts/build_kototech_pptx.py kototech-theme/template.md -o kototech-theme/kototech-brand-template.pptx
python kototech-theme/scripts/make_logo.py kototech-theme/assets/kototech-logo-source.jpg kototech-theme/assets
python kototech-theme/scripts/make_logo.py kototech-theme/assets/kototech-logo-source.jpg kototech-theme/assets --color FFFFFF --name kototech-logo-white
```

ロゴの作り直しには `pip install potracer cairosvg numpy pillow` が必要です。
ビルドスクリプトは `hackathon-2026/scripts/build_hackathon_slide_pptx.py` のテーマ生成を読み込んで使うので、
そちらを変えるとこのテーマにも反映されます。変えたら下のテストを流してください。

```
python -m pytest -q kototech-theme/scripts/test_build_kototech_pptx.py
```

## 注意

- 奈良市の市章とカラビナテクノロジーのロゴは、使用規定により加工しないでください（切り抜き、色変更、縦横比の変更、透過加工は不可）。
- フォントは Windows 標準（Bahnschrift / Segoe UI / 游ゴシック）です。macOS・Google スライドでは代替フォントになり、改行位置がずれます。
- `#4D9622` は `#1A832A` より明るいため、緑の上の小さい白文字は読みにくくなります。本番の投影環境で一度確認してください。

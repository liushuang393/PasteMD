# PasteMD
<p align="center">
  <img src="../../assets/icons/logo.png" alt="PasteMD" width="160" height="160">
</p>

<p align="center">
  <a href="https://github.com/RICHQAQ/PasteMD/releases">
    <img src="https://img.shields.io/github/v/release/RICHQAQ/PasteMD?sort=semver&label=Release&style=flat-square&logo=github" alt="Release">
  </a>
  <a href="https://github.com/RICHQAQ/PasteMD/releases">
    <img src="https://img.shields.io/github/downloads/RICHQAQ/PasteMD/total?label=Downloads&style=flat-square&logo=github" alt="Downloads">
  </a>
  <a href="../../LICENSE">
    <img src="https://img.shields.io/github/license/RICHQAQ/PasteMD?style=flat-square" alt="License">
  </a>
  <img src="https://img.shields.io/badge/Python-3.12%2B-3776AB?style=flat-square&logo=python&logoColor=white" alt="Python 3.12+">
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20Word%20%7C%20WPS-5e8d36?style=flat-square&logo=windows&logoColor=white" alt="Platform">
</p>

<p align="center">
  <a href="README.en.md">English</a>
  | 
  <a href="../../README.md">简体中文</a>
  |
  <a href="README.ja.md">日本語</a>
</p>

> 論文やレポートを書く際、AIツール(ChatGPTやDeepSeekなど)からコピーした数式がWordで文字化けしてしまう?MarkdownテーブルをExcelに正しく貼り付けられない?**PasteMDはこれらの問題を解決するために開発されました。**
> 
> <img src="../../docs/gif/atri/igood.gif"
     alt="I am good"
     width="100">

PasteMDは、クリップボードを監視し、MarkdownまたはHTMLリッチテキストをPandoc経由でDOCXに変換し、WordやWPSのカーソル位置に直接貼り付ける軽量なトレイアプリです。Markdownテーブルを認識してExcelに書式を保持したまま直接貼り付けることができ、WebページからコピーされたHTMLリッチテキスト(数式を除く)も認識します。

---

## 主な機能

### デモ動画

#### Markdown → Word/WPS

<p align="center">
  <img src="../../docs/gif/demo.gif" alt="MarkdownからWordへのデモ" width="600">
</p>

#### WebページのAI回答をコピー → Word/WPS
<p align="center">
  <img src="../../docs/gif/demo-html.gif" alt="HTMLリッチテキストのデモ" width="600">
</p>

#### Markdownテーブル → Excel
<p align="center">
  <img src="../../docs/gif/demo-excel.gif" alt="MarkdownテーブルからExcelへのデモ" width="600">
</p>

#### 書式プリセットの適用
<p align="center">
  <img src="../../docs/gif/demo-chage_format.gif" alt="書式設定のデモ" width="600">
</p>

### ワークフロー向上機能

- グローバルホットキー(デフォルト`Ctrl+Shift+B`)で最新のMarkdown/HTMLクリップボードスナップショットをDOCXとして貼り付け。
- Markdownテーブルを自動認識し、スプレッドシートに変換してExcelに貼り付け、太字/斜体/コード書式を保持。
- WebページからコピーされたHTMLリッチテキストを認識し、Word/WPSに変換/貼り付け。
- フォアグラウンドのターゲットアプリ(Word、WPS、またはExcel)を検出し、必要に応じて適切なプログラムを開く。
- 機能の切り替え、ログ表示、設定の再読み込み、更新確認のためのトレイメニュー。
- オプションのトースト通知と、すべての変換のバックグラウンドログ記録。

---

## AIウェブサイト互換性

以下の表は、人気のあるAIチャットサイトがMarkdownまたは直接HTMLコンテンツをコピーする際にPasteMDとどの程度うまく機能するかをまとめたものです。

| AIサービス | Markdownをコピー(数式なし) | Markdownをコピー(数式あり) | ページコンテンツをコピー(数式なし) | ページコンテンツをコピー(数式あり) |
|------------|----------------------------|-------------------------------|---------------------------------|-----------------------------------|
| Kimi | ✅ 完璧 | ✅ 完璧 | ✅ 完璧 | ⚠️ 数式が欠落 |
| DeepSeek | ✅ 完璧 | ✅ 完璧 | ✅ 完璧 | ✅ 完璧 |
| Tongyi Qianwen | ✅ 完璧 | ✅ 完璧 | ✅ 完璧 | ⚠️ 数式が欠落 |
| Doubao* | ✅ 完璧 | ✅ 完璧 | ✅ 完璧 | ✅ 完璧 |
| ChatGLM/Zhipu | ✅ 完璧 | ✅ 完璧 | ✅ 完璧 | ✅ 完璧 |
| ChatGPT | ✅ 完璧 | ⚠️ コードとして表示 | ✅ 完璧 | ✅ 完璧 |
| Gemini | ✅ 完璧 | ✅ 完璧 | ✅ 完璧 | ✅ 完璧 |
| Grok | ✅ 完璧 | ✅ 完璧 | ✅ 完璧 | ✅ 完璧 |
| Claude | ✅ 完璧 | ✅ 完璧 | ✅ 完璧 | ✅ 完璧 |

_*Doubaoは、数式を含むHTMLコンテンツをコピーする前に、ブラウザでクリップボード読み取り権限を付与する必要があります(URLバー近くのロックアイコンから設定)。_

凡例:
- ✅ **完璧** — 書式、スタイル、数式がそのまま保持されます。
- ⚠️ **コードとして表示** — 数式が生のLaTeXとして表示され、Word/WPS内で再構築する必要があります。
- ⚠️ **数式が欠落** — 数式が削除されます。数式エディタで手動で再構築してください。

テスト説明:
1. **Markdownをコピー** — ほとんどのAI応答の下にある「コピー」ボタンを使用(通常はMarkdown、時々HTML)。
2. **ページコンテンツをコピー** — AI応答を手動で選択してコピー(HTMLリッチテキスト)。

---

## 使い方

1. [リリースページ](https://github.com/RICHQAQ/PasteMD/releases/)から実行ファイルをダウンロード:
   - ~~**PasteMD_vx.x.x.exe** — ポータブルビルド、Pandocがインストールされ`PATH`からアクセス可能である必要があります。~~ (提供終了。必要な場合はソースからビルドしてください)
   - **PasteMD_pandoc-Setup.exe** — Pandocが同梱されたインストーラーで、すぐに使用できます。
2. Word、WPS、またはExcelを開き、貼り付けたい位置にカーソルを置きます。
3. MarkdownまたはHTMLリッチテキストをコピーし、グローバルホットキー(デフォルトは`Ctrl+Shift+B`)を押します。
4. PasteMDは以下を実行します:
   - MarkdownテーブルをExcelに送信(Excelが既に開いている場合)。
   - 通常のMarkdown/HTMLをDOCXに変換し、Word/WPSに挿入。
5. トレイの通知(およびオプションのトースト)で成功または失敗を確認します。

---

## 設定

初回起動時に`config.json`ファイルが作成されます。直接編集し、トレイメニューの**「設定/ホットキーを再読み込み」**を使用して変更を即座に適用します。

```json
{
  "hotkey": "<ctrl>+<shift>+b",
  "pandoc_path": "pandoc",
  "reference_docx": null,
  "save_dir": "%USERPROFILE%\\Documents\\pastemd",
  "keep_file": false,
  "notify": true,
  "enable_excel": true,
  "excel_keep_format": true,
  "no_app_action": "open",
  "md_disable_first_para_indent": true,
  "html_disable_first_para_indent": true,
  "html_formatting": {
    "strikethrough_to_del": true
  },
  "move_cursor_to_end": true,
  "Keep_original_formula": false,
  "language": "zh",
  "pandoc_filters": []
}
```

主要フィールド:

- `hotkey` — `<ctrl>+<alt>+v`のようなグローバルショートカット構文。
- `pandoc_path` — Pandocの実行可能ファイル名または絶対パス。
- `reference_docx` — Pandocが使用するオプションのスタイルテンプレート。
- `save_dir` — 生成されたDOCXファイルを保持する際に使用するディレクトリ。
- `keep_file` — 変換されたDOCXファイルを削除せずにディスクに保存。
- `notify` — 変換完了時にシステム通知を表示。
- `enable_excel` — Markdownテーブルを検出し、自動的にExcelに貼り付け。
- `excel_keep_format` — Excel内で太字/斜体/コードスタイルを保持しようとする。
- `no_app_action` — ターゲットアプリが検出されない場合のアクション。値: `open`(自動で開く)、`save`(保存のみ)、`clipboard`(ファイルをクリップボードにコピー)、`none`(何もしない)。デフォルト: `open`。
- `md_disable_first_para_indent` / `html_disable_first_para_indent` — 最初の段落スタイルを本文テキストに正規化。
- `html_formatting` — 変換前にHTMLリッチテキストをフォーマットするためのオプション。
  - `strikethrough_to_del` — 取り消し線~~を`<del>`タグに変換して適切にレンダリング。
- `move_cursor_to_end` — 挿入結果の末尾にカーソルを移動。
- `Keep_original_formula` — 元の数式を保持(LaTeXコード形式)。
- `language` — UI言語、`en`または`zh`。
- **`pandoc_filters`** — **✨ 新機能** - カスタムPandoc Filterリスト。`.lua`スクリプトまたは実行可能ファイルのパスを追加。フィルターはリスト順に実行されます。カスタム書式処理、特殊構文変換などでPandoc変換を拡張します。デフォルト: 空のリスト。例: `["%APPDATA%\\npm\\mermaid-filter.cmd"]`でMermaid図のサポート。

---

## 🔧 高度な機能: カスタムPandoc Filters

### Pandoc Filtersとは?

Pandoc Filtersは、変換中にドキュメントコンテンツを処理するプラグインプログラムです。PasteMDは、機能を拡張するために順次実行される複数のフィルターの設定をサポートしています。

### 使用例: Mermaid図のサポート

MarkdownでMermaid図を使用し、Wordに適切に変換するには、[mermaid-filter](https://github.com/raghur/mermaid-filter)を使用できます。

**1. mermaid-filterをインストール**

```bash
npm install --global mermaid-filter
```

*前提条件: [Node.js](https://nodejs.org/)がインストールされている必要があります*

<details>
<summary>⚠️ <b>トラブルシューティング: Chromeダウンロード失敗</b></summary>

mermaid-filterのインストールにはChromiumブラウザのダウンロードが必要です。自動ダウンロードが失敗した場合は、手動でダウンロードできます:

**ステップ1: 必要なChromiumバージョンを確認**

ファイルを確認: `%APPDATA%\npm\node_modules\mermaid-filter\node_modules\puppeteer-core\lib\cjs\puppeteer\revisions.d.ts`

次のような内容を見つけます:
```typescript
chromium: "1108766";
```
またはエラーメッセージ内で、例えば:

```bash
npm error Error: Download failed: server returned code 502. URL: https://npmmirror.com/mirrors/chromium-browser-snapshots/Win_x64/1108766/chrome-win.zip
```
`Win_x64/1108766`のようなバージョンを見つけます。

このバージョン番号をメモします(例: `1108766`)。

**ステップ2: Chromiumをダウンロード**

ステップ1で取得したバージョン番号に基づいて、対応するChromiumをダウンロード:

```
https://storage.googleapis.com/chromium-browser-snapshots/Win_x64/1108766/chrome-win.zip
```

(URL内の`1108766`をあなたのバージョン番号に置き換えてください)

**ステップ3: 指定されたディレクトリに解凍**

ダウンロードした`chrome-win.zip`を以下に解凍:

```
%USERPROFILE%\.cache\puppeteer\chrome\win64-1108766\chrome-win
```

(パス内の`1108766`をあなたのバージョン番号に置き換えてください)

解凍後、`chrome.exe`は以下に配置されます:
`%USERPROFILE%\.cache\puppeteer\chrome\win64-1108766\chrome-win\chrome.exe`

</details>

**2. PasteMDで設定**

オプション1: 設定UIから
- PasteMD設定を開く → 変換タブ → Pandoc Filters
- 「追加...」ボタンをクリック
- フィルターファイルを選択: `%APPDATA%\npm\mermaid-filter.cmd`
- 設定を保存

オプション2: 設定ファイルを編集
```json
{
  "pandoc_filters": [
    "%APPDATA%\\npm\\mermaid-filter.cmd"
  ]
}
```

**3. テスト**

以下のMarkdownをコピーしてPasteMDで変換:

~~~markdown
```mermaid
graph LR
    A[開始] --> B[処理]
    B --> C[終了]
```
~~~

Mermaid図は画像としてレンダリングされ、Word文書に挿入されます。

### その他のFilterリソース

- [公式Pandoc Filtersリスト](https://github.com/jgm/pandoc/wiki/Pandoc-Filters)
- [Lua Filtersドキュメント](https://pandoc.org/lua-filters.html)

---

## トレイメニュー

- 現在のグローバルホットキーを表示(読み取り専用)。
- ホットキーの有効/無効。
- 通知の切り替え、ターゲットアプリが検出されない場合のアクション設定、貼り付け後のカーソル移動の切り替え。
- Excel固有の機能と書式保持の有効/無効。
- 生成されたDOCXファイルの保持の切り替え。
- HTML書式設定: 取り消し線~~を`<del>`タグに変換して適切にレンダリング。
- Keep_original_formula: 元の数式をLaTeXコード形式で保持するかどうか。
- 保存ディレクトリを開く、ログを表示、設定を編集、ホットキーを再読み込み。
- 更新を確認し、インストールされているバージョンを表示。
- PasteMDを終了。

---

## ソースからビルド

推奨環境: Python 3.12 (64ビット)。

```bash
pip install -r requirements.txt
python main.py
```

パッケージビルド(PyInstaller):

```bash
pyinstaller --clean -F -w -n PasteMD
  --icon assets\icons\logo.ico
  --add-data "assets\icons;assets\icons"
  --add-data "pastemd\i18n\locales\*.json;pastemd\i18n\locales"
  --add-data "pastemd\lua;pastemd\lua"
  --hidden-import plyer.platforms.win.notification
  main.py
```

コンパイルされた実行ファイルは`dist/PasteMD.exe`に配置されます。

---

## ⭐ Star

すべてのStarが助けになります — PasteMDをより多くのユーザーと共有していただきありがとうございます。

<img src="../../docs/gif/atri/likeyou.gif"
     alt="like you"
     width="150">

[![Star History Chart](https://api.star-history.com/svg?repos=RICHQAQ/PasteMD&type=date&legend=top-left)](https://www.star-history.com/#RICHQAQ/PasteMD&type=date&legend=top-left)

---

## ☕ サポート&寄付


PasteMDが時間を節約できた場合は、作者にコーヒーをおごることを検討してください — あなたのサポートは、修正、機能強化、新しい統合の優先順位付けに役立ちます。

また、**PasteMDユーザーグループ**への参加も歓迎します:

<div align="center">
  <img src="../../docs/img/qrcode.jpg" alt="PasteMD QQグループQRコード" width="200" />
  <br>
  <sub>スキャンしてPasteMD QQグループに参加</sub>
</div>

<img src="../../docs/gif/atri/flower.gif"
     alt="give you a flower"
     width="150">

| Alipay | WeChat |
| --- | --- |
| ![Alipay](../../docs/pay/Alipay.jpg) | ![WeChat](../../docs/pay/Weixinpay.png) |

---

## ライセンス

このプロジェクトは[MITライセンス](../../LICENSE)の下でリリースされています。


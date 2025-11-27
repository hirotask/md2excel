# md2excel

Markdown形式のテキストをXLSX形式に変換するツールです。

## 機能ハイライト

- Markdown形式のテーブルをExcelで読み込める形式に変換
- 見出し1（#）で新しいシートを作成
- 見出し2（##）を見出しスタイル（太字・大きめフォント）で表示
- テーブル以外のテキスト（段落）も変換して表示
- 箇条書きと番号付きリストをセル内改行として処理
- テーブルのヘッダー行を太字で表示
- テーブルセルに罫線を自動適用
- Mermaid図をPNG画像としてExcelに埋め込み（オプション機能）

## Getting Started

### 方法1: Dockerを使用（推奨）

すべての依存関係がプリインストール済みで、最も簡単に始められます。

#### 1. Dockerイメージをビルド

```bash
docker build -t md2excel:latest .
```

#### 2. Markdownファイルを変換

```bash
# 基本的な使用方法
docker run -v $(pwd):/data md2excel:latest input.md -o output.xlsx

# サンプルファイルで試す
docker run -v $(pwd):/data md2excel:latest test_sample.md

# Mermaid図を含むファイル
docker run -v $(pwd):/data md2excel:latest test_mermaid.md -o output.xlsx

# ヘルプを表示
docker run md2excel:latest --help
```

**Dockerの利点:**
- ✅ 依存関係（Playwright、Chromium、日本語フォント）が全てプリインストール済み
- ✅ システムに依存関係をインストール不要
- ✅ どの環境でも同じように動作
- ✅ クリーンな実行環境

---

### 方法2: ローカル環境にインストール

#### 1. 基本インストール

```bash
uv pip install -e .
```

#### 2. Markdownファイルを変換

```bash
# 基本的な使用方法
uv run md2excel input.md

# Excel形式で出力
uv run md2excel input.md -o output.xlsx

# サンプルファイルで試す
uv run md2excel test_sample.md
```

#### 3. Mermaid図サポートの追加設定（オプション）

Mermaid図をExcelに埋め込む機能を使用する場合のみ必要です：

```bash
# システム依存関係のインストール（Linux/WSL）
sudo apt-get install -y libnspr4 libnss3 libasound2t64

# 日本語フォントのインストール（日本語を含むMermaid図の場合）
sudo apt-get install -y fonts-noto-cjk fonts-noto-cjk-extra

# Playwrightブラウザのインストール
uv run python -m playwright install chromium

# または、全ての依存関係を一括インストール
sudo uv run python -m playwright install-deps
```

**注意:**
- Mermaid機能を使用しない場合は、基本インストールのみで動作します
- 依存関係が不足している場合、Mermaid図は警告メッセージに置き換えられます

## 変換可能なMarkdown記述

### シート構造

- `# Sheet Name` - 新しいシートを作成
- `## Section Name` - シート内のセクション見出しとして表示（太字・フォントサイズ13）
- `### Subsection` - 無視される
- 通常のテキスト（段落） - テーブルの前後に配置可能、折り返しなしで表示

### テーブル形式

```markdown
| Header1 | Header2 |
|---------|---------|
| Cell1   | Cell2   |
| Cell3   | Cell4   |
```

### リスト処理

箇条書き:
```markdown
- Item1
- Item2
```
→ セル内で「・ Item1」「・ Item2」と改行表示

番号付きリスト:
```markdown
1. Step1
2. Step2
```
→ セル内で「1) Step1」「2) Step2」と改行表示

セル内で改行するには `<br>` タグを使用します。

### Mermaid図

Mermaid形式のフローチャートやダイアグラムを埋め込むことができます：

```markdown
# シート名

## フローチャート

\`\`\`mermaid
flowchart TD
  A[開始] --> B{判定}
  B -->|Yes| C[処理A]
  B -->|No| D[処理B]
  C --> E[終了]
  D --> E
\`\`\`
```

Mermaid図は自動的にPNG画像に変換され、Excelシートに埋め込まれます。

**対応する図の種類：**
- フローチャート（flowchart）
- シーケンス図（sequenceDiagram）
- ガントチャート（gantt）
- クラス図（classDiagram）
- 状態遷移図（stateDiagram）
- その他、Mermaidがサポートする全ての図

### テキストとテーブルの組み合わせ

テーブルの前後にテキストを配置できます：

```markdown
# シート名

## セクション見出し

ここに説明文を書けます。

| Header1 | Header2 |
|---------|---------|
| Data1   | Data2   |

テーブルの後にも説明を追加できます。
```

**スタイル適用：**
- 見出し（##）：太字、フォントサイズ13、折り返しなし、罫線なし
- 通常のテキスト：通常フォント、折り返しなし、罫線なし
- Mermaid図：PNG画像として埋め込み（幅600px）
- テーブルヘッダー：太字、折り返しあり、罫線あり
- テーブルデータ：通常フォント、折り返しあり、罫線あり

## 出力形式

拡張子`.xlsx`で保存すると、Office Open XML形式（Excel 2007以降）で出力されます。

## トラブルシューティング

### Mermaid図が表示されない

Mermaid図が「[Mermaid diagram - renderer not available]」と表示される場合：

1. システム依存関係がインストールされていることを確認：
   ```bash
   sudo apt-get install -y libnspr4 libnss3 libasound2t64
   ```

2. Chromiumブラウザがインストールされていることを確認：
   ```bash
   uv run python -m playwright install chromium
   ```

3. WSL環境の場合、追加のライブラリが必要な場合があります：
   ```bash
   sudo uv run python -m playwright install-deps
   ```

### Mermaid図の日本語が文字化けする

日本語テキストが正しく表示されない場合は、日本語フォントをインストールしてください：

```bash
# Noto CJKフォントのインストール（推奨）
sudo apt-get install -y fonts-noto-cjk fonts-noto-cjk-extra

# または、他の日本語フォント
sudo apt-get install -y fonts-ipafont fonts-ipaexfont
```

インストール後、再度変換を実行してください。ツールは以下のフォントを優先的に使用します：
- Noto Sans JP（推奨）
- Hiragino Sans
- Yu Gothic
- Meiryo

### テーブルが正しく検出されない

- テーブルのセパレーター行（`|---|---|`）とデータ行の間に空白行がないことを確認してください
- セパレーター行には必ず3文字以上のハイフン（`---`）を使用してください

## ライセンス

MIT License

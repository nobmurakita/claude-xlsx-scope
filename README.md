# xlsx-scope

Excel ファイル（.xlsx / .xlsm）の内容を CLI から読み取り、JSONL形式で出力するGoツール。
Claude Code 等の AI エージェントが Excel 資料（通常の表、Excel方眼紙の仕様書など）を構造的に読み取る用途を主眼とする。

## インストール

### Claude Code で使う

ワンライナーでインストール:

```bash
curl -fsSL https://raw.githubusercontent.com/nobmurakita/claude-xlsx-scope/main/install.sh | bash
```

macOS / Linux / Windows (Git Bash)で動作する。インストール後、Claude Code が Excel ファイルの読み取りが必要な場面で自動的に xlsx-scope を使用する。

### チャット, Cowork で使う

1. [Releases](https://github.com/nobmurakita/claude-xlsx-scope/releases) から最新の zip をダウンロード
2. Claude の設定で「機能 > コード実行とファイル作成」が有効になっていることを確認
3. [カスタマイズ > スキル](https://claude.ai/customize/skills) を開く
4. 「+」→「スキルを作成」→「スキルをアップロード」で zip をアップロード

スキルリストに追加されたら、Claude が Excel ファイルの読み取りが必要な場面で自動的に xlsx-scope を使用する。

出力形式やフィールドの詳細は [DESIGN.md](DESIGN.md) を参照。

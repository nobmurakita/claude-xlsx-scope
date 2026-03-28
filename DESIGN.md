# cc-read-excel 設計ドキュメント

Excel ファイル（.xlsx / .xlsm）の内容をCLIから出力するGoツール。
Claude CodeがExcel資料（通常の表、Excel方眼紙の仕様書など）を読み取る用途を主眼とする。

## コマンド構成と利用フロー

Claude Code からの典型的な利用フローは以下の通り:

1. **`info`** — ファイルの全体像を把握する（シート一覧、定義名）。どのシートを読むべきか判断する
2. **`scan`** — 対象シートの used_range を取得する（任意）。`cells` の `--range` を決定するために使う。図形の有無（`has_drawings`）も確認できる
3. **`cells`** — セルデータを取得する。先頭に `_meta` 行でレイアウト情報（列幅、デフォルト行高等）を出力し、続いてセルデータを出力する
4. **`shapes`** — 図形・フローチャートを取得する。図形間の接続関係を含めて構造を把握する（`scan` で `has_drawings: true` のシートに対して使用）
5. **`search`** — 特定の値やセル型を検索する。シート全体から条件に合うセルだけを抽出したい場合に使用する。`cells` で全体を取得するより効率的

`scan` を経由せず `info` → `cells` で直接データを取得することも可能。`cells` の `_meta` 行でレイアウト情報を取得でき、`scan` は大きいシートで used_range を事前に把握したい場合に有用

### `cc-read-excel info <file>`

**役割:** ファイルレベルの概要を把握する。シート一覧から対象シートを特定し、以降の `scan` / `cells` / `search` に渡す `--sheet` を決定する。

ファイルレベルの概要をJSON形式で出力する。

- シート一覧（名前、インデックス、種類、表示状態）
- 名前付き範囲（定義名）一覧（名前、スコープ、参照先）。`refer` は workbook.xml の値をそのまま出力する（削除済みシートへの参照や数式を含む定義名もそのまま）

**出力例:**

```json
{
  "file": "基本設計書.xlsx",
  "defined_names": [
    {"name": "マスタ", "scope": "Workbook", "refer": "Sheet1!$A$1:$D$100"},
    {"name": "入力範囲", "scope": "Sheet2", "refer": "Sheet2!$B$3:$F$20"}
  ],
  "sheets": [
    {"index": 0, "name": "表紙", "type": "worksheet"},
    {"index": 1, "name": "機能一覧", "type": "worksheet"},
    {"index": 2, "name": "非表示データ", "type": "worksheet", "hidden": true},
    {"index": 3, "name": "グラフ", "type": "chartsheet"}
  ]
}
```

**`sheets` 配列の各要素:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `index` | number | シートの0始まりインデックス |
| `name` | string | シート名 |
| `type` | string | シート種類（`worksheet`, `chartsheet` 等。workbook.xml.rels のリレーション種別から判定） |
| `hidden` | bool | 非表示の場合のみ `true` を出力（`hidden` / `veryHidden` を区別しない）。表示状態のシートでは省略 |

### `cc-read-excel cells <file>`

**役割:** セルの値と書式を取得する。先頭に `_meta` 行でレイアウト情報を出力し、続いてセルデータを出力する。`--range` で範囲を絞るか、`--range` なしで先頭からストリーミング取得する。内部的にワークシートXMLを自前でSAXパースし、`--limit` 到達時に即座に走査を打ち切る。

指定シートのセル情報を出力する。

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--sheet <name\|index>` | 対象シート（名前 or 0始まりインデックス） | 最初のシート |
| `--range <range>` | セル範囲。Excel記法で指定（例: `A1:H20`, `A:F`, `1:20`） | シートの使用範囲全体 |
| `--start <cell>` | 開始位置（例: `A51`）。行優先の走査順で指定セル以降のセルを出力する（指定セルの行の途中から開始し、次の行はA列から走査する）。`--range` との併用不可 | 先頭セル |
| `--include-empty` | 空セルも出力する | OFF（空セルはスキップ） |
| `--style` | 書式情報を出力する | OFF（書式は省略） |
| `--formula` | 数式文字列を出力する | OFF（数式は省略し値のみ出力） |
| `--limit <n>` | 出力セル数の上限 | 1000 |

- `--limit 0` で上限なし
- `--include-empty` 指定時は空セルも出力対象に含まれるため、対象セル数が大幅に増える。必要に応じて `--range` で範囲を絞るか `--limit` を調整すること

### `cc-read-excel scan <file>`

**役割:** 対象シートの使用範囲（used_range）を取得する。`cells` の `--range` を決定するために使う。

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--sheet <name\|index>` | 対象シート（名前 or 0始まりインデックス） | 最初のシート |

**出力例:**

```json
{"sheet": "機能一覧", "used_range": "A1:H200"}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `sheet` | string | シート名 |
| `used_range` | string | シートの使用範囲。空シートの場合は省略 |

- dimension（XMLの `<dimension>` 要素）があれば即座に返す（数十ms）
- dimension がない場合（Google Sheets 由来等）は全セル走査で算出する
- レイアウト情報（列幅、行高、デフォルトフォント等）は `cells` の `_meta` 行で取得可能

### `cc-read-excel search <file>`

**役割:** 特定の値やセル型をシート内から検索する。シート全体を `cells` して手元でフィルタするより効率的。例えば「特定のキーワードを含むセルの位置を特定 → その周辺を `cells --range` で取得」という使い方ができる。

セル値を検索し、一致するセルの情報を `cells` と同じ出力形式で出力する。

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--query <text>` | 検索文字列（部分一致） | なし |
| `--numeric <expr>` | 数値比較（例: `">100"`, `"100:200"`, `"=42"`） | なし |
| `--type <type>` | セルの型でフィルタ（`string`, `number`, `date`, `bool`, `formula`） | すべて |
| `--sheet <name\|index>` | 対象シート（名前 or 0始まりインデックス） | 最初のシート |
| `--range <range>` | セル範囲。Excel記法で指定（例: `A1:H20`） | シートの使用範囲全体 |
| `--start <cell>` | 開始位置（例: `A51`）。`cells` の `--start` と同様の走査順で指定セル以降を検索する。`--range` との併用不可 | 先頭セル |
| `--style` | 書式情報を出力する | OFF（書式は省略） |
| `--formula` | 数式文字列を出力する | OFF（数式は省略し値のみ出力） |
| `--limit <n>` | 出力セル数の上限 | 1000 |

- `--limit 0` で上限なし
- `--query`, `--numeric`, `--type` のうち少なくとも1つの指定が必須

**検索の挙動:**

- 文字列検索（`--query`）: `display`（表示文字列）と `value` の両方に対して大文字・小文字を区別しない部分一致検索を行う（いずれかにマッチすればヒット）。全角・半角は区別する。正規表現には対応しない。数値・日付・真偽値セルの `value` は出力時と同じ文字列表現に変換してから比較する（例: 数値 `1000` は `"1000"` として、日付は ISO 8601 形式として比較）
- 例: `--query "令和"` → 和暦表示のセルにヒット、`--query "2025-03"` → ISO 8601形式の `value` にヒット
- 数値比較（`--numeric`）: `number` 型のセルに対して比較演算を行う。`formula` 型のセルもキャッシュ値が数値の場合は対象とする。`date` 型のセルは対象外とする（日付をシリアル値で比較する用途は想定しない）
  - `">100"` — 100より大きい
  - `">=100"` — 100以上
  - `"<50"` — 50未満
  - `"100:200"` — 100以上200以下（範囲指定）
  - `"=42"` — 42と等しい（浮動小数点の誤差を考慮し、差の絶対値が 1e-9 以下なら等しいとみなす）
- 複数指定した場合はAND条件で絞り込む
- `--type` フィルタは `formula` 型セルのキャッシュ値の型も考慮する（例: `--type number` は数値キャッシュの数式セルにもヒットする）
- `--numeric` 単独指定時は数値セル（数式の数値キャッシュ含む）のみが対象、`--type` 単独指定時は該当型のセルをすべて出力
- `--numeric` の値はシェルの解釈を避けるためクォートが必要（例: `--numeric ">100"`, `--numeric "100:200"`）

## セルの出力構造

出力形式はJSONL固定。大きなファイルでも行単位でフィルタ・パイプ処理が可能。

### メタ情報（`_meta`）

cells の最初の行に出力される。シートのレイアウト基準値を含む。

```jsonl
{"_meta":true,"default_width":8.43,"default_height":15,"col_widths":{"B":24.5,"H":30}}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `_meta` | bool | 常に `true`。メタ情報行の識別用 |
| `default_width` | number | デフォルト列幅。XMLに未指定の場合は Excel 標準値 `9.140625` |
| `default_height` | number | デフォルト行高（ポイント単位） |
| `col_widths` | object | デフォルトと異なる列幅のマップ（キー: 列名、値: 幅）。差分がない場合は省略 |

### 行情報

行高や非表示など行レベルの情報がある場合、その行のセル出力の前に行情報を1行出力する。`_row` フィールドの有無でセル行と区別できる。行高がデフォルトかつ非表示でない行は行情報を省略する。

```jsonl
{"_row":1,"height":30}
{"cell":"A1","value":"タイトル"}
{"cell":"B1","value":"内容"}
{"cell":"A2","value":"データ"}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `_row` | number | 行番号 |
| `height` | number | 行高（デフォルト行高と異なる場合のみ出力） |
| `hidden` | bool | 行が非表示の場合のみ `true` を出力 |

### セルフィールド

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `cell` | string | セル座標（例: `B3`） |
| `value` | any | セルの値（文字列、数値、真偽値） |
| `display` | string | Excelでの表示文字列（`value` を文字列化した結果と異なる場合のみ出力。後述の詳細を参照） |
| `type` | string | `date` の場合のみ出力。他の型はJSONの値の型や `formula` フィールドの有無から推測可能なため省略する |
| `error` | bool | 値がExcelエラー（`#N/A`, `#REF!` 等）の場合のみ `true` を出力。数式セルのキャッシュ値がエラーの場合も含む |
| `merge` | string | 結合範囲（結合セルの場合のみ。例: `B3:D3`） |
| `formula` | string | 数式文字列（数式セルの場合のみ。例: `=SUM(D3:D9)`） |
| `link` | object | ハイパーリンク情報（リンクが設定されている場合のみ） |
| `hidden_col` | bool | 列が非表示の場合のみ `true` を出力 |
| `comment` | object | コメント/メモ（設定されている場合のみ） |

**型の判定ルール:**

- `type` 省略 + `value` が文字列 → string（エラー値 `#N/A`, `#REF!` 等も文字列として出力）
- `type` 省略 + `value` が数値 → number
- `type` 省略 + `value` が true/false → bool
- `type` 省略 + `formula` あり → formula（`value` はキャッシュ値）
- `type` 省略 + `value` なし → empty（`--include-empty` 時のみ出現）
- `type: "date"` → 日付（`value` は ISO 8601 文字列）

### 日付値

日付セルの `value` はISO 8601形式で出力する。

- 日付のみ: `"2025-03-15"`
- 日時: `"2025-03-15T10:30:00"`
- 時刻のみ: `"10:30:00"`

日付の判定は styles.xml の数値フォーマット情報に基づく。Excelは日付を内部的に数値（シリアル値）として保持しており、セルの表示形式が日付系フォーマット（`yyyy/mm/dd` 等）であれば日付セルとして扱う。認識できないカスタム表示形式の場合は数値セル（`number`）として出力する。

### `display` フィールド

`display` はセルがExcel上で表示する文字列をそのまま出力する。`value` をJSONエンコードした結果と同一の場合は省略する（例: 数値 `1000` → JSON表現 `"1000"` と比較。Go の `encoding/json` が生成する数値表現が基準）。bool値の場合、`value` は JSON の `true`/`false` だが Excel の表示は `TRUE`/`FALSE` であるため `display` を出力する。

**対象:**

- 数値フォーマット: 通貨（`¥1,000`）、パーセント（`50%`）、日付（`2025年3月15日`）、カスタム書式等
- 数式セル: 計算結果の表示文字列（例: `value` が `150` でも表示形式が通貨なら `display` は `"¥150"`）
- ハイパーリンク: セルの表示テキスト（リンク先URLとは異なる表示文字列がある場合）

**制限事項:**

- 条件付き書式（セル値に応じた動的なフォーマット）には非対応。セル自体の表示形式のみ適用する
- ロケール依存のフォーマット（`[$-ja-JP]` 等）は非対応
- 認識できないカスタム表示形式の場合、`display` は省略される（`value` のみ出力）

### 空セル（`--include-empty` 時）

```jsonl
{"cell":"C3","type":"empty"}
```

- `value` は省略する（`null` や空文字は出力しない）
- 書式フィールドはデフォルトと異なる場合のみ出力する（`--style` 指定時）

### 結合セル

- 結合領域の左上セルのみを出力する。結合に含まれる他のセルは出力しない（`--include-empty` 時も同様にスキップする）
- `--range` で結合領域の一部のみが含まれる場合、左上セルが範囲内なら出力する（左上セルが範囲外なら出力しない）

### セルの走査順序

セルは行優先順（row-major order）で走査する: A1 → B1 → C1 → ... → A2 → B2 → ... 。`--limit` による切り捨てや `next_cell` はこの走査順に基づく。

### `--range` の指定形式

以下の形式をサポートする:

- 矩形範囲: `A1:H20`
- 列のみ: `A:F`（行はExcelの使用範囲（`scan` の `used_range`）で補完）
- 行のみ: `1:20`（列はExcelの使用範囲（`scan` の `used_range`）で補完）
- 単一セル: `B5`（`B5:B5` と同等）
- 単一行: `3:3`
- 単一列: `A:A`

列のみ・行のみ指定では `used_range` から省略された行・列の上下限を補完する。空シート（`used_range` が空）の場合は空出力で正常終了する。

### 書式フィールド（`--style` 指定時のみ出力）

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `font` | object | フォント情報 |
| `fill` | object | 背景色 |
| `border` | object | 罫線 |
| `alignment` | object | 配置 |
| `rich_text` | array | リッチテキストラン（セル内の一部が異なる書式を持つ場合） |

### リッチテキスト

セル内の文字列の一部が異なる書式を持つ場合、`value` にはプレーンテキストを格納し、`rich_text` に書式付きランの配列を格納する。

```jsonl
{"cell":"B2","value":"本機能は必須とする。","rich_text":[{"text":"本機能は"},{"text":"必須","font":{"bold":true,"color":"#FF0000","size":14}},{"text":"とする。"}]}
```

- `value` は常にプレーンテキスト（検索・概要把握用）
- `rich_text` はセル内に書式の異なるランが存在する場合のみ出力
- 各ランの `font` はセルレベルの `font` オブジェクトと同じ構造。デフォルト値のフィールドは省略する。書式がセルのフォントと同一のランでは `font` 自体を省略する
- `--style` 未指定時は `rich_text` を省略

### 数式

数式セルでは `type` は出力せず、`formula` フィールドの有無で判定する。`formula` フィールドに数式文字列を格納し、`value` にはExcelが最後に保存した計算結果（キャッシュ値）が入る。

```jsonl
{"cell":"D10","value":150,"display":"¥150","formula":"=SUM(D3:D9)"}
```

- 数式の再計算は行わない（Excelエンジンではないため）
- キャッシュ値が存在しない場合、`value` は `null`
- 共有数式（shared formula）: ワークシートXMLの `<f>` 要素の文字列をそのまま格納する。共有数式の非プライマリセルでは `formula` が空になることがある（`value` はキャッシュ値を出力する）
- 配列数式（CSE数式）: 波括弧付きで格納する（例: `"{=SUM(A1:A10*B1:B10)}"`）。動的配列（スピル）の場合、スピル元セルのみに `formula` を格納し、スピル先セルは通常の値セルとして出力する

### エラー値

セルがExcelのエラー値を含む場合、`error: true` を出力し、`value` にエラー文字列をそのまま格納する。エラー値は文字列型として扱い、`type` は出力しない。

```jsonl
{"cell":"A1","value":"#N/A","error":true}
{"cell":"B2","value":"#REF!","error":true,"formula":"=Sheet1!A1"}
```

- エラー値: `#N/A`, `#REF!`, `#VALUE!`, `#DIV/0!`, `#NAME?`, `#NULL!`, `#NUM!`
- 数式セルの計算結果がエラーの場合も `error: true` を出力する
- `error` フラグにより、ユーザーが入力した `#N/A` という文字列と、Excelのエラー値を区別できる

### ハイパーリンク

セルにハイパーリンクが設定されている場合、`link` オブジェクトを出力する。

```jsonl
{"cell":"C5","value":"参照先ドキュメント","type":"string","link":{"url":"https://example.com/doc.pdf"}}
{"cell":"A1","value":"詳細はSheet2参照","type":"string","link":{"location":"Sheet2!A1"}}
```

- 外部URL → `url` フィールド
- シート内・ブック内リンク → `location` フィールド
- 両方が設定されている場合は両方出力

### コメント/メモ

セルにコメントまたはメモが設定されている場合、`comment` オブジェクトを出力する。レガシーコメント（`xl/comments.xml`）とスレッドコメント（`xl/threadedComments/`）の両方に対応する。

```jsonl
{"cell":"A1","value":"項目","comment":{"author":"田中太郎","text":"この項目は必須です"}}
{"cell":"B2","value":"値","comment":{"author":"佐藤花子","text":"確認お願いします","thread":[{"text":"確認しました","date":"2025-09-26T05:32:58.00"},{"text":"ありがとうございます","date":"2025-09-26T06:00:00.00","done":true}]}}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `author` | string | コメントの著者名。不明な場合は省略 |
| `text` | string | コメント本文 |
| `thread` | array | スレッドの返信（返信がある場合のみ出力） |

**`thread` 配列の各要素:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `author` | string | 返信の著者名。不明な場合は省略 |
| `text` | string | 返信本文 |
| `date` | string | 返信日時（ISO 8601形式）。不明な場合は省略 |
| `done` | bool | 解決済みの場合のみ `true` を出力 |

- コメントはシートの `.rels` から `comments` / `threadedComments` リレーションを辿って読み込む
- レガシーコメントとスレッドコメントが同一セルに存在する場合は統合する
- VML（`vmlDrawing.vml`）は視覚情報のみのためパースしない

### font オブジェクト

```json
{"name": "MS ゴシック", "size": 11, "bold": true, "italic": true, "strikethrough": true, "underline": "single", "color": "#FF0000"}
```

- セルのフォントがシートのデフォルトフォントと完全に一致する場合、`font` フィールド自体を省略する
- デフォルトフォントと異なるフィールドのみ出力する（例: デフォルトと name/size が同じで bold だけ異なる場合 → `{"bold": true}` のみ）
- 値がデフォルト（bold: false 等）のフィールドも省略する
- `italic`, `strikethrough` は bool 型。`false` の場合は省略する
- `underline` は文字列型で、値は `single`, `double`, `singleAccounting`, `doubleAccounting` のいずれか（下線なしの場合は省略）

### fill オブジェクト

ソリッド塗りつぶし（単色）の前景色のみ出力する。パターン塗りつぶしやグラデーションには対応しない。

```json
{"color": "#D9E2F3"}
```

### border オブジェクト

```json
{"top": {"style": "thin"}, "bottom": {"style": "thin", "color": "#FF0000"}, "left": {"style": "medium"}, "right": {"style": "thin"}, "diagonal_up": {"style": "thin"}}
```

- 罫線がないエッジは省略する。各エッジは `style` と `color` を持つオブジェクト
- エッジ: `top`, `bottom`, `left`, `right`, `diagonal_up`（左下→右上）, `diagonal_down`（左上→右下）
- `color` はデフォルト（黒）の場合は省略する
- スタイル値: `thin`, `medium`, `thick`, `dashed`, `dotted`, `double`, `hair`, `mediumDashed`, `dashDot`, `mediumDashDot`, `dashDotDot`, `mediumDashDotDot`, `slantDashDot`

### alignment オブジェクト

```json
{"horizontal": "center", "vertical": "center", "wrap": true, "indent": 2, "text_rotation": 90, "shrink_to_fit": true}
```

- `horizontal`: `left`, `center`, `right`, `fill`, `justify`, `distributed`。デフォルト（`general`）の場合は省略
- `vertical`: `top`, `center`, `bottom`, `justify`, `distributed`。デフォルト（`bottom`）の場合は省略
- `wrap`: bool。テキスト折り返し。`false` の場合は省略
- `indent`: number。インデントレベル。0の場合は省略
- `text_rotation`: number。テキストの回転角度（-90〜90、または255で縦書き）。0の場合は省略
- `shrink_to_fit`: bool。縮小して全体を表示。`false` の場合は省略

### 切り捨て通知

`cells` / `search` で `--limit` に到達した場合、最終行に切り捨て通知を出力する。`--limit 0`（上限なし）の場合や、対象セル数が `--limit` 未満の場合は出力されない。

```jsonl
{"_truncated":true,"next_cell":"A51"}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `_truncated` | bool | 常に `true` |
| `next_cell` | string | 打ち切られた次の開始位置。そのまま `--start` に渡して続きを取得する |

- この行は `_truncated` フィールドの有無で通常のセル行と区別できる
- `next_cell` を `--start` に渡して続きを取得する
- `search` で続きを取得する場合、同じフィルタ条件（`--query`, `--numeric`, `--type`）を再度指定する必要がある

### 出力例

```jsonl
{"_row":1,"height":30}
{"cell":"B1","value":"基本設計書","merge":"B1:H1","font":{"size":16,"bold":true},"alignment":{"horizontal":"center"}}
{"_row":2}
{"cell":"B2","value":"文書番号","merge":"B2:C2","fill":{"color":"#D9E2F3"}}
```

## 技術選定

- **言語:** Go
- **Excelパーサー:** 自前実装（ZIP + encoding/xml による直接パース）
- **CLIフレームワーク:** [cobra](https://github.com/spf13/cobra)

## エラーハンドリング

- ファイルが存在しない / 読み取れない → エラーメッセージを stderr に出力
- .xlsx / .xlsm 以外のファイル → 「.xlsx / .xlsm 形式のみ対応」のエラーメッセージ
- 存在しないシート名/インデックス → 利用可能なシート一覧をエラーメッセージに含める
- 不正なオプション値（range, numeric） → パース失敗箇所を示すエラーメッセージ。論理的に不正な範囲（例: 終端が始端より前）もエラーとする
- `--range` と `--start` の併用 → エラーメッセージを出力（排他オプション）
- パスワード保護されたファイル → 非対応としてエラーメッセージを出力
- 破損したファイル（不正なzip構造等） → エラーメッセージを出力
- ワークシート以外のシート（チャートシート等）を `scan` / `cells` / `search` で指定 → 「ワークシートのみ対応」のエラーメッセージ
- シートにセルが一つもない場合 → `scan` は `used_range` を省略する。`cells` は `_meta` 行のみ出力（セルなし）で正常終了する
- `search` で結果が0件の場合 → 空出力（0行のJSONL）で正常終了する
- 終了コード: 0=成功（検索結果なしも含む）、1=エラー。上記のエラーケースはすべて終了コード 1
- エラーメッセージは stderr に `cc-read-excel: <メッセージ>` の形式で出力する。stdout には常にJSONL（またはJSON）のみを出力する

## 設計方針

- 対応形式は .xlsx および .xlsm（.xls は非対応）
- ZIP 内の XML を自前で直接パースする（excelize 等の外部 Excel パーサーは使用しない）
- ワークシート XML は SAX（ストリーミング）パースで処理し、メモリ使用を最小限に抑える
- `cells` / `search` の出力形式はJSONL固定（1行1セルのストリーム出力）。`info` / `scan` はJSON形式で出力する
- デフォルトで空セルをスキップする（Excel方眼紙は空セルが大量にあるため）
- 書式フィールドはデフォルト値と異なる場合のみ出力し、出力サイズを抑える
- シートのデフォルトフォントは styles.xml の最初のフォント定義から取得する
- `cells` / `search` の各セルの `font` はデフォルトフォントとの差分のみ出力する（トークン効率のため）
- 列幅は `cells` の `_meta` 行で出力する（デフォルト値との差分のみ）。行高は `cells` の行情報（`_row`）で行が変わるごとに出力する（デフォルト行高との差分のみ）
- 色は可能な限り `#RRGGBB` 形式（HEX RGB）で出力する。テーマカラーは theme1.xml から RGB に変換し、tint 値がある場合は HSL 色空間で明度を調整して適用する
- 日付値はISO 8601形式で統一する
- 結合セルは左上セルのみ出力する
- `--limit` のデフォルトは1000（Claude Codeのコンテキスト窓を考慮）
- 出力は常にUTF-8エンコーディング
- セル値に含まれる制御文字（改行、タブ等）はJSON仕様に従いエスケープする（`\n`, `\t` 等）
- SKILL化を見据え、コマンドの出力はパイプ処理しやすい形式にする

## コマンド構成（その他）

### `cc-read-excel version`

バージョン情報をプレーンテキストで出力する。

```
cc-read-excel version 0.1.0
```

バージョン番号は `go build -ldflags` でビルド時に埋め込む。未設定の場合は `dev` を表示する。

### `cc-read-excel shapes <file>`

**役割:** シート上の図形（オートシェイプ、テキストボックス、コネクタ、グループ、画像）を取得する。フローチャートや図解の構造を、図形間の接続関係を含めて把握する。

各シートの drawing XML（`xl/drawings/drawingN.xml`）をパースし、図形情報をJSONL形式で出力する。

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--sheet <name\|index>` | 対象シート（名前 or 0始まりインデックス） | 最初のシート |
| `--limit <n>` | 出力図形数の上限 | 1000 |
| `--style` | 書式情報（塗りつぶし、枠線、フォント）を出力する | OFF（書式は省略） |
| `--extract-images <dir>` | 画像を指定ディレクトリに抽出する。未指定時は画像をスキップ | OFF（画像スキップ） |

- `--limit 0` で上限なし
- `--extract-images` 未指定時は `xdr:pic` 要素をスキップする（画像は出力に含まれない）

**出力例:**

```jsonl
{"_meta":true,"shape_count":8,"connector_count":3}
{"id":1,"type":"flowChartTerminator","name":"開始","text":"開始","cell":"B2:D3","z":0}
{"id":2,"type":"flowChartProcess","text":"データ取得","cell":"B5:D7","z":1}
{"id":3,"type":"flowChartDecision","text":"正常？","cell":"B9:D12","z":2}
{"id":4,"type":"flowChartProcess","text":"処理実行","cell":"F9:H11","z":3}
{"id":5,"type":"flowChartTerminator","text":"終了","cell":"B14:D15","z":4}
{"id":6,"type":"connector","from":1,"to":2,"connector_type":"straightConnector","arrow":"end","z":5}
{"id":7,"type":"connector","from":2,"to":3,"connector_type":"straightConnector","arrow":"end","z":6}
{"id":8,"type":"connector","from":3,"to":4,"label":"Yes","connector_type":"elbowConnector","arrow":"end","z":7}
```

`--style` 指定時:
```jsonl
{"id":1,"type":"flowChartTerminator","text":"開始","cell":"B2:D3","z":0,"fill":"#4A86E8","line":{"color":"#000000","style":"solid","width":1},"font":{"bold":true,"color":"#FFFFFF","size":11}}
```

#### メタ情報（`_meta`）

shapes の最初の行に出力される。

```jsonl
{"_meta":true,"shape_count":8,"connector_count":3}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `_meta` | bool | 常に `true`。メタ情報行の識別用 |
| `shape_count` | number | 図形の総数（コネクタを含む） |
| `connector_count` | number | コネクタの数（`shape_count` の内数） |

#### 図形フィールド

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `id` | number | 図形ID。drawing XML 内の出現順で1始まりの連番を割り当てる |
| `type` | string | 図形種別。後述の図形種別一覧を参照 |
| `name` | string | 図形名（Excel上で設定された名前。`xdr:nvSpPr/xdr:cNvPr` の `name` 属性）。省略不可 |
| `text` | string | 図形内テキスト（`a:txBody` 内の全 `a:t` を結合。複数段落は `\n` で結合）。テキストがない場合は省略 |
| `cell` | string | アンカー位置をセル範囲で表現（例: `B2:D4`）。`twoCellAnchor` / `oneCellAnchor` の列・行インデックスから算出。`absoluteAnchor` の場合は省略 |
| `z` | number | Z-order。drawing XML 内のアンカー要素の出現順で0始まり。大きいほど前面 |
| `rotation` | number | 回転角度（度単位、時計回り）。`a:xfrm` の `rot` 属性を60000で除算して度に変換。0の場合は省略 |
| `flip` | string | 反転。`"h"`（水平）、`"v"`（垂直）、`"hv"`（両方）。`a:xfrm` の `flipH` / `flipV` 属性から判定。なければ省略 |
| `rich_text` | array | リッチテキストラン。図形内テキストに書式の異なるランが存在する場合のみ出力。構造はセルの `rich_text` と同一 |

#### コネクタフィールド

コネクタ（`xdr:cxnSp`）は `type` が `"connector"` となり、以下の追加フィールドを持つ。

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `from` | number | 接続元の図形ID。接続情報がない場合は省略 |
| `to` | number | 接続先の図形ID。接続情報がない場合は省略 |
| `connector_type` | string | コネクタの形状。`a:prstGeom` の `prst` 属性値（`line`, `straightConnector1`, `bentConnector3`, `curvedConnector3` 等） |
| `arrow` | string | 矢印の位置。`"start"`, `"end"`, `"both"`, `"none"`。`a:ln` の `a:headEnd` / `a:tailEnd` の `type` 属性から判定。省略時は `"none"` |
| `label` | string | コネクタ上のテキスト。テキストがない場合は省略 |

コネクタの `from` / `to` は、drawing XML 内の `a:stCxn` / `a:endCxn` 要素の `id` 属性を参照する。この `id` は Excel が付与する図形IDであり、`shapes` コマンドが割り当てる連番 `id` とは異なる。パース時に Excel 図形IDから連番IDへのマッピングを行い、出力時は連番IDで参照する。接続先が drawing 内に見つからない場合は `from` / `to` を省略する。

#### 画像フィールド（`--extract-images` 指定時）

画像（`xdr:pic`）は `type` が `"picture"` となり、以下の追加フィールドを持つ。`--extract-images` 未指定時は画像要素自体がスキップされる。

```jsonl
{"id":10,"type":"picture","name":"図 1","cell":"B2:F8","z":5,"alt_text":"システム構成図","image":{"format":"png","width":640,"height":480,"size":45230,"path":"/tmp/shapes/image_1.png"}}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `alt_text` | string | 代替テキスト（`cNvPr` の `descr` 属性）。設定されていない場合は省略 |
| `image` | object | 画像メタデータ |

**`image` オブジェクト:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `format` | string | 画像形式（`png`, `jpeg`, `emf` 等。リレーション先のファイル拡張子から判定） |
| `width` | number | 画像の幅（ピクセル）。`a:ext` の `cx` をEMUからピクセルに変換（÷ 9525） |
| `height` | number | 画像の高さ（ピクセル）。`a:ext` の `cy` をEMUからピクセルに変換（÷ 9525） |
| `size` | number | ファイルサイズ（バイト）。ZIPエントリから取得 |
| `path` | string | 抽出先のファイルパス。`--extract-images` で指定したディレクトリ内に `image_1.png`, `image_2.jpg`, ... の形式で連番出力 |

画像ファイルは drawing の `.rels` から `blip` の `r:embed` 属性で参照されるリレーションIDを解決し、ZIP内の `xl/media/` 配下から抽出する。

#### グループフィールド

グループ（`xdr:grpSp`）は `type` が `"group"` となり、以下の追加フィールドを持つ。

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `children` | array | 子要素のID配列（出現順） |

グループの子要素は `parent` フィールドで親グループを参照する。

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `parent` | number | 親グループのID。トップレベルの図形では省略 |

- グループ内の子要素の `z` はグループ内での相対順序（0始まり）
- グループ自体の `z` はシートレベルでの重なり順
- ネストしたグループも同じ構造で再帰的に表現する
- グループの `cell` はグループ全体のアンカー範囲

#### 書式フィールド（`--style` 指定時）

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `fill` | string | 塗りつぶし色（`#RRGGBB` 形式）。`a:solidFill` の色を解決。塗りつぶしなしの場合は省略 |
| `line` | object | 枠線情報。`a:ln` から取得 |
| `font` | object | テキストのデフォルトフォント。図形内の `a:defRPr` または最初の `a:rPr` から取得 |

**`line` オブジェクト:**

```json
{"color": "#000000", "style": "solid", "width": 1.5}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `color` | string | 線の色（`#RRGGBB` 形式） |
| `style` | string | 線のスタイル。`solid`, `dash`, `dot`, `dashDot`, `lgDash`, `lgDashDot`, `sysDash`, `sysDot`, `sysDashDot`, `sysDashDotDot` |
| `width` | number | 線幅（ポイント単位）。`a:ln` の `w` 属性をEMUからポイントに変換（÷ 12700） |

枠線がない場合は `line` フィールド自体を省略する。

**`font` オブジェクト:**

セルの `font` オブジェクトと同じ構造（`name`, `size`, `bold`, `italic`, `strikethrough`, `underline`, `color`）。デフォルト値のフィールドは省略する。

#### 図形種別

`type` フィールドの値は以下のいずれか:

- **コネクタ**: 常に `"connector"`（`xdr:cxnSp` 要素に対応）
- **グループ**: 常に `"group"`（`xdr:grpSp` 要素に対応）
- **シェイプ**: `a:prstGeom` の `prst` 属性値をそのまま使用する（`xdr:sp` 要素に対応）

シェイプの `prst` 値の例:

| prst 値 | Excel上の名称 |
|---------|-------------|
| `rect` | 四角形 |
| `roundRect` | 角丸四角形 |
| `ellipse` | 楕円 |
| `triangle` | 三角形 |
| `diamond` | ひし形 |
| `flowChartProcess` | 処理 |
| `flowChartDecision` | 判断 |
| `flowChartTerminator` | 端子 |
| `flowChartPredefinedProcess` | 定義済み処理 |
| `flowChartDocument` | 書類 |
| `flowChartMultidocument` | 複数書類 |
| `flowChartManualInput` | 手操作入力 |
| `flowChartManualOperation` | 手操作 |
| `flowChartPreparation` | 準備 |
| `flowChartInternalStorage` | 内部記憶 |
| `flowChartDisplay` | 表示 |
| `flowChartMerge` | 合流 |
| `flowChartConnector` | 結合子 |
| `flowChartOffpageConnector` | 他ページ結合子 |

`a:prstGeom` が存在しない場合（カスタムジオメトリ `a:custGeom`）は `type` を `"customShape"` とする。

#### セル範囲の算出

`twoCellAnchor` の場合:
- `xdr:from` の `col` と `row`（0始まり）から開始セルを算出
- `xdr:to` の `col` と `row`（0始まり）から終了セルを算出
- セル範囲文字列を構築（例: col=1,row=1 → col=3,row=3 で `"B2:D4"`）
- EMUオフセット（`colOff`, `rowOff`）は無視する（セル単位の粗い位置で十分）

`oneCellAnchor` の場合:
- `xdr:from` の `col` と `row` から開始セルを算出
- `a:ext` の `cx`, `cy`（EMU単位のサイズ）は無視し、開始セルのみを `cell` フィールドに出力する（例: `"B2"`）

`absoluteAnchor` の場合:
- セル座標への変換にシートの列幅・行高情報が必要なため、`cell` フィールドは省略する

#### scan コマンドの拡張

`scan` の出力に `has_drawings` フィールドを追加する。シートの `.rels` ファイル内に drawing タイプのリレーションが存在する場合に `true` を出力する。

```json
{"sheet": "フロー図", "used_range": "A1:H200", "has_drawings": true}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `has_drawings` | bool | シートに図形が存在する場合のみ `true` を出力。図形がない場合は省略 |

drawing リレーションの有無のみを確認し、drawing XML 自体は読み込まない（パフォーマンスへの影響なし）。

#### 対応しない図形要素

| 要素 | 理由 |
|------|------|
| チャート（`xdr:graphicFrame` + chart URI） | 独自のXML体系（`c:chartSpace`）で複雑。別途対応を検討 |
| SmartArt（`xdr:graphicFrame` + diagram URI） | 独自のXML体系（`dgm:`）で複雑。Excel保存時に描画キャッシュとして `grpSp` に展開されるため、通常はグループとして読み取り可能 |
| VML図形（`vmlDrawing.vml`） | レガシー形式。DrawingML（`drawing.xml`）のみ対応 |

## 対応しない機能

以下の機能は意図的に対応しない。理由とともに記録する。

| 機能 | 理由 |
|------|------|
| .xls 形式の読み取り | OOXML（.xlsx/.xlsm）のみ対応。旧形式のファイルは事前に .xlsx へ変換して使用する |
| 数式の再計算 | Excelエンジンではないため再計算は行わない。キャッシュ値（最後に保存された計算結果）を出力する |
| パスワード保護されたファイル | 非対応 |
| 塗りつぶしのパターン・グラデーション | ソリッド塗りつぶし（単色）の前景色のみ対応。パターンやグラデーションは利用頻度が低く、出力が複雑になるため非対応 |
| 条件付き書式の評価 | セルの表示形式（`display`）は静的な数値フォーマットのみ適用する。条件付き書式は実行時の評価が必要であり、Excelエンジンなしでは正確に再現できない |
| 正規表現による検索 | `--query` は部分一致のみ。正規表現はシェル側の `grep` と組み合わせて実現する想定 |
| チャートシート・マクロシートの出力 | ワークシートのみ対応。`info` コマンドではシート種類として表示するが、`cells` / `search` の対象外とする |
| Excel書き込み・編集 | 読み取り専用ツールとして設計。書き込みは別ツールの責務とする |
| 複数シートの同時指定・全シート指定 | `--sheet` は常に単一シートのみ指定可能。`--sheet all` やワイルドカード等による複数シート・全シート指定には対応しない。複数シートを処理する場合はシートごとにコマンドを実行する。`info` でシート一覧を確認し、対象シートを特定する運用を想定。これはブック横断検索（`search --all-sheets` 等）も同様で、対応しない |
| メモリ使用量の上限 | 設けない。ワークシート XML は SAX パースでストリーミング処理するため、メモリ使用は共有文字列テーブルのサイズに依存する。メモリ制限やファイルサイズ制限は設定しない |

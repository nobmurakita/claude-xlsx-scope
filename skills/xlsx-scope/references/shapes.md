# shapes コマンド

```bash
xlsx-scope shapes [options] <file>
```

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--sheet <name\|index>` | 対象シート | 最初のシート |
| `--limit <n>` | 出力図形数の上限（0で無制限） | 1000 |
| `--style` | 書式情報を出力 | OFF |

出力例:
```jsonl
{"_meta":true,"shape_count":8,"connector_count":3}
{"id":1,"type":"roundRect","text":"処理A","cell":"B2:D4","pos":{"x":120,"y":80,"w":200,"h":60},"adj":{"adj1":16667},"z":0}
{"id":2,"type":"flowChartDecision","text":"条件分岐","cell":"B6:D8","pos":{"x":120,"y":200,"w":200,"h":80},"z":1}
{"id":3,"type":"connector","cell":"B4:B6","pos":{"x":220,"y":140,"w":0,"h":60},"from":1,"to":2,"from_idx":2,"to_idx":0,"connector_type":"bentConnector3","adj":{"adj1":50000},"arrow":"end","start":{"x":220,"y":140},"end":{"x":220,"y":200},"z":2}
{"id":4,"type":"wedgeRoundRectCallout","text":"注意","cell":"E2:G4","pos":{"x":300,"y":50,"w":150,"h":40},"callout_target":{"x":269,"y":75},"z":3}
```

- `pos`: ピクセル座標（96 DPI）。`{x, y, w, h}` で左上原点。グループ内の子要素では省略
- `start`/`end`: コネクタの始点・終点座標。`pos` と `flip` から算出
- `callout_target`: 吹き出しのポインタ先座標。wedge 系等の吹き出し形状でのみ出力
- `adj`: 図形の調整値（1/100000単位の比率）。角丸半径、台形の傾き等の形状パラメータ。例: `roundRect` の `adj1: 16667` は短辺の約16.7%が角丸半径

## 図形種別

- シェイプ: `rect`, `roundRect`, `ellipse`, `flowChartProcess`, `flowChartDecision`, `flowChartTerminator` 等
- 吹き出し: `wedgeRectCallout`, `wedgeRoundRectCallout` 等。`callout_target` でポインタ先を出力
- コネクタ: `type` は常に `"connector"`。`from`/`to` で接続先の図形IDを参照。`from_idx`/`to_idx` で接続ポイントのインデックス（図形上の接続位置、形状依存）。`start`/`end` で両端座標。`connector_type` でコネクタ形状。`adj` で屈曲・カーブの調整値
- グループ: `type` は `"group"`。`children` に子要素ID配列。子要素は `parent` で親を参照
- 画像: `type` は `"picture"`。`image_id` で `image` サブコマンドにより画像を取得可能

## 図形内テキスト

- `text`: プレーンテキスト（複数段落は `\n` で結合）
- `rich_text`: 書式の異なるランがある場合のみ出力（`--style` 指定時）
- コネクタのテキストは `label` フィールド

## 画像の確認方法

出力の `image_id` を使い、`image` サブコマンドで画像のバイナリを取得できる。

```bash
xlsx-scope image example.xlsx xl/media/image1.png
```

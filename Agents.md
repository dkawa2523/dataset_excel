
# Excel作業者が「ぶれなく登録」できるデータ収集パイプライン（Python + Excel + ClearML）仕様書 / 実装指示書

> **このMarkdownは VSCode 上で Codex / Copilot に渡して実装しやすい形**を意図しています。  
> 実装は「後で各処理を個別に実行・差し替え・改良できる」よう **CLIサブコマンド + モジュール分割**で構成します。  
> 解析・特徴量化のためではなく、**作業者が確実に登録するための仕組み**を優先します。

---

## 0. 背景と目的

### 背景（課題）
- Excel作業者は「その場限り」で入力・加工しがちで、**フォーマット統一 / 収集 / 共有**が破綻しやすい。
- NotebookやPythonコードを渡しても慣れておらず、**業務フローが変わる/手間が増える**理由で使われない。

### 目的（最優先）
- Excel作業者が **操作数を減らした状態**で、入力フォーマットが統一され、データが **ClearML Dataset に蓄積**される。
- 将来、f の選択や計算式が変わっても、作業者の手間を増やさず **再処理で追従**できる。

### 非目的（やらない）
- モデル学習や高度な解析機能の提供は目的ではない（登録の邪魔になるため）。
- 作業者に YAML / ClearML UI / Python を操作させない（例外は管理者のみ）。

---

## 1. 利用者と責務（“ぶれ”を防ぐ役割固定）

### 作業者（Excelユーザー）
- すること：テンプレExcelを開く → 入力 → **登録ボタン1つ**
- しないこと：YAML編集、Python実行、ClearML操作

### 管理者（テンプレ設計者）
- YAMLを設計/更新
- テンプレExcel生成
- ClearMLにテンプレ（Task/Artifact）登録

### システム（ClearML Agent / runner）
- 入力検証、ファイル収集、統一フォーマット生成、ClearML Dataset 新Version登録
- 登録状況（欠損率など）を ClearML Plots / Debug samples に可視化

---

## 2. 要件（ユーザー提示の[1]〜[3]）

### [1] データフォーマット作成
- YAML で列名/型/必須/入力制約を定義し、テンプレExcelを生成
- 多次元/別ブック/別ファイル（x,y,z,t,f）を **ファイルパス列で紐づけ**
- 外部ファイルから x,y,z,t,f を選択（候補列名、型、optional/skipを指定可能）
- 目的変数 f は複数列でもOK
- 簡易計算（+ - * /、係数、定数など）を YAML で指定できること

### [2] テンプレ登録とClearML連携、アドイン（ボタン実行）
- YAMLで ClearML の Dataset project/name を指定して登録
- YAML設定を ClearML Task の Configuration / Parameters に保存し、クローンして再利用可能
- テンプレExcelを ClearML Artifacts からダウンロード可能
- テンプレExcelには「登録ボタン」（＝アドイン/マクロ/外部exe呼び出し）を含める
- ボタン実行で以下を自動実行
  1. 入力検証＋外部ファイル処理＋統一フォーマットExcel生成
  2. 元Excel＋参照ファイル＋処理済みファイルを Dataset 新Version として登録
  3. クローン先プロジェクトに自動紐づく（テンプレ埋め込みメタで判定）
  4. Excel内に ClearMLプロジェクトリンク（データに影響しない）を表示
  5. 登録状況のグラフを Plots に、画像参照なら Debug samples に配置

### [3] ClearMLからまとまったデータのダウンロード
- Excelのリンクから ClearMLプロジェクトへ遷移してDL
- 後で f の選択/計算/特徴量化を変えたい場合：
  - Task をクローンし、Configuration/Parameters を変更して enqueue
  - Artifact/Plots/Debug samples が更新され、処理済みExcel（ボタン込み）をDL可能

---

## 3. “ぶれなく登録”させるための設計（追加提案：登録の成功率を最大化）

> **解析ではなく、登録完了まで迷わせない/失敗させない/諦めさせない**ための工夫。

### 3.1 Excel側の強制力（入力ブレを物理的に減らす）
- 必須列が埋まるまで「登録ボタンを無効」 or 実行時に即ブロック
- 列名/テーブル構造を編集できないように **シート保護**（入力セルのみ許可）
- enum は **プルダウン**（自由記述を極小化）
- 日時・数値は入力規則（format/範囲）

### 3.2 失敗時の“直しどころ”をExcelへ返す
- エラーは必ず「シート名!セル番地」と「理由」を一覧化
- エラー行をハイライト（条件付き書式 or 直接スタイル付与）
- 成功時は「受付番号（dataset id / version / timestamp）」をExcelに書き戻す  
  → “登録できた体験”を強くする

### 3.3 パス入力事故を減らす（現場離脱の主要因）
- ファイル存在チェック、拡張子チェック、サイズ上限、ネットワークパス到達性
- 失敗したら “どのパスがNGか” を明示

### 3.4 重複登録防止（かなり効く）
- 元Excel＋参照ファイルのハッシュを manifest に保存
- 同一ハッシュが過去にあれば警告（ただし登録は可能）
- 登録理由をドロップダウンで選択（「誤り修正」「再測定」「追加データ」など）

### 3.5 オフライン耐性（“失敗を作業者の責任”にしない）
- アップロード失敗時は outbox にパッケージ保存
- 次回ボタン実行時に自動再送

### 3.6 テンプレ古い問題を潰す
- テンプレに template_version/config_hash を埋め込み
- 実行時に ClearML 上の最新テンプレと比較し、古ければリンク表示（可能なら自動DL）

---

## 4. 推奨技術スタック

- CLI: `typer`（または `click`）
- YAML: `PyYAML`
- 設定スキーマ: `pydantic`
- Excel生成/編集: `openpyxl`
- DataFrame: `pandas`
- DataFrame検証: `pandera`（任意）
- 画像サムネ: `Pillow`
- 安全な簡易式評価: AST制限 eval（ホワイトリスト） or `numexpr`
- ClearML: `clearml` SDK
- テスト: `pytest` + `pytest-cov`
- 形式: `ruff`（lint） + `black`（format）+ `mypy`（任意）

> Excelからの実行は **VBAボタン → runner.exe（PyInstaller）** を推奨  
> 作業者PCにPythonを入れずに済むため採用率が上がる。

---

## 5. リポジトリ構成（後で個別実行/改修できる分割）

```

repo/
pyproject.toml
README.md
configs/
sample_config.yaml
scripts/
make_testdata.py
src/datapipeline/
**init**.py
cli.py

```
config/
  schema.py        # Pydantic models
  loader.py        # YAML load + migrate + hash

excel/
  template.py      # generate template.xlsx
  reader.py        # read filled template
  writer.py        # write status/errors back

validate/
  engine.py        # validate main table + file paths
  rules.py         # reusable rules

extract/
  readers.py       # file readers (csv/excel/parquet/image)
  mapper.py        # resolve x,y,z,t,f columns
  compute.py       # derived calculations (safe)
  normalize.py     # output normalized tables

pack/
  manifest.py      # list files + hashes + metadata
  packer.py        # build package folder

clearml_io/
  template_task.py # register template task/artifacts
  dataset_upload.py# upload new dataset version
  reporting.py     # plots/debug samples
```

tests/
test_config.py
test_template_roundtrip.py
test_validation.py
test_formula_eval.py
test_pack_manifest.py
test_offline_outbox.py
test_integration_local.py  # dry-run integration

````

---

## 6. CLI設計（サブコマンドで個別実行）

> **重要**：すべてのコマンドは `--dry-run` を実装し、ClearMLが無い環境でもテスト可能にする。

### 6.1 コマンド一覧
- `dp generate-template --config configs/sample_config.yaml --out out/template.xlsx`
- `dp validate-excel --config ... --excel filled.xlsx --out out/validated.xlsx`
- `dp pack-dataset --config ... --excel filled.xlsx --out out/package_dir/`
- `dp upload-clearml --config ... --package out/package_dir/`
- `dp run --config ... --excel filled.xlsx`  
  → validate → pack → upload（作業者ボタンはこれを叩く想定）
- `dp dev make-testdata --out out/testdata/`（開発用）
- `dp reprocess --task-id <id>`（ClearML Agentで再処理用の入口、任意）

### 6.2 `dp run` の入出力（作業者に見えるもの）
**入力**
- テンプレExcel（入力済み）
- 参照ファイル（パスで指定される）

**出力**
- Excelに結果を書き戻し（Statusシート）
- 成功時：ClearML dataset version のURL/ID を表示
- 失敗時：エラー一覧（セル位置付き）

---

## 7. YAML仕様（最小サンプル + 実装指示）

### 7.1 YAMLの基本構造（案）
- `schema_version`（必須）
- `template.id`, `template.version`
- `main_table`（Excel入力テーブル定義）
- `file_profiles`（外部ファイル読み取りと x,y,z,t,f マッピング）
- `derived`（簡易計算）
- `clearml`（登録先）
- `outputs`（生成物）

### 7.2 サンプル `configs/sample_config.yaml`
```yaml
schema_version: 1
template:
  id: "spacetime_measure"
  version: "1.0.0"

main_table:
  sheet: "Input"
  table_name: "Records"
  columns:
    - name: record_id
      dtype: str
      required: true
      auto: uuid4
    - name: operator
      dtype: str
      required: true
      enum: ["A", "B", "C"]
    - name: measured_at
      dtype: datetime
      required: true
      format: "%Y-%m-%d %H:%M"
    - name: condition
      dtype: str
      required: false
    - name: data_path
      dtype: path
      required: true
      role: file_link
      profile: "spacetime_csv"

file_profiles:
  spacetime_csv:
    reader:
      type: csv
      encoding: "utf-8"
    map:
      x: { candidates: ["x","X"], dtype: float }
      y: { candidates: ["y","Y"], dtype: float }
      z: { candidates: ["z","Z"], dtype: float, optional: true, on_missing: blank }
      t: { candidates: ["t","time"], dtype: float }
      f:
        candidates: ["f","force","pressure"]
        dtype: float
        allow_multiple: true
        on_missing: error
    derived:
      - name: f_over_t
        expr: "f / (t + 1e-9)"
        dtype: float

clearml:
  enable: false     # 開発用はfalse。実運用でtrueにする。
  template_project: "Templates/Spacetime"
  dataset_project: "Datasets/Spacetime"
  dataset_name: "spacetime_measure"
  tags: ["excel", "ingest"]
  latest_tag: "latest"

outputs:
  normalized:
    format: ["parquet", "xlsx"]
    long_table: true
  store_raw_inputs: true
  make_thumbnails: true
````

---

## 8. Excelテンプレ仕様（openpyxlで生成）

### 8.1 シート構成

* `Input`：作業者が入力する（テーブル化）
* `Instructions`：入力例/手順（短い・1画面）
* `Status`：検証結果/登録結果（作業者が見る）
* `_META`（隠し）：template_id/version、config_hash、ClearMLリンク、runnerバージョン

### 8.2 Inputテーブル（Records）

* 先頭行に説明（固定表示）
* Excelテーブル化（行追加しやすい）
* enum列はプルダウン
* required列は条件付き書式で未入力を色付け
* `data_path` は “パスの例” を表示（UNC含む）

### 8.3 Excelボタン（VBAで runner.exe を呼ぶ）

> 実運用は「VBAボタン→exe」が最小摩擦。
> 開発中は `dp run ...` を手で実行してOK。

#### VBA例（Windows想定）

```vb
Sub RegisterDataset()
    Dim wbPath As String
    wbPath = ThisWorkbook.FullName

    ' runner.exe はテンプレと同じフォルダ or 既定インストール先に置く
    Dim exePath As String
    exePath = ThisWorkbook.Path & "\runner.exe"

    Dim cmd As String
    cmd = """" & exePath & """ run --excel """ & wbPath & """"

    Shell cmd, vbNormalFocus
End Sub
```

---

## 9. 実装タスク（Codex向け指示）

### 9.1 まず作るMVP（Phase 0）

* [ ] YAMLロード + pydantic検証 + `config_hash` 生成
* [ ] `generate-template`：Excelテンプレ生成（Input/Instructions/Status/_META）
* [ ] `validate-excel`：必須/型/enum/日付/パス存在チェック、エラーをStatusへ書き戻し
* [ ] `pack-dataset`：

  * 元Excelコピー
  * 参照ファイルコピー
  * manifest.json（ファイル一覧・hash・行数・欠損率など）作成
* [ ] `run`：validate→pack（ClearMLは後回しでOK）
* [ ] `--dry-run`：ClearML無しでも完走し、outディレクトリが得られる

### 9.2 次に作る（Phase 1）

* [ ] 外部ファイル読み取り（csv/excel/parquet）
* [ ] x,y,z,t,f マッピング解決（候補列名）
* [ ] `derived` の簡易式評価（安全な式のみ）
* [ ] 正規化出力：long table（parquet/xlsx）

### 9.3 ClearML連携（Phase 2）

* [ ] テンプレTask登録（config.yaml, template.xlsx, runner.exe をArtifactへ）
* [ ] Dataset新Version登録（親Datasetを指定、latestタグ運用）
* [ ] Plots/Debug samples への登録状況出力（欠損率、行数、画像サムネ）

### 9.4 運用強化（Phase 3）

* [ ] 重複検知（hash一致）
* [ ] offline outbox（アップロード失敗時に保存→次回再送）
* [ ] テンプレ更新検知（古いテンプレ警告）

---

## 10. 安全な簡易式評価（必須仕様）

### 10.1 許可する式

* 演算子：`+ - * / ( )`
* 参照：列名（例：`f`, `t`）
* 定数：数値（例：`1e-9`）
* 許可関数（ホワイトリスト）：`abs`, `min`, `max`（必要なら）

### 10.2 禁止

* 属性アクセス、関数定義、import、os操作など
* 任意コード実行につながる記法

> 実装指示：`ast.parse` → ノード種別を検査 → 安全な環境で評価
> `eval` 直呼びは禁止。

---

## 11. テストデータ作成（開発者がすぐ回せること）

### 11.1 `scripts/make_testdata.py` の要件

出力先 `out/testdata/` に以下を生成：

* `configs/sample_config.yaml`（上記サンプルでも良い）
* `template.xlsx`（generate-templateで生成でも良い）
* `filled.xlsx`（入力済みExcel：3行程度）
* `linked/`：

  * `sample1.csv`（x,y,t,f列あり）
  * `sample2.csv`（z無し、optionalを試す）
  * `bad.csv`（列不足→エラー用）
  * `image1.png`（Debug samples想定、任意）

### 11.2 生成するCSV例

* sample1.csv: columns = x,y,t,f
* sample2.csv: columns = x,y,t,pressure（f候補の別名）
* bad.csv: columns = x,y（t,f欠落）

---

## 12. 実行テスト（ローカルで完結するテスト計画）

### 12.1 手動E2E（ClearMLなし / dry-run）

1. `dp dev make-testdata --out out/testdata`
2. `dp generate-template --config out/testdata/config.yaml --out out/testdata/template.xlsx`
3. `dp run --config out/testdata/config.yaml --excel out/testdata/filled.xlsx --dry-run --out out/run1`

**期待結果**

* out/run1/package_dir/ ができる
* manifest.json がある
* 正規化出力（parquet/xlsx）がある（Phase 1以降）
* filled.xlsx の Status シートが更新される（成功/エラーが見える）

### 12.2 自動テスト（pytest）

* `test_config.py`

  * YAMLがpydanticで検証される
  * config_hash が安定生成される
* `test_template_roundtrip.py`

  * generate-template → reader で列が一致する
* `test_validation.py`

  * 必須欠落 / enum違反 / 日付format違反 / パス不在 が検出される
* `test_formula_eval.py`

  * 許可式は通る、禁止ASTは落ちる
* `test_pack_manifest.py`

  * manifestにhash/サイズ/ファイル数が入る
* `test_offline_outbox.py`

  * アップロード失敗を模擬→outbox保存→再送で消化（ClearMLはモック）
* `test_integration_local.py`

  * `dp run --dry-run` 相当をPython関数で呼び、成果物存在を確認

---

## 13. 開発環境セットアップ手順（例）

```bash
# 例：uv or poetry のどちらでもOK。ここではpip例。
python -m venv .venv
source .venv/bin/activate  # Windowsは .venv\Scripts\activate

pip install -U pip
pip install pandas openpyxl pydantic PyYAML typer pytest pillow pyarrow
# ClearMLはPhase 2で
pip install clearml
```

実行：

```bash
dp dev make-testdata --out out/testdata
dp run --config out/testdata/config.yaml --excel out/testdata/filled.xlsx --dry-run --out out/run1
pytest -q
```

---

## 14. 受け入れ基準（Acceptance Criteria）

### 作業者視点（最重要）

* [ ] テンプレExcelは開いてすぐ何をすべきか分かる（Instructionsが短い）
* [ ] 入力→ボタン1回で登録（もしくはdry-runでパッケージ生成）できる
* [ ] エラーは「どこを直すか」セル番地で分かる
* [ ] 成功すると受付番号/リンクがExcelに残る
* [ ] ネットワーク不調でも outbox に残り、次回自動で再送できる

### 管理者視点

* [ ] YAML変更でテンプレ再生成できる
* [ ] fの定義変更は再処理タスクで吸収できる（作業者の手間増なし）
* [ ] ClearMLへテンプレ配布（Artifacts）とデータ蓄積（Dataset versioning）ができる（Phase 2）

---

## 15. Codex / Copilot に渡す実装プロンプト例

> 以下を VSCode の Codex / Copilot Chat に貼り付けて順に作らせると進めやすい。

### Prompt 1（MVPスキャフォールド）

* このREADMEの構成に従い `src/datapipeline` のパッケージを作り、
* `dp` CLI（typer）を用意し、
* `generate-template`, `validate-excel`, `pack-dataset`, `run`, `dev make-testdata` の空実装（骨組み）を作ってください。
* 依存は最小（pydantic, PyYAML, openpyxl, pandas, typer, pytest）。

### Prompt 2（テンプレ生成）

* YAMLの `main_table.columns` を読み取り、InputシートにExcelテーブルを生成。
* enumはプルダウン。requiredは未入力でハイライト。
* _METAシートに template_id/version/config_hash を書く。

### Prompt 3（検証とStatus書き戻し）

* 入力Excelを読み取り、必須/enum/日時format/パス存在を検証し、
* Statusシートにエラー一覧（sheet!cell + message）を出す。

### Prompt 4（テストデータ生成）

* scripts/make_testdata.py を作り、
* CSVとfilled.xlsxを自動生成して `dp dev make-testdata` で呼べるように。

### Prompt 5（外部ファイル処理）

* file_profiles を実装し、x,y,z,t,f を抽出して long table を生成。
* derived式を安全に評価する。

### Prompt 6（ClearML連携はPhase 2で追加）

* `--dry-run` が通った後で、ClearML upload を別モジュールに実装。

---

## 16. 補足：実装上の重要ポリシー

* すべてのI/O（Excel読み書き、ファイルコピー、ClearML API）は “端” に寄せる
  → 中央ロジックはテストしやすい純粋変換にする
* 例外は必ず “作業者向けエラー” に変換して Statusに出す
  → Python例外トレースを作業者に見せない
* Windowsパス/UNCに強くする（pathlib、文字コード、長いパス）
* 途中成果物は out/ に必ず残し、失敗原因を調べられるようにする

---

以上。

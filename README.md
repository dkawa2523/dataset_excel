# clearml-dataset-excel

YAML で「条件Excel（スカラー列＋計測ファイルパス列）」のテンプレートと、参照される計測ファイルの `x/y/z/t/f` へのマッピング・簡易計算・統計量（平均/最大/積分など）を定義し、

- 条件Excelテンプレの生成
- 記入済み条件Excel + 参照ファイルを 1つの canonical テーブルに正規化して出力
- 生成物 + 参照元ファイルを **すべて** ClearML Dataset の新バージョンとしてアップロード

を行う CLI ツールです。

## Install

```bash
python3 -m pip install -e .
```

## Quickstart (examples)

サンプルの計測CSVと、テンプレ（`condition_template.xlsm`）+アドイン用 `.bas` + パス入力済みの条件Excelを生成:

```bash
python3 examples/generate_example_files.py
```

YAML spec を検証:

```bash
clearml-dataset-excel template validate --spec examples/run.yaml
```

条件Excelテンプレを生成:

```bash
clearml-dataset-excel template generate --spec examples/run.yaml --overwrite
```

既にマクロ込みの `*.xlsm`（`.bas` Import 済み）を持っている場合、`--base-excel` でそれをベースにテンプレを生成できます（マクロ/追加シートを保持）:

```bash
clearml-dataset-excel template generate --spec examples/run.yaml --base-excel <base.xlsm> --overwrite
```

`addin.enabled: true` の場合、テンプレ生成と同時に以下も同じフォルダへ出力します。
- YAML spec のコピー（`addin.spec_filename`）
- VBA モジュール（`.bas`, `addin.vba_module_filename`）

Excel にマクロを組み込む場合は、VBE（Visual Basic Editor）で `.bas` を Import して `ClearMLDatasetExcel_Run` を実行してください。

`addin.embed_vba: true` の場合、`template generate` / `run` / `register` 時にテンプレ `.xlsm` へマクロを自動で埋め込みます（デフォルトは同梱の `vbaProject.bin` をコピー注入する方式のため、Excel/COM/UI 自動化不要・Win/Mac 共通）。  
`addin.vba_template_excel` を指定している場合は、そのテンプレ `.xlsm` から `vbaProject.bin` をコピーして組み込みます（同様に非UI、`--overwrite` で上書き可能）。

`addin.windows_mode: addin` の場合、上記に加えて Windows 用に以下も出力します。
- マクロ無しテンプレ（既定: `template.template_filename` の拡張子を `.xlsx` にしたもの。`addin.windows_template_filename` で変更）
- Excel アドイン（`.xlam`, `addin.windows_addin_filename`）

テンプレ配布用にまとめて zip を作りたい場合:

```bash
clearml-dataset-excel template package --spec <spec.yaml> --output <template_package.zip> --overwrite
```

zip の中に `mac/`（`.xlsm` + spec + `.bas`）と `windows/`（`.xlsx` + `.xlam` + spec + `runner.exe`(任意)）を作って配置します。

Windows では `.xlsx` を配布し、実行は `.xlam`（アドイン）側の `ClearMLDatasetExcel_Run` で行う運用を推奨します（テンプレ自体にマクロを入れないため、警告ダイアログや社内ポリシーの影響を受けにくい）。

### Windows 運用フロー（推奨: addinモード）

#### 1) 管理者: テンプレ登録（ClearML Dataset 作成）

```bash
clearml-dataset-excel register --spec <run.yaml>
```

- YAML の `clearml.dataset_project / clearml.dataset_name / clearml.output_uri` を使って Dataset を作成し、`template/`（テンプレ一式）をアップロードします。
- `clearml.output_uri: file://...` の場合、ローカルディレクトリは自動作成します（作成権限が必要です）。

（任意）作業者配布用に zip を作る:

```bash
clearml-dataset-excel template package --spec <run.yaml> --output template_package.zip --overwrite
```

#### 2) 作業者: セットアップ（Windows / 初回 or 更新時）

zip を使う場合は展開して `windows/` フォルダを使います。以下が **同一フォルダ** にある状態にします:
- `condition_template.xlsx`（マクロ無しテンプレ）
- `clearml_dataset_excel_addin.xlam`（Excel アドイン）
- spec（例: `run.yaml`）
- （任意）`clearml_dataset_excel_runner.exe`（Python無しで実行したい場合）

アドインのインストール（推奨: AddInsフォルダへコピー）:

```bash
clearml-dataset-excel addin install --xlam <clearml_dataset_excel_addin.xlam> --overwrite
```

既にインストール済みのアドインを更新/削除したい場合:
- 更新: `clearml-dataset-excel addin update --xlam <clearml_dataset_excel_addin.xlam>`
- 削除: `clearml-dataset-excel addin uninstall --name clearml_dataset_excel_addin.xlam`
- 既定の AddIns フォルダ確認: `clearml-dataset-excel addin locate`

（UIでやる場合）Excel → `ファイル` → `オプション` → `アドイン` → `管理: Excel アドイン` → `設定` → `参照` で `*.xlam` を追加して有効化します。

#### 3) 作業者: データ登録（Windows / 毎回）

1. `condition_template.xlsx` を開いて `Conditions` シートを埋め、同じフォルダに保存します（例: `conditions_filled.xlsx`）。
   - ファイルパス列は **相対パス推奨**（Excelと同じフォルダ基準）。参照元ファイルも Dataset にアップロードされます。
2. Excel のリボンに `ClearML` タブが出ていれば `Run` をクリックします。
   - タブが出ない場合は `マクロの表示` から `ClearMLDatasetExcel_Run` を実行してください。
3. 実行ログは同じフォルダの `clearml_dataset_excel_addin.log` に出ます（失敗時はここを確認）。

（コマンドで実行したい場合）:

- runner.exe あり（推奨）: `clearml_dataset_excel_runner.exe run --spec run.yaml --excel conditions_filled.xlsx`
- Python/CLI あり: `clearml-dataset-excel run --spec run.yaml --excel conditions_filled.xlsx`

成功すると Dataset の新バージョンが作成され、`processed/` に以下が含まれます（YAML の集計/派生列でスカラー化されたテーブルを含む）:
- `processed/consolidated.xlsx`
- `processed/conditions.csv`
- `processed/canonical.csv`

#### 4) 確認者: 生成物（consolidated.xlsx）を取得

ClearML UI からは Dataset の `Files` → `processed/consolidated.xlsx` をダウンロードできます。

Python でローカルに落としてパスを表示する例:

```bash
python3 -c "from clearml import Dataset; from pathlib import Path; ds=Dataset.get(dataset_project='<project>', dataset_name='<name>', alias='latest'); root=Path(ds.get_local_copy()); print(root); print((root/'processed'/'consolidated.xlsx').resolve())"
```

Windows で Python 無し配布（runner.exe）にしたい場合（例）:
1. Windows で `powershell -ExecutionPolicy Bypass -File scripts/windows/build_runner.ps1` を実行し、`dist/clearml_dataset_excel_runner.exe` を作る
2. `*.xlsx` / `*.xlam` / spec（例: `run.yaml`）と同じフォルダに `clearml_dataset_excel_runner.exe` を置く
3. `addin.command_windows` は exe を優先するコマンド（`examples/run.yaml`）を推奨

Windows の `ClearMLDatasetExcel_Run` は `cmd.exe /c` で CLI を起動し、ログを同じフォルダの `clearml_dataset_excel_addin.log` に出力します（実行に失敗した場合はまずログを確認してください）。
macOS の `ClearMLDatasetExcel_Run` は `/bin/zsh -lc` で CLI を起動し、ログを同じフォルダの `clearml_dataset_excel_addin.log` に出力します（実行に失敗した場合はまずログを確認してください）。
テンプレには `addin_version` を埋め込み、`ClearMLDatasetExcel_Run` 実行時にテンプレ/アドインのバージョン不一致を警告します（古いテンプレや別バージョンの `.xlam` 混在事故の検出）。

セキュリティ/ブロックの典型例:
- macOS: ダウンロード扱い（quarantine）でマクロが無効化される場合があります。`clearml-dataset-excel addin unquarantine --excel <file.xlsm>`（または `xattr -d com.apple.quarantine <file.xlsm>`）を試してください。
- Windows: インターネット由来のファイルは Office 側でマクロがブロックされることがあります。ファイルの `プロパティ` で `ブロックの解除`、または `信頼できる場所 (Trusted Location)` への配置を検討してください。
- 状況確認: `clearml-dataset-excel addin inspect --excel <file.xlsm/.xlam> --json`

`.bas` を直接埋め込みたい場合は以下を使います（この場合のみExcelの自動化が必要）:
- Windows: Excel COM（「VBA プロジェクト オブジェクト モデルへの信頼アクセス」が必要）
- macOS: Microsoft Excel + AppleScript UI 自動化（System Settings -> Privacy & Security -> Automation / Accessibility が必要）

Mac でも、作業者が `.bas` を Import 済みの `*.xlsm` を `run` に渡すと、その `*.xlsm` をベースに「空のテンプレ（`condition_template.xlsm`）」を再生成してアップロードするため、以降Datasetからダウンロードするテンプレにはマクロが含まれます。  
`agent reprocess` も、既存Datasetのテンプレ（`payload.json.template_excel`）が存在する場合はそれを再利用してマクロを保持します。

手動で組み込みたい場合は以下も使えます（推奨: `--template-excel` による `vbaProject.bin` コピー）:

```bash
clearml-dataset-excel addin embed --excel <template.xlsm> --overwrite
clearml-dataset-excel addin embed --excel <template.xlsm> --template-excel <macro_template.xlsm> --overwrite
```

アドイン用メタ情報と、マクロが埋め込まれているか（`vbaProject.bin` / `ClearMLDatasetExcel_Run`）を確認する:

```bash
clearml-dataset-excel addin inspect --excel <template.xlsm> --json
```

### macOS Excelで「マクロ一覧が空」のとき
`addin inspect` で `has_vba_project: true` / `has_clearml_macro: true` なのに、Excel の「マクロの表示」で一覧が空になる場合があります。

- まず `clearml-dataset-excel addin embed --excel <template.xlsm>` を再実行してください（旧テンプレの `vbaProject.bin` を修復します）
  - 修復済みでも上書きしたい場合は `--overwrite` を付けてください
  - 旧同梱 `vbaProject.bin` は「マクロ本体が private module 扱い」になっており、Excel の一覧に出ないことがありました（`addin embed` が自動で修復します）
- 実行時に `プロシージャの呼び出し、または引数が無効です` 等が出る場合は、同じフォルダの `clearml_dataset_excel_addin.log` を確認してください（古いテンプレは `addin embed` でマクロ本体も更新されます）
- まず Excel を完全終了（`Cmd+Q`）してから、`*.xlsm` を開き直し、セキュリティバーが出ている場合は「コンテンツの有効化 / マクロを有効化」を押してください
- 「マクロの表示」の `Macros in:` が `このブック`（または対象の `*.xlsm`）になっていることを確認してください
- 一覧が空でも、`ClearMLDatasetExcel_Run` を入力欄に手入力して実行できる場合があります
 - それでも動かない場合、`xattr -d com.apple.quarantine <template.xlsm>` で quarantine を消してから再度開いてください

テンプレ（spec+template+.bas）だけを ClearML Dataset として配布する:

```bash
clearml-dataset-excel register --spec examples/run.yaml
```

Mac等でテンプレにマクロを含めたい場合は、`--base-excel` を併用してください（マクロ/追加シートを保持したテンプレをアップロード）:

```bash
clearml-dataset-excel register --spec examples/run.yaml --base-excel <base.xlsm>
```

`run/register/agent reprocess` で作成されるテンプレExcel（`condition_template.xlsm`）の `Info` シートには、Dataset task の `dataset_id` と `clearml_web_url` を自動で記入します（クリックでWeb UIへ遷移）。

処理のみ（アップロードしない）:

```bash
clearml-dataset-excel run --spec examples/run.yaml --excel examples/conditions_filled.xlsm --no-upload
```

`--no-upload` の場合も、アップロード対象のディレクトリ構造を `processed/_clearml_stage` に作成します（`--stage-dir` で変更、`--overwrite-stage` で上書き）。

ClearML Dataset へアップロード（YAML の `clearml.*` を使用。全参照ファイルも含めてアップロード）:

```bash
clearml-dataset-excel run --spec examples/run.yaml --excel examples/conditions_filled.xlsm
```

出力は `processed/conditions.csv`, `processed/canonical.csv`, `processed/consolidated.xlsx` です。

## ClearML Agent reprocess (Phase 3)
ClearML Task の Configuration（`dataset_format_spec`）に保存された spec を使い、既存 Dataset を再処理して新バージョンを作成します。

```bash
clearml-dataset-excel agent reprocess
```

- デバッグ用途で、処理だけ行いステージングを保持する場合は `--no-upload` を使います（`--output-root` / `--stage-dir` / `--overwrite-stage` で出力先を調整できます）。
- `--dataset-id` 省略時は、Hyperparameters の `clearml_dataset_excel/dataset_id` を優先し、なければ `Task.id` を Dataset id とみなします。
- clone/enqueue で spec を差し替えたい場合、Hyperparameters の `dataset_format_spec/yaml`（YAML文字列）を編集すると、`agent reprocess` はそれを優先して読み込みます。
- Dataset ルートの `payload.json` から条件Excel（`condition_excel`）を特定します。
- Excel 内のパスが他環境の絶対パスで存在しない場合、まず `payload.json` の `path_map`（rawパス→dataset内ファイル）で解決し、なければ `input/` 配下から `basename` / `*_basename`（external 退避名）で探索します。
- 出力Datasetの `project` はデフォルトで「現在のTaskのProject（=clone先）」を優先します。`name` はベースDatasetの `name` を使います。
  - 出力先がベースDatasetと同じ `project/name` の場合は、そのDatasetの新バージョンとして作成します。
  - 出力先が異なる `project/name` の場合は、その `project/name` のDatasetを作成/更新します（clone先プロジェクトへ分岐）。
- アップロードされる spec YAML は、実際に作成した Dataset の `project/name` を `clearml.dataset_project/dataset_name` に反映して保存します（ダウンロードしたテンプレのアドイン実行が clone 先へ自動的に紐づく）。
- clone/enqueue で出力先を明示したい場合（スクリプト引数なしで変更したい場合）は、Task の Hyperparameters に以下を設定できます（任意）:
  - `clearml_dataset_excel/output_dataset_project`
  - `clearml_dataset_excel/output_dataset_name`
  - `clearml_dataset_excel/output_uri`
  - `clearml_dataset_excel/output_tags`（`["tag1","tag2"]` 形式 or `tag1,tag2`）
  - `run/register` で作成した Dataset task にはデフォルト値が自動で入るため、clone した Task 側で編集するだけでOKです。

### clone/enqueue を “実行可能” にする（Task雛形）
Dataset task を clone/enqueue で実行するには、Task の Script（repo/entry_point）設定が必要です。  
YAML の `clearml.execution` を設定すると、`run/register/agent reprocess` が作成する Dataset task に Script を設定します。
このとき Dataset task の Python requirements は、ローカルの `requirements.txt` を読み取って Task に保存します（clone/enqueue 先の Agent は保存された requirements を使って依存関係をインストールします）。

例（このリポジトリを agent から clone できる前提）:

```yaml
clearml:
  execution:
    repository: https://github.com/your-org/clearml_dataset_excel.git
    branch: main
    working_dir: .
    entry_point: clearml_agent_reprocess.py
```

- 依存関係は repo 直下の `requirements.txt` を使う想定です（必要に応じて編集）。

実行手順（例）:
1. `clearml.execution` を設定した spec で `register`（テンプレ配布）または `run`（データ登録）を実行し、Dataset task を作成
2. ClearML Web UI で Dataset task を Clone して、目的の Project（clone先）へ移す
3. clone した Task の Hyperparameters で `dataset_format_spec/yaml` を編集（必要なら `clearml_dataset_excel/output_*` も設定）
4. Task を queue に Enqueue（Agent 側は `clearml-agent daemon --queue <queue_name>` 等で待機）
5. Agent が `clearml_agent_reprocess.py` を実行し、clone先プロジェクトに Dataset を作成/更新

## YAML spec (v1)
- `condition.columns`: 条件Excelの列定義（型/必須）
- `files[]`: 条件Excelの「ファイルパス列」ごとに、計測ファイルから `x/y/z/t` と `f` を選択・型を設定（`mapping.axes` は `"time"` のような列名か `{source,type}` を指定可）
- `mapping.derived`: `+ - * /` の簡易式（列名参照）で派生列を作成
- `mapping.aggregates`: `mean/max/min/sum/trapz`（積分は `wrt: t`）を条件行ごとに計算して条件テーブルへ追加
- `addin.*`: Excel VBA から `run` を実行するための設定
  - 実行OS切り替え: `addin.target_os` と `addin.command_{mac,windows}`
  - 自動組み込み（テンプレ `.xlsm` にマクロを埋め込む）: `addin.embed_vba`（必要なら `addin.vba_template_excel` で差し替え）
  - Windows アドイン配布: `addin.windows_mode: addin` + `addin.windows_template_filename` + `addin.windows_addin_filename`
  - 注意: `addin.embed_vba: true`（bundled）/ `addin.windows_mode: addin` は `template.meta_sheet: _meta` 前提です
- `output.combine_mode`: 同一条件行で複数計測ファイルをどう結合するか（`auto|merge|append`）

## payload.json (v1)
`run/register/agent reprocess` のステージング/アップロードに含まれるメタ情報です（互換管理のため `payload_version` を持ちます）。

- `payload_version`: 現在 `1`
- `created_at`: 生成時刻（UTC ISO）
- `spec_path`: `spec/*.yaml` の相対パス
- `template_excel`: `template/*.xlsm` の相対パス
- `template_excel_windows`: `template/*.xlsx` の相対パス（`addin.windows_mode=addin` の場合）
- `addin_xlam_windows`: `template/*.xlam` の相対パス（`addin.windows_mode=addin` の場合）
- `template_spec`: `template/` に置いた spec の相対パス（VBA が同フォルダから参照するため、`addin.enabled: true` の場合）
- `vba_module`: `template/*.bas` の相対パス（ある場合）
- `runner_exe_windows`: `template/clearml_dataset_excel_runner.exe` の相対パス（ある場合。Windows の Python 無し実行用）
- `condition_excel`: `input/` 配下の条件Excel相対パス（`register` のみの場合は無い）
- `path_map`: 条件Excelに書かれた生パス（raw）→Dataset内のファイル相対パス（例: `input/external/000_x.csv`）
- `conditions_csv` / `canonical_csv` / `consolidated_excel`: `processed/` 配下の相対パス（`register` のみの場合は無い）

デバッグ用:

```bash
clearml-dataset-excel payload show --root <dataset_root>
clearml-dataset-excel payload validate --root <dataset_root>
```

再現性の “深い” 検証（spec+condition_excel+path_map を使って一時ディレクトリで再処理）:

```bash
clearml-dataset-excel payload validate --root <dataset_root> --deep
```

## Notes / Constraints
- `run` は **全ファイルをアップロードする前提**のため、`http(s)://` 等の URL は未対応です（ローカルパスのみ）。
- 複数の計測ファイルを 1 行で結合する場合、挙動は `output.combine_mode` に依存します（`merge` なら軸セット不一致はエラー、`auto`/`append` なら縦に追加）。
- Excel の数式（アドイン/カスタム関数含む）は実行しません。Excel 側で計算済みの「キャッシュ値」を読み取ります（未計算だと空扱いになる場合があります）。

## ClearML Reports
`run/register/agent reprocess` で作成された Dataset task の Plots には以下を出力します。
- Coverage（欠損率テーブル）
- Summary（数値列の統計テーブルとヒストグラム、入力ファイル拡張子の内訳など）
- Summary（条件テーブルのファイルパス列の記入率: `file_path_column_coverage`）
- Debug Samples（画像と、CSV/TSV先頭のテーブルサンプル）

## Legacy (manifest uploader)
旧仕様の manifest（`path` 列）から Dataset を作る機能も残しています。
- `clearml-dataset-excel manifest ...`（またはサブコマンドなしで従来互換）

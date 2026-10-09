# ビット列のFFT解析

100 MHzでサンプリングしたリングオシレータのビット列を、
1,000,000 bit × 1000区間に分けて解析します。G/Mは10進です。
既定では先頭の1 Gbit（1,000,000,000 bit）だけを解析し、それ以降の0パディングは無視します。
1 bitを1サンプルとして扱います。既定のbin形式では1バイトに数値0または1を1個ずつ保存します。
1 Mbit分のサンプルは1,000,000バイト（1 MB）、1 Gbit分は1,000,000,000バイト（1 GB）です。
全体は10秒、各区間は10 ms、FFTの周波数刻みは100 Hzになります。

## 実行

Python 3.10以上で実行してください。以下はWSLのターミナル（bash）用です。作業フォルダは `~/FFT_template` を例にしています。実際の配置先に合わせて変更してください。`fft_bits.py`、`requirements.txt`、`bin_files` をこのフォルダに配置してください。Windows側の現在のフォルダを使う場合は、`cd ~/FFT_template` を `cd /mnt/c/path/to/FFT_template` に置き換えてください。

### 仮想環境の準備（初回のみ）

作業フォルダへ移動して、WSL用の仮想環境 `.venv-wsl` を作成し、必要なライブラリを入れます。Windows用の仮想環境は使い回せないため、別名で作成します。Ubuntu系のWSLでPythonやvenvが未導入の場合は、先に `sudo apt update` と `sudo apt install -y python3 python3-venv python3-pip` を実行してください。
すでにWSLで作成した `.venv-wsl` がある場合は作成コマンドを省略できます。

```bash
cd ~/FFT_template
python3 -m venv .venv-wsl
./.venv-wsl/bin/python -m pip install -r requirements.txt
```

### 仮想環境を有効化して実行

WSLのターミナルを新しく開いたときは、作業フォルダへ移動して有効化します。

```bash
cd ~/FFT_template
source .venv-wsl/bin/activate
```

`fft_bits.py` と同じ場所にある `bin_files` フォルダへ解析するbinファイルを置いてください。
binファイルが1つだけなら、ファイル名を省略して実行できます。

```bash
./.venv-wsl/bin/python fft_bits.py
```

複数ある場合は、以下の `bits.bin` を `bin_files` 内の解析するファイル名に置き換えてください。
名前のみの指定は `bin_files` から読み込みます（実行時の作業フォルダには依存しません）。
別の場所のファイルは、絶対パスまたはディレクトリを含む相対パスで指定できます。
`--input-format text` の入力も、名前のみなら `bin_files` に置いてください。

```bash
cd ~/FFT_template
source .venv-wsl/bin/activate
python fft_bits.py bits.bin
```

有効化せずに実行する場合は、作業フォルダで仮想環境のPythonを直接指定できます。

```bash
./.venv-wsl/bin/python fft_bits.py bits.bin
```

有効化した仮想環境を終了するには、次を実行します。

```bash
deactivate
```

既定の入力は、ヘッダーなしで1バイトに数値0または1を1個ずつ保存したbinファイルです。
例えばバイト列 `00 01 01 00` を4サンプルの `0, 1, 1, 0` として読みます。
文字の「0」「1」（`0x30`・`0x31`）とは異なります。
0と1以外のバイト（改行やヘッダーを含む）はエラーになります。
拡張子から保存形式を自動判定しないため、収録時の形式に合わせて指定してください。
指定したサンプル数（`--chunk-bits` × `--chunks`）より短い場合はエラーになります。
指定数を超えるデータは読み込み・解析の対象外です。入力ファイル自体は変更しません。

```bash
# 標準の形式を明示する場合（指定を省略しても同じ）
./.venv-wsl/bin/python fft_bits.py bits.bin --input-format bytes

# 1バイトに8bitを詰めた別形式の場合のみ使用（上位bitから時系列順）
./.venv-wsl/bin/python fft_bits.py bits.bin --input-format packed

# 従来の0/1文字列の場合（空白・改行を除去、ASCII/UTF-8・BOMなし）
./.venv-wsl/bin/python fft_bits.py bits.txt --input-format text
```

`packed` 形式で下位bitから収録されている場合は `--bit-order little` を追加します。
この形式では1 Mbitは125,000バイトです。既定では先頭125,000,000バイトが1 Gbitの解析対象です。

入力ファイルを物理的に1000個作る必要はありません。順に読み込むので、
10億サンプル全体をメモリに載せずに処理できます。

設定を変更する場合（仮想環境の有効化は不要）:

```bash
./.venv-wsl/bin/python fft_bits.py bits.bin --sample-rate 100000000 --chunk-bits 1000000 --chunks 1000 --output results_ro1
```

## 出力

入力ファイル名に合わせて結果を保存します。例えば `bits.bin` の場合:

- `fft_results/bits.xlsx`: `Summary`（入力ファイルのパスと統計）、`Peaks`（区間ごとの結果）、`Average spectrum`（平均PSD）の3シート。
- `fft_results/bits_summary.json`: 各区間の最大ピーク周波数の平均・母分散・標準偏差、およびピークPSDの統計。
- `fft_results/bits_peaks.csv`: 区間ごとのピーク周波数（Hz）、ピークPSD（bit²/Hz）、1の割合。
- `fft_results/bits_average_spectrum.csv`: 全区間の片側パワースペクトル密度（PSD）の算術平均。

既存の仮想環境でも、実行前に `./.venv-wsl/bin/python -m pip install -r requirements.txt` を実行し、Excel出力用ライブラリを追加してください。
Excelの行数上限を超える表は、連番のシートに分割して保存します。

周波数の分散の単位はHz²です。分散は取得した区間全体を対象とする母分散（ddof=0）です。
「区間ごとのピーク周波数の平均」と「平均スペクトルのピーク周波数」は異なる値です。
後者は `average_spectrum_peak_frequency_hz` に記録します。

各区間の平均を引いてからHann窓を適用し、FFTを行います。
窓が不要なら `--window none` を指定してください。
位相の相殺を避けるため、複素FFTそのものではなく、窓のエネルギーで正規化したPSDを平均します。
ピーク探索は0 Hzを除き、50 MHzを含みます。ピーク位置はFFTビン単位です。
全0・全1など非DC成分がゼロの区間はピークを欠損値として統計から除外します。
非ゼロのノイズだけでも最大値は選ばれます。ピークの有意性判定は行いません。

検出値は観測された最大スペクトルピークであり、RO基本波の折り返しと断定するものではありません。
高調波の折り返しやピークの切り替わりも含む可能性があります。
この結果から元の300〜400 MHzの発振周波数を一意には復元できません。
複数のROはファイル名を指定して1つずつ解析してください。異なる入力名なら同じ出力フォルダに結果を保存できます。
同じ入力名を同じ出力フォルダで再実行すると、その入力の解析結果を上書きします。
別の場所にある同名の入力ファイルは、`--output` で出力フォルダを分けてください。


## テンプレートの構成

- fft_bits.py: ビット列のFFT解析とExcel・CSV・JSON出力。
- test_fft_excel.py: 小さな人工信号を使った出力テスト。
- requirements.txt: 必要なPythonライブラリ。
- bin_files/: 入力ファイルの配置先。
- fft_results/: 解析結果の既定の保存先。
- fft_results_review/: 別条件の解析結果などを保存するための空フォルダ。

データ・解析結果・仮想環境・キャッシュは含みません。各フォルダの.gitkeepは、GitHubでも空フォルダの構成を維持するための空ファイルです。

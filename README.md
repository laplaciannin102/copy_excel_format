# copy-excel-format

- author: laplaciannin102(Kosuke Asada)
- date: 2021/01/12
- latest version: 0.1.9

---

## Table of Contents

- [copy-excel-format](#copy-excel-format)
  - [Table of Contents](#table-of-contents)
  - [How to install](#how-to-install)
  - [概要](#概要)
    - [何をするプログラム？](#何をするプログラム)
    - [注意事項](#注意事項)
  - [入出力](#入出力)
    - [入力(Input)](#入力input)
    - [出力(Output)](#出力output)
  - [使用手順](#使用手順)
  - [Example](#example)
  - [Repository](#repository)
    - [Github](#github)
    - [PyPI](#pypi)

---

## How to install

```shell
pip install copy_excel_format
```

---

## 概要

### 何をするプログラム？

- たくさんのテーブル(DataFrameを想定)をたくさんの書式付きexcelシートとして出力する.

### 注意事項

- xlwingsを使用して並列処理する場合はexcelのインストール(Office)が必要.
- テスト等きちんと出来ていないのでバグがある可能性あり.

---

## 入出力

### 入力(Input)

- 複数のpandas.DataFrame.
- 書式のテンプレートとして使用したいexcelシート.

### 出力(Output)

- **書式付き**, **テーブルの値が入力済み**のシートが複数あるexcelファイル.
  - 書式はテンプレートexcelファイルのもの.
  - テーブルの値はpandas.DataFrameのもの.

---

## 使用手順

- 以下, JupyterNotebookなどのPython環境での操作を想定しています.

- copy-excel-formatモジュールをimportします.

```python
import copy_excel_format as cef
```

- テンプレートとなるexcelファイルを準備します.
  - 下記を実行すると, sampleのexcelファイルおよびディレクトリ構成を取得, 設定できます.
  - とりあえず試してみたい時にご使用ください.

```python
cef.load_sample_files()
```

- pandas.DataFrameを複数準備します.

```python
import pandas as pd
df0 = pd.DataFrame(...) # 1つ目のpandas.DataFrame
df1 = pd.DataFrame(...) # 2つ目のpandas.DataFrame
```

- テンプレートexcelファイルのパス, pandas.DataFrameオブジェクト, そのDataFrameの値を入力したいシートのシート名を引数にCopyExcelInfoHolderオブジェクトを作成し, list変数に格納します.

```python
"""
df0: 1つ目のpandas.DataFrame
df1: 2つ目のpandas.DataFrame
input_template_excel_path: テンプレートexcelファイルのパス
sheet_name0: 1つ目のsheet_name
sheet_name1: 2つ目のsheet_name
"""

# CopyExcelInfoHolderオブジェクトの作成
ceih0 = cef.CopyExcelInfoHolder(
    template_excel_path = input_template_excel_path,
    template_sheet_name = 'blank_template',
    output_sheet_name = sheet_name0,
    df = df0
)

ceih1 = cef.CopyExcelInfoHolder(
    template_excel_path = input_template_excel_path,
    template_sheet_name = 'blank_template',
    output_sheet_name = sheet_name1,
    df = df1
)

ceih_list = [ceih0, ceih1] # list変数
```

- copy_excel_format関数を実行します.
   - 下記は直列処理での実行例.
   - output_excel_pathには出力excelファイルのパスを与えます.

```python
# excel書式コピーを直列で実行
cef.copy_excel_format(
    ceih_list = ceih_list,
    output_excel_path = output_path + 'output_excel_sample.xlsx',
    cef_manual_set_rows = None,
    cef_force_dimension_copy = False,
    cef_debug_mode = True,
    write_index = False,
    write_header = False,
    copy_values = False
)
```

[end]

---

## Example

- sample ipynb: [Githubリンク](https://github.com/laplaciannin102/copy_excel_format/blob/master/examples/copy_excel_format_sample.ipynb)

- sample program:

```python

# --------------------------------------------------------------------------------
# Load modules
# --------------------------------------------------------------------------------
## copy-excel-format module
import copy_excel_format as cef

## other modules
import gc
import numpy as np
import pandas as pd
import random
import time
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta


# --------------------------------------------------------------------------------
# Configure
# --------------------------------------------------------------------------------
# random seed
np.random.seed(57)
random.seed(57)


# --------------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------------
# paths
## directory paths
input_path = './input/'
output_path = './output/'
interm_path = './intermediate/'

## file paths
input_template_excel_path = input_path + 'input_template_excel_sample.xlsx'
input_header_csv_path = input_path + 'input_header_df_sample.csv'


# --------------------------------------------------------------------------------
# Load sample files
# --------------------------------------------------------------------------------
cef.load_sample_files()


# --------------------------------------------------------------------------------
# Functions
# --------------------------------------------------------------------------------
# header dataframe
header_df = pd.read_csv(input_header_csv_path)

# get_sample_df
def get_sample_df(n_rows=10, header_df=header_df):
    """
    sampleデータを作成する関数.
    
    Args:
        n_rows: int, optional(default=10)
            データ部分のDataFrameの行数.
        
        header_df: pandas.DataFrame
            ヘッダー部分のDataFrame
    """
    col1_samples = ['hoge', 'fuga', 'poyo', 'gray', None]
    
    sample_df = pd.DataFrame()
    sample_df['No.'] = range(n_rows)
    sample_df['No.'] = sample_df['No.'] + 1
    sample_df['date'] = [datetime(2020, 11, 20) + relativedelta(days=jj*7) for jj in range(n_rows)]
    sample_df['col1'] = random.choices(col1_samples, k=n_rows)
    sample_df['col2'] = np.random.randint(0, 10, size=n_rows)
    sample_df['col3'] = np.random.randint(100, 200, size=n_rows)
    sample_df['col4'] = random.choices(col1_samples, k=n_rows)
    sample_df['col5'] = np.random.randint(0, 10, size=n_rows)
    sample_df['col6'] = np.random.randint(100, 200, size=n_rows)
    
    # headerをつける
    tmp_name = random.choice(['hoge', 'fuga', 'poyo'])
    tmp_header_df = header_df.copy()
    tmp_header_df = tmp_header_df.replace('name: <name>', 'name: ' + tmp_name)
    
    sample_df = tmp_header_df.append(sample_df)
    
    return sample_df


# --------------------------------------------------------------------------------
# excel書式コピー準備
# --------------------------------------------------------------------------------
# テンプレートのexcelパスとシート名とDataFrameをセット
# DataFrameの数. シート数も同じ数.
n_df = 10

# CopyExcelInfoHolderオブジェクトインスタンスのリスト
ceih_list = []

# ceih_listというCopyExcelInfoHolderオブジェクトインスタンスのリストを作成しておく
for ii in range(n_df):
    
    tmp_sheet_name = 'sheet' + str(ii+1).zfill(3)
    tmp_df = get_sample_df(
        n_rows = np.random.randint(10, 28)
    )

    ceih = cef.CopyExcelInfoHolder(
        template_excel_path = input_template_excel_path,
        template_sheet_name = 'blank_template',
        output_sheet_name = tmp_sheet_name,
        df = tmp_df
    )
    
    ceih_list += [ceih]
    
    del ceih
    gc.collect()


# --------------------------------------------------------------------------------
# Execute
# --------------------------------------------------------------------------------
# excel書式コピーを直列で実行
cef.copy_excel_format(
    ceih_list = ceih_list,
    output_excel_path = output_path + 'output_excel_sample.xlsx',
    cef_manual_set_rows = None,
    cef_force_dimension_copy = False,
    cef_debug_mode = True,
    write_index = False,
    write_header = False,
    copy_values = False
)

# excel書式コピーを並列で実行1(1つの関数で実行)
cef.copy_excel_format_parallel(
    ceih_list = ceih_list,
    output_excel_path = output_path + 'output_excel_sample_parallel001.xlsx',
    tmp_output_excel_dir_path = interm_path + 'tmp_output_excel/',
    parallel_method = 'multiprocess',
    n_jobs = None,
    copy_sheet_method = 'xlwings',
    sorted_sheet_names_list = None,
    del_tmp_dir = True,
    n_seconds_to_sleep = 1,
    cef_manual_set_rows = None,
    cef_force_dimension_copy = False,
    cef_debug_mode = True,
    write_index = False,
    write_header = False,
    copy_values = False
)

# excel書式コピーを並列で実行2(2つの関数に分けて実行)
# 並列処理を行い, 一時的な書式設定済みのexcelファイルを出力する.
cef.output_temporary_excel_parallel(
    ceih_list = ceih_list,
    tmp_output_excel_dir_path = interm_path + 'tmp_output_excel/',
    parallel_method = 'multiprocess',
    n_jobs = None,
    cef_manual_set_rows = None,
    cef_force_dimension_copy = False,
    cef_debug_mode = True,
    write_index = False,
    write_header = False,
    copy_values = False
)

# 一時的に出力した複数のexcelファイルをまとめて複数シートを持つ1つのexcelファイルとする.
cef.copy_excel_format_from_temporary_files(
    ceih_list = ceih_list,
    output_excel_path = output_path + 'output_excel_sample_parallel002.xlsx',
    tmp_output_excel_dir_path = interm_path + 'tmp_output_excel/',
    copy_sheet_method = 'xlwings',
    sorted_sheet_names_list = None,
    del_tmp_dir = True,
    n_seconds_to_sleep = 1
)

```

---

## Repository

### Github

- [https://github.com/laplaciannin102/copy_excel_format](https://github.com/laplaciannin102/copy_excel_format)

### PyPI

- [https://pypi.org/project/copy_excel_format/](https://pypi.org/project/copy_excel_format/)


# copy-excel-format

- author: laplaciannin102(Kosuke Asada)
- date: 2020/12/23

---

## Table of Contents

- [copy-excel-format](#copy-excel-format)
  - [Table of Contents](#table-of-contents)
  - [How to install](#how-to-install)
  - [概要](#概要)
    - [何をするプログラム？](#何をするプログラム)
    - [注意事項](#注意事項)
  - [入出力](#入出力)
    - [入力](#入力)
    - [出力](#出力)
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

- たくさんのテーブル(DataFrameを想定)をたくさんの書式付きexcelシートとして出力.

### 注意事項

- xlwingsを使用して並列処理する場合はexcelのインストール(Office)が必要.

---

## 入出力

### 入力

- 複数のpandas.DataFrame
- 書式のテンプレートとして使用したいexcelシート

### 出力

- 書式付きでテーブルの値が入力済みシートが複数あるexcelファイル

---

## Example

- sample ipynb:
  - [https://github.com/laplaciannin102/copy_excel_format/blob/master/examples/src/copy_excel_format_sample.ipynb](https://github.com/laplaciannin102/copy_excel_format/blob/master/examples/src/copy_excel_format_sample.ipynb)

```python

# Load modules
import sys, os
import gc
import copy
import numpy as np
import pandas as pd

import random

import time
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

# copy_excel_format
from copy_excel_format import *

# random seed
np.random.seed(57)
random.seed(57)

# file paths
input_path = '../input/'
output_path = '../output/'
interm_path = '../intermediate/'

input_template_excel_path = input_path + 'input_template_excel_sample.xlsx'
input_header_csv_path = input_path + 'input_header_df_sample.csv'

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

    ceih = CopyExcelInfoHolder(
        template_excel_path = input_template_excel_path,
        template_sheet_name = 'blank_template',
        output_sheet_name = tmp_sheet_name,
        df = tmp_df
    )
    
    ceih_list += [ceih]
    
    del ceih
    gc.collect()

# Execute
# excel書式コピーを直列で実行
copy_excel_format(
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
copy_excel_format_parallel(
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
output_temporary_excel_parallel(
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
copy_excel_format_from_temporary_files(
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


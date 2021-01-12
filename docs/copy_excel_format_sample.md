# Readme

- author: laplaciannin102(Kosuke Asada)
- date: 2021/01/12
- latest version: 0.1.9

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

- [README.md](https://github.com/laplaciannin102/copy_excel_format/blob/master/README.md)参照.

---

## 参考

- [openpyxlライブラリ](https://pypi.org/project/openpyxl/)
- [xlwingsドキュメント](https://docs.xlwings.org/ja/latest/#)

- [PythonでExcelシートを別のワークブックにコピーする方法](https://www.it-swarm-ja.tech/ja/python/python%E3%81%A7excel%E3%82%B7%E3%83%BC%E3%83%88%E3%82%92%E5%88%A5%E3%81%AE%E3%83%AF%E3%83%BC%E3%82%AF%E3%83%96%E3%83%83%E3%82%AF%E3%81%AB%E3%82%B3%E3%83%94%E3%83%BC%E3%81%99%E3%82%8B%E6%96%B9%E6%B3%95/831860805/)

# Load modules

## copy-excel-format module


```python
# copy_excel_format
import copy_excel_format as cef
```

## other modules


```python
import os
import gc
import numpy as np
import pandas as pd
import random
import time
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
```


```python
gc.collect()
```




    40




```python

```

# Configures

## random seed


```python
np.random.seed(57)
random.seed(57)
```


```python

```

# Constants

## paths

## directory paths


```python
input_path = './input/'
output_path = './output/'
interm_path = './intermediate/'
```

## file paths


```python
input_template_excel_path = input_path + 'input_template_excel_sample.xlsx'
input_header_csv_path = input_path + 'input_header_df_sample.csv'
```


```python

```

# Load sample files


```python
cef.load_sample_files()
```

    ********************************************************************************
    make directory
    path: ./input/
    
    make directory
    path: ./output/
    
    make directory
    path: ./intermediate/
    
    ********************************************************************************
    
    
    
    ********************************************************************************
    copy excel file
    path: ./input/input_template_excel_sample.xlsx
    ********************************************************************************
    copy csv file
    path: ./input/input_header_df_sample.csv
    ********************************************************************************
    

# Functions

## get_sample_df


```python
header_df = pd.read_csv(input_header_csv_path)
header_df.head()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>No.</th>
      <th>date</th>
      <th>col1</th>
      <th>col2</th>
      <th>col3</th>
      <th>col4</th>
      <th>col5</th>
      <th>col6</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>1</th>
      <td>NaN</td>
      <td>name: &lt;name&gt;</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>2</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>3</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>4</th>
      <td>No.</td>
      <td>date</td>
      <td>col1</td>
      <td>col2</td>
      <td>col3_4_5</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>col6</td>
    </tr>
  </tbody>
</table>
</div>




```python
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
```

### get_sample_df example


```python
sample_df = get_sample_df()
sample_df.head(3)
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>No.</th>
      <th>date</th>
      <th>col1</th>
      <th>col2</th>
      <th>col3</th>
      <th>col4</th>
      <th>col5</th>
      <th>col6</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>1</th>
      <td>NaN</td>
      <td>name: poyo</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>2</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
  </tbody>
</table>
</div>




```python
sample_df.tail()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>No.</th>
      <th>date</th>
      <th>col1</th>
      <th>col2</th>
      <th>col3</th>
      <th>col4</th>
      <th>col5</th>
      <th>col6</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>5</th>
      <td>6</td>
      <td>2020-12-25 00:00:00</td>
      <td>fuga</td>
      <td>8</td>
      <td>158</td>
      <td>gray</td>
      <td>2</td>
      <td>133</td>
    </tr>
    <tr>
      <th>6</th>
      <td>7</td>
      <td>2021-01-01 00:00:00</td>
      <td>None</td>
      <td>8</td>
      <td>101</td>
      <td>gray</td>
      <td>8</td>
      <td>199</td>
    </tr>
    <tr>
      <th>7</th>
      <td>8</td>
      <td>2021-01-08 00:00:00</td>
      <td>gray</td>
      <td>1</td>
      <td>198</td>
      <td>gray</td>
      <td>5</td>
      <td>137</td>
    </tr>
    <tr>
      <th>8</th>
      <td>9</td>
      <td>2021-01-15 00:00:00</td>
      <td>None</td>
      <td>7</td>
      <td>148</td>
      <td>poyo</td>
      <td>2</td>
      <td>130</td>
    </tr>
    <tr>
      <th>9</th>
      <td>10</td>
      <td>2021-01-22 00:00:00</td>
      <td>hoge</td>
      <td>9</td>
      <td>101</td>
      <td>None</td>
      <td>0</td>
      <td>123</td>
    </tr>
  </tbody>
</table>
</div>




```python

```

# excel書式コピー準備

## テンプレートのexcelパスとシート名とDataFrameをセット


```python
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
```


```python
# ceih_listというCopyExcelInfoHolderオブジェクトインスタンスのリストを作成しておく
ceih_list[:3]
```




    [<copy_excel_format.excel_module.CopyExcelInfoHolder at 0x222467a6fd0>,
     <copy_excel_format.excel_module.CopyExcelInfoHolder at 0x222466fdf40>,
     <copy_excel_format.excel_module.CopyExcelInfoHolder at 0x22235bf0ca0>]




```python
# ceih_listの中身を表示
[print('template_excel_path:{}'.format(ii.template_excel_path)) for ii in ceih_list[:3]]
```

    template_excel_path:./input/input_template_excel_sample.xlsx
    template_excel_path:./input/input_template_excel_sample.xlsx
    template_excel_path:./input/input_template_excel_sample.xlsx
    




    [None, None, None]




```python
# ceih_listの中身を表示
[print('template_sheet_name:{}'.format(ii.template_sheet_name)) for ii in ceih_list[:3]]
```

    template_sheet_name:blank_template
    template_sheet_name:blank_template
    template_sheet_name:blank_template
    




    [None, None, None]




```python
# ceih_listの中身を表示
[print('output_sheet_name:{}'.format(ii.output_sheet_name)) for ii in ceih_list[:3]]
```

    output_sheet_name:sheet001
    output_sheet_name:sheet002
    output_sheet_name:sheet003
    




    [None, None, None]




```python

```


```python
# ceih_listの中身を表示
[print('*' * 80 + '\ndf.head(3):{}'.format(ii.df.head(3)) + '\n' + '*' * 80 + '\n\n') for ii in ceih_list[:3]]
```

    ********************************************************************************
    df.head(3):   No.        date col1 col2 col3 col4 col5 col6
    0  NaN         NaN  NaN  NaN  NaN  NaN  NaN  NaN
    1  NaN  name: poyo  NaN  NaN  NaN  NaN  NaN  NaN
    2  NaN         NaN  NaN  NaN  NaN  NaN  NaN  NaN
    ********************************************************************************
    
    
    ********************************************************************************
    df.head(3):   No.        date col1 col2 col3 col4 col5 col6
    0  NaN         NaN  NaN  NaN  NaN  NaN  NaN  NaN
    1  NaN  name: fuga  NaN  NaN  NaN  NaN  NaN  NaN
    2  NaN         NaN  NaN  NaN  NaN  NaN  NaN  NaN
    ********************************************************************************
    
    
    ********************************************************************************
    df.head(3):   No.        date col1 col2 col3 col4 col5 col6
    0  NaN         NaN  NaN  NaN  NaN  NaN  NaN  NaN
    1  NaN  name: poyo  NaN  NaN  NaN  NaN  NaN  NaN
    2  NaN         NaN  NaN  NaN  NaN  NaN  NaN  NaN
    ********************************************************************************
    
    
    




    [None, None, None]




```python
# ceih_listの中身を表示
[print('*' * 80 + '\ndf.tail(3):{}'.format(ii.df.tail(3)) + '\n' + '*' * 80 + '\n\n') for ii in ceih_list[:3]]
```

    ********************************************************************************
    df.tail(3):   No.                 date  col1 col2 col3  col4 col5 col6
    12  13  2021-02-12 00:00:00  None    6  147  fuga    6  157
    13  14  2021-02-19 00:00:00  fuga    6  126  fuga    2  155
    14  15  2021-02-26 00:00:00  gray    5  109  gray    9  143
    ********************************************************************************
    
    
    ********************************************************************************
    df.tail(3):   No.                 date  col1 col2 col3  col4 col5 col6
    18  19  2021-03-26 00:00:00  None    4  115  hoge    6  192
    19  20  2021-04-02 00:00:00  hoge    0  134  None    2  134
    20  21  2021-04-09 00:00:00  poyo    1  127  gray    0  194
    ********************************************************************************
    
    
    ********************************************************************************
    df.tail(3):   No.                 date  col1 col2 col3  col4 col5 col6
    11  12  2021-02-05 00:00:00  gray    8  187  hoge    0  190
    12  13  2021-02-12 00:00:00  fuga    0  100  gray    8  161
    13  14  2021-02-19 00:00:00  fuga    0  159  None    1  198
    ********************************************************************************
    
    
    




    [None, None, None]




```python
print(len(ceih_list))
```

    10
    


```python

```

# excel書式コピーを直列で実行

## 出力ファイル名定義


```python
output_excel_path = output_path + 'output_excel_sample.xlsx'
output_excel_path
```




    './output/output_excel_sample.xlsx'



## 実行


```python
start = time.time()
```


```python
# copy_excel_format関数を実行
cef.copy_excel_format(
    ceih_list = ceih_list,
    output_excel_path = output_excel_path,
    cef_manual_set_rows = None,
    cef_force_dimension_copy = False,
    cef_debug_mode = True,
    write_index = False,
    write_header = False,
    copy_values = False
)
```

    ********************************************************************************
    sheet name: sheet001
    
    to write df to sheet end.
    elapsed time: 0.0 s
    
    to copy cell format end.
    elapsed time: 7.5 s
    
    to copy format end.
    elapsed time: 7.5 s
    
    ********************************************************************************
    
    ********************************************************************************
    template_excel_path:./input/input_template_excel_sample.xlsx
    template_sheet_name:blank_template
    output_sheet_name:sheet001
    ********************************************************************************
    ********************************************************************************
    sheet name: sheet002
    
    to write df to sheet end.
    elapsed time: 0.0 s
    
    to copy cell format end.
    elapsed time: 7.6 s
    
    to copy format end.
    elapsed time: 7.6 s
    
    ********************************************************************************
    
    ********************************************************************************
    template_excel_path:./input/input_template_excel_sample.xlsx
    template_sheet_name:blank_template
    output_sheet_name:sheet002
    ********************************************************************************
    ********************************************************************************
    sheet name: sheet003
    
    to write df to sheet end.
    elapsed time: 0.0 s
    
    to copy cell format end.
    elapsed time: 7.4 s
    
    to copy format end.
    elapsed time: 7.4 s
    
    ********************************************************************************
    
    ********************************************************************************
    template_excel_path:./input/input_template_excel_sample.xlsx
    template_sheet_name:blank_template
    output_sheet_name:sheet003
    ********************************************************************************
    ********************************************************************************
    sheet name: sheet004
    
    to write df to sheet end.
    elapsed time: 0.0 s
    
    to copy cell format end.
    elapsed time: 7.4 s
    
    to copy format end.
    elapsed time: 7.4 s
    
    ********************************************************************************
    
    ********************************************************************************
    template_excel_path:./input/input_template_excel_sample.xlsx
    template_sheet_name:blank_template
    output_sheet_name:sheet004
    ********************************************************************************
    ********************************************************************************
    sheet name: sheet005
    
    to write df to sheet end.
    elapsed time: 0.0 s
    
    to copy cell format end.
    elapsed time: 7.2 s
    
    to copy format end.
    elapsed time: 7.2 s
    
    ********************************************************************************
    
    ********************************************************************************
    template_excel_path:./input/input_template_excel_sample.xlsx
    template_sheet_name:blank_template
    output_sheet_name:sheet005
    ********************************************************************************
    ********************************************************************************
    sheet name: sheet006
    
    to write df to sheet end.
    elapsed time: 0.0 s
    
    to copy cell format end.
    elapsed time: 7.4 s
    
    to copy format end.
    elapsed time: 7.4 s
    
    ********************************************************************************
    
    ********************************************************************************
    template_excel_path:./input/input_template_excel_sample.xlsx
    template_sheet_name:blank_template
    output_sheet_name:sheet006
    ********************************************************************************
    ********************************************************************************
    sheet name: sheet007
    
    to write df to sheet end.
    elapsed time: 0.0 s
    
    to copy cell format end.
    elapsed time: 7.6 s
    
    to copy format end.
    elapsed time: 7.6 s
    
    ********************************************************************************
    
    ********************************************************************************
    template_excel_path:./input/input_template_excel_sample.xlsx
    template_sheet_name:blank_template
    output_sheet_name:sheet007
    ********************************************************************************
    ********************************************************************************
    sheet name: sheet008
    
    to write df to sheet end.
    elapsed time: 0.0 s
    
    to copy cell format end.
    elapsed time: 7.6 s
    
    to copy format end.
    elapsed time: 7.6 s
    
    ********************************************************************************
    
    ********************************************************************************
    template_excel_path:./input/input_template_excel_sample.xlsx
    template_sheet_name:blank_template
    output_sheet_name:sheet008
    ********************************************************************************
    ********************************************************************************
    sheet name: sheet009
    
    to write df to sheet end.
    elapsed time: 0.0 s
    
    to copy cell format end.
    elapsed time: 7.6 s
    
    to copy format end.
    elapsed time: 7.6 s
    
    ********************************************************************************
    
    ********************************************************************************
    template_excel_path:./input/input_template_excel_sample.xlsx
    template_sheet_name:blank_template
    output_sheet_name:sheet009
    ********************************************************************************
    ********************************************************************************
    sheet name: sheet010
    
    to write df to sheet end.
    elapsed time: 0.0 s
    
    to copy cell format end.
    elapsed time: 7.4 s
    
    to copy format end.
    elapsed time: 7.5 s
    
    ********************************************************************************
    
    ********************************************************************************
    template_excel_path:./input/input_template_excel_sample.xlsx
    template_sheet_name:blank_template
    output_sheet_name:sheet010
    ********************************************************************************
    

## 処理時間確認


```python
cef.get_elapsed_time(start)
```

    elapsed time: 75.3 s
    




    75.29727053642273


elapsed time: 75.3 s
75.29727053642273

```python

```

# excel書式コピーを並列で実行1(1つの関数で実行)

## cpu数確認


```python
print('cpu_count:{}'.format(str(os.cpu_count())))
```

    cpu_count:8
    

## 出力ファイル名定義


```python
output_excel_path = output_path + 'output_excel_sample_parallel001.xlsx'
output_excel_path
```




    './output/output_excel_sample_parallel001.xlsx'



## 一時的出力ディレクトリ名定義


```python
tmp_output_excel_dir_path = interm_path + 'tmp_output_excel/'
tmp_output_excel_dir_path
```




    './intermediate/tmp_output_excel/'



## 実行


```python
start = time.time()
```


```python
# copy_excel_format関数の並列版を実行
cef.copy_excel_format_parallel(
    ceih_list = ceih_list,
    output_excel_path = output_excel_path,
    tmp_output_excel_dir_path = tmp_output_excel_dir_path,
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
```

## 処理時間確認


```python
cef.get_elapsed_time(start)
```

    elapsed time: 80.8 s
    




    80.78811526298523


elapsed time: 80.8 s
80.78811526298523

```python

```


```python

```

# excel書式コピーを並列で実行2(2つの関数に分けて実行)

## 出力ファイル名定義


```python
output_excel_path = output_path + 'output_excel_sample_parallel002.xlsx'
output_excel_path
```




    './output/output_excel_sample_parallel002.xlsx'



## 一時的出力ディレクトリ名定義


```python
tmp_output_excel_dir_path = interm_path + 'tmp_output_excel/'
tmp_output_excel_dir_path
```




    './intermediate/tmp_output_excel/'



## 実行


```python
start = time.time()
```


```python
# 並列処理を行い, 一時的な書式設定済みのexcelファイルを出力する.
cef.output_temporary_excel_parallel(
    ceih_list = ceih_list,
    tmp_output_excel_dir_path = tmp_output_excel_dir_path,
    parallel_method = 'multiprocess',
    n_jobs = None,
    cef_manual_set_rows = None,
    cef_force_dimension_copy = False,
    cef_debug_mode = True,
    write_index = False,
    write_header = False,
    copy_values = False
)
```


```python
# 一時的に出力した複数のexcelファイルをまとめて複数シートを持つ1つのexcelファイルとする.
cef.copy_excel_format_from_temporary_files(
    ceih_list = ceih_list,
    output_excel_path = output_excel_path,
    tmp_output_excel_dir_path = tmp_output_excel_dir_path,
    copy_sheet_method = 'xlwings',
    sorted_sheet_names_list = None,
    del_tmp_dir = True,
    n_seconds_to_sleep = 1
)
```

## 処理時間確認


```python
cef.get_elapsed_time(start)
```

    elapsed time: 88.5 s
    




    88.5116548538208


elapsed time: 88.5 s
88.5116548538208

```python

```

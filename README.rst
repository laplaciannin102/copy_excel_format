copy-excel-format
=================

-  author: laplaciannin102(Kosuke Asada)
-  date: 2020/12/23

--------------

Table of Contents
-----------------

-  `copy-excel-format`_

   -  `Table of Contents`_
   -  `How to install`_
   -  `概要`_

      -  `何をするプログラム？`_
      -  `注意事項`_

   -  `入出力`_

      -  `入力`_
      -  `出力`_

   -  `Example`_
   -  `Repository`_

      -  `Github`_
      -  `PyPI`_

--------------

How to install
--------------

.. code:: shell

   pip install copy_excel_format

--------------

概要
----

何をするプログラム？
~~~~~~~~~~~~~~~~~~~~

-  たくさんのテーブル(DataFrameを想定)をたくさんの書式付きexcelシートとして出力.

注意事項
~~~~~~~~

-  xlwingsを使用して並列処理する場合はexcelのインストール(Office)が必要.

--------------

入出力
------

入力
~~~~

-  複数のpandas.DataFrame
-  書式のテンプレートとして使用したいexcelシート

出力
~~~~

-  書式付きでテーブルの値が入力済みシートが複数あるexcelファイル

--------------

Example
-------

-  sample ipynb:

   -  https://github.com/laplaciannin102/copy_excel_format/blob/master/examples/src/copy_excel_format_sample.ipynb

\```python # Load modules import sys, os import gc import copy import
numpy as np import pandas as pd

import random

import time from datetime import datetime, timedelta from
dateutil.relativedelta import relativedelta

copy_excel_format
=================

from copy_excel_format import \*

random seed
===========

np.random.seed(57) random.seed(57)

file paths
==========

input_path = ‘../input/’ output_path = ‘../output/’ interm_path =
‘../intermediate/’

input_template_excel_path = input_path +
‘input_template_excel_sample.xlsx’ input_header_csv_path = input_path +
‘input_header_df_sample.csv’

header dataframe
================

header_df = pd.read_csv(input_header_csv_path)

get_sample_df
=============

def get_sample_df(n_rows=10, header_df=header_df): """
sampleデータを作成する関数.

::

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

テンプレートのexcelパスとシート名とDataFrameをセット
====================================================

DataFrameの数. シート数も同じ数.
================================

n_df = 10

CopyExcelInfoHolderオブジェクトインスタンスのリスト
===================================================

ceih_list = []

ceih_listというCopyE
====================

.. _copy-excel-format: #copy-excel-format
.. _Table of Contents: #table-of-contents
.. _How to install: #how-to-install
.. _概要: #概要
.. _何をするプログラム？: #何をするプログラム
.. _注意事項: #注意事項
.. _入出力: #入出力
.. _入力: #入力
.. _出力: #出力
.. _Example: #example
.. _Repository: #repository
.. _Github: #github
.. _PyPI: #pypi
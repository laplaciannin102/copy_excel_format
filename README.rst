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

Repository
----------

Github
~~~~~~

-  https://github.com/laplaciannin102/sample_annn_pkg

PyPI
~~~~

-  https://test.pypi.org/project/sample_annn_pkg/

.. _copy-excel-format: #copy-excel-format
.. _Table of Contents: #table-of-contents
.. _How to install: #how-to-install
.. _概要: #概要
.. _何をするプログラム？: #何をするプログラム
.. _注意事項: #注意事項
.. _入出力: #入出力
.. _入力: #入力
.. _出力: #出力
.. _Repository: #repository
.. _Github: #github
.. _PyPI: #pypi
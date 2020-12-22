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

## Repository

### Github

- [https://github.com/laplaciannin102/sample_annn_pkg](https://github.com/laplaciannin102/sample_annn_pkg)

### PyPI

- [https://test.pypi.org/project/sample_annn_pkg/](https://test.pypi.org/project/sample_annn_pkg/)


#!/usr/bin/env python3
# -*- coding: utf-8 -*-


"""
@author: Kosuke Asada
@date: 2020/12/23
@version: 0.1.9

sample data setを読み込む.

@history:
    2021/01/12:
        初期版作成.
"""



# --------------------------------------------------------------------------------
# Load modules
# --------------------------------------------------------------------------------

import os
import shutil



# --------------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------------

# module path
module_path = os.path.dirname(__file__)

# directory paths
input_path = './input/'
output_path = './output/'
interm_path = './intermediate/'

# file names
input_template_excel_fname = 'input_template_excel_sample.xlsx'
input_header_csv_fname = 'input_header_df_sample.csv'



# --------------------------------------------------------------------------------
# Functions
# --------------------------------------------------------------------------------


def pro_makedirs(dir_path):
    """
    ディレクトリを作成する.指定のディレクトリが存在しない場合のみ作成する.
    深い階層のディレクトリを指定した場合、途中階層のディレクトリも全て作成する.
    
    Arg:
        dir_path: str
            ディレクトリのパス.
    """
    dir_path = str(dir_path)
    if not os.path.isdir(dir_path):
        os.makedirs(dir_path)
    else:
        pass


def load_sample_files():
    """
    sampleデータをロードする.
    必要なディレクトリを作成する.
    """

    print('*' * 80)

    # 必要なディレクトリを作成する.
    for dir_path in [input_path, output_path, interm_path]:
        pro_makedirs(dir_path)

        print('make directory')
        print('path: {}'.format(dir_path))
        print()
    
    print('*' * 80)
    print()
    print()
    print()
    print('*' * 80)

    # 必要なファイルをコピーする.
    # excelのコピー
    from_path0 = '{}/sample_data/{}'.format(module_path, input_template_excel_fname)
    to_path0 = '{}{}'.format(input_path, input_template_excel_fname)
    shutil.copyfile(from_path0, to_path0)
    print('copy excel file')
    print('path: {}'.format(to_path0))
    print('*' * 80)

    # csvのコピー
    from_path1 = '{}/sample_data/{}'.format(module_path, input_header_csv_fname)
    to_path1 = '{}{}'.format(input_path, input_header_csv_fname)
    shutil.copyfile(from_path1, to_path1)
    print('copy csv file')
    print('path: {}'.format(to_path1))
    print('*' * 80)

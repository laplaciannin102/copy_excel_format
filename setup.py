# -*- coding: utf-8 -*-

# --------------------------------------------------------------------------------
# Load modules
# --------------------------------------------------------------------------------

import os, sys
from setuptools import setup, find_packages



# --------------------------------------------------------------------------------
# read files
# --------------------------------------------------------------------------------

# get current working directory
# current_path = os.path.abspath(os.path.dirname(__file__))

# read files

# requirements.txt
# requirements_path = os.path.join(current_path, 'requirements.txt')
requirements_path = './requirements.txt'
with open(file=requirements_path, mode='r', encoding='utf-8') as f:
    requirements_list = f.readlines()

# README.rst
# readme_path = os.path.join(current_path, 'README.rst')
readme_path = './README.rst'
with open(file=readme_path, mode='r', encoding='utf-8') as f:
    readme_txt = f.read()

# LICENSE
# license_path = os.path.join(current_path, 'LICENSE')
license_path = './LICENSE'
with open(file=license_path, mode='r', encoding='utf-8') as f:
    license_txt = f.read()



# --------------------------------------------------------------------------------
# setup
# --------------------------------------------------------------------------------

setup(
    name='copy_excel_format',
    version='0.1.9',
    description='copy excel format',
    long_description=readme_txt,
    author='Kosuke Asada',
    author_email='laplaciannin102@gmail.com',
    install_requires=requirements_list,
    url='https://github.com/laplaciannin102/copy_excel_format',
    license=license_txt,
    # packages=find_packages(exclude=('tests', 'docs')),
    packages=[
        'copy_excel_format',
        'copy_excel_format/datasets'
    ],
    package_dir={
        'copy_excel_format': 'copy_excel_format'
    },
    package_data={
        'copy_excel_format': [
            'datasets/sample_data/*.csv',
            'datasets/sample_data/*.xlsx'
        ]
    },
    test_suite='tests'
)


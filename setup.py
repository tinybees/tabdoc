# coding=utf-8

"""
MIT License

Copyright (c) 2018 Tiny Bees

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

"""
from setuptools import find_packages, setup

from tabdoc import __version__

setup(name='tabdoc',
      version=__version__,
      description='tabular datasets to excel,word,pdf',
      long_description=open('README.md').read(),
      long_description_content_type='text/markdown',
      author='TinyBees',
      author_email='a598824322@qq.com',
      url='https://github.com/tinybees/tabdoc',
      packages=find_packages(exclude=("tests",)),
      include_package_data=True,
      entry_points={},
      install_requires=['tablib>=0.12.1',
                        'python-docx>=0.8.10',
                        'path<=16.2.0',
                        'aelog>=1.0.3',
                        'reportlab<=3.5.68',
                        'openpyxl>=2.4.11,<=3.1.5'],
      python_requires=">=3.6",
      keywords="tabular,datasets,excel,word,pdf",
      license='MIT',
      classifiers=[
          'Development Status :: 4 - Beta',
          'Intended Audience :: Developers',
          'License :: OSI Approved :: MIT License',
          'Natural Language :: Chinese (Simplified)',
          'Operating System :: POSIX :: Linux',
          'Operating System :: Microsoft :: Windows',
          'Operating System :: MacOS :: MacOS X',
          'Topic :: Software Development :: Libraries :: Python Modules',
          'Topic :: Utilities',
          'Programming Language :: Python',
          'Programming Language :: Python :: 3.6',
          'Programming Language :: Python :: 3.7',
          'Programming Language :: Python :: 3.8',
          'Programming Language :: Python :: 3.9']
      )

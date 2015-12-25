#!/usr/bin/env python
# -*- coding:utf-8 -*-

import os
import sys
import pyexcel as pe
from pyexcel.ext import xlsx, xls

#class Api(object):


def code_gen(base_dir):
    # if not base_dir in sys.path:
    #     sys.path.append(base_dir)

    for filename in os.listdir(base_dir):
        if filename.endswith(('xlsx')):
            print filename
            # 加载excel
            sheet = pe.get_sheet(file_name=os.path.join(base_dir, filename))

            # row_range
            for r in sheet.row_range():
                # note
                if sheet.cell_value(r, 0).startswith('*') or sheet.cell_value(r, 0) == '':
                    continue
                for c in sheet.column_range():
                    # empty
                    if sheet.cell_value(r, c) == '':
                        continue
                    # identifier
                    # for c == 0:
                    #
                    # print (sheet.cell_value(r, c))


if __name__ == '__main__':
    print os.getcwd()
    code_gen(os.getcwd())

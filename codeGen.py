#!/usr/bin/env python
# -*- coding:utf-8 -*-

import os
import sys
import pyexcel as pe
from pyexcel.ext import xlsx, xls

api_file_obj = open('api.txt', 'w+')
service_file_obj = open('service.txt', 'w+')

arg_type = {0: 'Str', 1: 'Int'}
arg_required = {0: '', 1: 'required=True'}
arg_missing = {'-': ''}


def write_api(api_name, args_list):
    # args perse
    args = {}
    for i in range(0, len(args_list)):
        if i % 4 == 0:
            key = args_list[i]
            args[key] = []
            continue

        args[key].append(args_list[i])

    # gen args comment
    args_comment = ''
    for arg in args:
        if args_comment != '':
            args_comment += '&'
        args_comment += '{arg}?='.format(arg=arg)

    # @api('/xxxx')
    api_file_obj.write('    @api("/{apiname}")\r\n'.format(apiname=api_name))
    # def xxxx(self):
    api_file_obj.write('    def {apiname}(self):\r\n'.format(apiname=api_name))
    # comment
    api_file_obj.write('        """\r\n')
    api_file_obj.write('            /webapi/pre?fix/{apiname}?{argscomment}\r\n'.format(
        apiname=api_name, argscomment=args_comment
    ))
    api_file_obj.write('        """\r\n\r\n')
    # args_spec
    api_file_obj.write('        args_spec = {\r\n')
    for arg, value in args.items():
        api_file_obj.write("            '{arg}': fields.{type}({required}{missing}),\r\n".format(
            arg=arg,
            type=arg_type[value[0]],
            required=arg_required[value[1]],
            missing='' if value[2] in arg_missing else '{space}missing={value}'.format(
                space='' if arg_required[value[1]] is '' else ', ',
                value=value[2]),
        ))
    api_file_obj.write('        }\r\n')
    api_file_obj.write('        arg = args_parsr.parse(args_spec)\r\n\r\n')
    api_file_obj.write('        return srv.{srvname}(arg)\r\n\r\n\r\n'.format(srvname=api_name))

def write_service(api_name, args_list):
    # def xxx(args):
    service_file_obj.write('def {apiname}(args):\r\n'.format(apiname=api_name))
    # args
    args_name = []
    for i in range(0, len(args_list)):
        if i % 4 == 0:
            args_name.append(args_list[i])
    for arg in args_name:
        service_file_obj.write('    if {arg} in args:\r\n'.format(arg=arg))
        service_file_obj.write('        {arg} = args[{arg}]\r\n\r\n'.format(arg=arg))


API_FIELD = '*Api'
ARGS_FIELD = '*Args'
api_field = {API_FIELD: 'api', ARGS_FIELD: 'args'}
api_value = []


def gen_by_api():
    for api in api_value:
        write_api(api[API_FIELD][0], api[ARGS_FIELD])
        write_service(api[API_FIELD][0], api[ARGS_FIELD])


def code_gen(base_dir):
    for filename in os.listdir(base_dir):
        if filename.endswith(('xlsx')):
            print filename
            # 加载excel
            sheet = pe.get_sheet(file_name=os.path.join(base_dir, filename))

            # row_range
            per_api = {}
            for r in sheet.row_range():
                if sheet.cell_value(r, 0) != '':
                    head_cell = sheet.cell_value(r, 0)
                # parse field
                if head_cell.startswith('*'):
                    # note
                    if not head_cell in api_field:
                        continue

                    if head_cell == API_FIELD and len(per_api):
                        api_value.append(per_api.copy())
                        per_api.clear()

                    if sheet.cell_value(r, 0) != '':
                        per_api[head_cell] = []

                for c in sheet.column_range():
                    if c == 0: continue
                    if sheet.cell_value(r, c) == '':
                        continue
                    per_api[head_cell].append(sheet.cell_value(r, c))

            api_value.append(per_api.copy())
            per_api.clear()

            gen_by_api()


if __name__ == '__main__':
    code_gen(os.getcwd())

    api_file_obj.close()

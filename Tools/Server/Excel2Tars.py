# -*- coding:gbk -*-

import xlrd
import sys
import time

def open_excel(file= 'file.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception, e:
        print str(e)

def excel_table_byindex(file= 'file.xls',colnameindex=0,by_index=0):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows
    ncols = table.ncols
    colnames =  table.row_values(colnameindex)
    list =[]
    for rownum in range(1,nrows):
        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
                list.append(app)
    return list

def print_jce_line_begin(struct_list, line, level):
    for i in range(1, level):
        struct_list.append("    ")
    struct_list.append(line)
    struct_list.append("\n")
def print_jce_module_begin(struct_list, module_name, level):
    for i in range(1, level):
        struct_list.append("    ")
    struct_list.append("module ")
    struct_list.append(module_name)
    struct_list.append("\n")
    for i in range(1, level):
        struct_list.append("    ")
    struct_list.append("{" )
    struct_list.append("\n")
def print_jce_module_end(struct_list, level):
    for i in range(1, level):
        struct_list.append("    ")
    struct_list.append("};")
    struct_list.append("\n")
def print_jce_struct_begin(struct_list, struct_name, level):
    for i in range(1, level):
        struct_list.append("    ")
    struct_list.append("struct ")
    struct_list.append("T")
    struct_list.append(struct_name)
    struct_list.append("\n")
    for i in range(1, level):
        struct_list.append("    ")
    struct_list.append("{" )
    struct_list.append("\n")

def print_jce_struct_end(struct_list, level):
    for i in range(1, level):
        struct_list.append("    ")
    struct_list.append("};")
    struct_list.append("\n")
def print_jce_field(struct_list, field_num, field_type, field_name, field_comment, level):
    for i in range(1, level):
        struct_list.append("    ")
    struct_list.append(str(field_num))
    struct_list.append("    ")
    struct_list.append("optional")
    struct_list.append(" ")
    struct_list.append(field_type)
    for i in range(1, 3):
        struct_list.append("    ")
    struct_list.append(field_name)
    struct_list.append(";")
    for i in range(1, 3):
        struct_list.append("    ")
    struct_list.append("//")
    struct_list.append(unicode(field_comment))
    struct_list.append("\n")





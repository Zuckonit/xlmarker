#!/usr/bin/env python
# encoding: utf-8

from optparse import OptionParser
from glob import glob
import os.path as osp
import sys
import xlrd
import xlutils.copy


reload(sys)
sys.setdefaultencoding('utf-8')


UNIQUE_KEY = u'企业详细名称'
COL_MARK_NAME = u'是否标记'

def get_key_col(row0, key):
    """获取key所在的列"""
    for i, v in enumerate(row0):
        if v == key:
            return i
    return -1


def get_unique_key_col(row0):
    return get_key_col(row0, UNIQUE_KEY)


def get_mark_key_col(row0):
    return get_key_col(row0, COL_MARK_NAME)


def loader(f):
    """load excel to data structure
        {"Company": [(row, col_for_mark), (row2, col2]}
    """
    fd = xlrd.open_workbook(f)
    table = fd.sheet_by_index(0)  #only one table of each xls
    row0 = table.row_values(0)
    rows = table.nrows
    cols = table.ncols
    d = {}

    key_col = get_unique_key_col(row0)
    if key_col == -1:
        raise Exception(u"Excel sheet should have item {0}".format(UNIQUE_KEY))
    mark_col = get_mark_key_col(row0)
    should_add_mark_label = True if mark_col == -1 else False  #add mark label
    mark_col = cols if mark_col == -1 else mark_col

    for row in xrange(1, rows):
        v = table.row_values(row)
        k = str(v[key_col]).strip()
        if d.has_key(k):
            d.get(k).append((row, mark_col))
        else:
            d[k]= [(row, mark_col),]
    return fd, table, d, should_add_mark_label


def get_intersection(data1, data2):
    keys1 = data1.keys()
    keys2 = data2.keys()
    u = list([k for k in keys1 if k in keys2])
    return u


def mark(interactions, filename, fd, table, data, lb, sign='1', xf=0, offset=0):
    """fd is a file handler of excel"""
    _type = 2 if isinstance(sign, int) else 1
    has_same = False
    cur_cols = table.ncols
    for k in interactions:
        pos = data.get(k)
        if not pos:
            continue
        for p in pos:
            has_same = True
            table.put_cell(p[0], p[1]-offset, _type, sign, xf)
            cur_cols = p[1] - offset
    if has_same and lb:
        table.put_cell(0, cur_cols, 1, COL_MARK_NAME, xf)

    if has_same:
        workBook = xlutils.copy.copy(fd)
        workBook.save(filename)


def mark_processor1(fd1, tb1, d1, f2, lb1):
    fd2, tb2, d2, lb2 = loader(f2)
    len_com = len(d2.keys())
    print '\n----------------------------------------------'
    print 'start to get the interactions companies...'
    print '{0} has {1} companies'.format(osp.basename(f2), len_com)
    interactions = get_intersection(d1, d2)
    len_inter = len(interactions)
    mark(interactions, f1, fd1, tb1, d1, lb1)
    mark(interactions, f2, fd2, tb2, d2, lb2)
    print '{0} handle over, marked {1} companies'.format(osp.basename(f1), len_inter)
    print '----------------------------------------------\n'
    return len_com, len_inter


def mark_processor(f1, f2):
    fd1, tb1, d1, lb1 = loader(f1)
    return mark_processor1(fd1, tb1, d1, f2, lb1)
    

def optargs(arg):
    parser = OptionParser(usage="usage: %prog", version="%prog 0.0.2")
    parser.add_option('-f', "--file", dest='file', metavar='FILE', 
            help='参照表(A表)')
    parser.add_option('-d', '--directory', dest='directory', metavar='DIR',
            help='待标记表格所在目录(B表目录)')
    return parser.parse_args(arg)


if __name__ == '__main__':
    options, args = optargs(sys.argv)
    f1 = options.file
    d  = options.directory
    pname = osp.basename(sys.argv[0])
    
    if not f1 or not d:
        print 'type {0} -h to get help'.format(pname)
        sys.exit(1)
    if not f1 or not osp.isfile(f1):
        print '{0} not exist or not a file'.format(f1)
        sys.exit(1)
    if not d or not osp.isdir(d):
        print '{0} not exist or not a directory'.format(d)
        print sys.exit(1)

    companies_for_mark = 0
    marked = 0

    fd1, tb1, d1, lb1 = loader(f1)
    for f in glob('{0}/*'.format(d)):
        f = f.replace('\\', '/')
        len_com, len_inter = mark_processor1(fd1, tb1, d1, f, lb1)
        companies_for_mark += len_com
        marked += len_inter

    print '==============Statitics===================='
    print 'TOTAL COMPANIES: {0}'.format(companies_for_mark)
    print 'TOTAL MARKED: {0}'.format(marked)
    print '==========================================='

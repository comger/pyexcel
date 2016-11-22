# -*- coding:utf-8 -*-
"""
    author comger@gmail.com
    example for excel template render
"""
from excel_template import ExcelTemplate

template = 'abc.xls'

et = ExcelTemplate(template)

headermap = {'company':u'comp', 'addon':'2016-11-07', 'all':'100'}
table1 = []
for i in range(10):
    item = dict(no=i, name='name'+str(i), day='day'+str(i), num=i*10, price=i*12, dw='m', fj=i*12*i*10, bz='')
    table1.append(item)

table = []
for r in range(10):
    arr = []
    for c in range(10):
        arr.append('{0}_{1}'.format(r,c))

    table.append(arr)

et.render('output.xls', ht=headermap, table1=table1,table2=table1, table3=table1, table=table)




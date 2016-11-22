# -*- coding:utf-8 -*-
"""
    author comger@mgail.com
    logic for excel
"""
import re
from xlwt.Style import xf_dict,XFStyle
from xlrd import open_workbook 
from xlutils.copy import copy
from xlutils.styles import Styles

partter = r'{{([^\s]+)}}'
for_partter = r'{%Foreach.([^\s]+)%}'
for_end = r'{%Forend.([^\s]+)%}'
fix_start = r'{%FixStart.([^\s]+)%}'
fix_end = r'{%FixEnd.([^\s]+)%}'
table_partter = r'{%Table.([^\s]+)%}'


vertical_align = {
    0:'top',
    1:'middle',
    2:'bottom',
    3:'auto',
    4:'auto'
}

hor_align = {
    0:'auto',
    1:'left',
    2:'center',
    3:'right',
    4:'auto',
    5:'auto',
    6:'auto',
    7:'auto'
}

class ExcelTemplate(object):
    def __init__(self, template_path):
        
        self.notformatbook = open_workbook(template_path)
        self.sheetslen = len(self.notformatbook.sheets())
        self.rbook =  open_workbook(template_path,formatting_info=True)       
        self.styles = Styles(self.rbook)
        self.conf = {}


    def get_config(self, idx=0):
        conf = self.conf.get(idx,{}) 
        if not conf:
            rsheet = self.notformatbook.sheet_by_index(idx)
            startrowindex = rsheet._first_full_rowx
            first_allnull_rowx = startrowindex+1
            for row in range(first_allnull_rowx,rsheet.nrows):
                if rsheet.row_types(row).count(1)==0:
                    first_allnull_rowx = row
                    break

                
            if first_allnull_rowx == rsheet.nrows - 1:
                first_allnull_rowx += 1

            first_nonull_rowx = first_allnull_rowx
            for row in range(first_allnull_rowx,rsheet.nrows):
                if rsheet.row_types(row).count(1)>0:
                    first_nonull_rowx = row
                    break
            
            conf['first_allnull_rowx'] = first_allnull_rowx
            conf['first_nonull_rowx'] = first_nonull_rowx

            fields = []
            startrowindex = first_allnull_rowx - 1
            if rsheet.row_types(startrowindex).count(0) == 0:
                """ 如果本行没有空单元 """
                for cell in rsheet.row(startrowindex):
                    fields.append(cell.value)
            else:
                rowids = range(startrowindex,first_allnull_rowx)
                rowids.reverse()
                for col in range(rsheet.ncols):
                    f = None
                    for rowidx in rowids:
                        f = rsheet.cell_value(rowidx,col)
                        if f:break
                        
                    fields.append(f)


            conf['fields'] = fields
            self.conf[idx] = conf

        return conf
            
    def parse_data(self,idx,dataset):
        conf = self.get_config(idx=idx)
        newlst = []
        if dataset:
            for row in dataset[idx]:
                arr = []
                for f in conf['fields']:
                    arr.append(row.get(f,''))

                newlst.append(arr)

        return newlst
    

    def get_template(self,idx,sheetStartIdx):
        '''
            获取表头模板标签位置及标签内容
        '''
        temp,for_temp,fix_temp,table_temp = {},{},{},{}
        for_index = []
        
        rsheet = self.rbook.sheet_by_index(idx)
        for i in range(0, rsheet.nrows):
            for col,cell in enumerate(rsheet.row(i)):
                key = '%d_%d' %(i,col)
                m = re.findall(partter,cell.value)
                if m and '.' in m[0]:
                    temp[key] = dict(old=cell.value,match=m)

                m = re.findall(for_partter,cell.value)
                if m:
                    for_temp[m[0]]= dict(start= dict(row=i, col=col))

                
                m = re.findall(for_end,cell.value)
                if m:
                    for_index.append(m[0])
                    for_temp[m[0]]['end']=dict(row=i, col=col)
                    fields = {}
                    if col == for_temp[m[0]]['start']['col'] and  (i-1) >= for_temp[m[0]]['start']['row']:
                        #foreach 竖向循环
                        for f_col,f_cell in enumerate(rsheet.row(i-1)):
                            f_m = re.findall(partter,f_cell.value)
                            if f_m and not '.' in f_m[0]:
                                f_key = '%d_%d' %(i-1,f_col)
                                fields[f_key] = f_m[0]


                        for_temp[m[0]]['fields'] = fields


                    if i == for_temp[m[0]]['start']['row'] and  (cel-1) >= for_temp[m[0]]['start']['col']:
                        #foreach 竖向循环
                        for f_col,f_cell in enumerate(rsheet.row(i)):
                            f_m = re.findall(partter,f_cell.value)
                            if f_m and not '.' in f_m[0]:
                                f_key = '%d_%d' %(i-1,f_col)
                                fields[f_key] = f_m[0]
                        
                        for_temp[m[0]]['fields'] = fields

                m = re.findall(fix_start,cell.value)
                if m:
                    fix_temp[m[0]]= dict(start= dict(row=i, col=col))


                m = re.findall(fix_end,cell.value)
                if m:
                    fix_temp[m[0]]['end']=  dict(row=i, col=col)
                    for_index.append(m[0]) 


                m = re.findall(table_partter,cell.value)
                if m:
                    table_temp[m[0]] = {'start':dict(row=i, col=col)}
                    for_index.append(m[0]) 

                
        
        conf = dict(header=temp, for_temp=for_temp,for_index=for_index, fix_temp=fix_temp, table_temp=table_temp)
        return conf 

    
    def move_row(self, ws, args, from_row, to_row, idx=0):
        rsheet = self.rbook.sheet_by_index(idx)
        for col,cell in enumerate(rsheet.row(from_row)):
            m = re.findall(partter,cell.value)
            body = cell.value
            if m and '.' in m[0]:
                key,name = m[0].split('.')
                body = args.get(key,{}).get(name,'')
                
            xf = self.rbook.xf_list[cell.xf_index]
            ws.write(to_row,col,body,self.conv_xf_xfstyle(xf))
            ws.write(from_row,col,'')
            
            

    def del_row(self, ws, row, idx=0):
        rsheet = self.rbook.sheet_by_index(idx)
        for col,cell in enumerate(rsheet.row(row)):
            ws.write(row,col)
            

    def render(self, target, **kwargs):
        '''
            按序渲染excel 模板
        '''

        wb = copy(self.rbook)
        for idx in range(0, 1):
            sheetStartIdx = self.get_config(idx=idx)['first_allnull_rowx'] - 1
            conf = self.get_template(idx,sheetStartIdx)
            ht = conf['header']
            ws = wb.get_sheet(idx)

            for key,val in ht.items():
                row,col = key.split('_')

                body = val['old']
                for m in val['match']:
                    name,key = m.split('.')
                    
                    body=body.replace('{{%s}}' % (str(m)), str(kwargs.get(name,{}).get(key, '')) or '')
                
                row,col = int(row),int(col)
                cell = self.rbook.sheet_by_index(idx).cell(row,col)
                xf = self.rbook.xf_list[cell.xf_index]
                ws.write(row,col,body,self.conv_xf_xfstyle(xf))
           
            fix_temp_map = conf.get('fix_temp',{})
            for_index = conf.get('for_index',())
            for key,fix_temp in fix_temp_map.items():
                move_row = 0
                for forkey in for_index[0:for_index.index(key)]:
                    move_row = move_row + len(kwargs.get(forkey,()))
                    
                self.del_row(ws, fix_temp['start']['row'])
                self.del_row(ws, fix_temp['end']['row'])
                for row in range(fix_temp['start']['row']+1,fix_temp['end']['row']):
                    self.move_row(ws, kwargs,row,row+move_row,idx)


            for_temp_map = conf.get('for_temp',{})
            for key, for_temp in for_temp_map.items():
                fields = for_temp.get('fields',{})
                need_add_for_tag = 0
                
                for front_key in for_index:
                    if key==front_key:
                        break


                    need_add_for_tag = need_add_for_tag+len(kwargs.get(front_key,()))+for_temp['end']['row']-for_temp['end']['row']-1
                
                if for_temp['end']['row']-for_temp['start']['row']>1:
                    need_add_for_tag = need_add_for_tag-1


                if need_add_for_tag>0:
                    m_row = for_temp['start']['row']-1
                    self.move_row(ws, kwargs,m_row,m_row+need_add_for_tag+1,idx)
                    self.del_row(ws, for_temp['start']['row']+1)
                    self.del_row(ws, for_temp['start']['row'])
                    self.del_row(ws, for_temp['end']['row'])

                for pos,fname in fields.items():
                    row,col = pos.split('_')
                    row,col = int(row),int(col)
                    cell = self.rbook.sheet_by_index(idx).cell(row,col)
                    xf = self.rbook.xf_list[cell.xf_index]
                    for i,item in enumerate(kwargs.get(key,())):
                        val = item.get(fname,'')
                        
                        ws.write(row+i+need_add_for_tag ,col,val,self.conv_xf_xfstyle(xf))

            table_temp_map = conf.get('table_temp',{})
            for key, table_temp in table_temp_map.items():
                pos = table_temp['start']
                move_row = 0
                cell = self.rbook.sheet_by_index(idx).cell(pos['row'],pos['col'])
                xf = self.rbook.xf_list[cell.xf_index]
                for forkey in for_index[0:for_index.index(key)]:
                    move_row = move_row + len(kwargs.get(forkey,()))

                table_data = kwargs.get(key,())
                for r, row_data in enumerate(table_data):
                    for c,cell_val in enumerate(row_data):
                        ws.write(pos['row']+r+move_row, pos['col']+c, cell_val,self.conv_xf_xfstyle(xf))
                
        wb.save(target)

    def conv_xf_xfstyle(self, xf):
        wtxf = XFStyle()

        wtf = wtxf.font
        rdf = self.rbook.font_list[xf.font_index]
        wtf.height = rdf.height
        wtf.italic = rdf.italic
        wtf.struck_out = rdf.struck_out
        wtf.outline = rdf.outline
        wtf.shadow = rdf.outline
        wtf.colour_index = rdf.colour_index
        wtf.bold = rdf.bold #### This attribute is redundant, should be driven by weight
        wtf._weight = rdf.weight #### Why "private"?
        wtf.escapement = rdf.escapement
        wtf.underline = rdf.underline_type #### 
        # wtf.???? = rdf.underline #### redundant attribute, set on the fly when writing
        wtf.family = rdf.family
        wtf.charset = rdf.character_set
        wtf.name = rdf.name
        # 
        # protection
        #
        wtp = wtxf.protection
        rdp = xf.protection
        wtp.cell_locked = rdp.cell_locked
        wtp.formula_hidden = rdp.formula_hidden
        #
        # border(s) (rename ????)
        #
        wtb = wtxf.borders
        rdb = xf.border
        wtb.left   = rdb.left_line_style
        wtb.right  = rdb.right_line_style
        wtb.top    = rdb.top_line_style
        wtb.bottom = rdb.bottom_line_style 
        wtb.diag   = rdb.diag_line_style
        wtb.left_colour   = rdb.left_colour_index 
        wtb.right_colour  = rdb.right_colour_index 
        wtb.top_colour    = rdb.top_colour_index
        wtb.bottom_colour = rdb.bottom_colour_index 
        wtb.diag_colour   = rdb.diag_colour_index 
        wtb.need_diag1 = rdb.diag_down
        wtb.need_diag2 = rdb.diag_up
        #
        # background / pattern (rename???)
        #
        wtpat = wtxf.pattern
        rdbg = xf.background
        wtpat.pattern = rdbg.fill_pattern
        wtpat.pattern_fore_colour = rdbg.pattern_colour_index
        wtpat.pattern_back_colour = rdbg.background_colour_index


        wta = wtxf.alignment
        rda = xf.alignment
        wta.horz = rda.hor_align

        wta.vert = rda.vert_align

        wta.dire = rda.text_direction
        wta.rota = rda.rotation
        wta.wrap = rda.text_wrapped
        wta.shri = rda.shrink_to_fit
        wta.inde = rda.indent_level

        return wtxf


    def writeRow(self,sheet,rowindex,rowdata,xlsStyleDict,idx=0):
        for col,o in enumerate(rowdata):
            if col in xlsStyleDict:
                style = xlsStyleDict.get(col)
                sheet.write(rowindex,col,o,style)
            else:
                cell = self.rbook.sheet_by_index(idx).cell(rowindex,col)
                xf = self.rbook.xf_list[cell.xf_index]
                sheet.write(rowindex,col,o,self.conv_xf_xfstyle(xf))

    


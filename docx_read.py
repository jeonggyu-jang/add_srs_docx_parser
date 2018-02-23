import os
import docx
from docx.document import Document
from docx.oxml.table import *
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml.numbering import CT_NumPr
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill, Color
from konlpy.tag import Kkma
from konlpy.utils import pprint
import datetime

os.chdir('C:\\Users\\mocolab\\PycharmProjects\\add_parser_docx_ver') #docx file directory
doc = docx.Document('Use_case 작성.docx')
global sv_tblId #shared variable for table ID numbering
sv_tblId = 0

class Paragraph_DS:
    def __init__(self,text,ilvl=0,ccff=None,name=None):
        self.text = text
        self.ilvl = ilvl
        self.ccff = ccff
        self.name = name
        self.tree = None
        self.kkma_pos()
    def kkma_pos(self):
        kkma=Kkma()
        self.tree = kkma.pos(self.text)

class Tbl_DS:
    def __init__(self,tblId,newCell=None):
        self.ttlNmbRow = 0
        self.ttlNmbCol = 0
        self.cells = []
        self.tblId = tblId
        if newCell != None :
            self.insert_cell(newCell)
    def insert_cell(self,newCell):
        newCellDict = {'row':newCell.row,'col':newCell.col,'cell':newCell}
        self.cells.append(newCellDict)
        if self.ttlNmbRow < newCell.row:
            self.ttlNmbRow = newCell.row
        elif self.ttlNmbCol < newCell.col:
            self.ttlNmbCol = newCell.col
    def find_cell(self,row,col):
        for i in range(len(self.cells)):
            row_t = self.cells[i].get('row')
            col_t = self.cells[i].get('col')
            if row == row_t and col == col_t :
                return self.cells[i].get('cell')

class Cell_DS:
    def __init__(self,affTbl,row,col,newPrgrph=None,newTbl=None):
        self.prgrphs = []
        self.row = row
        self.col = col
        self.tbls = []
        self.affTbl = affTbl
        if newPrgrph != None :
            self.insert_prgrph(newPrgrph)
        if newTbl != None :
            self.insert_tbl(newTbl)
        self.affilate_tbl()
    def insert_prgrph(self,newPrgrph):
        self.prgrphs.append(newPrgrph)
    def insert_tbl(self,newTbl):
        self.tbls.append(newTbl)
    def affilate_tbl(self):
        self.affTbl.insert_cell(self)


class DS_rule_checker:
    def __init__(self,text,ilvl=0,row=0,col=0,ccff=0):
        self.text = text
        self.row = row
        self.col = col
        self.ilvl = ilvl
        self.ccff = ccff
    def parsing(self):
        kkma = Kkma()
        tree = kkma.pos(self.text)
        '''
        grammar = """
        NP: {<N.*>*<Suffix>?}   # Noun phrase
        VP: {<V.*>*}            # Verb phrase
        AP: {<A.*>*}            # Adjective phrase
        """
        parser = nltk.RegexpParser(grammar)
        chunks = parser.parse(tree)
        chunks.draw()
        '''
        print(tree)
        # translated_seq ,check_seq= self.Translate(tree)
        # composed_seq = self.Compose_unit(translated_seq)
        '''
        print(translated_seq)
        print(check_seq)
        print(composed_seq)
        '''
        for k in range(len(tree)):
            if tree[k][1][0] == 'J' or tree[k][1][0] == 'E' or tree[k][1][0] == 'M' :
                print(tree[k],end=" ")

        # post_processed_seq = self.Comp_OP_post_processing(composed_seq)
        '''
        for j in range(len(translated_seq)):
            print(translated_seq[j],end=" ")
        print("\n")
        '''

        # template_number = self.WhatTemplate(post_processed_seq)
        #print(template_number)
        '''
        for j in range(len(post_processed_seq)):
            print(post_processed_seq[j], end=" ")
        print("    -- Template[{}] \n".format(template_number))
        '''

        #excel output
        #self.changeNum(post_processed_seq)
        #cond_sen, cond_time_cond_sen, act_sen, act_time_cond_sen = self.devideCondAct(post_processed_seq)

class Tree:
    def __init__(self, data, parent=None):
        self.data = data
        self.parent  = parent
        self.children = list()
    def new_child_append(self,new_child):
        print(new_child,end="-----")
        self.children.append(new_child)
    def print_children(self):
        for i in range(len(self.children)):
            print("%s - %s"%(self.data,self.children[i].data))
        for i in range(len(self.children)):
            self.children[i].print_children()

def find_ilvl_val(root):
    if isinstance(root,CT_NumPr):
        ilvl_val = root.ilvl.val
        return ilvl_val+1
    else :
        for diver in root.iterchildren():
            temp = find_ilvl_val(diver)
            if temp != 0:
                ilvl_val = temp
                return ilvl_val
    return 0

def iter_block_items(parent):
    """
    Yield each paragraph and table child within *parent*, in document order.
    Each returned value is an instance of either Table or Paragraph. *parent*
    would most commonly be a reference to a main Document object, but
    also works for a _Cell object, which itself can contain paragraphs and tables.
    """
    if isinstance(parent, Document): #The type of root is determined.
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
        print("iter_parent is _Cell")
    elif isinstance(parent,CT_Tc):
        parent_elm = parent
    else:
        raise ValueError("something's not right")
    cell_color_filled_flag = 0
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            ilvl_val = find_ilvl_val(child)
            #print(ilvl_val)
            #print("iter_child is CT_P")
            yield Paragraph(child, parent), child, ilvl_val, cell_color_filled_flag
            cell_color_filled_flag = 0
            #for cchild in child.iterchildren():
                #print("\t",end="")
                #print(type(cchild))
                #print("\t  ",end="")
        elif isinstance(child, CT_Tbl):
            #print(child.tblStyle_val)
            #print("iter_child is CT_Tbl")
            yield Table(child, parent), child, 0, cell_color_filled_flag
            cell_color_filled_flag = 0
        elif isinstance(child,CT_TcPr):
            for tcpr in child.iterchildren():
                if "shd" in str(tcpr) :
                    cell_color_filled_flag = 1

def table_parsing(root,affTbl):
    global sv_tblId
    row_n=0
    for child in root.iterchildren():
        if isinstance(child,CT_Row):
            row_n += 1
            col_n = 0
            for cell in child.iterchildren():
                if isinstance(cell,CT_Tc):
                    col_n += 1
                    temp_inst_cell = Cell_DS(affTbl,row_n,col_n)
                    for block, root_tbl, ilvl_val,ccff in iter_block_items(cell):
                        if "Paragraph" in str(type(block)):
                            #print("[%d,%d]"%(row_n,col_n),end=" ")
                            #for i in range(ilvl_val):
                                #if i != ilvl_val-1 : print("     ",end="")
                                #else : print("   L__",end="")
                            if ccff == 1 :
                                temp_inst_cell.insert_prgrph(Paragraph_DS(block.text,ccff=ccff,ilvl=ilvl_val))
                                #print ('__%s ||%s||'%(str(ilvl_val),block.text))
                            else :
                                temp_inst_cell.insert_prgrph(Paragraph_DS(block.text,ilvl=ilvl_val))
                                #print ('__%s %s'%(str(ilvl_val),block.text))
                        elif "Table" in str(type(block)):
                            sv_tblId += 1
                            temp_inst_tbl = Tbl_DS(sv_tblId)
                            temp_inst_cell.insert_tbl(temp_inst_tbl)
                            table_parsing(root_tbl)


def print_out_srs(srs):
    for i in range(len(srs)):
        if isinstance(srs[i],Paragraph_DS) :
            for k in range(srs[i].ilvl):
                if k != (srs[i].ilvl)-1 : print("     ",end="")
                else : print("   L__",end="")
            if srs[i].ccff == 1 :
                print ('__%s ||%s||'%(str(srs[i].ilvl),srs[i].text),end=" ")
                print ('_%s_'%(srs[i].tree))
            else :
                print ('__%s %s'%(str(srs[i].ilvl),srs[i].text),end=" ")
                print ('_%s_'%(srs[i].tree))
        elif isinstance(srs[i],Tbl_DS) :
            print_out_table(srs[i])

def print_out_table(table):
    # if (table.ttlNmbRow * table.ttlNmbCol) != len(table.cells):
    #     print("[ERROR] [tblId = %d] table has a incorrect the number of cells "%(table.tblId))
    #     print("[ERROR] %d != %d"%((table.ttlNmbRow * table.ttlNmbCol),len(table.cells)))
    for i in range(len(table.cells)):
        for j in range(len(table.cells[i].get('cell').prgrphs)):
            print("[%d,%d]"%(table.cells[i].get('row'),table.cells[i].get('col')),end=" ")
            for k in range(table.cells[i].get('cell').prgrphs[j].ilvl):
                if k != table.cells[i].get('cell').prgrphs[j].ilvl - 1 : print("     ",end="")
                else : print("   L__",end="")
            if table.cells[i].get('cell').prgrphs[j].ccff == 1 :
                print ('__%s ||%s||'%(str(table.cells[i].get('cell').prgrphs[j].ilvl),table.cells[i].get('cell').prgrphs[j].text),end="")
                print ('_%s_'%(table.cells[i].get('cell').prgrphs[j].tree))
            else :
                print ('__%s %s'%(str(table.cells[i].get('cell').prgrphs[j].ilvl),table.cells[i].get('cell').prgrphs[j].text),end="")
                print ('_%s_'%(table.cells[i].get('cell').prgrphs[j].tree))
        for j in range(len(table.cells[i].get('cell').tbls)):
            print_out_table(table.cells[i].get('cell').tbls[j])


def srs_parsing():
    global sv_tblId
    srs = [] # all srs block list
    tbl_list = [] # list only for table block
    for block,root, ilvl_val,ccff in iter_block_items(doc): #ccff : cell_color_filled_flag
        if "Paragraph" in str(type(block)):
            if ccff == 1 :
                srs.append(Paragraph_DS(block.text,ccff=ccff,ilvl=ilvl_val))
                #print ('__%s ||%s||'%(str(ilvl_val),block.text))
            else :
                srs.append(Paragraph_DS(block.text,ilvl=ilvl_val))
                #print ('__%s %s'%(str(ilvl_val),block.text))
        elif "Table" in str(type(block)):
            sv_tblId = sv_tblId + 1
            temp_inst_tbl = Tbl_DS(sv_tblId)
            srs.append(temp_inst_tbl)
            tbl_list.append(temp_inst_tbl)
            table_parsing(root,temp_inst_tbl)
    #print(srs)
    print_out_srs(srs)

srs_parsing()
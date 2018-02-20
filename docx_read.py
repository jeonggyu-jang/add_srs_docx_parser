import os
import docx
from docx.document import Document
from docx.oxml.table import *
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml.numbering import CT_NumPr

os.chdir('C:\\Users\\mocolab\\PycharmProjects\\add_parser_docx_ver') #docx file directory
doc = docx.Document('Use_case 작성.docx')

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

def table_parsing(root):
    row_n=0
    for child in root.iterchildren():
        if isinstance(child,CT_Row):
            row_n += 1
            col_n = 0
            for cell in child.iterchildren():
                if isinstance(cell,CT_Tc):
                    col_n += 1
                    for block, root_tbl, ilvl_val,ccff in iter_block_items(cell):
                        if "Paragraph" in str(type(block)):
                            print("[%d,%d]"%(row_n,col_n),end=" ")
                            for i in range(ilvl_val):
                                if i != ilvl_val-1 : print("     ",end="")
                                else : print("   L__",end="")
                            if ccff == 1 :
                                print ('__%s ||%s||'%(str(ilvl_val),block.text))
                            else :
                                print ('__%s %s'%(str(ilvl_val),block.text))
                        elif "Table" in str(type(block)):
                            table_parsing(root_tbl)


def srs_parsing():
    for block,root, ilvl_val,ccff in iter_block_items(doc): #ccff : cell_color_filled_flag
        if "Paragraph" in str(type(block)):
            if ccff == 1 :
                print ('__%s ||%s||'%(str(ilvl_val),block.text))
            else :
                print ('__%s %s'%(str(ilvl_val),block.text))

        elif "Table" in str(type(block)):
            table_parsing(root)

srs_parsing()
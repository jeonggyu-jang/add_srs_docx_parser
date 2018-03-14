import os
from gensim.models import Word2Vec
import docx
from docx.document import Document
from docx.oxml.table import *
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml.numbering import CT_NumPr
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill, Color
from konlpy.tag import Kkma, Twitter
from konlpy.utils import pprint
import datetime
from template import *
req_printout = 0
pos_printout = 1 #print out pos tree
os.chdir('C:\\Users\\mocolab\\PycharmProjects\\add_parser_docx_ver') #docx file directory
#doc = docx.Document('Use_case 작성.docx')
doc = docx.Document('표적관리_SRS.docx')
global sv_tblId #shared variable for table ID numbering
sv_tblId = 0

def passive_check(tree):
    violation_flag=0
    passive_DB = ['되','어지','받','당하','이','히','리','기']
    #print(tree)
    for i in range(len(tree)):
        for j in range(len(passive_DB)):
            if passive_DB[j] in tree[i][0] and ('N' != tree[i][1][0] and 'J' != tree[i][1][0] ):
                print("\t[WARNING!!] Passive_rule checking violation!!",end=" ")
                print(": There is \"%s\" in \"%s\"[%d]"%(passive_DB[j],tree[i][0],i))
                violation_flag = 1
    return violation_flag

def findReqId(tbl):
    return 0

def tokenizePrgrph_unitN(prgrph): #Selecting just unit Noun
    tokenized_prgrph = []
    for i in range(len(prgrph.tree)):
        if prgrph.tree[i][1] == 'NNG' or prgrph.tree[i][1] == 'NNP' or prgrph.tree[i][1] == 'OL':
            tokenized_prgrph.append(prgrph.tree[i][0])
    return tokenized_prgrph

def tokenizePrgrph_comNoun_kkma(prgrph): #Selecting just compound Noun
    tokenized_prgrph = []
    reqId = prgrph.reqId
    i=0
    while i <= (len(prgrph.tree)-1):
        comNoun = str()
        if prgrph.tree[i][1][0:2] == 'NN' or prgrph.tree[i][1] == 'OL' or prgrph.tree[i][0] == '-' or prgrph.tree[i][1] == 'XSN' or prgrph.tree[i][1] == 'JKG': #have to add 'NR' !!!!!!!!!!!
            try :
                if prgrph.tree[i-1][1][0] == 'J' and prgrph.tree[i+1][1] == 'XSV':
                    newDict = {'word':prgrph.tree[i][0] + '하다','reqId':reqId}
                    tokenized_prgrph.append(newDict)
                else :
                    comNoun += prgrph.tree[i][0]
                    for j in range(i+1,len(prgrph.tree)):
                        if prgrph.tree[j][1][0:2] != 'NN' and prgrph.tree[j][1] != 'OL' and prgrph.tree[j][0] != '-' and prgrph.tree[j][1] != 'XSN' and prgrph.tree[j][1] != 'JKG':
                            #print(comNoun)
                            i = j
                            break
                        elif j+1 == len(prgrph.tree):
                            comNoun += prgrph.tree[j][0]
                            i = j
                        else :
                            comNoun += prgrph.tree[j][0]
                    newDict = {'word':comNoun,'reqId':reqId}
                    tokenized_prgrph.append(newDict)
            except :
                comNoun += prgrph.tree[i][0]
                for j in range(i+1,len(prgrph.tree)):
                    if prgrph.tree[j][1][0:2] != 'NN' and prgrph.tree[j][1] != 'OL' and prgrph.tree[j][0] != '-' and prgrph.tree[j][1] != 'XSN' and prgrph.tree[j][1] != 'JKG':
                        #print(comNoun)
                        i = j
                        break
                    elif j+1 == len(prgrph.tree):
                        comNoun += prgrph.tree[j][0]
                        i = j
                    else :
                        comNoun += prgrph.tree[j][0]
                newDict = {'word':comNoun,'reqId':reqId}
                tokenized_prgrph.append(newDict)
        i += 1
    return tokenized_prgrph

def tokenizePrgrph_comNoun_twitter(prgrph): #Selecting just compound Noun
    tokenized_prgrph = []
    reqId = prgrph.reqId
    i=0
    while i <= (len(prgrph.t_tree)-1):
        comNoun = str()
        if (prgrph.t_tree[i][1] == 'Noun' or prgrph.t_tree[i][0] == '-' or prgrph.t_tree[i][1] == 'Alpha' \
                or prgrph.t_tree[i][0] == '/' or prgrph.t_tree[i][1] == 'Number') \
                and (prgrph.t_tree[i][0] != '와' and prgrph.t_tree[i][0] != '다음') :#have to add 'NR' !!!!!!!!!!!
            try :
                if prgrph.t_tree[i+1][0] == '하다':
                    newDict = {'word':prgrph.t_tree[i][0] + '하다','reqId':reqId}
                    tokenized_prgrph.append(newDict)
                else :
                    comNoun += prgrph.t_tree[i][0]
                    for j in range(i+1,len(prgrph.t_tree)):
                        if (prgrph.t_tree[j][1] != 'Noun' and prgrph.t_tree[j][0] != '-' and prgrph.t_tree[j][1] != 'Alpha' \
                                and prgrph.t_tree[j][0] != '/' and prgrph.t_tree[j][1] != 'Number')\
                                or (prgrph.t_tree[j][0] == '와' and prgrph.t_tree[j][0] == '다음') :#have to add 'NR' !!!!!!!!!!!
                            #print(comNoun)
                            i = j
                            break
                        elif j+1 == len(prgrph.t_tree):
                            comNoun += prgrph.t_tree[j][0]
                            i = j
                        else :
                            comNoun += prgrph.t_tree[j][0]
                    newDict = {'word':comNoun,'reqId':reqId}
                    tokenized_prgrph.append(newDict)
            except :
                comNoun += prgrph.t_tree[i][0]
                for j in range(i+1,len(prgrph.t_tree)):
                    if (prgrph.t_tree[j][1] != 'Noun' and prgrph.t_tree[j][0] != '-' and prgrph.t_tree[j][1] != 'Alpha' \
                            and prgrph.t_tree[j][0] != '/' and prgrph.t_tree[j][1] != 'Number')\
                            or (prgrph.t_tree[j][0] == '와' and prgrph.t_tree[j][0] == '다음') :#have to add 'NR' !!!!!!!!!!!
                        #print(comNoun)
                        i = j
                        break
                    elif j+1 == len(prgrph.t_tree):
                        comNoun += prgrph.t_tree[j][0]
                        i = j
                    else :
                        comNoun += prgrph.t_tree[j][0]
                newDict = {'word':comNoun,'reqId':reqId}
                tokenized_prgrph.append(newDict)
        i += 1
    return tokenized_prgrph

def tokenizePrgrph_N_XSV(prgrph): #Selecting just N + XSV
    tokenized_prgrph = []
    for i in range(len(prgrph.tree)-1):
        if prgrph.tree[i][1][0:2] == 'NN' and prgrph.tree[i+1][1] == 'XSV':
            tokenized_prgrph.append(prgrph.tree[i][0] + '하다')
    return tokenized_prgrph

def collectPrgrph(srs,tokenized_srs): #all paragraphs and tables parsing
    for i in range(len(srs)):
        if isinstance(srs[i],Paragraph_DS):
            tokenized_srs.append(tokenizePrgrph(srs[i]))
        elif isinstance(srs[i],Tbl_DS):
            for j in range(len(srs[i].cells)):
                collectPrgrph(srs[i].cells[j].get('cell').prgrphs,tokenized_srs)
                collectPrgrph(srs[i].cells[j].get('cell').tbls,tokenized_srs)
    return tokenized_srs

def collect_SRS_Prgrph(srs,tokenized_srs,tblflag=0):
    for i in range(len(srs)):
        if isinstance(srs[i],Paragraph_DS) and tblflag==1:
            tokenized_srs.append(tokenizePrgrph_comNoun_twitter(srs[i]))
            #tokenized_srs.append(tokenizePrgrph_N_XSV(srs[i]))
        elif isinstance(srs[i],Tbl_DS) and tblflag==0: #only the (2,2)cell is selected in table (is not in table)
            try :
                if srs[i].find_cell(2,1).prgrphs[0].text == '요구사항' :
                    collect_SRS_Prgrph(srs[i].find_cell(2,2).prgrphs,tokenized_srs,tblflag=1)
                    collect_SRS_Prgrph(srs[i].find_cell(2,2).tbls,tokenized_srs,tblflag=1)
            except :
                pass
        elif isinstance(srs[i],Tbl_DS) and tblflag==1: #all cells of the table in table are selected
            try :
                for j in range(len(srs[i].cells)):
                    cell_t = srs[i].cells[j].get('cell')
                    collect_SRS_Prgrph(cell_t.prgrphs,tokenized_srs,tblflag=1)
                    collect_SRS_Prgrph(cell_t.tbls,tokenized_srs,tblflag=1)
            except :
                pass
    return tokenized_srs

def _count(t):
    return t[1]
def _word(t):
    return t[0]
def makeReqIdDic(reqIdDic,tokenized_srs):
    for i in tokenized_srs:
        for j in i:
            word = j.get('word')
            reqId = j.get('reqId')

def makeDic_w2v(srs): #tokenized_srs를 튜플리스트로 바꾸었기때문에 makeDic_w2v에서 사용하기위해 word만 추출하는 코드 추가해야함 - jjg 03/14
    tokenized_srs=[]
    #tokenized_srs = collectPrgrph(srs,tokenized_srs)
    tokenized_srs = collect_SRS_Prgrph(srs,tokenized_srs)
    print("[tokenized srs]")
    for i in tokenized_srs:
        if i != [] :
            print(i)
    #print(tokenized_srs)
    embedding_model = Word2Vec(tokenized_srs,size=100,window=3,min_count=1,workers=4,iter=100,sg=1)
    embedding_model.init_sims(replace=True)
    embedding_model.save("word2vec_result.model")
    #print(embedding_model)
    #print(embedding_model.wv)
    #print(embedding_model.wv.vocab)
    #print(embedding_model.most_similar(positive=["표적"], topn=100))
    word_hist=[]
    for w in embedding_model.wv.vocab:
        word_hist_instance = (w,embedding_model.wv.vocab[w].count)
        word_hist.append(word_hist_instance)
    word_hist.sort(key=_count)
    for i in word_hist:
        print(i)

def makeDic(srs):
    reqIdDic = []
    tokenized_srs=[]
    tokenized_srs4w2v = []
    tokenized_srs = collect_SRS_Prgrph(srs,tokenized_srs)
    for i in tokenized_srs:
        list_t = list()
        for j in i:
            word_t = j.get('word')
            list_t.append(word_t)
            print(word_t,j.get('reqId'))
        tokenized_srs4w2v.append(list_t)
    embedding_model = Word2Vec(tokenized_srs4w2v,size=100,window=3,min_count=1,workers=4,iter=100,sg=1)
    embedding_model.init_sims(replace=True)
    embedding_model.save("word2vec_result.model")
    word_hist=[]
    for w in embedding_model.wv.vocab:
        word_hist_instance = (w,embedding_model.wv.vocab[w].count)
        word_hist.append(word_hist_instance)
    word_hist.sort(key=_count)
    print("- - - - - [Dictionary] - - - - -")
    reqIdDic = makeReqIdDic(reqIdDic,tokenized_srs)
    print(reqIdDic)
    print("- - - - - [Histogram] - - - - -")
    for i in word_hist:
        print(i)

class Paragraph_DS:
    def __init__(self,text,ilvl=0,ccff=None,name=None,reqId=None,affCell=None):
        self.reqId= reqId
        self.text = text
        self.ilvl = ilvl
        self.ccff = ccff
        self.name = name
        self.tree = None
        self.t_tree = None
        self.twitter_pos()
        self.kkma_pos()
        self.violation_flag = 0
        self.affCell = affCell
        # self.violation_flag = passive_check(self.t_tree)
        # if self.violation_flag == 0 :
        #     self.violation_flag = passive_check(self.tree)
    def kkma_pos(self):
        kkma=Kkma()
        try :
            self.tree = kkma.pos(self.text)
        except:
            if self.tree == None :
                self.tree = []
    def twitter_pos(self):
        twitter = Twitter()
        try:
            self.t_tree = twitter.pos(self.text,stem=True)
        except:
            pass

class Tbl_DS:
    def __init__(self,tblId,newCell=None,reqId=None,affCell=None):
        self.reqId = reqId
        self.ttlNmbRow = 0
        self.ttlNmbCol = 0
        self.cells = []
        self.tblId = tblId
        self.affCell = affCell
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
    def InitReqId(self):
        for i in range(len(self.cells)):
            cell_t = self.cells[i].get('cell')
            if cell_t.prgrphs[0].text == '식 별 자':
                row_t = self.cells[i].get('row')
                col_t = self.cells[i].get('col')
                reqId_cell = self.find_cell(row_t,col_t+1)
                self.reqId = reqId_cell.prgrphs[0].text
                for j in self.cells:
                    j.get('cell').InitReqId()
                return reqId_cell.prgrphs[0].text
        if self.affCell != None:
            self.reqId = self.affCell.reqId


class Cell_DS:
    def __init__(self,affTbl,row,col,newPrgrph=None,newTbl=None,reqId=None):
        self.reqId = reqId
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
    def InitReqId(self):
        self.reqId = self.affTbl.reqId
        for i in self.prgrphs:
            i.reqId = self.reqId
        for i in self.tbls:
            i.reqId.InitReqId()

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
                                temp_inst_cell.insert_prgrph(Paragraph_DS(block.text,ccff=ccff,ilvl=ilvl_val,affCell=temp_inst_cell))
                                #print ('__%s ||%s||'%(str(ilvl_val),block.text))
                            else :
                                temp_inst_cell.insert_prgrph(Paragraph_DS(block.text,ilvl=ilvl_val,affCell=temp_inst_cell))
                                #print ('__%s %s'%(str(ilvl_val),block.text))
                        elif "Table" in str(type(block)):
                            sv_tblId += 1
                            temp_inst_tbl = Tbl_DS(sv_tblId,affCell=temp_inst_cell)
                            temp_inst_cell.insert_tbl(temp_inst_tbl)
                            table_parsing(root_tbl,temp_inst_tbl)

def print_out_srs(srs):
    for i in range(len(srs)):
        if isinstance(srs[i],Paragraph_DS) :
            for k in range(srs[i].ilvl):
                if k != (srs[i].ilvl)-1 : print("     ",end="")
                else : print("   L__",end="")
            if srs[i].ccff == 1 :
                print ('__%s ||%s||'%(str(srs[i].ilvl),srs[i].text))
                if req_printout == 1:
                    print('_%s_'%(srs[i].reqId))
                if pos_printout == 1:
                    print ('_%s_'%(srs[i].tree))
                    print ('_%s_'%(srs[i].t_tree))
                if srs[i].violation_flag == 1 :
                    print('L__[Rule Violation!!]')
            else :
                print ('__%s %s'%(str(srs[i].ilvl),srs[i].text))
                if req_printout == 1:
                    print('_%s_'%(srs[i].reqId))
                if pos_printout == 1:
                    print ('_%s_'%(srs[i].tree))
                    print ('_%s_'%(srs[i].t_tree))
                if srs[i].violation_flag == 1 :
                    print('L__[Rule Violation!!]')
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
                print ('__%s ||%s||'%(str(table.cells[i].get('cell').prgrphs[j].ilvl),table.cells[i].get('cell').prgrphs[j].text))
                if req_printout == 1:
                    print('_%s_'%(table.cells[i].get('cell').prgrphs[j].reqId))
                if pos_printout == 1:
                    print ('_%s_'%(table.cells[i].get('cell').prgrphs[j].tree))
                    print ('_%s_'%(table.cells[i].get('cell').prgrphs[j].t_tree))
                if table.cells[i].get('cell').prgrphs[j].violation_flag == 1 :
                    print('L__[Rule Violation!!]')
            else :
                print ('__%s %s'%(str(table.cells[i].get('cell').prgrphs[j].ilvl),table.cells[i].get('cell').prgrphs[j].text))
                if req_printout == 1:
                    print('_%s_'%(table.cells[i].get('cell').prgrphs[j].reqId))
                if pos_printout == 1:
                    print ('_%s_'%(table.cells[i].get('cell').prgrphs[j].tree))
                    print ('_%s_'%(table.cells[i].get('cell').prgrphs[j].t_tree))
                if table.cells[i].get('cell').prgrphs[j].violation_flag == 1 :
                    print('L__[Rule Violation!!]')
        for j in range(len(table.cells[i].get('cell').tbls)):
            print_out_table(table.cells[i].get('cell').tbls[j])

def Init_SRS_ReqId(tbl_list):
    for i in tbl_list:
        i.InitReqId()

def srs_parsing():
    global sv_tblId
    srs = [] # all srs block list
    tbl_list = [] # list only for table block
    for block,root, ilvl_val,ccff in iter_block_items(doc): #ccff : csell_color_filled_flag
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
    Init_SRS_ReqId(tbl_list)
    #print(srs)
    print_out_srs(srs)
    #makeDic_w2v(srs)
    makeDic(srs)

srs_parsing()

#! /usr/bin/env python
# -*- coding: utf-8 -*-
#

# Copyright 2014 Ηλιάδης Ηλίας
"""
Read Calc's spreadsheets.

Read Calc's spreadsheets. Can be used as a module: it provides the
class LOspreadData which is simply a list of lists initialized with the
contents of the spreadsheet stored in a (typically .ods) file (passed as an
argument).

idea and most of code got from:
http://code.activestate.com/recipes/436066-read-openoffice-spreadsheet-as-list-of-lists-witho/

USAGE (in command line):
odsrearder file.ods
(prints the data in the .ods)

As a loadable class:
The class must receive a full path name of an ods.


"""

#TODO Convert main class to include globals as locals 
#in order to avoid the conflicts with other globals
#TODO Add Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
#TODO Add cell formula

import sys

class ODSReaderError(Exception):
    pass

import xml.parsers.expat
import zipfile

row=[]
cell=u''
rept=u'table:number-columns-repeated'
last_repeat_col=0
incol=False
compact=False
str_strip=False
sheetcounter=0

def copyandtrim(l, trim):
    a = l[:]
    if trim:
        x=range(len(a))
        x.reverse()
        for i in x:
            if a[i]=="":
                del a[i]
            else:
                break
    return a

# 3 handler functions
def start_element(name, attrs):
    global row, cell, rept, last_repeat_col, incol, compact, sheet, sheetcounter
    if name=="table:table":
        sheets[str(sheetcounter)] = {"name" : attrs['table:name'],"sheetdata":[]}
        sheetcounter += 1
        if sheetcounter>0:#aka >0
            #print sheets[str(sheetcounter-1)]
            pass
        return
    if name!="table:table-cell":
        return
    if incol:#not sure about what it does
        raise ODSReaderError("double cell start")
    incol=True
    cell=u""
    if attrs.has_key(rept):
        last_repeat_col = int(attrs[rept])
    else:
        last_repeat_col = 0

def end_element(name):
    global row, cell, rept, last_repeat_col, incol, compact, str_strip, sheet, sheetcounter
    if name=="table:table-cell":
        if not incol:#not sure about what it does
            raise ODSReaderError("double cell end")
        incol=False
        # add the contents to the row
        if str_strip:
            row.append(cell.strip())
        else:
            row.append(cell)
        # print "append to row %d, col %d : %s" % (len(tabla),len(row),cell)
        # manage the repeater
        if last_repeat_col > 1:
            row.extend([cell]*(last_repeat_col-1))
        
    elif name=="table:table-row":
        l = copyandtrim(row,compact)
        if l == []:
            row = []            
            sheets[str(sheetcounter-1)]["sheetdata"].append(l)
            #print 'row',row
            return        
        sheets[str(sheetcounter-1)]["sheetdata"].append(l)
        row = []

def char_data(data):
    global row, cell, rept, last_repeat_col, incol
    if incol:
        cell += data


def read_and_parse(inFileName):
    p = xml.parsers.expat.ParserCreate("UTF-8")
    p.StartElementHandler = start_element
    p.EndElementHandler = end_element
    p.CharacterDataHandler = char_data
    zf = zipfile.ZipFile(inFileName, "r")
    all = zf.read("content.xml")
    # Start the parse.
    p.returns_unicode=1
    p.Parse(all)    
    zf.close()    


class LOspreadData(list):
    """LOspreadData: a=LOspreadData("file",trim=True,strip=False)

the class LOspreadData which is simply a list of lists initialized with the
contents of the spreadsheet stored in a (typically .ods) file (passed as an
argument). Note: there is no validity analysis on the data.
Garbage in, garbage out, or unexepected execptions.
For now it lists only the calculated text, not the formulas.

If trim is true, multiple void cell at the end of a row and void rows are
trimmed out; otherwise, all the cells are reported.

If strip is true, every cell content is stripped of blanks.
    """

    def __init__(self, fname,trim=True,strip=False):
        global row, cell, rept, last_repeat_col, incol, compact, str_strip, sheets, sheetcounter        
        row=[]
        sheets = {}
        incol=False
        cell=u''
        last_repeat_col=0
        compact=trim
        str_strip=strip
        # ok, do the hard work
        read_and_parse(fname)
        #print "after read"        
        #print last_repeat_col
        #print sheets

        for val in sheets:
            #print '==sheets=='
            sheets[val]['rows'] = len(sheets[val]["sheetdata"])
            sheets[val]['cols'] = len(max(sheets[val]["sheetdata"]))
            #print  sheets[val]
            #print len(sheets[val]["sheetdata"])
            #print len(max(sheets[val]["sheetdata"]))
        #list.__init__(self, tabla)

    def num_rows(self, sheetname = u'', sheetnum = '0'):
        '''Returns max number of rows for the spcified sheet.
        
Sheet can be accessed either by sheetname (defaults to u'')
or by sheetnumber (defaults to '0')
First checks for sheetname and if the parapemter is empty returns
the rows num of sheetnum 
(which defaults to '0')
        '''
        if sheetname == u'':
            return sheets[sheetnum]['rows']
        for k, v in sheets.items():
            if sheets[k]['name'] == sheetname:
                return sheets[k]['rows']
        #TODO return error or blank?'''

    def num_cols(self, sheetname = u'', sheetnum = '0'):
        '''Returns max number of columns for the spcified sheet.
        
Sheet can be accessed either by sheetname (defaults to u'')
or by sheetnumber (defaults to '0')
First checks for sheetname and if the parapemter is empty returns
the cols num of sheetnum 
(which defaults to '0')
        '''
        if sheetname == u'':
            return sheets[sheetnum]['cols']
        for k, v in sheets.items():
            if sheets[k]['name'] == sheetname:
                return sheets[k]['cols']
        #TODO return error or blank?'''

    def get_row(self, rownum = 0, sheetname = u'', sheetnum = '0'):
        '''Returns a list with cell values (strings) in a specified row.
        
        If trim is true list may contain up to the last nonempty cell thus
        rest of cells up to num_cols must be added manually as empty strings.
        This is done in order to avoid unnecessary empty strings.
        Can be changed or added as a parameter in a later varsion.
        '''
        if sheetname == u'':
            return self._rowlist(sheets[sheetnum])
        for k, v in sheets.items():
            if sheets[k]['name'] == sheetname:
                return self._rowlist(sheets[k])
        #TODO return error or blank?'''



    def _rowlist(self,whichsheet):
        '''Internal. Requires the sheet dict. 
        
        '''
        return whichsheet['sheetdata']


if __name__=="__main__":

    if len(sys.argv)==2:
        loods = LOspreadData(sys.argv[1])
        print 'second sheet cols', loods.num_cols(sheetname =u'Φύλλο1')
        print loods.get_row(1,sheetname =u'Φύλλο1')
    else:
        print >> sys.stderr, "Usage: %s <OO_calc_file>" % sys.argv[0]
        sys.exit(1)
        print "====exiting==="
        for l in loods:
            a = ['"%s"' % i for i in l]
            print ",".join(a)
        print '===exited=='
    sys.exit(0)

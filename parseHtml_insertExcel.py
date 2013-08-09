#! /usr/bin/env python
#coding:utf-8
import urllib
import httplib
import time
import re
from win32com.client import constants, Dispatch
from BeautifulSoup import BeautifulSoup



class EasyExcel:
    def __init__(self, filename=None):
        self.xlApp = Dispatch('Excel.Application')
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            print "please input the filename"

    def close(self):
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp
     
     
    def getCell(self, sheet, row, col):
        #"Get value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value
   
    def getRange(self, sheet, row1, col1, row2, col2):
        #"return a 2d array (i.e. tuple of tuples)"
        sht = self.xlApp.Worksheets(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value
    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()
    def setCell(self, sheet, row, col, value):
        #"set value of one cell"
        self.xlBook.Worksheets(sheet).Cells(row, col).Value = value
    def getsheetname(self,sheet):

        return self.xlBook.Worksheets(sheet).name
    def changsheetname(self,sheet,name):

        self.xlBook.Worksheets(sheet).name=name
        
        
def parseTop500():
    urlbase = r"http://www.fortunechina.com/china500/"
    xls = EasyExcel(r'D:\fortune.xls')
 
    for num in range(1,100,1):
        
        url = urlbase+str(num)+r"/2013"
        print url
        #url = r"http://www.fortunechina.com/china500/1/2013/ZHONG-GUO-SHI-YOU-HUA-GONG-GU-FEN-YOU-XIAN-GONG-SI"
        try:
            page = urllib.urlopen(url)
            body = page.read() 
            page.close()
        except Exception:
            print "catch" + url
            continue
        
        try:

            
            rank_pattern = re.compile(r'排名第.*?位')
            rank_res = rank_pattern.findall(body)
            if rank_res is not None:
                print rank_res[0]
                text = rank_res[0]
                rank = text.lstrip(r"排名第").rstrip(r"位")
                print rank 
            else:
                print url
            
            name_pattern = re.compile(r'<title>.*?中国500')
            name_res = name_pattern.findall(body)
            if name_res is not None:
                print name_res[0]
                text = name_res[0]
                text = text.lstrip("<title>").rstrip(" (中国500")
                res = text.split(" ")
                ch_name =  res[0]
        
                eng_name = text.lstrip(ch_name).lstrip(" ")
            else:
                print url
                 
            ceo_name_pattern = re.compile(r'董事长.<.strong><.span>.*?\n.*?\n.*?\n.*?</span>')
            ceo_name_res = ceo_name_pattern.findall(body)
            if ceo_name_res is not None:
                text =  ceo_name_res[0]
                res = text.split("\n")
                print res[-2]
                ceo_name =  res[-2].lstrip(' ').rstrip(" ")
                #ceo_name = ceo_name_res[0][8:-7]
            else:
                print url
            
            emp_count_pattern = re.compile(r"员工数.*?\n.*?\n.*?\n.*?\n.*?<")
            emp_count_res = emp_count_pattern.findall(body)
            if emp_count_res != []:
                text = emp_count_res[0]
                res = text.split("\n")
    
                emp_count = res[-2].lstrip(" ").rstrip(" ")
            else:
                print "error: emp " + url
    
            pro_value_pattern = re.compile(r'资产</strong><.span>.*?\n.*?\n.*?txt-14">\n.*?\n')
            pro_value_res = pro_value_pattern.findall(body)
            if pro_value_res is not None:
                #pro_value = pro_value_res[0][-15:-7]
                text = pro_value_res[0]
                res = text.split("\t")
    
                pro_value =  res[-1].rstrip('\n')  
            else:
                print url
                
            stock_value_pattern = re.compile(r'市值</strong><.span>.*?\n.*?txt-14.>.*?</span>')
            stock_value_res = stock_value_pattern.findall(body)
            if stock_value_res is not None:
                text = stock_value_res[0]
    
                res = text.split("<")
                res = res[-2].split(">")
                stock_value =  res[-1]
            else:
                print url
    
            website_pattern = re.compile(r"方网站..*?\n.*?'.target")
            website_res = website_pattern.findall(body)
            if website_res != []:
                #website = website_res[0].lstrip("f='").rstrip("'")
                text = website_res[0]
                res = text.split(" ")
    
                website=  res[-2].lstrip("href='http://").rstrip("'")
            else:
                print "website" +  url
                
            xls.setCell('Sheet1', rank,1,rank.decode('utf-8'))
            xls.setCell('Sheet1', rank,2,ch_name.decode('utf-8'))
            xls.setCell('Sheet1', rank,3,eng_name.decode('utf-8'))
            xls.setCell('Sheet1', rank,4,ceo_name.decode('utf-8'))
            xls.setCell('Sheet1', rank,5,emp_count.decode('utf-8'))
            xls.setCell('Sheet1', rank,6,pro_value.decode('utf-8'))
            xls.setCell('Sheet1', rank,7,stock_value.decode('utf-8'))
            xls.setCell('Sheet1', rank,8,website.decode('utf-8'))        
        except Exception,e:
            print e
            filename = r'D:\grapdata.log'
            f = open(filename, "w")
            f.write(url)
            f.write(body)
            f.write(e.__str__())
            f.close()
    
            continue
        
    xls.save()  
    xls.close()

def parseHtml():
    url = "http://www.fortunechina.com/fortune500/c/2013-07/16/2013C500.htm"
    page = urllib.urlopen(url)
    body = page.read() 
    page.close()
    soup = BeautifulSoup(body)


    tag1 = soup.find('a', href="http://www.fortunechina.com/china500/1/2013")
    

    tag2 = tag1.findParent('td')
    print tag2
    
    tag3 = tag2.findPreviousSibling('td')
    tag = tag3.findPreviousSibling('td')
    top = tag2.findParent('tr')
    print top
    company = top
    while company is not None:
        print company.contents[0].text + company.contents[2].text 
        suburl = company.contents[2].contents[0]['href']
        
        subpage = urllib.urlopen(suburl)
        subbody = subpage.read() 
        subpage.close()
        subsoup = BeautifulSoup(subbody)
        subtop = subsoup.find(text="资产")
        print subtop
        
        next = subtop.findNext('span')
        print next.text.lstrip(" ").rstrip(" ")
        company = company.findNextSibling('tr')
    
def parseStock():

    xls = EasyExcel(r'D:\stock.xlsx')
    
    body = "nothing"
    for i in range(5005,5055,1):
        
        try:
            url = r"http://quotes.money.163.com/f10/gszl_id.html#11a01"
            id =  300001 + i -5005

            print id
            
            url = url.replace("id", str(id))
            
            xls.setCell('Sheet1', i,1,id)   
                 
            page = urllib.urlopen(url)
            print url
            body = page.read() 
            page.close()
            soup = BeautifulSoup(body)
    
    
            tag1 = soup.find('td',text = "公司全称")
            tag_name = tag1.findNext('td').text
            print tag_name
            xls.setCell('Sheet1', i,2,tag_name.decode('utf-8'))
            
            tag1 = soup.find('td',text = "组织形式")
            tag_name = tag1.findNext('td').text
            print tag_name
            xls.setCell('Sheet1', i,3,tag_name.decode('utf-8'))
        
            tag1 = soup.find('td',text = "董事长")
            tag_name = tag1.findNext('td').text
            print tag_name
            xls.setCell('Sheet1', i,4,tag_name.decode('utf-8'))
            
            tag1 = soup.find('td',text = "职工总数")
            tag_name = tag1.findNext('td').text
            print tag_name
            xls.setCell('Sheet1', i,5,tag_name.decode('utf-8'))
            
            tag1 = soup.find('td',text = "公司网址")
            tag_name = tag1.findNext('td').text
            print tag_name
            xls.setCell('Sheet1', i,6,tag_name.decode('utf-8'))
            xls.save()  
        except Exception:
            filename = r'D:\grapdata.log'
            f = open(filename, "w")
            f.write(str(i))
            f.write(url)
            f.write(body)

            f.close()
    xls.save()  
    xls.close()      

#    tag2 = tag1.findParent('td')
#    print tag2
#    
#    tag3 = tag2.findPreviousSibling('td')
#    tag = tag3.findPreviousSibling('td')
#    top = tag2.findParent('tr')
#    print top
#    company = top
#    while company is not None:
#        print company.contents[0].text + company.contents[2].text 
#        suburl = company.contents[2].contents[0]['href']
#        
#        subpage = urllib.urlopen(suburl)
#        subbody = subpage.read() 
#        subpage.close()
#        subsoup = BeautifulSoup(subbody)
#        subtop = subsoup.find(text="资产")
#        print subtop
#        
#        next = subtop.findNext('span')
#        print next.text.lstrip(" ").rstrip(" ")
#        company = company.findNextSibling('tr')
    
    
        
if __name__ == '__main__':
    parseStock()

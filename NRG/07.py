from Bio import Entrez
import win32com.client as win32
import os
import math
#import datetime
import time
#testing git push

Entrez.email = 'mkoh510@gmail.com'
listedDatabase = 'pubmed'

def verifyExist(location):
	if os.path.isfile(location):
		print("entering: ", location)
		ExcelApp = win32.Dispatch('Excel.Application')
		ExcelApp.Visible = True
		return ExcelApp.Workbooks.Open(location)
def wordAppend(sheet, count, first, last):
	wordString = ''
	word1 = sheet.Cells(count, 1).Value
	word2 = sheet.Cells(count, 2).Value
	word3 = sheet.Cells(count, 3).Value
	if word1: wordString+=str(word1)
	if word2: wordString+=str(" "+word2)
	if word3: wordString+=str(" "+str(math.floor(word3)))
	return wordString.lstrip(' ')
def resultCount(dbName, querString, retmax):
	handle = Entrez.esearch(db = listedDatabase, term = query, retmax = "5000")
	tmpRec = Entrez.read(handle)
	handle.close()
	return tmpRec
def newPage(pageNum, queryString):
		wb.Worksheets.Add(Before = None, After = wb.Sheets(wb.Worksheets.Count))
		wb.Worksheets(pageNum).Name = queryString


try: 
	root = os.getcwd()
	path = os.path.join(root, 'Neuregulin.xlsx')
	
	wb = verifyExist(path)
	ws = wb.Worksheets('PubMed')
	date = str(time.ctime(int(time.time())))
	#print(date)
	ws.Cells(1, 1).Value = 'updated: ' + date

	lastRow = ws.Cells(ws.Rows.Count, "A").End(-4162).Row+1
	header_labels = ('Item', 'Id', 'PubDate', 'ePubDate', 'Source', 'AuthorList', 'LastAuthor', 'Title', 'Volume', 'Issue', 'Pages', 'LangList', 'NlmUniqueID', 'ISSN', 'ESSN', 'PubTypeList', 'RecordStatus', 'PubStatus', 'ArticleIds', 'DOI', 'History', 'medline', 'received', 'revised', 'accepted', 'entrez', 'References', 'HasAbstract', 'PmcRefCount', 'FullJournalName', 'ELocationID', 'SO')
			
	#print(ws.Cells(2, 4))
	for i in range(2, lastRow):
		query = wordAppend(ws, i, 2, lastRow)
		record = resultCount(listedDatabase, query, "40")

		#print(query, record["Count"])
		ws.Cells(i, 4).Value = record["Count"]

		if not query in [sheet.Name for sheet in wb.Sheets]:
			print(query, "isn't on the page list. making new page")
			newPage(i, query)	
			for indx, val in enumerate(header_labels):
				wb.Worksheets(query).Cells(1, indx + 1).Value = val

		tempHandle = Entrez.esummary(db = listedDatabase, id = record["IdList"])
		tempRecord = Entrez.read(tempHandle)
		tempHandle.close()


		sheetRow = 1
		for iRow, val1 in enumerate(tempRecord):
			for jCol, val2 in enumerate(val1):
				wb.Worksheets(query).Cells(iRow+2, jCol+1).Value = str(val1[val2])
		else: 
			pass

except Exception as e: print("exception:", e)
finally:
	ws = None
	wb = None
	xl = None

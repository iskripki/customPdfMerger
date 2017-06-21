'''
Project: Custom PDF Merger
Date: 07.06.2017
Name: Irina Skripkina
E-mail: irina.skripkina@toroaluminum.ca

'''

from appJar import gui
from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
from shutil import copyfile
import os, sys, glob, shutil
import subprocess
import re


listBP = [];
listDC = [];
listGR = [];
listYR = [];
listSR = [];
listPS = [];
allLists1 = [];
allLists2 = [];


#GUI customization goes here:
def press(btn):

	if btn=="Cancel":
		app.stop()
	else:
		strDir1 = "C:\work\\target1" #app.getEntry('target1')
		strDir2 = "C:\work\\target2" #app.getEntry('target2')
		name = app.getEntry('fileName')	
		dest1 = "C:\work\dest1" #app.getEntry('dest')
		dest2 = "C:\work\dest2"
		file_list = [] #unsorted file list

		
		if (len(name) != 8):
			app.setLabel("msg", "Invalid file number, try again.")
			app.clearEntry("fileName")

		

		else:
			#os.chdir(strDir1) #navigate to target direcotry	

			for file in [pdf for pdf in os.listdir(strDir1)      #fill the list with required files
			if ((pdf.endswith(".pdf")) and (name in pdf) and ("WP" not in pdf))]:
				file_list.append(file)

			for file in [pdf for pdf in os.listdir(strDir2)      #fill the list with required files
			if ((pdf.endswith(".pdf")) and (name in pdf) and ("WP" not in pdf))]:
				file_list.append(file)

			if (not file_list):
				app.setLabel("msg", "No files found, try again.")
				app.clearEntry("fileName")

			else:
				merger = PdfFileMerger() #new merging object

				for i in file_list:      #merge pdfs
				#	merger.append(i)
					twoLet = i[0] + i[1]
					if(twoLet.lower() == 'bp'):
						listBP.append(i)
					elif(twoLet.lower() == 'dc'):
						listDC.append(i)
					elif(twoLet.lower() == 'gr'):
						listGR.append(i)
					elif(twoLet.lower() == 'yr'):
						listYR.append(i)
					elif(twoLet.lower() == 'sr'):
						listSR.append(i)
					elif(twoLet.lower() == 'ps'):
						listPS.append(i)
					else:
						print("Error processing files")

				allLists1 = [listBP, listDC, listGR]
				for i in allLists1:
					if(len(i) > 1):
						pageMerge(i, strDir1, name)

				allLists2 = [listYR, listSR, listPS]
				for i in allLists2:
					if(len(i) > 1):
						pageMerge(i, strDir2, name)



				#STARTING OVER HERE
				####################################################################
				listBP.clear()
				listGR.clear()
				listDC.clear()
				listYR.clear()
				listSR.clear()
				listPS.clear()
				file_list.clear()
				allLists1.clear()
				allLists2.clear()



				for file in [pdf for pdf in os.listdir(strDir1)      #fill the list with required files
				if ((pdf.endswith(".pdf")) and ((name in pdf) or ("_UPDATED" in pdf)) and (pdf[10].lower() + pdf[11].lower() != 'pg' ))]:
					file_list.append(file)

				for file in [pdf for pdf in os.listdir(strDir2)      #fill the list with required files
				if ((pdf.endswith(".pdf")) and ((name in pdf) or ("_UPDATED" in pdf)) and (pdf[10].lower() + pdf[11].lower() != 'pg' ))]:
					file_list.append(file)


				for i in file_list:
					twoLet = i[0] + i[1]
					if(twoLet.lower() == 'bp'):
						listBP.append(i)
					elif(twoLet.lower() == 'dc'):
						listDC.append(i)
					elif(twoLet.lower() == 'gr'):
						listGR.append(i)
					elif(twoLet.lower() == 'yr'):
						listYR.append(i)
					elif(twoLet.lower() == 'sr'):
						listSR.append(i)
					elif(twoLet.lower() == 'ps'):
						listPS.append(i)
					else:
						print("Error processing files")

				
				file_list.clear()
				allLists1 = [listBP, listDC, listGR, listYR, listSR, listPS]
				for i in allLists1:
					if(len(i) > 1):
						for j in i:
							if("_UPDATED" in j):
								file_list.append(j)
					if(len(i) == 1):
						file_list.append(i[0])


				sortedList1=[]	
				sortedList2=[]
				os.chdir(strDir1)
				for i in file_list:
					if("bp" in i.lower()):
						sortedList1.append(i)
				for i in file_list:
					if("dc" in i.lower()):
						sortedList1.append(i)
				for i in file_list:
					if("gr" in i.lower()):
						sortedList1.append(i)				

				finalPDF = PdfFileMerger()
				for i in sortedList1:
					shutil.copy2(i, dest2)
					finalPDF.append(i)




				os.chdir(strDir2)
				for i in file_list:
					if("yr" in i.lower()):
						sortedList2.append(i)
				for i in file_list:
					if("sr" in i.lower()):
						sortedList2.append(i)
				for i in file_list:
					if("ps" in i.lower()):
						sortedList2.append(i)

				for i in sortedList2:
					shutil.copy2(i, dest2)
					finalPDF.append(i)



				finalPDF.write(dest1 + "\\" +"WP" + name + ".pdf") #create new WP file with all pdfs
				finalPDF.close()

				#clear all the lists for future use

				listBP.clear()
				listGR.clear()
				listDC.clear()
				listYR.clear()
				listSR.clear()
				listPS.clear()
				file_list.clear()
				allLists1.clear()
				sortedList1.clear()
				sortedList2.clear()


				app.setLabel("msg", "Success.")

				for pdf in glob.glob("*_UPDATED.pdf"):
					os.remove(pdf)

				os.chdir(strDir1)
				for pdf in glob.glob("*_UPDATED.pdf"):
					os.remove(pdf)


 
def pageMerge(listPG, strDir, pNum):

	os.chdir(strDir)
	newPDF = PdfFileMerger()

	
	pgsInserted = -1			#for shifting purposes
	for i in listPG:
		if ("PG" not in i.upper()):
			newPDF.append(i) #base
			listPG.remove(i)
	for i in listPG:
		pageNum = int(i[12] + i[13])
		
		if(pageNum == 0):
			newPDF.merge(0, i)
		elif(pageNum == 99):
			newPDF.append(i)
		else:			
			newPDF.merge(pageNum + pgsInserted, i)
		
		pgsInserted += 1

	newPDF.write((i[0] + i[1]).upper() + pNum + "_UPDATED.pdf")
	newPDF.close()



app = gui("Custom PDF Merger", "700x200")
app.addLabel("fileName", "Enter package number to merge (Ex: 12345678):", 3, 0)              # Row 2,CoGRmn 0
app.addEntry("fileName", 3, 1)                     # Row 2,CoGRmn 1
app.addButtons(["Merge", "Cancel"], press, 5, 0, 2) # Row 3,CoGRmn 0,Span 2
app.addLabel("msg", "", 6, 0)             # Row 1,CoGRmn 0
app.go()







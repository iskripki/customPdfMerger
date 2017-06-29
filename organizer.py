


'''
Project: Custom PDF Merger
Date: 07.06.2017
Name: Irina Skripkina
E-mail: irina.skripkina@toroaluminum.ca

'''

from appJar import gui
from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
from shutil import copyfile
from subprocess import check_output
import os, sys, glob, shutil
import subprocess
import re


listGR = [];
listVR = [];
listPS = [];
listBP = [];

allLists1 = [];
allLists2 = [];  



#GUI customization goes here:
def press(btn):

	if btn=="Cancel":
		app.stop()
	else:

		tempDir = r"C:\\work\test\tempp"
		if not os.path.exists(tempDir):
			os.mkdir(tempDir)

		strDir1 = r"C:\\work\test\source1"			#r'\\tal-fs01\\GalzingReports\\'
		strDir2 = r"C:\\work\test\source2"			#users directory
		
		name = app.getEntry('fileName')	        
		dest1 = r"C:\work\test\dest1" #app.getEntry('dest')
		dest2 = r"C:\work\test\dest2"
		file_list = [] #unsorted file list

		
		if (len(name) != 8):
			app.setLabel("msg", "Invalid file number, try again.")
			app.clearEntry("fileName")		

		else:
			
			os.chdir(strDir1) #navigate to target direcotry	

			for file in [pdf for pdf in os.listdir(strDir1)      #fill the list with required files
			if ((pdf.endswith(".pdf")) and (name in pdf))]:
				file_list.append(file)
				shutil.copy2(file, tempDir)
			
			os.chdir(strDir2)
			for file in [pdf for pdf in os.listdir(strDir2)      #fill the list with required files
			if ((pdf.endswith(".pdf")) and (name in pdf))]:
				file_list.append(file)
				shutil.copy2(file, tempDir)
			
			
			if (not file_list):
				app.setLabel("msg", "No files found, try again.")
				app.clearEntry("fileName")

			else:
				
				
				os.chdir(tempDir)
				for i in file_list:      
				
					twoLet = i[0] + i[1]
					if(twoLet.lower() == 'gr'):
						listGR.append(i)
					elif(twoLet.lower() == 'vr'):
						listVR.append(i)
					elif(twoLet.lower() == 'ps'):
						listPS.append(i)
					elif(twoLet.lower() == 'bp'):
						listBP.append(i)				
					else:
						print("Error processing files")

				allLists1 = [listGR, listVR, listPS, listBP]
				for i in allLists1:
					if(len(i) > 1):
						pageMerge(i, tempDir, name)					


				#STARTING OVER HERE
				####################################################################
				

				for i in allLists1:
					i.clear()				
				file_list.clear()
				allLists1.clear()				


				for file in [pdf for pdf in os.listdir(tempDir)      #fill the list with required files
				if ((pdf.endswith(".pdf")) and (name in pdf))]:
					file_list.append(file)				
				

				for i in file_list:
					twoLet = i[0] + i[1]
					if(twoLet.lower() == 'gr'):
						listGR.append(i)
					elif(twoLet.lower() == 'vr'):
						listVR.append(i)
					elif(twoLet.lower() == 'ps'):
						listPS.append(i)
					elif(twoLet.lower() == 'bp'):
						listBP.append(i)
					else:
						print("Error processing files")				
				

				sortedList1=[]	
				
				os.chdir(tempDir)
				for i in file_list:
					if("gr" in i.lower()):
						sortedList1.append(i)
				for i in file_list:
					if("vr" in i.lower()):
						sortedList1.append(i)
				for i in file_list:
					if("ps" in i.lower()):
						sortedList1.append(i)
				for i in file_list:
					if("bp" in i.lower()):
						sortedList1.append(i)
							

				finalPDF = PdfFileMerger()
				for i in sortedList1:
					shutil.copy2(i, dest2)
					finalPDF.append(i)	

				

				finalPDF.write(dest1 + "\\" +"WP" + name + ".pdf") #create new WP file with all pdfs
				finalPDF.close()

				#clear all the lists for future use

				listGR.clear()
				listVR.clear()
				listPS.clear()
				listBP.clear()
				file_list.clear()
				allLists1.clear()
				sortedList1.clear()

				app.setLabel("msg", "Success.")	

				for file in [pdf for pdf in os.listdir(tempDir)]:
					os.remove(file)						
				


 
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


	newPDF.write((i[0] + i[1]).upper() + pNum + "new.pdf")
	newPDF.close()

	for file in [pdf for pdf in os.listdir(strDir)      
	if ((pdf.endswith(".pdf")) and (pdf[10].lower() + pdf[11].lower() == 'pg' ))]:
		os.remove(file)

	os.remove((i[0] + i[1]).upper() + pNum + ".pdf")
	os.rename(((i[0] + i[1]).upper() + pNum + "new.pdf"), ((i[0] + i[1]).upper() + pNum + ".pdf"))


app = gui("Custom PDF Merger", "700x200")
app.addLabel("fileName", "Enter package number to merge (Ex: 12345678):", 3, 0)              # Row 2,CoGRmn 0
app.addEntry("fileName", 3, 1)                     # Row 2,CoGRmn 1
app.addButtons(["Merge", "Cancel"], press, 5, 0, 2) # Row 3,CoGRmn 0,Span 2
app.addLabel("msg", "", 6, 0)             # Row 1,CoGRmn 0
app.go()




#7124



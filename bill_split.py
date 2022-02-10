import sys 
import os 
from termcolor import colored
from openpyxl import Workbook, load_workbook


'''GLOBAL VARIABLES'''
excelFile = None 
taxAmount = 0
currentCell = None 
workbook = None

os.system('color')

def Input():
	''' Function takes care of the input part of the program'''

	global excelFile
	global taxAmount
	if len(sys.argv) == 1:
		excelFile = input("Enter the path to the file or just the file name if the file is in the same directory: \n")
		taxAmount = float(input("Enter the tax amount:\n"))
	else: 
		excelFile = sys.argv[1]
		taxAmount = float(sys.argv[2])

def ExcelPart():
	'''Function takes care of the excel part of the program which includes the items, etc''' 
	global workbook
	workbook = load_workbook(filename = excelFile)
	sheet = workbook.active 
	
	
# Code to take the input from the user. It can be in the call itself or the user will be asked explicitly
while(True):
	try: 
		Input()
	except: 
		print(colored("\nERROR: Please make sure that you have entered the right values\n", 'red'))
		sys.argv = [sys.argv[0]]
		continue
	break

ExcelPart()


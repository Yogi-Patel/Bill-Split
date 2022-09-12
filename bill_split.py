import sys 
import os 
from termcolor import colored
from openpyxl import Workbook, load_workbook
from tabulate import tabulate 

'''GLOBAL VARIABLES'''
excelFile = None 
taxAmount = 0
currentCell = None 
workbook = None
main_dict = {}
subTotal = 0
os.system('color')

def Input():
	''' Function takes care of the input part of the program'''

	global excelFile
	global taxAmount
	global saveToFile

	if len(sys.argv) == 1:
		excelFile = input("Enter the path to the file or just the file name if the file is in the same directory: \n")
		taxAmount = float(input("Enter the tax amount:\n"))
	else: 
		excelFile = sys.argv[1]
		taxAmount = sys.argv[2].split("+")
		tax_sum = 0
		for x in taxAmount:
			tax_sum += float(x)

		taxAmount = tax_sum


def ExcelPart():
	'''Function takes care of the excel part of the program which includes the items, etc''' 
	
	global workbook
	global main_dict
	global subTotal

	workbook = load_workbook(filename = excelFile)
	sheet = workbook.active 
	
	
	numberOfEntities = min(len(sheet['A']), len(sheet['B']), len(sheet['C']))

	# Find out all the people who are part of this bill
	people = set()
	for a in sheet['B']:
		codes = set(a.value.split(','))
		codes = [x.strip().lower() for x in codes]
		people.update(codes)

	# Discard items that are not part of the bill
	people.discard("")
	people.discard("Split")
	people.discard("split")


	# Doing the calculations each row at a time 
	for a in range(numberOfEntities):
		codes = sheet["B"+str(a+1)].value.split(',')
		codes = [x.strip().lower() for x in codes]
		
		subTotal = subTotal + float(sheet['C' + str(a+1)].value)

		# Replace the word split with all the people in the bill 
		if "split" in codes or "Split" in codes: 
			codes = list(people)

		# Remove empty string from the list of people 
		if "" in codes:
			try:
				while True:
						codes.remove("")
			except ValueError:
				pass

		num_codes = len(codes)

		# Add items and prices to the corresponding person's total
		for b in codes:
			if b.strip() == "":
				continue
			if b not in main_dict: 
				main_dict[b] = [[], 0]
			itemPriceForPerson = float(sheet["C"+str(a+1)].value) / num_codes 
			main_dict[b][0].append([sheet["A"+str(a+1)].value, itemPriceForPerson])
			main_dict[b][1] = main_dict[b][1] + itemPriceForPerson
		

	
	for a in main_dict.keys():
		print(colored("\nThe total for the person: "+a, "cyan"))
		print(tabulate(main_dict[a][0], headers = ['Item', "Price (For this person)"]))
		print()
		totalForPerson = main_dict[a][1] + main_dict[a][1]/subTotal*taxAmount
		print(colored("Sub-Total: "+str(main_dict[a][1]), "green"))
		print(colored("Tax: "+str(totalForPerson - main_dict[a][1]), "green"))
		print()
		print(colored("Total: "+str(totalForPerson), "green"))
		



	
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




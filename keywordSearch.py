#------------------------------------------------------------------------------
# Author: 	Joel Gurnett
# Desc: 	find standings for keywords
# Date:		October 18, 2020
#------------------------------------------------------------------------------
import getopt, sys
from googleapi import google
from openpyxl import load_workbook
from datetime import date

def main(argv):
	inputfile = ""
	site = ""
	standing = {}
	try: 
		opts, args = getopt.getopt(argv, "hi:", ["ifile="])
	except getopt.GetoptError:
		sys.exit(2)
	for opt, arg in opts:
		if opt == '-h':
			print('you have reached the help menu')
			sys.exit()
		elif opt in ('-i', '--ifile'):
			inputfile = arg
			print('input file is: ' + inputfile)
	
	workbook = load_workbook(filename="sample.xlsx")
	sheet = workbook.active
	site = sheet.title
	
	query = ""
	file = open(inputfile, 'r')
	Lines = file.readlines()
	row = 2
	col = 2
	while sheet.cell(row=1, column=col).value != None:		
		col = col + 1
	sheet.cell(row=row-1, column=col).value = date.today().strftime("%Y-%m-%d") 
	
	query = sheet.cell(row=row, column=1).value
	while query != None:
		numPage = 3
		print("searching for: " + query + "...")
		searchResults = google.search(query, numPage)
		sheet['A'+str(row)] = query
		count = 0
		for result in searchResults:
			count = count + 1
			try:
				if site in result.link:
					page = 1
					position = count
					if count > 9:
						page = int(count / 10) + 1 
						position = count % 10
					sheet.cell(row=row, column=col).value = count
					standing[query] = "Page: " + str(page) + " position: " + str(position)
					print(standing[query])
					print()
					break
			except:
				print("error!")
		if query not in standing:
			sheet.cell(row=row, column=col).value = 'Not Found'
			print('Query: ' + query + " not found!\n")
		row = row + 1
		query = sheet.cell(row=row, column=1).value

	workbook.save(filename="sample.xlsx")


if __name__ == "__main__":
	main(sys.argv[1:])

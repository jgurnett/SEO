# SEO
To keep track of SEO rankings

## Dependencies 
* openpyxl
    * `pip install openpyxl`
* googleapi
    * `pip install git+https://github.com/abenassi/Google-Search-API`

## Terminology
* WorkBook - The Excel file
* Sheet - the WorkSheet within the Excel file

## Set Up
You need to have your excel file set up correctly for this to work
* Name the Sheet your target website (e.g. example.com)
* Put your keywords in Column A starting from row 2
* use a dash (-) to skip a line to add an empty row

## Usage
* Run the program
	* `./keywordSearch.py -i filename.xlsx`
* show help menu
	* `./keywordSearch.py -h`


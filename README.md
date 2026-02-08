# Automating-consolidation-and-data-downloads
Automating the download of monthly scheme portfolio Axis banks's mutual fund details and consolidating them into single sheet  
READ ME
Axis Bank MF automation :
By providing month (MM) and year (YY) as input, this project automates the download of the monthly portfolio from Axis bank’s website
Consolidate them into a single ‘cosolidated excel’ file
Data model :
By providing the month and for the download, this project fetch the corresponding file from the website url. And ouputs the downloaded monthly portfolio file and consolidated excel sheet containing following details
•	AMC Name        
•	Scheme Name     
•	Instrument Name 
•	ISIN            
•	Market Value
•	pct Portfolio   
•	Reporting Date  
•	Instrument Type 

Assumptions:
•	Since Axis bank’s website doesn’t have any hidden api/JSON endpoint, simple request function is used to fetch data.
•	The automation has been done exlcusively for the axis bank’s monthly portfolio, whereas other AMC’s portfolios can also be added.
•	If other AMC’s websites contain JSON endpoint/java script rendering selenium based navigation could be used.
•	The monthly portfolio of Axis MF excel file uses uniform formatting across sheets
•	File date varies between 28–31 → auto-detected
•	Assuming every equit/debt instrument has a certain list keywords on their instrument name
•	It is assumed that in the monthly portfolio file of axis bank’s MF portfolio every instrument are not in bold sentence

Automation approach:
•	Try multiple date formats to locate valid portfolio file
•	Download Excel automatically from the url
•	Loop through all scheme sheets
•	Extract clean instrument rows, and their corresponding ISIN, market share, percentage portfolio
•	Inject AMC name column
•	Classify instruments by type
•	Consolidate into a single DataFrame
•	Export to Excel

Tools Used:
•	requests — file download
•	openpyxl — Excel parsing
•	pandas — data processing

Steps to run the code:
Use a local python IDE to run the code 
Place the .py in a new folder, open that folder in python compiler
Install dependencies (if not already installed) :
Python -m pip install pandas openpyxl requests 
Run the python code in a new folder
Give proper valid inputs for Month and year according MM and YY format respectively 

Output: 
Downloaded monthly portfolio excel sheet and a consolidated excel sheet containing all the instruments details, in the same folder that u chose.



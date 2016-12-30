import win32com.client as win32

# Opens xls, deletes all irrelevant columns in ALL sheets (22), saves as xlsx (there is a problem with that that)

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False

for i in range(2007,2004,-1):

	wb = excel.Workbooks.Open(r'C:\Users\Konstantinos\Documents\GIT\Python Projects\OU Line Data Testing\files\all-euro-data-%d-%d.xls' % (i, i+1) )

	to_keep = ['Div', 'Date', 'HomeTeam', 'AwayTeam', 'FTHG', 'FTAG', 'B365H', 'B365A', 'BbAv>2.5', 'BbAv<2.5']

	for n in range(22) :

		ws = wb.Sheets(n+1)

		counter = 1

		for r in range(75):

			print(ws.Cells(1, counter).Value, counter)

			if ws.Cells(1, counter).Value not in to_keep:

				print(ws.Cells(1, counter).Value, counter, "DEL")

				ws.Columns(counter).Delete()

			else:
				counter += 1




	wb.SaveAs(r'C:\Users\Konstantinos\Documents\GIT\Python Projects\OU Line Data Testing\files\all-euro-data-%d-%d-FIXED.xlsx' % (i, i+1), FileFormat=win32.constants.xlOpenXMLWorkbook)		# FileFormat is needed to convert file to .xlsx
	excel.Application.Quit()

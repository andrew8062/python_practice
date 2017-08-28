from win32com.client import DispatchEx

class Excel_Comparison:
	def __init__(self, file_new, file_old ):
		self.file_new = file_new
		self.file_old = file_old
	def compare(self):
		self.excel = DispatchEx('Excel.Application')
		self.excel.Visible = False
		print'excle initial setup'
		print 'new file ',self.file_new
		print 'old file ',self.file_old 
		wbB=self.excel.Workbooks.Open(self.file_old)
		wbA=self.excel.Workbooks.Open(self.file_new)
		print 'open excels'
		wsA = wbA.worksheets("Sheet0")
		wsB = wbB.worksheets("Sheet0")
		wsA_used_row = wsA.usedrange.rows.count
		wsB_used_row = wsB.usedrange.rows.count
		#only compare first 9 columns
		used_col = 9
		print 'initialize variables'
		print 'start comparing'
		for i in xrange(1, wsA_used_row+1):
			for j in xrange(1, used_col+1):
				#if bits number is different
				if wsA.Cells(i,j).Value != wsB.Cells(i,j).Value:
					wsA.Cells(i,j).Interior.Colorindex = 4
		print 'end comparing'
		wbB.Close(False)
		self.excel.Visible = True
	def quit(self):
		self.excel.Appkication.Quit()

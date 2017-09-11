from win32com.client import DispatchEx
import logging

class Excel_Comparison:
	def __init__(self, file_new, file_old ):
		logging.basicConfig(level = logging.INFO)
		self.logger = logging.getLogger(__name__)
		self.logger.info('excel comparing start')
		self.file_new = file_new
		self.file_old = file_old
	def compare(self):
		self.excel = DispatchEx('Excel.Application')
		self.excel.Visible = False
		self.logger.info('excel initial setup')
		
		self.logger.info('new file '+self.file_new)
		self.logger.info('old file '+self.file_old)
		wbB=self.excel.Workbooks.Open(self.file_old)
		wbA=self.excel.Workbooks.Open(self.file_new)
		self.logger.info('open excels')
		wsA = wbA.worksheets("Sheet0")
		wsB = wbB.worksheets("Sheet0")
		wsA_used_row = wsA.usedrange.rows.count
		wsB_used_row = wsB.usedrange.rows.count
		#only compare first 9 columns
		used_col = 9
		self.logger.info('initialize variables')
		self.logger.info('start comparing')
		for i in xrange(1, wsA_used_row+1):
			for j in xrange(1, used_col+1):
				#if bits number is different
				if wsA.Cells(i,j).Value != wsB.Cells(i,j).Value:
					wsA.Cells(i,j).Interior.Colorindex = 4
		self.logger.info('end comparing')
		wbB.Close(False)
		self.excel.Visible = True
	def quit(self):
		self.excel.Appkication.Quit()

import xlrd 
import xlwt
import pyexcel

class Excel:
	"""Класс который делает список из файли xls
	
	>>>a = Excel('./excel/каталог_9.xls', code=0, producer=1, articul=2, photo=3)
	>>>b = Excel('./excel/фото_с_сервера.xls', code=0, producer=1, photo=3)

	"""
	def __init__(self, file_name, **kwargs):
		self.file = file_name
		self.name_column = kwargs
		self.common_list = None

	def create_list(self):
		rb = xlrd.open_workbook(self.file, formatting_info=True)
		sheet = rb.sheet_by_index(0)
		count = 0
		common_list = []
		for rownum in range(sheet.nrows):
		    row = sheet.row_values(rownum)
		    if self.name_column.get('code'):
		    	row[self.name_column['code']] = str(row[self.name_column['code']]).split('.')[0]
		    common_list.append(row)

		self.common_list = common_list
		return common_list

	def write_excel(self, file):
		file_for_writer = xlwt.Workbook()
		add_list_with_table = file_for_writer.add_sheet('Result')
		row = 0
		coll = 0
		while row < len(self.common_list):
			while coll < len(self.common_list[row]):
				add_list_with_table.write(row, coll, self.common_list[row][coll])
				coll += 1
			coll = 0  
			row +=1
		file_for_writer.save(file)

	
	def create_list_pyexcel(self):
		""" Создает массив

		Метож работает с помощью библиотеки pyexcel. 
		Форматы:
			-xls
			-xlsx
		"""
		self.common_list = pyexcel.get_array(file_name=self.file)
		return self.common_list

	def write_excel_pyexcel(self, file):
		""" Сохроняет в новый excel

		Форматы:
			-xls
			-xlsx
		"""

		pyexcel.save_as(array=self.common_list, dest_file_name=file)

import datetime
import os.path
import sys
import random
import time
import re 
import optparse
import imaplib
import email
import base64
import argparse

sys.path.append("/home/ivan/Документы/local/python/excel")
from Excel import Excel
		
class FolderException(FileNotFoundError): pass

class Core:
	"""Общие данные для всех imap
	"""

	def __init__(self, server, port, login, password):
		self.server = server
		self.port = port
		self.login = login
		self.password = password



class MailServer(Core):
	"""Сощдаем обэект для работы с mail.ru
	"""

	def __init__(self, server, port, login, password):
		super(MailServer, self).__init__(server, port, login, password)
		self.connector  = None
	
	def connect_server(self):
		"""
		Подключение к почте
		"""
		connector = imaplib.IMAP4_SSL(host=self.server, port=self.port)
		connector.login(self.login, self.password)
		self.connector = connector
		return connector





	def input_massage(self): 
		"""Функция ищет входящие сообщения
		"""

		#Выбираем ящик входящие сообщения
		status, select_data = self.connector.select('INBOX')
		
		# Общее количество сообщений в ящике
		nmessages = select_data[0].decode('utf-8')

		# Выполняет поиск и возвращает UID писем
		result, data = self.connector.uid('search', None, "(UNSEEN)")

		latest_email_uid = data[0].split()

		need_email_id = []
		for massege in latest_email_uid:
			result, data = self.connector.uid('fetch', massege, '(RFC822)')
			raw_email = data[0][1].decode('utf-8')
			email_message = email.message_from_string(raw_email)
			date_time = [*email.utils.parsedate(email_message['Date'])][:-2]

			name_output = list(email.utils.parseaddr(email_message['From']))
			if name_output[0].find('UTF-8?B') > 0:
				name_output[0] = self.decode_msg(name_output[0])
			print("Письмо от имя: {0[0]}, почта: {0[1]}".format(name_output))
			print("Дата и время прибытия: ", datetime.datetime(*date_time))
			if email.utils.parseaddr(email_message['From'])[1] == 'ivan_1995i@mail.ru':
				need_email_id.append(data[0][1])

		return need_email_id
		






	def save_attachment(self, msg, download_folder):
		"""Скачиваем приложение в письме

		Функция юерет из сообщение приложения
		и скачивает их в установленную папку
		"""

		for attech in msg:
			mail = email.message_from_bytes(attech)
			if mail.is_multipart():
			    for part in mail.walk():
			        content_type = part.get_content_type()
			        filename = part.get_filename()
			        if filename:
			            # Нам плохого не надо, в письме может быть всякое барахло
			            download_folder = download_folder if download_folder[-1] == '/' else download_folder + '/'
			            print('Старое имя файла в письме:', filename)
			            try:
			            	if not os.path.exists(download_folder):
			            		raise FolderException()
			            	name, exe = filename.split('.')
			            except FolderException as e:
			            	print('Не сущенствует такого пути или файла: %s'%download_folder)
			            	break
			            except ValueError as err:
			            	print('Было присьмо с кирилицой')
			            	filename = self.decode_msg(filename)
			            	name, exe = filename.split('.')
			            	print('Имя после декодирования: ', filename)
			            # name, exe = filename.split('.')
			            random_int = str(random.randint(0, 10000))
			            all_filename = "{}_{}.{}".format(name, random_int, exe)
			            with open(download_folder + all_filename, 'wb') as new_file:
			                new_file.write(part.get_payload(decode=True))
			                print("Приложение успешно загружено ", all_filename)

		
	def decode_msg(self, msg):
		"""Декодирует кирилицу

			Метод переводит из кодировки base64 в UTF-8
		"""
		rm_utf = msg.replace('UTF-8?B', '')
		decode = base64.b64decode(rm_utf).decode("UTF-8")
		return decode




	def close_connection(self):
		"""
		Close the connection to the IMAP server
		"""

		self.connector.close()    	


class ExcelStock(Excel):
	""" Класс для обновления наличия

	Класс унаследован от Excel и создан
	для обработки изменения наличия
	"""



	def equal(self, dict_stock):
		""" Внести данные наличия по КОДАМ

		Метод сверяется по кодам основного файла
		и новых поступлени на склад и обновляет у основного 
		файла наличие, если коды совпали.
		"""

		count = 1
		while count < len(self.common_list):
			try:
				row_code = int(self.common_list[count][self.name_column['code']])
				row_stock = int(self.common_list[count][self.name_column['stock']])
				new_stock = dict_stock.get(row_code, 'None')
				if new_stock != 'None':
					print('old', self.common_list[count][self.name_column['stock']])
					if new_stock == 0:
						self.common_list[count][self.name_column['stock']] = 0
					elif 0 < new_stock  < 10:
						self.common_list[count][self.name_column['stock']] = 9
					else:
						self.common_list[count][self.name_column['stock']] = 11
					print('new', self.common_list[count][self.name_column['stock']])
				count += 1
			except ValueError as err:
				print("Code error: %s"%row_code)
				count += 1
				continue


	def create_stock_dict(self):
		"""Создаем словарь

		Создаем словарь из новых постулений
		которые мы скачали из сообщений
		{код: наличие }

		"""

		count = 0
		lenght = len(self.common_list)
		raw_code = False
		my_dict = dict()
		while count != lenght:
			code = self.common_list[count][self.name_column['code']]
			stock = self.common_list[count][self.name_column['stock']]
			if raw_code: 
				if code:
					try:
						my_dict.update({int(code): int(stock)})
					except ValueError as e:
						count += 1
						continue
			elif str(code).lower().strip() == 'код':
				raw_code = True
			count += 1

		return my_dict







def console_managment():
	"""Менеджмент аргументов

	Обрабатывает аргументы с помощью argparse
	"""

	parser = argparse.ArgumentParser("Модуль для подключения к почте\n")

	# nargs='modificator' 
	# Modificators бывают + (не ограниченное число параметров возвращает список)
	# Цифра которая бы говорила сколько он ожидает 
	# ? - необязательный параметр 
	parser.add_argument('-l', '--login', required=True, \
		help='Введите логин от ящика.')
	parser.add_argument('-p', '--password', required=True, \
		help='Введите пароль от ящика.')
	parser.add_argument('-f', '--folder', default='.', \
		help='Нужно ввести полный путь до папки в которой хотите сохранить приложения в письме.')
	parser.add_argument('-s', '--server', required=True, \
		help='Введите imap (протокол прикладного уровня для доступа к электронной почте) сервера.')
	parser.add_argument('--port', default='993', \
		help='Введите порт для подключение через imap.')
	# parser.add_argument('-k', '--file', type=open, help='Введите aфайл для чтения.')

	return parser






def folder_check(folder='./attachments/'):
	"""Проверяет папку

	Файлы не имеющие рассширение xlsx or 
	xls удаляются
	"""
	files_dowanload = os.listdir(folder)
	files_excel = []
	if files_dowanload:
		for file in files_dowanload:
			if file.endswith('.xls') or file.endswith('.xlsx'):
				files_excel.append(file)
			else:
				print(file)
				os.remove(folder + file)
	else:
		return None

	return files_excel


def main():
	"""Запускает манипуляции с excel

	Функция начинает свою работу уже после
	того как скачаются приложение из 
	сообщений. Здесь мы запускаем объекты
	для их обработки

	"""
	list_files = folder_check()
	if list_files:
		for excel in list_files:
			main_excel = ExcelStock(code=0, producer=1, \
				articul=2, stock=10, file_name='./main_excel/каталог_11.xlsx')
			stock_excel = ExcelStock(code=0, producer=1, \
				articul=2, stock=4, file_name='./attachments/' + excel)
			main_excel.create_list_pyexcel()
			stock_excel.create_list_pyexcel()
			dict_stock = stock_excel.create_stock_dict()
			main_excel.equal(dict_stock)
			main_excel.write_excel('./main_excel/каталог_11.xlsx')
			#main_excel.write_excel_pyexcel('./main_excel/Обработанный_каталог.xlsx')
			time.sleep(3)
			os.remove('./attachments/' + excel)
	else:
		print('Папка с файлами пуста. Это может значит', \
			'что нет новых писем на почте или какой-то баг.') 







if __name__ == '__main__':
	parser = console_managment()
	newspace = parser.parse_args()
	try:
		mail = MailServer(newspace.server, newspace.port, newspace.login, newspace.password)
		mail.connect_server()
		file = mail.input_massage()
		mail.save_attachment(file, download_folder=newspace.folder)
	except Exception as err:
		raise(err)
	finally:
		mail.close_connection()
	main()

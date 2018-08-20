import datetime
import os.path
import sys
import time 
import optparse
import imaplib
import email
import argparse





class FolderException(FileNotFoundError): pass

class Core:
	def __init__(self, server, port, login, password):
		self.server = server
		self.port = port
		self.login = login
		self.password = password






class MailServer(Core):

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

		# Выбираем ящик входящие сообщения
		status, select_data = self.connector.select('INBOX')
		# Общее количество сообщений в ящике
		nmessages = select_data[0].decode('utf-8')

		# Выполняет поиск и возвращает UID писем
		result, data = self.connector.uid('search', None, "ALL")
		
		# for message in data[0].split(): 
		# 	result, data = self.connector.uid('fetch', message, '(RFC822)')
		# 	raw_email = data[0][1].decode('utf-8')
		# 	email_message = email.message_from_string(raw_email)
		# 	print(email.utils.parseaddr(email_message['From']))
		# 	print(email.utils.parsedate(email_message['Date']))


		latest_email_uid = data[0].split()[-1]
		result, data = self.connector.uid('fetch', latest_email_uid, '(RFC822)')
		raw_email = data[0][1].decode('utf-8')
		
		#Вывод данных отправителя
		email_message = email.message_from_string(raw_email)
		date_time = [*email.utils.parsedate(email_message['Date'])][:-2]
		print("Письмо от имя: {0[0]}, почта: {0[1]}".format(email.utils.parseaddr(email_message['From'])))
		print("Дата и время прибытия: ", datetime.datetime(*date_time))

		return  data[0][1]

	def close_connection(self):
		"""
		Close the connection to the IMAP server
		"""

		self.connector.close()


	def save_attachment(self, msg, download_folder):
		"""Скачиваем приложение в письме

		Функция юерет из сообщение приложения
		и скачивает их в установленную папку
		"""

		mail = email.message_from_bytes(msg)
		if mail.is_multipart():
		    for part in mail.walk():
		        content_type = part.get_content_type()
		        filename = part.get_filename()
		        if filename:
		            # Нам плохого не надо, в письме может быть всякое барахло
		            download_folder = download_folder if download_folder[-1] == '/' else download_folder + '/'
		            try:
		            	if not os.path.exists(download_folder):
		            		raise FolderException()
		            except FolderException as e:
		            	print('Не сущенствует такого пути или файла: %s'%download_folder)
		            	break
		            with open(download_folder + part.get_filename(), 'wb') as new_file:
		                new_file.write(part.get_payload(decode=True))
		                print("Приложение успешно загружено")







def console_managment():
	"""Менеджмент аргументов

	Обрабатывает аргументы с помощью argparse
	"""

	parser = argparse.ArgumentParser("Модуль для подключения к почте\n")

	# nargs='modificator' 
	# Modificators бывают + (не ограниченное число параметров возвращает список)
	# Цифра которая бы говорила сколько он ожидает 
	# ? - необязательный параметр 
	parser.add_argument('-l', '--login', required=True, help='Введите логин от ящика.')
	parser.add_argument('-p', '--password', required=True, help='Введите пароль от ящика.')
	parser.add_argument('-f', '--folder', default='.', help='Нужно ввести полный путь до папки в которой хотите сохранить приложения в письме.')
	parser.add_argument('-s', '--server', required=True, help='Введите imap (протокол прикладного уровня для доступа к электронной почте) сервера.')
	parser.add_argument('--port', default='993', help='Введите порт для подключение через imap.')
	# parser.add_argument('-k', '--file', type=open, help='Введите aфайл для чтения.')

	return parser


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

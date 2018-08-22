import smtplib
from email.mime.text import MIMEText
from email.header import Header


def smtp_send(smtp_host, smtp_port, smtp_login, smtp_password, send_to, message_text):
	msg = MIMEText(message_text, 'plain', 'utf-8')
	msg['Subject'] = Header("Наличие ООО \"Прайм Тулс\"", 'utf-8')
	msg['From'] = smtp_login
	msg["To"] = send_to
	smtpObj = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=10)
	smtpObj.login(smtp_login, smtp_password)
	smtpObj.sendmail(smtp_login, send_to, msg.as_string())
	smtpObj.quit()




if __name__ == '__main__':
	smtp_host = "smtp.mail.ru"
	smtp_port = "465"
	smtp_login = "ivan_1995i@mail.ru"
	smtp_password = "7454308"
	send_to = "stock@kvam.ru"

	message_text = """Добрый день! Прошу отправить актуальные остатки на этот почтовый адрес. 

	Просьба не изменять нумерацию столбиков уникальных кодов и наличия в файле Exel, 
	и отправлять только свежий файл с обновленным наличием. 
	Это связано с корректной работой нашей автоматизированной системы. 
	Спасибо! 

	С уважением, Компания ООО "Прайм Тулс"
	"""
	smtp_send(smtp_host, smtp_port, smtp_login, smtp_password, send_to, message_text)

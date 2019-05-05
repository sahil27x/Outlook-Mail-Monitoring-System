import pythoncom
import win32com.client as win32
import logging
import logging.config
from CheckMailAndPerformAction import checkMail
import traceback
from threading import Timer
import time
import datetime
from BatchJobSummary import getSummaryDetails
from BatchJobSummary import getJobDetails
import email_handler
from ReadingProperties import getProperty
import pdb
import sys 

logFile= getProperty("MasterConfigurationSection","LoggingConfigFile")
logging.config.fileConfig(logFile)
logger = logging.getLogger('MonitorMail')
logger.info('Mailbox Monitoring Started')


def sendBatchJobStatus():
	logger.info('sendBatchJobStatus:: START')
	statusEmailFlag = getProperty("StatusEmails","SendStatusEmailFlag")
	if statusEmailFlag == 'True':
		logger.info('Sending Status Email')
		email_handler.send_status_mail()
	else:
		logger.info('Status Email Flag is False')
	
	logger.info('sendBatchJobStatus:: END')

def ProcessEmails():
	outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
	inbox = outlook.GetDefaultFolder(6)
	folderName = getProperty("MasterConfigurationSection","folderName")
	folder = inbox.Folders[folderName]
	print('Processing Emails...')
	for i in range(1,folder.Items.Count+1):
		mailItem = folder.Items[i]
		if '_MailItem' in repr(mailItem.__class__) and mailItem.UnRead:
			bSuccess = checkMail(mailItem)
			#Mark email as read
			if bSuccess:
				mailItem.UnRead = False
				
	print('Emails Processed...')
	print('Sending Processing Status...')
	sendBatchJobStatus()
	print('Processing Status Sent...')


try:
	ProcessEmails()
	
except Exception as err:
	logger.error('Error Accessing Mailbox')
	logger.error(traceback.format_exc(err))
	todaysDate = datetime.datetime.now().strftime('%m/%d/%y %H:%M:%S')
	email_handler.send_error_mail("System Error","N/A",todaysDate,'Error Accessing Mailbox',traceback.format_exc(err),None, None)

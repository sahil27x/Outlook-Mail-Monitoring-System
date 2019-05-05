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

logFile= getProperty("MasterConfigurationSection","LoggingConfigFile")
logging.config.fileConfig(logFile)
logger = logging.getLogger('MonitorMail')
logger.info('Mailbox Monitoring Started')


class Handler_Class(object):

    def OnNewMailEx(self, receivedItemsIDs):
        # RecrivedItemIDs is a collection of mail IDs separated by a ",".
        # Sometimes more than 1 mail is received at the same moment.
        
		for ID in receivedItemsIDs.split(","):
			mail = outlook.Session.GetItemFromID(ID)
			if '_MailItem' in repr(mail.__class__):
				bSuccess = checkMail(mail)
				#Mark email as read
			if bSuccess:
				mail.UnRead = False
		print("\n**Mailbox monitoring is active, waiting for emails.**")
		

def sendBatchJobStatus():
	logger.info('sendBatchJobStatus:: START')
	statusEmailFlag = getProperty("StatusEmails","SendStatusEmailFlag")
	if statusEmailFlag == 'True':
		logger.info('Sending Status Email')
		email_handler.send_status_mail()
	else:
		logger.info('Status Email Flag is False')
	
	statusEmailFrequency = getProperty("StatusEmails","EmailFrequencyInSeconds")
	t=Timer(float(statusEmailFrequency),sendBatchJobStatus)
	t.start()
	logger.info('sendBatchJobStatus:: END')
		
try:
	
	outlook = win32.DispatchWithEvents("Outlook.Application", Handler_Class)
	
	statusEmailFrequency = getProperty("StatusEmails","EmailFrequencyInSeconds")
	t=Timer(float(statusEmailFrequency),sendBatchJobStatus)
	t.start()	
	print('Mailbox Monitoring Started...')
	#and then an infinit loop that waits from events.
	pythoncom.PumpMessages()
	
except Exception as err:
	logger.error('Error Accessing Mailbox')
	logger.error(traceback.format_exc(err))
	todaysDate = datetime.datetime.now().strftime('%m/%d/%y %H:%M:%S')
	email_handler.send_error_mail("System Error","N/A",todaysDate,'Error Accessing Mailbox',traceback.format_exc(err),None, None)

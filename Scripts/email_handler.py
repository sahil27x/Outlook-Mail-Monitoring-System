import re
import io
import smtplib
import ConfigParser
import logging
import logging.config
import win32com.client as win32
import os
from ReadingProperties import getProperty
from BatchJobSummary import getCurrentCount

logFile= getProperty("MasterConfigurationSection","LoggingConfigFile")
logging.config.fileConfig(logFile)
logger = logging.getLogger('email_handler')

errorEmailTemplatePath=getProperty('EmailTemplatesSection', 'ErrorEmailTemplatePath')
statusEmailTemplatePath=getProperty('EmailTemplatesSection', 'StatusEmailTemplatePath')

def send_email(subject,recipient,body,attachment):
	logger.debug('send_email:: START')
	try:
		import pythoncom
		pythoncom.CoInitialize()
		outlook = win32.Dispatch('outlook.application')
		mail = outlook.CreateItem(0)
		
		if recipient:
			addressToVerify = recipient
		else:
			globalRecipient=getProperty('SMTPSection','ToEmail')
			addressToVerify = globalRecipient
		
		#Loop through all the emails (semi-colon separated) and then check the format for each one of them
		for emailAddress in addressToVerify.split(';'):
			if emailAddress and emailAddress.strip() != '':
				logger.info('send_mail:: Validating the email address format')
				match = re.match('^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$', emailAddress.strip())
				if match == None:
					print('Incorrect email address format. Email not sent.')
					logger.error('send_mail:: incorrect email address format.') 
					return

		mail.To = addressToVerify
		mail.Subject = subject
		mail.HTMLBody = body
		if attachment:
			if os.path.exists(attachment):
				mail.Attachments.Add(attachment)
			else:
				logger.error('send_email:: Attachment not found. Sending without attachment.')
		mail.Send()
	except Exception as error:
		#Send Email via SMTP
		logger.error('send_email:: Failed to send email via Outlook. Sending email via SMTP')
		SMTPHost=getProperty('SMTPSection', 'SMTPHost')
		SMTPPort=getProperty('SMTPSection', 'SMTPPort')
		fromEmail=getProperty('SMTPSection','fromEmail')
		host =  SMTPHost
		port = SMTPPort
		username = fromEmail
		password = '*******'
		email_conn = smtplib.SMTP(host,port)
		email_conn.ehlo()

		email_conn.starttls()

		email_conn.login(username,password)
		email_conn.sendmail(subject,recipient,body,attachment)
		email_conn.quit()		
	logger.debug('send_email:: END')						

def replace_words(templateString, data):
	for key, val in data.items():
		templateString = templateString.replace(key, val)
	return templateString
	
def send_error_mail(appName,jobName, dateTime, error, description, recipient,attachment): 
	logger.debug('send_error_mail:: START')
	errorDetails = {} 
	errorDetails["appName"] = appName
	errorDetails["jobName"] = jobName
	errorDetails["dateTime"] = dateTime
	errorDetails["errorSummary"] = error
	errorDetails["errorDescription"] = description

	#Read the input parameters & replace them in HTML content to set that replaced content as body of the email
	if os.path.exists(errorEmailTemplatePath):
		t = io.open(errorEmailTemplatePath, 'r', encoding='utf-8-sig')
		tempstr = t.read()
		t.close()
		body = replace_words(tempstr, errorDetails)
		send_email("Email Monitoring Error",recipient,body,attachment)
	else:
		logger.error('send_error_mail:: Template File Not Found.')
	
	logger.debug('send_error_mail:: END')
								
def send_status_mail():
	logger.debug('send_status_mail:: START')
	batchJobStatus = {} 
	summaryGraphPath=getProperty('BatchJobSummary','summaryGraphPath')
	summaryGraphFileName=getProperty('BatchJobSummary','summaryGraphFileName')
	includeSummaryGraph = getProperty('BatchJobSummary','includeSummaryGraph')
	
	if includeSummaryGraph=='True':
		if os.path.exists(summaryGraphPath + summaryGraphFileName):
			batchJobStatus["summaryGraphFileName"] = str(summaryGraphFileName)
			batchJobStatus["<summaryGraph>"] = "<img src='cid:summaryGraphFileName'/>"
		else:
			logger.info('send_status_mail:: Summary Graph File Not Found.')
			batchJobStatus["summaryGraphFileName"]= ""
			batchJobStatus["<summaryGraph>"]=""
	else:
		logger.info('send_status_mail:: Include Summary Graph File is switched off in configuration file.')
		batchJobStatus["summaryGraphFileName"]= ""
		batchJobStatus["<summaryGraph>"]=""
		
		
	#Get the job details from the schedule
	jobDetailsList = getCurrentCount()
	includeJobSchedule = getProperty('BatchJobSummary','includeJobSchedule')
	
	if includeJobSchedule=='True' and jobDetailsList is not None and len(jobDetailsList)>0 : 
		#create HTML table for the Job Details and append to the batchJobStatus object
		jobDetailsTable= "<h3>Today's Batch Job Schedule</h3> <table> <tr style='background: #20B2AA;'><td>&nbsp;Application Name&nbsp;</td><td>&nbsp;Job Name&nbsp;</td><td>&nbsp;Expected Run Time &nbsp;</td><td>&nbsp;Actual Run Time&nbsp;</td><td>&nbsp;Job Status&nbsp;</td></tr>"
		for objJobDetails in  jobDetailsList:
			if objJobDetails.slaFlag == "F":
				jobDetailsTable = jobDetailsTable + " <tr style='color: #ff0000;'><td>" + objJobDetails.appName + "</td><td>" + objJobDetails.jobName +  "</td><td>" + objJobDetails.expectedRunTime +  "</td><td>" + objJobDetails.actualRunTime +  "</td><td>" + objJobDetails.jobStatus + "</td></tr>"	
			elif objJobDetails.actualRunTime != "":
				jobDetailsTable = jobDetailsTable +  " <tr style='color: #006400;'><td>" + objJobDetails.appName + "</td><td>" + objJobDetails.jobName +  "</td><td>" + objJobDetails.expectedRunTime +  "</td><td>" + objJobDetails.actualRunTime +  "</td><td>" + objJobDetails.jobStatus + "</td></tr>"
			else:
				jobDetailsTable = jobDetailsTable +  " <tr><td>" + objJobDetails.appName + "</td><td>" + objJobDetails.jobName +  "</td><td>" + objJobDetails.expectedRunTime +  "</td><td>" + objJobDetails.actualRunTime +  "</td><td>" + objJobDetails.jobStatus + "</td></tr>"
		jobDetailsTable += " </table>"
		batchJobStatus["<jobDetailsTable>"] = jobDetailsTable
	else:
		logger.info('send_status_mail:: No Batch Job Schedule Details and Current Details Found')
		batchJobStatus["<jobDetailsTable>"] = ""
		
	#Read the input parameters & replace them in HTML content to set that replaced content as body of the email
	if os.path.exists(statusEmailTemplatePath):
		if (batchJobStatus["<jobDetailsTable>"] is not "" or batchJobStatus["summaryGraphFileName"] is not ""):
			t = io.open(statusEmailTemplatePath, 'r', encoding='utf-8-sig')
			tempstr = t.read()
			t.close()
			body = replace_words(tempstr, batchJobStatus)
			if batchJobStatus["summaryGraphFileName"] is not "":
				send_email("Email Monitoring System Status",None,body,summaryGraphPath + summaryGraphFileName)
			else:
				send_email("Email Monitoring System Status",None,body,None)
		else:
			send_email("Email Monitoring System Status",None,"Mailbox monitoring is active. No batch job schedule available.",None)
	else:
		logger.error('send_status_mail:: Template File Not Found.')
	logger.debug('send_status_mail:: END')

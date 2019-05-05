from bs4 import BeautifulSoup
#import win32com.client as win32
from openpyxl.styles import Alignment
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
import datetime
import time
import os
from ReadingProperties import getProperty 
import logging
import logging.config
import traceback
import sys
import pdb
import unicodedata
import pandas as pd

ll = []

class AlertHandler():
	
	def __init__(self,mail,fileName):
		self.mail=mail
		self.fileName=fileName
		#self.parsehandler()
		
		
		
	def parsehandler(self):
		try:
			
			logger.info('AlertHandler:parsehandler(): Started')
			program_code=""
			test_code=""
			Mail_recieved_time=""
			sub_list = self.mail.Subject.split()
			#print self.mail.SentOn #changed by dileep
			#print self.mail.ReceivedTime #changed by dileep
			Mail_recieved_time = self.mail.ReceivedTime #changed by dileep

			ts1 = sub_list[1]
			ts1 =  ts1.encode('ascii','ignore')
			tc1 = sub_list[4]
			tc1 =  tc1.encode('ascii','ignore')
			mbody = self.mail.Body.encode('ascii', 'ignore').split()
			m_desc = mbody[6:9]
			s_desc = " ".join(m_desc)
			tv='NA'

			if self.mail.Attachments.Count > 0:
				for attachment in self.mail.Attachments:
					if attachment.FileName.endswith(".xml") or attachment.FileName.endswith(".bin"):
						attachment.SaveAsFile(os.getcwd() + '\\' + attachment.FileName)
						handler = open(attachment.FileName,"r")
						contents= handler.read()
						soup = BeautifulSoup(contents,'xml')

						tag = soup.CandidateAttribute
						tag_value = tag['CandidateID']
						tag_value = tag_value.encode('ascii', 'ignore')
						tv = tag_value if len(tag_value)>0 else "NA"

					else:
						print("Invalid Attachment")





			print("this", ts1, tc1, s_desc, Mail_recieved_time,tv)

			df = pd.read_csv("C:\Users\SARORA003\Desktop\mmoutput.csv",dtype={1:'str'})
			#df.dropna(how='all')
			print(str(tc1) in df["Transmission ID"].values)



			if str(tc1) in df["Transmission ID"].values:
				count = df.loc[df['Transmission ID'] == str(tc1), 'Count'].iloc[0] + 1
			else:
				count=1

			print(count)

			if count > 1:
				df.loc[df["Transmission ID"] == tc1, 'Count'] = count
				df.loc[df["Transmission ID"] == tc1, 'Last Mail Time'] = Mail_recieved_time
				df.loc[df["Transmission ID"] == tc1, 'Transmission Status'] = ts1
				df.loc[df["Transmission ID"] == tc1, 'Error'] = s_desc
				df.loc[df["Transmission ID"] == tc1, 'CandidateID'] = tv
				print(df.head())
				df.to_csv("C:\Users\Desktop\mmoutput.csv", index=False)

			else:

				df = df.append({"Count":count,"Transmission ID":tc1,"Transmission Status":ts1,"Error":s_desc,"Last Mail Time":Mail_recieved_time,"CandidateID":tv}, ignore_index=True)
				print(df.head())
				df.to_csv("C:\Users\Desktop\mmoutput.csv",index=False)

	
	

try:
	mailItem = sys.argv[1]
	datetime.datetime.now().strftime('%m/%d/%y')
	currDate=datetime.datetime.now().strftime('%m%d%Y')# variable appends with file name for creating new file with current date
	#currDate=self.mail.SentOn
	fileName=getProperty('Section','fileName')+currDate+".xlsx"
	logger = logging.getLogger('Execution')
	alert = AlertHandler(mailItem,fileName)
	alert.parsehandler()
	
	
except Exception as err:
	logger.error('AlertHandler: Error in running alert handler process')
	logger.error(traceback.format_exc(err))
	

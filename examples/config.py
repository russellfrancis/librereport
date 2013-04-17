#!/usr/bin/env python
# -*- coding: utf-8 -*-
#description     :Config example for LibreReport
#author          :Brandon Tylke (brandon.tylke@lauren.com)
#date            :2013-04-01
#version         :0.0.1
#license         :GPL v3 http://www.gnu.org/licenses/gpl.txt
#python_version  :2.7.1  
#==============================================================================

##########################
#                        #
#   LibreReport Config   #
#                        #
##########################

class config:

    def __init__(self):
        
        # default path to any image files you would want to include in your reports
        self.site_files = ''
        
        # default path to save completed reports
        self.completed_report_path = './completed/'
        
        #default path to where your report python scripts live
        self.code_root = ''
        
        #full path to the DocumentConverter.py script
        self.doc_converter_path = ''
        
        #log report activity (these get extremely large)
        self.savelog = True
        
        #location of log files
        self.log_root = './logs/'
        #Windows default path if unspecified
        #log_root = os.path.join(os.environ['USERPROFILE'],'.librereport', 'logs')
        #Linux default path if unspecified
        #log_root = os.path.join('/var/log/','librereport')
        
        #this is the path to the LibreOffice bundled version of python
        #it is a long-term goal that this will not be necessary 
        self.python_cmd = ''
        
        #database config options
        self.db_server = ''
        self.db_dbname = 'example.db'
        self.db_user = ''
        self.db_password = ''
        self.db_port = ''
        
        #PEP: 249 compatible module for your database
        # set this to None and no default database connection will be established
        # you can create your own connections as needed in your reports
        self.db_module = 'sqlite3'
        
        #email config options
        self.smtp_server = 'smtp.gmail.com'
        self.smtp_port = 587
        self.smtp_username = ''
        self.smtp_password = ''
        self.default_mail_sender = 'your.name@example.org'
        self.default_mail_subject = ''
        self.default_mail_message = ''
        
        #what is the maximum time any report should take to run
        # None to disable
        self.max_report_runtime = 14400 # 4 hours
        
        
        #optional job id if you are using the db to pass parameters from another system 
        self.job_id = ''
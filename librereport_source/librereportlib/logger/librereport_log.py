#!/usr/bin/env python
# -*- coding: utf-8 -*-
#description     :Logs output from LibreReport.
#author          :Brandon Tylke (brandon.tylke@lauren.com)
#date            :2013-04-01
#version         :0.0.1
#license         :GPL v3 http://www.gnu.org/licenses/gpl.txt
#python_version  :2.7.1  
#==============================================================================

import time, os, sys
import errno

# http://www.youtube.com/watch?v=eusMzC7Rx7M
class itslog:
    def __init__(self, savelog=True, log_root=None):

        #we use savelog instead of a dummy function that returns so we can turn logging on and off during a report run
        #if something fails, turn it on, log it and then turn it back off
        self.savelog = savelog

        # If we don't have a default log file path provided, use some defaults.
        if log_root == None:
            if sys.platform == 'win32':
                log_root = os.path.join(os.environ['USERPROFILE'],'.librereport', 'logs')
            else:
                log_root = os.path.join('/var/log/','librereport')

        try:
            os.makedirs(log_root)
        except OSError, exception:
            if exception.errno != errno.EEXIST:
                raise

        if log_root not in sys.path:
            sys.path.insert(0, log_root)

        #log filename
        year = str(time.localtime()[0])
        month = str(time.localtime()[1])
        day = str(time.localtime()[2])
        
        if len(day) == 1:
            day = "0" + day
        
        if len(month) == 1:
            month = "0" + month
        
        now = year + '-' + month + '-' + day

        self.logfile = os.path.join(log_root, now + "_librereportlog.txt")

    def log(self, text, time_stamp=True):
        if self.savelog == True:
            text = str(text)
            if text[-1:] == '\n':
                text = text[:-1]
            
            text = text.replace("\n","\n\t")
            
            if time_stamp:
                text = str(time.ctime()) + '\n\t' + text + '\n'
            
            print text
            
            try:
                openlog=open(self.logfile,'a')
                openlog.write(text)
                openlog.close()
            except:
                pass
    
    def new_session(self):
        if self.savelog == True:
            text = "\n\n\n\n--------------------------------------------\n"
            text = text + str("  LibreReport Started: " + time.ctime() + "\n")
            text = text +  "--------------------------------------------\n"
            
            try:
                text = text + str("  PID: " + str(os.getpid()) + "\n")
                text = text +  "--------------------------------------------\n"
            except StandardError, e:
                print "Get PID Error:", e
                pass
            
            self.log(text, False)

    def start_logging(self):
        self.savelog = True

    def stop_logging(self):
        self.savelog = False

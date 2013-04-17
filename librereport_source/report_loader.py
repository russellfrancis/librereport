#!/usr/bin/env python
# -*- coding: utf-8 -*-
#description     :Utility to setup initial db connection and retrieve report parameters.
#author          :Brandon Tylke (brandon.tylke@lauren.com)
#date            :2013-04-01
#version         :0.0.1
#license         :GPL v3 http://www.gnu.org/licenses/gpl.txt
#python_version  :2.7.1  
#==============================================================================

import os
import sys

# We don't want to use 'from librereportlib import X' all the time,
# so we add the librereportlib directory to the search path
import librereportlib
libpath = librereportlib.__path__[0]
if libpath not in sys.path:
    sys.path.insert(0, libpath)
del librereportlib


# import the sql connection utility
from db_connect import sql_connection

# import the report engine module
from report_engine import report_engine

#It's log!  (For me, that never gets old.)
import logger



if __name__ == "__main__":
    params = {'report_name':None}
    
    if len(sys.argv) < 2:
        print "Enter report path and optionally a report number:"
        report_path = raw_input()
        
        if report_path == '':
            print "USAGE: python %s </path/to/report>" % sys.argv[0]
            sys.exit(255)
        else:
            sys.argv.append(report_path)
    
        print "Enter report number: (optional)"
        report_num = raw_input()
        
        if report_num == '':
            #not required
            pass
        else:
            sys.argv.append(report_num)
    
    
    if str(sys.argv[1]).lower() != 'none':
        try:
            full_report_path = sys.argv[1]
            report_path, report_name = os.path.split(full_report_path)
            
            #get the base report name
            params['report_name'] = report_name.replace('.py','')
            
            #add the report path
            sys.path.insert(0, report_path )
            
        except StandardError, e:
            print "Unable to get report path. Error:", e
            pass
    else:
        full_report_path = None
        
    
    ####################
    # get our settings #
    ####################
    
    #get the config
    from config import config
    conf = config()
    
    
    #start logging
    log = logger.itslog(conf.savelog, conf.log_root)
    log.new_session()
    
    
    if len(sys.argv) > 2:
        conf.job_id = sys.argv[2]
    else:
        conf.job_id = None
    
    #import the db connection
    sql = None
    if conf.db_module != None:
        sql = sql_connection(conf.db_module, conf.db_dbname)
    
    #get parameters from the db connection
    if sql != None and conf.job_id != None:
        params = sql.get_job_details(conf.job_id)
    
    log.log("Params: %s" % (str(params)))
    
    
    engine = report_engine(params, conf, log, sql)
    
    
    #insert the report path
    sys.path.insert(0, os.path.join(conf.code_root, 'reports'))
    
    report = None
    try:
        exec("import %s as report" % (params['report_name']))
        
        log.log("Got report file: %s.py" % (params['report_name']))
            
    except StandardError, e:
        print "Got Error Loading Report:", e
        sys.exit(1)
    
    #run the report script
    report.run(engine)
    
    engine.log.log("Shutting Down")
    
    #close sql connections, cleanup temporary files, etc.
    engine.cleanup()
    
    engine.log.log("Run Complete - Waiting on watch thread to finish.")
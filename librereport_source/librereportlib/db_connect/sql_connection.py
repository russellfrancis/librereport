#!/usr/bin/env python
# -*- coding: utf-8 -*-
#description     :Creates SQL connections for LibreReport.
#author          :Brandon Tylke (brandon.tylke@lauren.com)
#date            :2013-04-01
#version         :0.0.1
#license         :GPL v3 http://www.gnu.org/licenses/gpl.txt
#python_version  :2.7.1  
#==============================================================================


#############################
#                           #
#  NaviGate SQL Connection  #
#                           #
#############################
 
try:
    import json
except ImportError:
    #RHEL 5, Python 2.4.x grumble... grumble...
    import simplejson as json

class sql_connection:
    """
    
    This is a basic sql connection
    db_module should conform to PEP: 249
    http://www.python.org/dev/peps/pep-0249/
    
    """
    
    def __init__(self, db_module, db_dbname, db_server=None, db_user=None, db_password=None, db_port=None):
        
        self.db_server =    db_server
        self.db_dbname =    db_dbname
        self.db_user =      db_user
        self.db_password =  db_password
        self.db_port =      db_port
        self.db_module =    db_module
        self.connection =   None
        
        #PEP: 249 compatible module for your database
        dbinterface = None
        self.dbinterface = None
        if self.db_module is not None:
            exec("import %s as dbinterface" % self.db_module)
            self.dbinterface = dbinterface
            
            if self.db_module == 'psycopg2':
                import psycopg2.extras
                from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT
        
    def connect(self):
        
        connect_string = "database='%s'" % self.db_dbname
        
        if self.db_server is not None:
            connect_string = "%s, host='%s'" % (connect_string, self.db_server)
        
        if self.db_user is not None:
            connect_string = "%s, user='%s'" % (connect_string, self.db_user)
        
        if self.db_password is not None:
            connect_string = "%s, password='%s'" % (connect_string, self.db_password)
        
        if self.db_port is not None:
            connect_string = "%s, port='%s'" % (connect_string, self.db_port)

        #this is a hackish way to get around a limitation I hit when connecting with sqlite3
        exec("self.connection = self.dbinterface.connect(%s)" % connect_string)
        
        if self.db_module == 'psycopg2':
            self.connection = psycopg2.extras.DictConnection
            self.connection.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
        else:
            self.connection.row_factory = self.dict_factory
        
        return self.connection
    
    #http://stackoverflow.com/questions/811548/sqlite-and-python-return-a-dictionary-using-fetchone
    def dict_factory(self, cursor, row):
        d = {}
        for idx,col in enumerate(cursor.description):
            d[col[0]] = row[idx]
        return d
    
    
    def cursor(self, connection=None):
        if connection is None:
            #do we have a connection yet?
            if self.connection is None:
                self.connect()
                
            connection = self.connection
        
        #sqlite3.ProgrammingError: SQLite objects created in a thread can only be used in that same thread.
        try:
            new_cursor = connection.cursor()
        except StandardError, e:
            #sqlite3 does not like sharing connections across threads
            if 'thread' in str(e).lower():
                try:
                    self.connect()
                    connection = self.connection
                    new_cursor = connection.cursor()
                except:
                    print "Bailing on getting a cursor.  Something is wrong with the connection."
                    raise
            else:
                raise
        
        return new_cursor
    
    def commit(self, connection=None):
        if connection is None:
            connection = self.connection
        
        return connection.commit()
    
    def close(self, connection=None):
        if connection is None:
            connection = self.connection
        
        return connection.close()
    
    # I often use this table to pass parameters from external application to LibreReport
    def get_job_details(self, job_id):
        if self.connection is None:
            self.connect()
        
        #CREATE TABLE reports
        #(
        #  id bigserial NOT NULL,
        #  date_entered timestamp without time zone,
        #  parameters text,
        #  pa_id integer,
        #  mt_id integer,
        #  status character varying,
        #  filename text,
        #  final_output character varying(25),
        #  recipients text,
        #  mimetype character varying(255),
        #  original_filename character varying(255),
        #  CONSTRAINT reports_pkey PRIMARY KEY (id )
        #)
        #WITH (
        #  OIDS=FALSE
        #);
        
        cur = self.cursor(self.connection)

        #let the user know that we have the parameters
        cur.execute("UPDATE reports SET status='grabbed_parameters' WHERE id = %s;" % job_id)
        self.connection.commit()
        
        cur.execute("SELECT parameters FROM reports WHERE id = %s;" % job_id)
        report_parameters = cur.fetchone()['parameters']
        json_params = json.loads(report_parameters)
        
        return json_params
    
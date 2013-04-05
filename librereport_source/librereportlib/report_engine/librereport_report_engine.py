#!/usr/bin/env python
# -*- coding: utf-8 -*-
 
# import the needed modules
import zipfile
from time import time, sleep
from random import randint
from xml.sax.saxutils import quoteattr
import shutil
import os
import errno
import signal
import tempfile
from PIL import Image
from decimal import Decimal, getcontext
import subprocess
#for threading
from threading import Thread
import sys
import stat


#for mail
#http://docs.python.org/library/email-examples.html
#http://stackoverflow.com/questions/3362600/how-to-send-email-attachments-with-python
import smtplib
# For guessing MIME type based on file name extension
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.Utils import COMMASPACE, formatdate
from email import Encoders

#we should look into consolidating this and not need a second call
#import DocumentConverter

class report_engine:
    """LibreReport Engine
    
    'Mail Merge' for reports.  ;)
    
    """
    
    def __init__(self, params, conf, log, sql=None, start_thread=True):
        #info that we need to run
        self.params = params #report specific parameters
        self.sql = sql #sql connection
        self.conf = conf #config vars
        self.log = log #config vars
        
        #global to hold the main xmldata
        self.contentxml = ""
        self.stylesxml = ""
        self.manifestxml = ""
        
        self.originalchartobjects = {}
        #self.newchartobjects = {}
        
        
        self.document_name = ""
        self.document_extension = ""
        self.temp_document = ""
        self.temp_directory = ""
        self.output_documents = []
        self.image_sizes = {}

        try:
            os.makedirs(self.conf.completed_report_path)
        except OSError, exception:
            if exception.errno != errno.EEXIST:
                raise
        
        #conversion process information
        self.soffice_proc = None
        
        #this tracks when the report started running (we don't let any report run longer than 4 hours)
        self.start_time = time();
        
        if start_thread:
            #start the thread that watches for a user cancel
            self.run_thread = True
            t = Thread(target=self.watch_for_cancel)
            t.start()

    
    
    
    
    
    
    
    
    def find_chart(self, chart_name, content=None):
        
        if content == None:
            content = self.contentxml
        
        #<draw:frame draw:style-name="fr1" draw:name="success_fail_percent_chart" text:anchor-type="paragraph" svg:x="0.0898in" svg:y="0.1752in" svg:width="3.1756in" svg:height="3.0161in" draw:z-index="0"><draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/><draw:image xlink:href="./ObjectReplacements/Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>
        
        #brute force
        chart_start = -1
        depth = 1
        
        while chart_start == -1:
            chart_start = content.find('<draw:frame draw:style-name="fr%s" draw:name="%s"' % (str(depth), chart_name))
            depth += 1
            
            #you should never get this high
            if depth > 100:
                break
        
        if chart_start == -1: 
            self.log.log("chart_name %s" % chart_name)
            self.log.log(content)
        
        chart_end = content[chart_start:].find('</draw:frame>') + chart_start + len('</draw:frame>')
        
        original_chart_content = content[chart_start:chart_end]
        
        #get the original filename
        #self.log.log(original_chart_content)
        chart_object_name = original_chart_content.split('<draw:object xlink:href="./')[1].split('"')[0]
        
        self.log.log(chart_object_name)
        
        #content = content.replace(original_chart_content,new_chart_content)
        
        return original_chart_content, chart_object_name
    

    def process_chart(self, chart_name, content, column_order, sql='', preformatted_data=None, content_replacements={}):
        assert type(column_order) == tuple
        
        if self.sql == None:
            self.kill_engine('No SQL connection provided.')
        
        if preformatted_data == None:
            cur = self.sql.cursor()
            #self.log.log('sql: %s' % str(sql))
            try:
                if type(sql) == tuple:
                    cur.execute(sql[0], sql[1])
                else:
                    cur.execute(sql)
                rows = cur.fetchall()
            except StandardError, e:
                self.log.log("Got Error: %s on sql: %s" % (str(e), sql))
                sys.exit(255)
        else:
            rows = preformatted_data
        
        #find the chart in the provided document content
        original_chart_content, original_chart_object_name = self.find_chart(chart_name, content)
        
        #get the content for the original chart object
        object_contentxml = self.originalchartobjects[original_chart_object_name]["content.xml"]
        
        #do any title find/replace work here
        if content_replacements != {}:
            for key in content_replacements.keys():
                object_contentxml = object_contentxml.replace(key, content_replacements[key])
        
        object_stylesxml = self.originalchartobjects[original_chart_object_name]["styles.xml"]
        object_metaxml = self.originalchartobjects[original_chart_object_name]["meta.xml"]
        object_objectreplacements = self.originalchartobjects[original_chart_object_name]["ObjectReplacements"]
        
        #find the data table for the chart
        #they seem to be statically named "local-table"
        
        self.log.log('object_contentxml ' + object_contentxml)
        #section_start, section_end, has_subsection = self.find_table('local-table', object_contentxml, False)
        self.log.log('found local-table')
        #original_chart_object_data_table_content = self.get_content(section_start, section_end, object_contentxml)
        original_chart_object_data_table_content = object_contentxml.split('<table:table table:name="local-table">')[1].split('</table:table>')[0]
        self.log.log('found table content: %s' % original_chart_object_data_table_content)
        
        final_chart_object_data_table_content = ""
        
        first_row = True
        
        self.log.log('rows %s' % rows)
        for row in rows:
            temp_content = ''
            self.log.log('row %s' % row)
            if first_row:
                #start a new row
                temp_content += '<table:table-header-columns><table:table-column/></table:table-header-columns><table:table-columns><table:table-column/></table:table-columns><table:table-header-rows><table:table-row>'
                    
                first_column = True
                for column in column_order:
                    if first_column:
                        #the value of this column is ignored
                        temp_content += '<table:table-cell><text:p/></table:table-cell>'
                        first_column = False
                    else:
                        self.log.log('column %s' % column)
                        self.log.log('str(row[column]) %s' % str(row[column]))
                        temp_content += '<table:table-cell office:value-type="string"><text:p>%s</text:p></table:table-cell>' % (str(row[column]))
                    
                #end the header row
                temp_content += '</table:table-row></table:table-header-rows>'
                #start the data rows
                temp_content += '<table:table-rows>'
                
                first_row = False
            else:
                
                #start a new row
                temp_content += '<table:table-row>'
                
                first_column = True
                for column in column_order:
                    self.log.log('data column %s' % column)
                    self.log.log('data str(row[column]) %s' % str(row[column]))
                    if first_column:
                        temp_content += '<table:table-cell office:value-type="string"><text:p>%s</text:p></table:table-cell>' % (str(row[column]))
                        first_column = False
                    else:
                        temp_content += '<table:table-cell office:value-type="float" office:value="%s"><text:p>%s</text:p></table:table-cell>' % (str(row[column]), str(row[column]))
                    
                #end the row
                temp_content += '</table:table-row>'
                
            #add the row to the final content
            self.log.log("temp_content %s" % temp_content)
            final_chart_object_data_table_content += temp_content

        final_chart_object_data_table_content += '</table:table-rows>'
        
        self.log.log("original_chart_object_data_table_content %s" % original_chart_object_data_table_content)
        self.log.log("")
        self.log.log("final_chart_object_data_table_content %s" % final_chart_object_data_table_content)
        
        #chart data table replacement
        object_contentxml = object_contentxml.replace(original_chart_object_data_table_content, final_chart_object_data_table_content)
                
        #get a random number for a new chart name
        new_chart_object_number = str(randint(10000,99999)) + str(int(time()))
        new_chart_object_name = 'Object %s' % new_chart_object_number
        #create a replacement name for the main document
        new_chart_name = chart_name + new_chart_object_number
        
        ##add the object to the self.newchartobjects var, so it is added to the final zip file
        #self.newchartobjects[new_chart_object_name]["content.xml"] = object_contentxml
        #self.newchartobjects[new_chart_object_name]["styles.xml"] = object_stylesxml
        #self.newchartobjects[new_chart_object_name]["meta.xml"] = object_metaxml
        #self.newchartobjects[new_chart_object_name]["ObjectReplacements"] = object_objectreplacements
        
        #add the chart object to the temp doc
        ziparchiveout = zipfile.ZipFile(self.temp_document, "a")
        ziparchiveout.writestr("%s/content.xml" % (new_chart_object_name), object_contentxml)
        ziparchiveout.writestr("%s/styles.xml" % (new_chart_object_name), object_stylesxml)
        ziparchiveout.writestr("%s/meta.xml" % (new_chart_object_name), object_metaxml)
        ziparchiveout.writestr("ObjectReplacements/%s" % (new_chart_object_name), object_objectreplacements)
        self.log.log("Added chart TO ZIP: %s" % new_chart_object_name)
        ziparchiveout.close()
        
        #update manifest
        new_manifest_entry = """
        <manifest:file-entry manifest:media-type="application/x-openoffice-gdimetafile;windows_formatname=&quot;GDIMetaFile&quot;" manifest:full-path="ObjectReplacements/%s"/>
        <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="%s/content.xml"/>
        <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="%s/styles.xml"/>
        <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="%s/meta.xml"/>
        <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.chart" manifest:full-path="%s/"/>
        """ % (new_chart_object_name, new_chart_object_name, new_chart_object_name, new_chart_object_name, new_chart_object_name)
        
        self.manifestxml = self.manifestxml.replace("</manifest:manifest>","") #remove the closing tag
        self.manifestxml += new_manifest_entry + "\n</manifest:manifest>" #add the new entry and put the closing tag back
        
        
        #change the name in the main document and update the object it references
        final_chart_content = original_chart_content.replace(original_chart_object_name, new_chart_object_name).replace(chart_name, new_chart_name)
        return_content = self.replace_section_content(original_chart_content, final_chart_content, '', content, False)
        
        return return_content
    
    
    def find_table_end(self, content, open_string='<table:table ', close_string='</table:table>'):
        table_search_end = 0
        has_subtable = False
        
        tag_depth = 1
        
        #content = content[content.find(open_string)+len(open_string):]
        #table_search_end += content.find(open_string)+len(open_string)
        
        while tag_depth > 0:
            
            if content.find(open_string) != -1 and content.find(open_string) < content.find(close_string):
                table_search_end += content.find(open_string)+len(open_string)
                content = content[content.find(open_string)+len(open_string):]
                tag_depth += 1
                has_subtable = True
                
            elif content.find(close_string) != -1 and (content.find(close_string) < content.find(open_string) or content.find(open_string) == -1):
                table_search_end += content.find(close_string)+len(close_string)
                content = content[content.find(close_string)+len(close_string):]
                tag_depth -= 1
        
        table_search_end -= len(close_string)
        
        return table_search_end, has_subtable
    
        
    def find_table(self, table_name, content=None, strip_table_header=True, remove_footer_row=False):
        
        if content == None:
            content = self.contentxml
        
        table_start = 0
        table_end = 0
        has_subtable = False
        
        table_search = '<table:table table:name="%s" table:style-name="' % (table_name)
        
        table_start = content.find(table_search)
        if table_start == -1:
            self.log.log("Cannot find table: %s" % table_name)
            self.log.log("In Content: %s" % content)
        else:
            table_start += content[table_start:].find('>')
            #we don't want to include the table, just its content
    #        table_start += len(table_search)
            
            if content[table_start:].find("<table:table-row>") != -1 or content[table_start:].find("<table:table-row table:style-name=") != -1:
                #this is terrible - it could be either of these
                possible_table_start_short_tag = content[table_start:].find("<table:table-row>")
                possible_table_start_long_tag = content[table_start:].find("<table:table-row table:style-name=")
                if (possible_table_start_short_tag != -1 and possible_table_start_short_tag < possible_table_start_long_tag) or possible_table_start_long_tag == -1:
                    table_start += possible_table_start_short_tag
                else:
                    table_start += possible_table_start_long_tag
                    table_start += content[table_start:].find('>')
            else:
                print "Cannot find row start"
            
            table_end, has_subtable = self.find_table_end(content[table_start:])
            #because it was X chars after our starting position
            table_end += table_start
            
            header_check_content = content[table_start:table_end]
            
            #if the table has a header, we won't include it.
            if header_check_content.find('</table:table-header-rows>') != -1 and strip_table_header:
                
                #we don't want to start at the header of the sub table
                if has_subtable and header_check_content.find('</table:table-header-rows>') > header_check_content.find('<table:table '):
                    pass
                else:
                    table_start += header_check_content.find('</table:table-header-rows>')
                    table_start += len('</table:table-header-rows>')

            if remove_footer_row:
                remove_footer_content = content[table_start:table_end]
                
                #[::-1] is extended slice syntax and has the effect of reversing the string, thus finding the last occurance of "<table:table-row>"
                position_from_end = remove_footer_content[::-1].find("<table:table-row>"[::-1])
                
                #strip the opening tag
                position_from_end += len('<table:table-row>')
                
                #remove the last row of this table, we will treat it like a footer
                table_end = table_end - position_from_end

        return table_start, table_end, has_subtable


        
    def find_section_end(self, content, open_string='<text:section ', close_string='</text:section>'):
        section_search_end = 0
        has_subsection = False
        
        tag_depth = 1
        
        #content = content[content.find(open_string)+len(open_string):]
        #section_search_end += content.find(open_string)+len(open_string)
        
        while tag_depth > 0:
            
            if content.find(open_string) != -1 and content.find(open_string) < content.find(close_string):
                section_search_end += content.find(open_string)+len(open_string)
                content = content[content.find(open_string)+len(open_string):]
                tag_depth += 1
                has_subsection = True
                
            elif content.find(close_string) != -1 and (content.find(close_string) < content.find(open_string) or content.find(open_string) == -1):
                section_search_end += content.find(close_string)+len(close_string)
                content = content[content.find(close_string)+len(close_string):]
                tag_depth -= 1
        
        section_search_end -= len(close_string)
        
        return section_search_end, has_subsection
        
        
    def find_section(self, section_name, content=None):
        if content == None:
            content = self.contentxml
        
        section_start = -1
        section_end = 0
        has_subsection = False
        
        depth = 1
        
        while section_start == -1:
            section_search = '<text:section text:style-name="Sect%s" text:name="%s">' % (str(depth), section_name)
            section_start = content.find(section_search)
            depth += 1
            
            #you should never get this high
            if depth > 100:
                break
        
        if section_start == -1: 
            self.log.log("section_name %s" % section_name)
            self.log.log(content)
        
        #we don't want to include the section, just its content
        section_start += len(section_search)
        #print "section_start", section_start
        
        section_end, has_subsection = self.find_section_end(content[section_start:])
        #because it was X chars after our starting position
        section_end += section_start
        #print "section_end", section_end
        
        return section_start, section_end, has_subsection

    def clear_section(self, section_name):
        section_start, section_end, has_subsection = self.find_section(section_name)
        original_section_content = self.get_section_content(section_start, section_end)
        self.contentxml = self.contentxml.replace(original_section_content,'')
        return

    def update_content(self, section_content, new_section_content):
        self.contentxml = self.contentxml.replace(section_content, new_section_content)
        return

    def find_next_section(self, position):
        next_section_name = None
        section_start = 0
        section_end = 0
        has_subsection = False
        
        if position != 0:
            #go to the end of the current section
            position = position + len("</text:section>")
        
        search_content = self.contentxml[position:]
        
        next_section_start = search_content.find('<text:section ')
        
        if next_section_start != -1:
            search_content = search_content[next_section_start:]
            
            #find name start
            next_section_start = search_content.find('text:name="') + len('text:name="')
            search_content = search_content[next_section_start:]
            # " = the end of next section name definition
            next_section_name = search_content[:search_content.find('"')]
            
            section_start, section_end, has_subsection = self.find_section(next_section_name, self.contentxml)
            
        return next_section_name, section_start, section_end, has_subsection

    #we must specify at least one document variable to look for to know we have the correct row
    def get_spreadsheet_data_rows(self, document_variables, content=None):
        #self.log.log('get_spreadsheet_data_rows')
        if content == None:
            content = self.contentxml
        
        tmp_content = content
        
        row_start = 0
        row_end = 0
        
        found_var = False
        found_end = False
        
        tmp_row_start = 0
        tmp_row_end = 0
        
        while not found_var and tmp_row_start != -1 and tmp_row_end != -1:
            tmp_content = tmp_content[tmp_row_end:]
            tmp_row_start = tmp_content.find('<table:table-row table:style-name="')
            tmp_row_end = tmp_content[tmp_row_start:].find('</table:table-row>')
            tmp_row_end += len('</table:table-row>')
            tmp_row_end += tmp_row_start
            
            #self.log.log("tmp_row_start %s" % str(tmp_row_start))
            #self.log.log("tmp_row_end %s" % str(tmp_row_end))
            
            self.log.log("document_variables %s" % str(document_variables))
            
            for doc_var in document_variables:
                if found_var == False:
                    self.log.log("doc_var %s" % str(doc_var))
                    if doc_var in tmp_content[tmp_row_start:tmp_row_end]:
                        found_var = True
                        row_start += tmp_row_start
                        row_end = row_start + tmp_row_end
                        
                        tmp_content = tmp_content[tmp_row_end:]
                        self.log.log("found doc_var %s" % str(doc_var))
            
            if found_var == False:
                row_start += tmp_row_end
        
        if found_var:
            #start where we left off
            tmp_row_start = 0
            
            self.log.log("Part 2")
            while not found_end and tmp_row_start != -1:
                tmp_content = tmp_content[tmp_row_start:]
                self.log.log("Full tmp_content:" + tmp_content)
                
                #get the end of the next tag
                next_row_tag_end = tmp_content.find(">")
                self.log.log("next_row_tag_end %s" % str(next_row_tag_end))
                self.log.log("\nSearch tmp_content: "+tmp_content[:next_row_tag_end])
                 
                if 'table:table table:name="' in tmp_content[:next_row_tag_end]:
                    self.log.log("Found table:table:table table:name=")
                    found_end = True
                    row_end += tmp_row_start - len('</table:table>')
                elif 'table:number-rows-repeated' in tmp_content[:next_row_tag_end]:
                    self.log.log("Found table:number-rows-repeated")
                    found_end = True
                    row_end += tmp_row_start
                else:
                    if tmp_content.find('</table:table-row>') != -1 and tmp_content.find('</table:table-row>') < tmp_content.find('</table:table>'):
                        tmp_row_start = tmp_content.find('</table:table-row>')
                        if tmp_row_start != -1:
                            tmp_row_start += len('</table:table-row>')
                    else:
                        tmp_row_start = tmp_content.find('</table:table>')
                        if tmp_row_start != -1:
                            tmp_row_start += len('</table:table>')

                    
                
        
        self.log.log("row_start %s" % str(row_start))
        self.log.log("row_end %s" % str(row_end))

        if found_var and found_end:
            return content[row_start:row_end]
        else:
            return False
        

    def get_section_content(self, section_start, section_end, content=None):
        if content == None:
            content = self.contentxml
        
        return content[section_start:section_end]
    
    def replace_section_content(self, original_content, replacement_content, section_name='', content=None, strip_section_marker=True):
        if content == None:
            content = self.contentxml
        
        if strip_section_marker:
            original_content = '<text:section text:style-name="Sect1" text:name="' + section_name + '">' + original_content + '</text:section>'
        
        content = content.replace(original_content,replacement_content)
        
        return content

    def update_section_content(self, data, substitutions, content=None):
        assert type(substitutions) == tuple
        
        #there were a few instances that substitutions were sent as a single tuple
        #this should save some programmer frustration
        if type(substitutions[1]) != tuple:
            substitutions = [substitutions]
        
        if content == None:
            content = self.contentxml
        
        for sub in substitutions:
            replacement_value = ''
            try:
                #quoteattr feels the need to add "quotes", hence the [1:-1]
                if data[sub[1]] != None:
                    replacement_value = str(data[sub[1]])
            except StandardError, e:
                self.log.log("Got error in update_section_content: %s" % str(e))
                self.log.log("data: %s" % str(data))
                self.log.log("substitutions: %s" % str(substitutions))
                self.log.log("sub: %s" % str(sub))
                self.log.log("content: %s" % str(content))
            
            #handle unicode issues
            replacement_value = unicode(replacement_value, errors='ignore')
            
            replacement_data = quoteattr(replacement_value)[1:-1]
            
            content = content.replace(sub[0],replacement_data)
        
        return content

    ################################################################################################################################################################################################################################
    #
    # a function that takes the section name, document xml content, sql connection and tuple of variable/column substitutions
    # optional parameter image_details is a tuple of defined image name and column name pair for image replacement within the section content
    # optional parameter display_no_data_message will insert "No Data Entered" into the section content if no records are returned 
    #
    ################################################################################################################################################################################################################################

    def process_section(self, section_name, content, sql, substitutions, image_details=None, display_no_data_message=True):
        assert type(substitutions) == tuple
        
        if self.sql == None:
            self.kill_engine('No SQL connection provided.')
        
        cur = self.sql.cursor()
        
        try:
            if type(sql) == tuple:
                cur.execute(sql[0], sql[1])
            else:
                cur.execute(sql)
        except StandardError, e:
            self.log.log("Got Error: %s on sql: %s" % (str(e), sql))
            sys.exit(255)
        #self.log.log('section_name: %s' % str(section_name))
        #self.log.log('content: %s' % str(content))
        section_start, section_end, has_subsection = self.find_section(section_name, content)
        
        original_section_content = self.get_section_content(section_start, section_end, content)
    
        final_section_content = ""
        
        for row in cur.fetchall():
            #self.log.log("Next " + section_name + ": " + time.ctime())
            temp_content = ''
            temp_content += self.update_section_content(row, substitutions, original_section_content)
            
            #image replacement
            if image_details != None and type(image_details) == tuple and len(image_details) == 2:
                image_name, image_column_name = image_details
                if row[image_column_name] != None:
                    new_image_name = self.add_image(os.path.join(self.conf.site_files, row[image_column_name]))
                    if new_image_name != None:
                        temp_content = self.replace_image(image_name, new_image_name, temp_content)
            
            final_section_content += temp_content
        
        if final_section_content == "" and display_no_data_message == True:
            self.log.log('No Data Entered for %s' % str(section_name))
            final_section_content = '<text:p text:style-name="Standard">No Data Entered</text:p>'
        
        return_content = self.replace_section_content(original_section_content, final_section_content, section_name, content)
        
        return return_content

    ################################################################################################################################################################################################################################
    #
    # a function that takes the table name, document xml content, sql connection and tuple of variable/column substitutions
    # optional parameter image_details is a tuple of defined image name and column name pair for image replacement within the table row content
    #
    ################################################################################################################################################################################################################################

    def process_table(self, table_name, content, sql, substitutions, image_details=None, remove_footer_row=False):
        assert type(substitutions) == tuple
        
        if self.sql == None:
            self.kill_engine('No SQL connection provided.')
        
        cur = self.sql.cursor()
        #self.log.log('sql: %s' % str(sql))
        try:
            if type(sql) == tuple:
                cur.execute(sql[0], sql[1])
            else:
                cur.execute(sql)
        except StandardError, e:
            self.log.log("Got Error: %s on sql: %s" % (str(e), sql))
            sys.exit(255)
        #self.log.log('table_name: %s' % str(table_name))
        #self.log.log('content: %s' % str(content))
        section_start, section_end, has_subsection = self.find_table(table_name, content, remove_footer_row=remove_footer_row)
        
        original_section_content = self.get_section_content(section_start, section_end, content)
    
        final_section_content = ""
        
        for row in cur.fetchall():
            #self.log.log("Next " + table_name + ": " + time.ctime())
            #self.log.log('processing returned rows')
            temp_content = ''
            temp_content += self.update_section_content(row, substitutions, original_section_content)
            
            #image replacement
            if image_details != None and type(image_details) == tuple and len(image_details) == 2:
                image_name, image_column_name = image_details
                if row[image_column_name] != None:
                    new_image_name = self.add_image(os.path.join(self.conf.site_files, row[image_column_name]))
                    if new_image_name != None:
                        temp_content = self.replace_image(image_name, new_image_name, temp_content)
            
            final_section_content += temp_content

        #self.log.log('done processing returned rows')
        
        if final_section_content == "":
            self.log.log('No Data Entered for %s' % str(table_name))
            final_section_content = '<text:p text:style-name="Standard">No Data Entered</text:p>'

        return_content = self.replace_section_content(original_section_content, final_section_content, table_name, content, False)
        
        return return_content

    def process_spreadsheet(self, content, sql, substitutions, search_variable=None):
        assert type(substitutions) == tuple
        
        if self.sql == None:
            self.kill_engine('No SQL connection provided.')
        
        cur = self.sql.cursor()
        #self.log.log('sql: %s' % str(sql))
        try:
            if type(sql) == tuple:
                cur.execute(sql[0], sql[1])
            else:
                cur.execute(sql)
        except StandardError, e:
            self.log.log("Got Error: %s on sql: %s" % (str(e), sql))
            sys.exit(255)
        
        if search_variable == None:
            if type(substitutions[0]) != tuple:
                search_variable = list(substitutions[0])
            else:
                #this passes a list of variables to look for in the xml
                search_variable = []
                for sub in substitutions:
                    search_variable.append(sub[0]) #this makes the assumption that this varaible occurs on the first row to be swapped out
        elif type(search_variable) == str:
            search_variable = list(search_variable)
            
        original_section_content = self.get_spreadsheet_data_rows(search_variable, content)
        #self.log.log('original_section_content' + str(original_section_content)) 
        final_section_content = ""
        
        for row in cur.fetchall():
            #self.log.log('processing returned rows')
            temp_content = ''
            temp_content += self.update_section_content(row, substitutions, original_section_content)
            
            final_section_content += temp_content



        #self.log.log('final_section_content' + str(final_section_content)) 
        #self.log.log('content' + str(content)) 

        return_content = self.replace_section_content(original_section_content, final_section_content, '', content, False)
        
        return return_content

    def replace_image(self, image_section_name, new_image_name, content=None):
        
        if content == None:
            content = self.contentxml
        
        #<draw:frame draw:style-name="fr1" draw:name="people_image" text:anchor-type="paragraph" svg:width="2.2173in" svg:height="2.2098in" draw:z-index="2"><draw:image xlink:href="Pictures/10000201000001000000010091D48B04.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>
        
        #brute force
        if content.find('<draw:frame draw:style-name="fr1" draw:name="%s"' % (image_section_name)) != -1:
            image_start = content.find('<draw:frame draw:style-name="fr1" draw:name="%s"' % (image_section_name))
        elif content.find('<draw:frame draw:style-name="fr2" draw:name="%s"' % (image_section_name)) != -1:
            image_start = content.find('<draw:frame draw:style-name="fr2" draw:name="%s"' % (image_section_name))
        elif content.find('<draw:frame draw:style-name="fr3" draw:name="%s"' % (image_section_name)) != -1:
            image_start = content.find('<draw:frame draw:style-name="fr3" draw:name="%s"' % (image_section_name))
        
        self.log.log('')
        self.log.log(content)
        image_end = content[image_start:].find('</draw:frame>') + image_start + len('</draw:frame>')
        
        old_image_content = content[image_start:image_end]
        
        #get the old filename
        self.log.log(old_image_content)
        old_image_name = old_image_content.split('xlink:href="Pictures/')[1].split('"')[0]
        #get the old width
        old_image_width = old_image_content.split('svg:width="')[1].split('in"')[0]
        old_image_width_string = 'svg:width="%sin"' % (old_image_width)
        old_image_width = Decimal(old_image_width)
        #get the old height
        old_image_height = old_image_content.split('svg:height="')[1].split('in"')[0]
        old_image_height_string = 'svg:height="%sin"' % (old_image_height)
        old_image_height = Decimal(old_image_height)
        
        #find/replace
        new_image_content = old_image_content.replace(old_image_name, new_image_name)
        
        #get new image ratio
        w, h = self.image_sizes[new_image_name]
        
        getcontext().prec = 2
        
        #keep ration for the new image
        if w > h:
            new_image_width = old_image_width
            new_image_height = (Decimal(h) / Decimal(w)) * old_image_width
        else:
            new_image_height = old_image_height
            new_image_width = (Decimal(w) / Decimal(h)) * old_image_height
        
        #update image sizes
        new_image_content = new_image_content.replace(old_image_width_string, 'svg:width="%sin"' % (new_image_width))
        new_image_content = new_image_content.replace(old_image_height_string, 'svg:height="%sin"' % (new_image_height))
        
        content = content.replace(old_image_content,new_image_content)
        
        return content

    def add_image(self, image_full_path):
        try:
            #image_full_path = os.path.join(image_path, image_name)
            self.log.log("opening image: %s" % image_full_path)
            image_file = open(image_full_path,'rb')
            image_data = image_file.read()
            image_file.close()
            
            image_extension = os.path.splitext(image_full_path)[1] # '.ext'
            image_extension = image_extension.replace('.','') # 'ext'
            
            new_image_name = str(randint(10000,99999)) + str(int(time())) + '.' +image_extension
            
            #get image size, so we can calculate ratios when adding the image
            try:
                im = Image.open(image_full_path)
                self.image_sizes[new_image_name] = im.size
                
                try:
                    #for nonstandard images
                    resampled_image = os.path.join(self.temp_directory, new_image_name)
                    
                    im.save(resampled_image)
                    
                    #read resampled image
                    image_file = open(resampled_image,'rb')
                    image_data = image_file.read()
                    image_file.close()
                except StandardError, e:
                    self.log.log("could not save resampled image, got error: %s" % str(e))
                    
            except:
                self.image_sizes[new_image_name] = [1,1]
            
            ziparchiveout = zipfile.ZipFile(self.temp_document, "a")
            ziparchiveout.writestr("Pictures/%s" % (new_image_name), image_data)
            self.log.log("Added image TO ZIP: %s" % new_image_name)
            ziparchiveout.close()
            
            #print "update manifest"
            new_manifest_entry = '<manifest:file-entry manifest:media-type="image/%s" manifest:full-path="Pictures/%s"/>' % (image_extension, new_image_name)
            self.manifestxml = self.manifestxml.replace("</manifest:manifest>","") #remove the closing tag
            self.manifestxml += new_manifest_entry + "\n</manifest:manifest>" #add the new entry and put the closing tag back
            return new_image_name
        except StandardError, e:
            self.log.log('add_image() got error: %s' % str(e))
            return None
        

    def open_document(self, document_name=None):
        
        if document_name == None:
            for try_path in (os.path.join(self.conf.code_root, 'reports'), './'):
                if os.path.isfile(os.path.join(try_path, '%s.odt' % (self.params['report_name']))):
                    document_name = os.path.join(try_path, '%s.odt' % (self.params['report_name']))
                    
                elif os.path.isfile(os.path.join(try_path, '%s.ods' % (self.params['report_name']))):
                    document_name = os.path.join(try_path, '%s.ods' % (self.params['report_name']))

        self.log.log("Got report document: %s" % (document_name))
        
        self.document_name = document_name
        
        self.document_extension = os.path.splitext(self.document_name)[1]
        
        # get content xml data from OpenDocument template file
        ziparchive = zipfile.ZipFile(self.document_name, "r")
        self.contentxml = ziparchive.read("content.xml")
        self.stylesxml = ziparchive.read("styles.xml")
        self.manifestxml = ziparchive.read("META-INF/manifest.xml")
        
        self.contentxml = unicode(self.contentxml, errors='ignore')
        self.stylesxml = unicode(self.stylesxml, errors='ignore')
        self.manifestxml = unicode(self.manifestxml, errors='ignore')
        
        #get chart objects, if any
        for line in self.manifestxml.split('>'):
            if 'application/vnd.oasis.opendocument.chart' in line:
                chart_object_name = line.split('manifest:full-path="')[1].split('/"')[0]
                self.log.log("chart_object_name: %s" % chart_object_name)
                self.originalchartobjects[chart_object_name] = {}
        
        #if we found some charts, get their data
        if self.originalchartobjects != {}:
            for object_name in self.originalchartobjects.keys():
                self.originalchartobjects[object_name]["content.xml"] = ziparchive.read(object_name + "/content.xml")
                self.originalchartobjects[object_name]["styles.xml"] = ziparchive.read(object_name + "/styles.xml")
                self.originalchartobjects[object_name]["meta.xml"] = ziparchive.read(object_name + "/meta.xml")
                self.originalchartobjects[object_name]["ObjectReplacements"] = ziparchive.read("ObjectReplacements/" + object_name)
                
                # <manifest:file-entry manifest:media-type="application/x-openoffice-gdimetafile;windows_formatname=&quot;GDIMetaFile&quot;" manifest:full-path="ObjectReplacements/Object 1"/>
                #<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="Object 1/content.xml"/>
                #<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="Object 1/styles.xml"/>
                #<manifest:file-entry manifest:media-type="text/xml" manifest:full-path="Object 1/meta.xml"/>
                #<manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.chart" manifest:full-path="Object 1/"/>
        
        self.temp_directory = tempfile.mkdtemp()
        #print "self.temp_directory", self.temp_directory
        
        self.temp_document = os.path.join(self.temp_directory, str(randint(10000,99999)) + str(int(time())) + "_output.odt")
        
        #create a master temp document here
        #http://stackoverflow.com/questions/4653768/overwriting-file-in-ziparchive
        exclude_files = ["content.xml", "META-INF/manifest.xml", "styles.xml"]
        
        zipwrite = zipfile.ZipFile(self.temp_document, 'w')
        
        for item in ziparchive.infolist():
            #print "item.filename", item.filename
            if item.filename not in exclude_files:
                data = ziparchive.read(item.filename)
                zipwrite.writestr(item, data)
            #else:
                #print "skipped item.filename", item.filename
        
        #new doc
        zipwrite.close()
        
        #original doc
        ziparchive.close()

    def save_document(self):
        self.contentxml = self.contentxml.replace('-----newline-----','<text:line-break/>')
        self.stylesxml = self.stylesxml.replace('-----newline-----','<text:line-break/>')

        self.output_documents.insert(0, os.path.join(self.temp_directory, str(randint(10000,99999)) + str(int(time())) + "_output.odt"))

        #create output file
        shutil.copy2(self.temp_document, self.output_documents[0])

        ziparchiveout = zipfile.ZipFile(self.output_documents[0], "a")
        ziparchiveout.writestr("content.xml", self.contentxml)
        ziparchiveout.writestr("styles.xml", self.stylesxml)
        ziparchiveout.writestr("META-INF/manifest.xml", self.manifestxml)
        ziparchiveout.close()

        # The process which is running DocumentConverter.py / ooPool needs to be able to read this file.
        # And execute processes within the parent directory, so lets make an effort to adjust the permissions appropriately

        # 755
        parent_path = os.path.abspath(os.path.join(self.output_documents[0], '..'))
        os.chmod(parent_path, stat.S_IRWXU | stat.S_IRGRP | stat.S_IXGRP | stat.S_IROTH | stat.S_IXOTH)

        # 644
        os.chmod(self.output_documents[0], stat.S_IRUSR | stat.S_IWUSR | stat.S_IRGRP | stat.S_IROTH)

        return self.output_documents[0]
        
    def cleanup(self):
        #remove primary temp file
        try:
            os.unlink(self.temp_document)
        except:
            pass
        
        #removes all temp files
        for doc in self.output_documents:
            try:
                os.unlink(doc)
            except:
                pass
        
        #remove the temp dir
        try:
            shutil.rmtree(self.temp_directory, True)
        except:
            pass
        
        #close open sql connections
        if self.sql != None:
            self.sql.close()
        
        #stop the watch thread
        self.run_thread = False
        
    def save_pdf(self):
        return self.convert_report('pdf')
        
    def convert_report(self, output_type='pdf'):
        if output_type not in ["txt", "csv", "pdf", "html", "odt", "doc", "rtf", "ods", "xls", "odp", "ppt", "swf"]:
            self.cleanup()
            sys.exit("export format not available")
        
        temp_doc = self.save_document()
        self.log.log("Temporary ODF Document: %s" % temp_doc)
        file_path, original_file = os.path.split(temp_doc)
        filename = original_file+'.'+output_type
        temp_output_file = os.path.join(self.conf.completed_report_path, filename)
        
        #let the user know that we are creating their output file
        if self.sql != None:
            try:
                if output_type=='pdf':
                    cur = self.sql.cursor()
                    cur.execute("UPDATE reports SET status='generating_pdf' WHERE id = %s;" % self.conf.job_id)
                    self.sql.commit()
                elif output_type=='xls':
                    cur = self.sql.cursor()
                    cur.execute("UPDATE reports SET status='generating_xls' WHERE id = %s;" % self.conf.job_id)
                    self.sql.commit()
                else:
                    cur = self.sql.cursor()
                    cur.execute("UPDATE reports SET status='generating_report' WHERE id = %s;" % self.conf.job_id)
                    self.sql.commit()
                
            except StandardError, e:
                self.log.log("Unable to update report status.  Got error: %s" % str(e))
                pass
        
        #if the output extension is the same as the input file, just copy - otherwise oOConvert
        if self.document_extension == os.path.splitext(temp_output_file)[1]:
            self.log.log("Final OUTPUT: %s" % temp_output_file)
            shutil.copy2(temp_doc, temp_output_file)
        else:
            self.log.log("Final DOC_CONVERTER OUTPUT: %s" % temp_output_file)
            #for now we are calling the external converter - eventually we need to incorporate the doc converter, but it needs psycopg2
            #call the process
            self.log.log("command: %s" % self.conf.python_cmd)
            self.log.log("convert: %s" % self.conf.doc_converter_path)
            self.log.log("input f: %s" % temp_doc)
            self.log.log("output : %s" % temp_output_file)
            self.soffice_proc = subprocess.Popen([self.conf.python_cmd, self.conf.doc_converter_path, temp_doc, temp_output_file])
            #start a thread to watch for a user cancelling the job
            
            #wait for it to finish - we use Popen instead of call(0 so i can kill the process if the user cancels the report
            self.soffice_proc.wait()
            #it is complete
            self.soffice_proc = None
        
        if self.sql != None:
            try:
                #let the user know that we have completed the conversion
                cur.execute("UPDATE reports SET status='complete', filename='%s', original_filename='%s.%s', mimetype='application/%s' WHERE id = %s;" % (filename, self.params['report_name'], output_type, output_type, self.conf.job_id))
                self.sql.commit()
                
            except StandardError, e:
                self.log.log("Unable to update report status.  Got error: %s" % str(e))
                pass
        
        
        return temp_output_file
    
    def send_mail(self, send_from, send_to, files=[]):
        assert type(send_to)==list
        assert type(files)==list
    
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = COMMASPACE.join(send_to)
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = self.conf.default_mail_subject
    
        self.log.log("Composing email to: %s" % msg['To'])
    
        msg.attach( MIMEText(self.conf.default_mail_message) )
    
        for f in files:
            self.log.log("Attaching file: %s" % f)
            part = MIMEBase('application', "octet-stream")
            part.set_payload( open(f,"rb").read() )
            Encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
            msg.attach(part)
        
        self.log.log("Email server auth...")
        
        mailServer = smtplib.SMTP(self.conf.smtp_server,self.conf.smtp_port)
        mailServer.ehlo()
        mailServer.starttls()
        mailServer.ehlo()
        mailServer.login(self.conf.smtp_username, self.conf.smtp_password)
        mailServer.sendmail(send_from, send_to, msg.as_string())
        mailServer.close()

        self.log.log("Email sent...")
    
################################################################
#
#   Report inspector bits...
#
################################################################
    
    def report_inspector_find_sections(self):
        sections = []
        duplicate_sections = []
        
        content = self.contentxml
        
        #<table:table table:name="
        #<text:section text:style-name="Sect1" text:name="
        
        while content.find('<text:section ') != -1:
            content = content[content.find('<text:section ')+len('<text:section '):]
            content = content[content.find('text:name="')+len('text:name="'):]
            if content[:content.find('"')] in sections:
                if content[:content.find('"')] not in duplicate_sections:
                    duplicate_sections.append(content[:content.find('"')])
            else:
                sections.append(content[:content.find('"')])
        
        return sections, duplicate_sections
    
    def report_inspector_find_tables(self):
        tables = []
        duplicate_tables = []
        
        content = self.contentxml
        
        #<table:table table:name="
        #<text:section text:style-name="Sect1" text:name="
        
        while content.find('<table:table table:name="') != -1:
            content = content[content.find('<table:table table:name="')+len('<table:table table:name="'):]
            if content[:content.find('"')] in tables:
                if content[:content.find('"')] not in duplicate_tables:
                    duplicate_tables.append(content[:content.find('"')])
            else:
                tables.append(content[:content.find('"')])
                
        return tables, duplicate_tables
    
    def report_inspector_find_variables(self):
        variables = []
        duplicate_variables = []
        
        content = self.contentxml
        
        #<table:table table:name="
        #<text:section text:style-name="Sect1" text:name="
        
        while content.find('$') != -1:
            content = content[content.find('$')+1:]
            if '$'+content[:content.find('$')+1] in variables:
                if '$'+content[:content.find('$')+1] not in duplicate_variables:
                    duplicate_variables.append('$'+content[:content.find('$')+1])
            else:
                variables.append('$'+content[:content.find('$')+1])
                
            content = content[content.find('$')+1:]
        
        return variables, duplicate_variables
    
    def json_list_to_pg_array(self, json_array):
        """will not work for text arrays"""
        new_array = ",".join(json_array)
        new_array = "'{" + new_array + "}'"
        
        return new_array
    
    def kill_engine(self, reason):
        self.log.log("Killing process. Reason: " + str(reason))
        os.kill(int(os.getpid()), signal.SIGTERM)
        sys.exit(255)
        return

################################################################
#
#   Threading bits...
#
################################################################
    def watch_for_cancel(self):

        if self.sql == None:
            self.log.log("You must provide an active sql connection to watch for a remote kill of the current report.")
        else:
            cur = self.sql.cursor()
            
            #run FOREVER
            while self.run_thread:
                try:
                    cur.execute("SELECT deleted FROM reports WHERE id = %s;" % self.conf.job_id)
                    data = cur.fetchone()
                    
                    self.log.log("(Watch Thread) Report was cancelled/deleted: %s" % str(data['deleted']))
                    
                    if data['deleted'] == True:
                        try:
                            #kill libreffice processes
                            if self.soffice_proc != None:
                                #tell the libreoffice process to stop
                                self.soffice_proc.kill()
                        except StandardError, e:
                            self.log.log("Got error when trying to kill LibreOffice process: %s" % (str(e)))
                            
                        #close any open sql connections
                        self.sql.close()
                        
                        #kill the python process
                        #sys.exit("user cancelled report")
                        self.log.log("User cancelled the report.")
                        #give it 30 seconds to cleanup and kill this process
                        #sleep(30)
                        self.kill_engine("user cancelled report")
                except StandardError, e:
                    self.log.log("Watch thread threw an error: %s" % (str(e)))
                    self.log.log("Stopping the watch thread.")
                    self.run_thread = False
                    
                
                if self.conf.max_report_runtime != None and (time() - self.start_time) > self.conf.max_report_runtime:
                    self.log.log("We've run for more than %s seconds. I am killing the watch thread." % (self.conf.max_report_runtime))
                    self.run_thread = False
                    sleep(30)
                    self.kill_engine("report ran too long.")
                
                sleep(10)
                
        return 
        
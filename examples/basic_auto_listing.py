#!/usr/bin/env python
# -*- coding: utf-8 -*-
#description     :Example report script.  Lists auto makes and models.
#author          :Brandon Tylke (brandon.tylke@lauren.com)
#date            :2013-04-01
#version         :0.0.1
#license         :GPL v3 http://www.gnu.org/licenses/gpl.txt
#python_version  :2.7.1  
#==============================================================================

def run(engine):
    
    engine.open_document()
    
    sql = """
          SELECT
              gd.title AS make,
              gdm.title AS model
          FROM
              global_details gd
              INNER JOIN global_details AS gdm
              ON gdm.parent_global_details_id = gd.id;
    """
    
    engine.contentxml = engine.process_table('autos_tables', engine.contentxml, sql, (
                                                            ("$make$",'make'),
                                                            ("$model$",'model')
                                                            ))
    
    
    engine.convert_report('odt')
    
    #enable this option to create a PDF
    #Note: You will need LibreOffice running for this to work
    #engine.save_pdf()
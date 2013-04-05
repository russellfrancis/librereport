#!/usr/bin/env python
# -*- coding: utf-8 -*-

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
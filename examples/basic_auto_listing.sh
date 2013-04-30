#this will run the example report
#make sure you have LibreOffice v4 or higher running and listening on port 2002.

#this assumes that LibreOffice is installed in /opt/libreoffice4.0/
#  sudo /opt/libreoffice4.0/program/soffice.bin --nologo --nofirststartwizard --headless --norestore --invisible "--accept=socket,host=localhost,port=2002,tcpNoDelay=1;urp;StarOffice.ServiceManager" 

python ../librereport_source/report_loader.py ./basic_auto_listing.py

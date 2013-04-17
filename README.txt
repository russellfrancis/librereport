LibreReport is a python-based open source reporting system for applications.

License: GPLv3








Alternate File Creation: (PDF. XLS, DOC, etc.)
LibreReports requires LibreOffice v4 or higher to produce file formats other than ODT and ODS.


Assuming that LibreOffice is installed in /opt/libreoffice4.0/, you can run the following command to create a background process to generate alternate formats.
sudo /opt/libreoffice4.0/program/soffice.bin --nologo --nofirststartwizard --headless --norestore --invisible "--accept=socket,host=localhost,port=2002,tcpNoDelay=1;urp;StarOffice.ServiceManager" &


For productions systems, we highly recommend using oopool as it scales to support thousands of users.

About oopool:
An OpenOffice/LibreOffice connection pool which will delegate incoming document manipulation requests from clients across a cluster of worker office instances satisfying requests. The pool takes care of starting and stopping office instances based on load as well as distributing the incoming requests across available resources. 

http://oopool.laureninnovations.com/
https://bitbucket.org/brandontylke/oopool

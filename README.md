# VbFcgi
FCGI Host/Server Framework for Visual Basic 6 (VB6)/COM FCGI Applications
# Introduction
VbFcgi is a framework that allows you to easily get VB6 code onto the web. It was developed against Nginx, but should work with any web server that implements the FCGI spec.
# Included Executables
There are 3 main components of the VbFcgi framework:
1. VbFcgi.dll - This is the main framework library that includes all of the code for listening and responding to FCGI requests from the web server, as well as parsing out records for FCGI parameters, HTTP cookies, etc...
2. VbFcgiHost.exe - This is main executable file that will spawn FCGI listeners (multiple listeners support for load balancing) and monitor for terminated listeners that need respawning. It also acts as a shutdown co-ordinator.
3. VbFcgiApp.dll - This is the FCGI Application code. Typically this will be the only file you need to work with in order to create your FCGI application (the VbFcgi.dll and VbFcgiHost.exe projects will hopefully be reusable from project to project). The version inclded here is just a very basic proof-of-concept that will send an HTML page upstream with a table of the FCGI parameters that were received.

While the above DLLs are COM ActiveX libraries, you do NOT need to register them with regsvr32 since this code uses Olaf Schmidt's registration-free DirectCOM library. No harm will be done if you do register the DLLs though, it's just not required.

Also included is a binary build of Nginx with a basic configuration to support a single FCGI host server listener on localhost:9100. This is inclded for the sake of convenience and to demonstrate a minimal configuration. You should have your own properly configured Nginx (or other web server) running in most cases.

Lastly, I've also bundled Olaf Schmidt's excellent vbRichClient5 library (http://www.vbrichclient.com/), again for the sake of convenience. You can always get the latest version from the vbRichClient5 website.
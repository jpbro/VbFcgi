# VbFcgi
Create new web apps using your VB6 coding knowledge, or get your existing VB6 client/server apps on the web with this FCGI Host/Server Framework for Visual Basic 6 (VB6) ActiveX/COM DLL FCGI Applications!

# Introduction
VbFcgi is a framework that allows you to easily get VB6 code onto the web. It was developed against Nginx, but should work with any web server that implements the FCGI spec.

# Included Binaries
There are 3 main components of the VbFcgi framework:
1. **VbFcgiLib.dll** - This is the main framework library that includes all of the code for listening and responding to FCGI requests from the web server, as well as parsing out records for FCGI parameters, HTTP cookies, etc... This file should be included with every distribution of your FCGI application.
2. **VbFcgiHost.exe** - This is main executable file that will spawn FCGI listeners as a broker between your webserver and your FCGI application. It includes support for running multiple listeners on sequential ports for load balancing, and it also monitors for terminated listeners that need respawning. Lastly, it also acts as a shutdown co-ordinator for all running FCGI listener instances. This file should be included with every distribution of your FCGI application.
3. **VbFcgiApp.dll** - This is the demo FCGI Application code. The version included here is a very basic proof-of-concept that will send an HTML page upstream with a table of the FCGI parameters that were received, also demonstrates the basic usage of cookies and HTTP query parameters  This file should **not** be included when distributing your own FCGI application! Instead you should create your own version as described in the *Creating your own FCGI Application* section below.

While the above DLLs are COM ActiveX libraries, you do NOT need to register them with regsvr32 when deploying to users since this code uses Olaf Schmidt's registration-free DirectCOM library. You should however register the above DLLs on your development machine.

Also included is a binary build of Nginx with a basic configuration to support a single FCGI host server listener on localhost:9100. This is included for the sake of convenience and to demonstrate a minimal configuration. You should have your own properly configured Nginx (or other web server) running in most cases.

Lastly, I've also bundled Olaf Schmidt's excellent vbRichClient5 library (http://www.vbrichclient.com/), again for the sake of convenience. You can always get the latest version from the vbRichClient5 website.

# Demo Usage
1. If you don't already have a web server running, start nginx from the command-line by going to the .\VbFcgi\bin\nginx folder and then rnning the nginx.exe command. If you already have a web server running, make sure it is configured to pass *.fcgi requests from the browser upstream to 127.0.0.1 port 9100.
2. From the command line, start VbFcgiHost.exe with the following command: vbfcgihost.exe /host 127.0.0.1 /port 9100 /spawn 1
3. Open your browser and go to http://127.0.0.1/vbfcgiapp.fcgi - you should see the HTML response from the demo FCGI application.

# Creating your own FCGI Application
You can use the included VbFcgiApp source code as a starting point - all the work is done in the IFcgiApp_ProcessRequest method, so give it a thorough review.

In order to write your own FCGI application from scratch, you must:

1. Start a new ActiveX DLL project in VB6.
2. Change the name of the project from "Project1" to "MyFcgiApp" (or whatever name you would like it to have).
3. Change the name of "Class1" to "CFcgiApp".
4. Add a reference to VbFcgiLib from the Projects menu > References.
5. In the General section of the "CFcgiApp" class, type; Implements VbFcgiLib.IFcgiApp
6. Select "IFcgiApp" from the drop down list in code view. It will create the IFcgiApp_ProcessRequest method for you.
7. Code your app in the IFcgiApp_ProcessRequest method (the rest of the f*cking owl).
8. Build your DLL app. 
9. Make a copy of the built DLL and change the extension to .fcgi.
10. Move the .fcgi file to the same folder as the VbFcgiHost.exe and VbFcgiLib.dll files.

**NOTE:** You do **not** need to register your FCGI application DLL, nor VbFcgiLib.dll when distributing it as registration-free instantiation is used by this framework.

When you subsequently run the VbFcgiHost.exe, it will use your .fcgi as a "plugin" (of sorts) for responding to correspondng FCGI requests. For example, typing http://localhost/myapp.fcgi will cause VbFcgiHost to create an instance of the CFcgiApp class from the myapp.fcgi DLL stored in the same folder, and then it will call IFcgiApp_ProcessRequest in that class.

Enjoy!
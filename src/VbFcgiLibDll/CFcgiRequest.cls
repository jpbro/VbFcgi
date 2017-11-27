VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFcgiRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mo_Http As VbFcgiLib.CHttp
Private mo_Fcgi As VbFcgiLib.CFcgi
Private mo_Upstream As VbFcgiLib.CFcgiUpstream

Friend Property Set FcgiUpstream(po_Upstream As VbFcgiLib.CFcgiUpstream)
   Set mo_Upstream = po_Upstream
   Set mo_Http.FcgiUpstream = po_Upstream
   Set mo_Fcgi.FcgiUpstream = po_Upstream
End Property

Public Property Get Fcgi() As VbFcgiLib.CFcgi
   Set Fcgi = mo_Fcgi
End Property

Public Property Get Http() As VbFcgiLib.CHttp
   Set Http = mo_Http
End Property

Private Sub Class_Initialize()
   Set mo_Http = New VbFcgiLib.CHttp
   Set mo_Fcgi = New VbFcgiLib.CFcgi
End Sub

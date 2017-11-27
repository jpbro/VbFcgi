VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFcgi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mo_FcgiUpstream As VbFcgiLib.CFcgiUpstream

Friend Property Set FcgiUpstream(po_FcgiUpstream As VbFcgiLib.CFcgiUpstream)
   Set mo_FcgiUpstream = po_FcgiUpstream
End Property

Public Property Get Stdin() As VbFcgiLib.CFcgiStdIn
   Set Stdin = mo_FcgiUpstream.Stdin
End Property

Public Property Get Params() As VbFcgiLib.CFcgiParams
   Set Params = mo_FcgiUpstream.Params
End Property

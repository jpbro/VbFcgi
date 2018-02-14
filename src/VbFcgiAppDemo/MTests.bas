Attribute VB_Name = "MTests"
Option Explicit

Public Sub TestForm()
   Dim lo_Form As New frmImages
   
   Load lo_Form
   
   lo_Form.Show 'vbModal
   
   'Unload lo_Form
End Sub

Public Sub TestSimulator()
   Dim lo_Sim As VbFcgiLib.CSimulator
   
   Set lo_Sim = New VbFcgiLib.CSimulator
   
   lo_Sim.SimulateRequest "http://localhost/vbfcgiapp.fcgi?json_getdata=1", New VbFcgiApp.CFcgiApp
End Sub

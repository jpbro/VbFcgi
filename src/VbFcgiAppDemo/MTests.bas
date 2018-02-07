Attribute VB_Name = "MTests"
Option Explicit

Public Sub TestForm()
   Dim lo_Form As New frmImages
   
   Load lo_Form
   
   lo_Form.Show 'vbModal
   
   'Unload lo_Form
End Sub


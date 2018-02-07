VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImages 
   Caption         =   "Form1"
   ClientHeight    =   3380
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   3380
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2320
      Left            =   3660
      ScaleHeight     =   2280
      ScaleWidth      =   2850
      TabIndex        =   1
      Top             =   150
      Width           =   2890
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3690
      Top             =   2490
      _ExtentX        =   953
      _ExtentY        =   953
      BackColor       =   -2147483643
      ImageWidth      =   800
      ImageHeight     =   600
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":6F85
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImages.frx":185B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   2840
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   3370
   End
End
Attribute VB_Name = "frmImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This form simulates part of a legacy VB6 application where all logic is programmed
' directly into the form. We can interact with a hidden copy of the form from our VBFCGI application
' and send dynamically generated data downstream to the browser.
' In the case of this demo, we will send an image whenever the browser user clicks an item in an option list.

Private Sub Form_Load()
   With Me.List1
      .AddItem "Blue Hills"
      .AddItem "Sunset"
      .AddItem "Water Lilies"
      
      .ListIndex = 0
   End With
End Sub

Private Sub List1_Click()
   ' Change the Picture in the PictureBox when list box item is clicked
   Me.Picture1.AutoRedraw = True
   Set Me.Picture1.Picture = Me.ImageList1.ListImages.Item(Me.List1.ListIndex + 1).Picture
   Me.Picture1.AutoRedraw = False
End Sub

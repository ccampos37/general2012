VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmError 
   Caption         =   "Error"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichError 
      Height          =   3780
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   6668
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmError.frx":0000
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

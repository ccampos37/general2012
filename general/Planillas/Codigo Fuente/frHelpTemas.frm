VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frHelpTemas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pago de Gratificaciones"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog xPrinter 
      Left            =   5400
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   330
      Left            =   4455
      TabIndex        =   3
      Top             =   4110
      Width           =   1440
   End
   Begin RichTextLib.RichTextBox Help 
      Height          =   3255
      Left            =   180
      TabIndex        =   2
      Top             =   765
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   5741
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      RightMargin     =   1
      AutoVerbMenu    =   -1  'True
      FileName        =   "C:\Entsoft\hlp001.rtf"
      TextRTF         =   $"frHelpTemas.frx":0000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pago de Gratificaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   690
      TabIndex        =   1
      Top             =   300
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pago de Gratificaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   0
      Left            =   705
      TabIndex        =   0
      Top             =   315
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frHelpTemas.frx":09E0
      Top             =   105
      Width           =   480
   End
End
Attribute VB_Name = "frHelpTemas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub COMMAND1_CLICK()
 On Error GoTo ErrCancelar
    xPrinter.CancelError = True
    xPrinter.ShowPrinter
    MsgBox Printer.Width & Printer.Height
    'Help.SelPrint Printer.hDC
    
ErrCancelar:
    Exit Sub
End Sub


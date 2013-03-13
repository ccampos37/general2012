VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form RptReporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver Documento"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   13620
   Begin VB.Frame Fr1 
      Height          =   915
      Left            =   120
      TabIndex        =   1
      Top             =   7890
      Width           =   3165
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Exporta"
         Height          =   690
         Index           =   11
         Left            =   1110
         Picture         =   "RptReporte.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   870
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Cancelar"
         Height          =   690
         Index           =   12
         Left            =   2070
         Picture         =   "RptReporte.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Imprimir"
         Height          =   690
         Index           =   3
         Left            =   180
         Picture         =   "RptReporte.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   825
      End
   End
   Begin RichTextLib.RichTextBox RBox1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   13996
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"RptReporte.frx":0CC6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   4320
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "RptReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim warchi As String

Public Property Let BArchi(pdata)
   warchi = Trim(pdata)
End Property


Private Sub cmdBotones_Click(Index As Integer)
  Dim j As Integer
  
  On Error GoTo nimpre
  
  Select Case Index
    Case 3   'Impresion
        Printer.Print RBox1.Text
        Printer.EndDoc
    
'       Call aImpresora(warchi)
    Case 11  ' exportar
       Dialog1.DialogTitle = "Exporta Datos"
       Dialog1.DefaultExt = "*.txt"
       Dialog1.Filter = "(*.txt)|*.txt|(*.*)|*.*"
       Dialog1.ShowSave
       RBox1.SaveFile Dialog1.FileName, rtfRTF
    Case 12  'Cerrar
      Unload Me
  End Select
  
nimpre:
   If Err Then
      MsgBox "No existe Impresora en linea...Verifique!!", vbInformation, MsgTitle
      Err = 0
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
  MostrarForm Me, "C2"
 ' RBox1.Font.Name = "courier new"
 ' RBox1.Font.Size = 9
  RBox1.RightMargin = 20000
  RBox1.FileName = warchi
  DoEvents
  
End Sub


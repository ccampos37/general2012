VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frPrevBol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preliminar de Boletas de Remuneraciones"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   Icon            =   "frPrevBol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   4845
      Width           =   1560
   End
   Begin MSDataGridLib.DataGrid DGPrev 
      Height          =   4650
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   8202
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         SizeMode        =   1
         ScrollGroup     =   2
         Size            =   2
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frPrevBol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents RsPrev As ADODB.Recordset
Attribute RsPrev.VB_VarHelpID = -1

Private Sub Command1_Click()
    FormatDG
End Sub

Private Sub Form_Load()
    Set RsPrev = New ADODB.Recordset
    RsPrev.Open "CalcInput", DBSYSTEM, adOpenStatic
    Set DGPrev.DataSource = RsPrev
    FormatDG
End Sub

Public Sub FormatDG()
    Dim X As Integer
    With DGPrev
        For X = 3 To .Columns.Count - 1
            .Columns(X).NumberFormat = "##,##0.00"
            .Columns(X).Alignment = dbgRight
        Next
    End With
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RsPrev = Nothing
End Sub

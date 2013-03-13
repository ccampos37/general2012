VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frProgVacGrp 
   Caption         =   "Programación de Vacaciones en Grupo"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7080
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frProgVacGrp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DGTrabs 
      Height          =   3300
      Left            =   120
      TabIndex        =   8
      Top             =   1770
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5821
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   0
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
      Caption         =   "Trabajadores a Programar"
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
         MarqueeStyle    =   4
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5670
      TabIndex        =   7
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Programar"
      Height          =   375
      Left            =   5670
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmSelecTrab 
      Caption         =   "Seleccion (F5)"
      Height          =   990
      Left            =   6000
      Picture         =   "frProgVacGrp.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1770
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo Vacacional Correspondiente a:"
      Height          =   810
      Left            =   120
      TabIndex        =   0
      Top             =   885
      Width           =   6840
      Begin VB.OptionButton Option2 
         Caption         =   "Año Siguiente"
         Height          =   285
         Left            =   4365
         TabIndex        =   3
         Top             =   390
         Width           =   2025
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Actual"
         Height          =   285
         Left            =   2520
         TabIndex        =   2
         Top             =   390
         Value           =   -1  'True
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Programación del Año"
         Height          =   195
         Left            =   285
         TabIndex        =   1
         Top             =   405
         Width           =   1560
      End
   End
   Begin VB.Label Label2 
      Caption         =   $"frProgVacGrp.frx":0D0C
      Height          =   480
      Left            =   1140
      TabIndex        =   4
      Top             =   180
      Width           =   5520
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   420
      Picture         =   "frProgVacGrp.frx":0DA0
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "frProgVacGrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSTRABS As New ADODB.Recordset

Private Sub FORM_KEYDOWN(KEYCODE As Integer, SHIFT As Integer)
    If KEYCODE = vbKeyF5 Then CMSELECTRAB_CLICK
End Sub

Private Sub FORM_LOAD()
    Option1.Caption = "Año Actual (" & Year(Date) & ")"
    Option2.Caption = "Año Siguiente (" & Year(Date) + 1 & ")"
    RSTRABS.Open "_TMPSELECT", DBSYSTEM, adOpenStatic, adLockReadOnly
    Set DGTrabs.DataSource = RSTRABS
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSTRABS = Nothing
End Sub

Private Sub CMSELECTRAB_CLICK()
    Dim RSDELS As New ADODB.Recordset
    frSelect.Show 1
    RSTRABS.Requery
    Set DGTrabs.DataSource = RSTRABS
End Sub


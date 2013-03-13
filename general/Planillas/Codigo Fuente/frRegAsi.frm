VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frRegAsi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Asistencia"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "frRegAsi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleMode       =   0  'User
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DGFiltro 
      Height          =   2430
      Left            =   1275
      TabIndex        =   8
      Top             =   165
      Visible         =   0   'False
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   4286
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483633
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   3
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
      Caption         =   "Trabajadores Seleccionados"
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
         AllowRowSizing  =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmSelecTrab 
      Caption         =   "Seleccionar (F5)"
      Height          =   1080
      Left            =   120
      Picture         =   "frRegAsi.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   180
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Especificaciones"
      Height          =   1455
      Left            =   165
      TabIndex        =   0
      Top             =   2790
      Width           =   5160
      Begin VB.OptionButton Option1 
         Caption         =   "Ingreso con Centro de Costo"
         Height          =   210
         Left            =   2595
         TabIndex        =   10
         Top             =   825
         Width           =   2385
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   300
         Left            =   3690
         TabIndex        =   6
         Top             =   352
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   53018625
         CurrentDate     =   36699
      End
      Begin MSComCtl2.DTPicker XFechaIni 
         Height          =   300
         Left            =   1140
         TabIndex        =   5
         Top             =   337
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   53018625
         CurrentDate     =   36699
      End
      Begin VB.OptionButton Op1 
         Caption         =   "Ingresar por Trabajador"
         Height          =   240
         Index           =   0
         Left            =   345
         TabIndex        =   2
         Top             =   810
         Value           =   -1  'True
         Width           =   2070
      End
      Begin VB.OptionButton Op1 
         Caption         =   "Ingresar por Concepto"
         Height          =   255
         Index           =   1
         Left            =   345
         TabIndex        =   1
         Top             =   1065
         Width           =   2130
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   2730
         TabIndex        =   3
         Top             =   390
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   615
      Left            =   360
      TabIndex        =   11
      Top             =   4320
      Width           =   4500
      Begin VB.CommandButton cmContinuar 
         Caption         =   "&Continuar"
         Height          =   405
         Left            =   585
         TabIndex        =   13
         Top             =   45
         Width           =   1410
      End
      Begin VB.CommandButton cmSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   405
         Left            =   2475
         TabIndex        =   12
         Top             =   45
         Width           =   1410
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Todos los trabajadores de la empresa seleccionada"
      Height          =   510
      Left            =   1275
      TabIndex        =   9
      Top             =   450
      Width           =   4035
   End
End
Attribute VB_Name = "frRegAsi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##TMPSELECT
'MODIFICADO  10-07-2001
'ESTADO:OK
Option Explicit
Dim RSAUX As New ADODB.Recordset
Private Sub CMCONTINUAR_CLICK()
    If Not ExisteTablaAux(" [##TMPSELECT" & VGL_COMPUTER & "] ") Then MsgBox "SELECCIONE LOS TRABAJADORES": Exit Sub
    If Option1.Value Then
        MsgBox "No presenta este modulo en su Versión del Sistema", vbInformation
    Else
        If xFechaFin.Value < xFechaIni.Value Then
            MsgBox "La Fecha de Inicio debe ser mayor o igual a la Fecha Final", vbCritical
            Exit Sub
        End If
        frAdAsis.Show 1
    End If
End Sub
Private Sub CMSALIR_CLICK()
    Unload Me
End Sub
Private Sub CMSELECTRAB_CLICK()
    REGSELECT.FECHACESEMAX = xFechaFin.Value
    REGSELECT.FECHAINIMAX = xFechaFin.Value
    REGSELECT.FECHAINI = xFechaIni.Value
    REGSELECT.USARFECHACESE = True
    frSelect.Show 1
    REGSELECT.USARFECHACESE = True
    Set RSAUX = Nothing
    RSAUX.Open " [##TMPSELECT" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic
    Set DGFiltro.DataSource = RSAUX
    If RSAUX.RecordCount = 0 Then
        DGFiltro.Visible = False
    Else
        DGFiltro.Visible = True
    End If
    RSAUX.ActiveConnection = Nothing
End Sub
Private Sub FORM_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = vbKeyF5 Then CMSELECTRAB_CLICK
End Sub
Private Sub Form_Load()
    xFechaIni.Value = Date
    xFechaFin.Value = Date
End Sub

Private Sub Form_Resize()
If Me.Width < 5565 Then Exit Sub
If Me.Height < 5415 Then Exit Sub
'me.scaleHeigth=5010
'me.scaleWith=5445
'***********************************************
Frame1.TOP = Me.ScaleHeight - 2220
Frame1.Left = Me.ScaleWidth - 5280
'***********************************************
Frame2.TOP = Me.ScaleHeight - 690
Frame2.Left = Me.ScaleWidth - 5085
'*********************************************
DGFiltro.Height = Me.ScaleHeight - 2580
DGFiltro.Width = Me.ScaleWidth - 1410
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSAUX = Nothing
End Sub


VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRRHH002 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de Derechohabientes"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "frmRRHH002.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7080
   Begin Crystal.CrystalReport Reporte 
      Left            =   6645
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmTitulo 
      Appearance      =   0  'Flat
      Caption         =   "Titulo del Reporte (Generar)"
      Height          =   285
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4065
      Width           =   2295
   End
   Begin VB.TextBox xTitulo 
      Height          =   735
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   28
      Top             =   4410
      Width           =   3180
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Listado General"
      Height          =   390
      Left            =   5175
      TabIndex        =   27
      Top             =   4590
      Width           =   1650
   End
   Begin VB.CommandButton cmSelecTrab 
      Caption         =   "Seleccion (F5)"
      Height          =   990
      Left            =   3600
      Picture         =   "frmRRHH002.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Seleccione aquí los trabajadores de quienes obtendrá sus derechohabientes"
      Top             =   4185
      Width           =   945
   End
   Begin VB.Frame Frame5 
      Caption         =   "Incluir"
      Height          =   1590
      Left            =   4725
      TabIndex        =   21
      Top             =   2310
      Width           =   2100
      Begin VB.OptionButton Option8 
         Caption         =   "Individuales"
         Height          =   195
         Left            =   375
         TabIndex        =   24
         Top             =   1245
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Agrupados"
         Height          =   195
         Left            =   375
         TabIndex        =   23
         Top             =   915
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CheckBox xInfoTrab 
         Caption         =   "Información del Trabajador"
         Height          =   375
         Left            =   375
         TabIndex        =   22
         Top             =   405
         Value           =   1  'Checked
         Width           =   1410
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cerrar"
      Height          =   390
      Left            =   5175
      TabIndex        =   20
      Top             =   5055
      Width           =   1650
   End
   Begin VB.Frame Frame4 
      Caption         =   "Por Fecha de Nacimiento"
      Height          =   1590
      Left            =   255
      TabIndex        =   13
      Top             =   2310
      Width           =   4320
      Begin MSComCtl2.DTPicker xFecha 
         Height          =   315
         Left            =   2445
         TabIndex        =   16
         Top             =   450
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62062593
         CurrentDate     =   36870
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmRRHH002.frx":1194
         Left            =   1095
         List            =   "frmRRHH002.frx":119E
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tiempo transcurrido con la fecha actual"
         Height          =   195
         Left            =   255
         TabIndex        =   18
         Top             =   915
         Width           =   2805
      End
      Begin VB.Label xTiempo 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   255
         TabIndex        =   17
         Top             =   1155
         Width           =   3810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nacidos"
         Height          =   195
         Left            =   255
         TabIndex        =   14
         Top             =   495
         Width           =   585
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Situación"
      Height          =   2070
      Left            =   4725
      TabIndex        =   9
      Top             =   105
      Width           =   2100
      Begin VB.OptionButton Option6 
         Caption         =   "Ambos"
         Height          =   270
         Left            =   360
         TabIndex        =   12
         Top             =   1230
         Width           =   1200
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Baja"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   885
         Width           =   1320
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Activos"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   465
         Value           =   -1  'True
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sexo"
      Height          =   2055
      Left            =   2460
      TabIndex        =   5
      Top             =   105
      Width           =   2100
      Begin VB.OptionButton Option3 
         Caption         =   "Ambos"
         Height          =   195
         Left            =   330
         TabIndex        =   8
         Top             =   1215
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Femenino"
         Height          =   225
         Left            =   330
         TabIndex        =   7
         Top             =   825
         Width           =   1545
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Masculino"
         Height          =   270
         Left            =   330
         TabIndex        =   6
         Top             =   420
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Incluir en el reporte"
      Height          =   2025
      Left            =   255
      TabIndex        =   0
      Top             =   120
      Width           =   2070
      Begin VB.CheckBox Check4 
         Caption         =   "Gestante"
         Height          =   195
         Left            =   255
         TabIndex        =   4
         Top             =   1530
         Width           =   1260
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Concubino"
         Height          =   195
         Left            =   255
         TabIndex        =   3
         Top             =   1175
         Width           =   1260
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Conyuge"
         Height          =   195
         Left            =   255
         TabIndex        =   2
         Top             =   820
         Width           =   1260
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Hijos"
         Height          =   195
         Left            =   255
         TabIndex        =   1
         Top             =   465
         Value           =   1  'Checked
         Width           =   1260
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Por edades"
      Height          =   390
      Left            =   5175
      TabIndex        =   19
      Top             =   4125
      Width           =   1650
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Selecciones los Trabajadores"
      Height          =   195
      Left            =   2460
      TabIndex        =   26
      Top             =   5265
      Width           =   2085
   End
End
Attribute VB_Name = "frmRRHH002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMSELECTRAB_CLICK()
    REGSELECT.USARFECHACESE = False
    frSelect.Show 1
End Sub

Private Sub CMTITULO_Click()
    xTitulo.Text = "LISTADO DE " & ARMATITULO
End Sub

Private Sub Command1_Click()
    If (Check1.Value + Check2.Value + Check3.Value + Check4.Value) = 0 Then
        MsgBox "DEBE DE SELECCIONAR POR LO MENOS UN TIPO DE DERECHOHABIENTE", vbInformation
        Exit Sub
    End If
    If xTitulo.Text = "" Then
        MsgBox "NO SE PUEDE IMPRIMIR SI NO CONTIENE UN TÍTULO EN EL REPORTE", vbInformation
        Exit Sub
    End If
    Dim A As String
    A = ARMATITULO()
   ' MSGBOX FRAME1.TAG
    With Reporte
        .Reset
        'EXIT SUB
        If Frame1.Tag <> "" Then .SelectionFormula = Frame1.Tag
        .ReportFileName = REGSISTEMA.REPORTES & "PLNAS004.RPT"
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XTITULO='" & xTitulo.Text & "'"
        .WindowTitle = "PLNAS004- POR EDADES"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .DataFiles(2) = App.PATH & "\BDAUXCOM.MDB"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        If .Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub FORM_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = vbKeyF5 Then CMSELECTRAB_CLICK
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 0
    xFecha.Value = Date
    xFecha.MaxDate = Date
End Sub

Private Sub XFECHA_CHANGE()
    Dim A As Integer, B As Integer, C As Integer
    TiempoTrans xFecha.Value, Date, A, B, C
    xTiempo.Caption = A & " AÑOS, " & B & " MESES Y " & C & " DÍAS"
End Sub

Public Function ARMATITULO() As String
    Dim CAD As String, FILTRO As String
    CAD = ""
    If Check1.Value = 1 Then
        CAD = " HIJOS"
        FILTRO = "0"
    End If
    If Check2.Value = 1 Then
        CAD = CAD & " - CONYUGE"
        FILTRO = FILTRO & IIf(FILTRO = "", "", ",") & "1"
    End If
    If Check3.Value = 1 Then
        CAD = CAD & " - CONCUBINO"
        FILTRO = FILTRO & IIf(FILTRO = "", "", ",") & "2"
    End If
    If Check4.Value = 1 Then
        CAD = CAD & " - GESTANTE"
        FILTRO = FILTRO & IIf(FILTRO = "", "", ",") & "3"
    End If
    If CAD = " HIJOS - CONYUGE - CONCUBINO - GESTANTE" Then
        CAD = " TODOS LOS DERECHOHABIENTES"
        FILTRO = ""
    Else
        FILTRO = "({FAMILIAR.VINCULO} IN [" & FILTRO & "])"
    End If
    If Not Option3.Value Then
        If Option1.Value Then
            CAD = CAD & " DE SEXO MASCULINO"
            FILTRO = FILTRO & IIf(FILTRO = "", "{FAMILIAR.SEXO}=0", " AND {FAMILIAR.SEXO}=0")
        Else
            CAD = CAD & " DE SEXO FEMENINO"
            FILTRO = FILTRO & IIf(FILTRO = "", "{FAMILIAR.SEXO}=1", " AND {FAMILIAR.SEXO}=1")
        End If
    End If
    If Not Option6.Value Then
        If Option4.Value Then
            CAD = CAD & " DE SITUACIÓN ACTIVO"
            FILTRO = FILTRO & IIf(FILTRO = "", "{FAMILIAR.SITUACION} = 0", " AND {FAMILIAR.SITUACION} = 0")
        Else
            CAD = CAD & " DE SITUACIÓN DE BAJA"
            FILTRO = FILTRO & IIf(FILTRO = "", "{FAMILIAR.SITUACION} = 1", " AND {FAMILIAR.SITUACION} = 1")
        End If
    End If
    FILTRO = FILTRO & IIf(FILTRO = "", "", " AND ") & "{FAMILIAR.FECHANAC}" & IIf(Combo1.ListIndex = 0, "<", ">") & " DATE(" & xFecha.Year & "," & xFecha.Month & "," & xFecha.Day & ")"
    Frame1.Tag = FILTRO
    ARMATITULO = CAD
End Function

Private Sub XTITULO_GOTFOCUS()
    If xTitulo.Text = "" Then CMTITULO_Click
End Sub


VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form MantSCat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sub Categorias Eventos"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport RptBanco 
      Left            =   1560
      Top             =   2070
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "MantSCat.frx":0000
      Height          =   3795
      Left            =   75
      TabIndex        =   0
      Top             =   930
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   6694
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      Caption         =   "Sub Categorias Eventos"
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "MantSCat.frx":0015
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   $"MantSCat.frx":0457
      ForeColor       =   &H8000000E&
      Height          =   585
      Left            =   750
      TabIndex        =   1
      Top             =   120
      Width           =   3720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000001&
      Height          =   825
      Left            =   15
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "MantSCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSBANCOS As New ADODB.Recordset
Dim REGACT As REGWIN
Private Sub FORM_ACTIVATE()
    ActivarTools REGACT
End Sub
Private Sub FORM_KEYDOWN(KEYCODE As Integer, SHIFT As Integer)
If KEYCODE = vbKeyEscape Then
    Unload Me
End If
End Sub
Private Sub FORM_LOAD()
If Not REGSISTEMA.VALRRHH Then Exit Sub
Dim SQLSTR As String
SQLSTR = "SELECT COD_SBCA AS CODIGO, DES_SBCA AS SUBCATEGORIA FROM SUBCATEGORIA"
    RSBANCOS.Open SQLSTR, DBADMINPER, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = RSBANCOS
    With REGACT
        .BUSCAR = False
        .EDITAR = False
        .ELIMINAR = True
        .FILTRAR = False
        .IMPRIMIR = True
        .NUEVO = False
        .PRELIMINAR = True
    End With
End Sub
Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSBANCOS = Nothing
End Sub
Public Sub COMANDOTOOLBAR(ByVal COMANDO As String)
    Select Case UCase(COMANDO)
        Case "IMPRIMIR", "PRELIMINAR"
            With RptBanco
                .ReportFileName = REGSISTEMA.REPORTES & "PLAN0006.RPT"
                .DataFiles(0) = REGSISTEMA.PATHEMPRESA & "\PLANILLA.MDB"
                If UCase(COMANDO) = "IMPRIMIR" Then
                    .Destination = crptToPrinter
                Else
                    .Destination = crptToWindow
                    .WindowState = crptMaximized
                    .WindowShowPrintBtn = True
                    .WindowShowRefreshBtn = True
                    .WindowShowSearchBtn = True
                    .WindowShowPrintSetupBtn = True
                    .WindowTitle = "PLAN0006 - TABAJADORES AFILIADOS A LOS BANCOS"
                End If
                .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
                .Formulas(2) = "XHORA='" & Format(Time, "HH:MM") & "'"
                .PrintReport
            End With
    End Select
End Sub


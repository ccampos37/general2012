VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmAyuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "r"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6090
   Begin VB.CommandButton CmdIng 
      Caption         =   "&Selec."
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   675
      Left            =   1755
      Picture         =   "FrmAyuda.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3255
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "S&alir"
      CausesValidation=   0   'False
      Height          =   675
      Left            =   3435
      Picture         =   "FrmAyuda.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3255
      Width           =   775
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2415
      Left            =   150
      TabIndex        =   0
      Top             =   735
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      DefColWidth     =   167
      HeadLines       =   1
      RowHeight       =   17
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Tclave"
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
         DataField       =   "tdescri"
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
            Alignment       =   2
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   135
      TabIndex        =   1
      Top             =   0
      Width           =   5775
      Begin VB.TextBox TxFiltro 
         Height          =   300
         Left            =   720
         TabIndex        =   3
         Text            =   "TxFiltro"
         Top             =   225
         Width           =   1815
      End
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmAyuda.frx":0884
         Left            =   3600
         List            =   "FrmAyuda.frx":088E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   1935
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar :"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   270
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCod As String ' Tcod
Public cC As String   ' Tclave
Public cD As String   ' Tdescri
Dim adodc1 As ADODB.Recordset
Dim nCom As Integer
Dim nCursor As Integer

Private Sub CmbOrden_Click()
TxFiltro = ""
nCom = CmbOrden.ListIndex
Set adodc1 = New ADODB.Recordset
Select Case nCom
Case 0
    adodc1.Open "SELECT Tclave,Tdescri FROM TABAYU WHERE TCOD = '" & cCod & "' ORDER BY TCLAVE", VGcnx, adOpenStatic
Case 1
    adodc1.Open "SELECT Tclave,Tdescri FROM TABAYU WHERE TCOD = '" & cCod & "' ORDER BY TDESCRI", VGcnx, adOpenStatic
End Select
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh
CarObj (cCod)
DataGrid1.SetFocus
End Sub

Private Sub CmdIng_Click()      ' Seleccionar
If adodc1.RecordCount > 0 Then
    cC = adodc1("tclave")
    cD = adodc1("tdescri")
    CmdSalir_Click
End If
End Sub

Private Sub CmdSalir_Click()    ' Salir
Unload Me
End Sub

Private Sub DataGrid1_DblClick()
CmdIng_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
    If Len(TxFiltro) - 1 > 0 Then
        TxFiltro = Left(TxFiltro, Len(TxFiltro) - 1)
    Else
        TxFiltro = ""
    End If
    KeyAscii = 0
ElseIf KeyAscii <> 13 Then
    TxFiltro = TxFiltro & Chr(KeyAscii)
End If
End Sub

Private Sub Form_Activate()
TxFiltro = ""
CmbOrden.ListIndex = 0
DataGrid1.SetFocus
End Sub

Private Sub Form_Load()
central Me                                 ' Centra el Formulario
Init_ControlDataGrid DataGrid1
DataGrid1.ClearFields                       ' Limpia las Columnas
Set adodc1 = New ADODB.Recordset
adodc1.Open "SELECT Tclave,Tdescri FROM TABAYU WHERE TCOD = '" & cCod & "' ORDER BY TCLAVE", VGcnx, adOpenStatic
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh
CarObj (cCod)                               ' Objetos
cC = "": cD = ""
End Sub

Private Sub CarObj(cTabla As String)        ' Carga Objetos
DataGrid1.Columns(0).Locked = True
DataGrid1.Columns(0).WrapText = True
DataGrid1.Columns(0).Alignment = dbgCenter
DataGrid1.Columns(0).Caption = "            CODIGO"
DataGrid1.Columns(1).Caption = "    DESCRIPCION"
DataGrid1.Columns(0).Locked = False
DataGrid1.Columns(0).WrapText = False

Select Case cTabla
Case "13"
      Me.Caption = "DISTRITOS"
Case "58"
      Me.Caption = "TERRITORIO"
Case "59"
      Me.Caption = "RUTA"
Case "60"
      Me.Caption = "SEGMENTO"
Case "61"
      Me.Caption = "UBICACION DE SEGMENTO"
Case "67"
      Me.Caption = "TIPO DE CLIENTE"
Case "62"
      Me.Caption = "GIRO DEL NEGOCIO"
Case "22"
      Me.Caption = "TIPO DE VENTA"
Case "28"
      Me.Caption = "ZONAS"
Case "08"
      Me.Caption = "TIPO DE ARTICULOS"
Case "25"
      Me.Caption = "TARJETAS DE CREDITO"
End Select
End Sub
Private Sub TxFiltro_Change()           ' Filtro
If adodc1.RecordCount > 0 Then
    If Trim(TxFiltro) <> "" Then
        nCursor = adodc1.Bookmark
        adodc1.AbsolutePosition = 1
    
        adodc1.MoveFirst
        If CmbOrden.ListIndex = 0 Then
            adodc1.Find "TCLAVE like '" & Trim(UCase(TxFiltro)) & "*'"
        Else
            adodc1.Find "TDESCRI like '" & Trim(UCase(TxFiltro)) & "*'"
        End If
        If adodc1.EOF Then adodc1.AbsolutePage = nCursor
    End If
End If
End Sub

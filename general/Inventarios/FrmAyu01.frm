VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmAyu01 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "r"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6735
   Begin VB.CommandButton CmdIng 
      Caption         =   "&Selec."
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   672
      Left            =   72
      Picture         =   "FrmAyu01.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3744
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "S&alir"
      CausesValidation=   0   'False
      Height          =   636
      Left            =   972
      Picture         =   "FrmAyu01.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3780
      Width           =   775
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2772
      Left            =   72
      TabIndex        =   0
      Top             =   936
      Width           =   6588
      _ExtentX        =   11615
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      DefColWidth     =   167
      HeadLines       =   1
      RowHeight       =   18
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      Height          =   855
      Left            =   84
      TabIndex        =   1
      Top             =   12
      Width           =   6576
      Begin VB.TextBox TxFiltro 
         Height          =   288
         Left            =   1200
         TabIndex        =   3
         Text            =   "TxFiltro"
         Top             =   360
         Width           =   2280
      End
      Begin VB.ComboBox CmbOrden 
         Height          =   288
         ItemData        =   "FrmAyu01.frx":0884
         Left            =   4416
         List            =   "FrmAyu01.frx":088E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2076
      End
      Begin VB.Label Label32 
         Caption         =   "Filtro :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   120
         TabIndex        =   5
         Top             =   396
         Width           =   1092
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   3696
         TabIndex        =   4
         Top             =   396
         Width           =   852
      End
   End
End
Attribute VB_Name = "FrmAyu01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Adoreg1 As ADODB.Recordset
Public cCod As String ' Tcod
Public cC As String   ' Tclave
Public cD As String   ' Tdescri
Dim nCom As Integer

Private Sub CmbOrden_Click()
TxFiltro = ""
nCom = CmbOrden.ListIndex
Set Adoreg1 = New ADODB.Recordset
Select Case nCom
Case 0
    Adoreg1.Open "SELECT Tclave,Tdescri FROM TABAYU WHERE TCOD = '" & cCod & "' ORDER BY TCLAVE", VGcnx, adOpenStatic
    Label32.Caption = "CODIGO :"
Case 1
    Adoreg1.Open "SELECT Tclave,Tdescri FROM TABAYU WHERE TCOD = '" & cCod & "' ORDER BY TDESCRI", VGcnx, adOpenStatic
    Label32.Caption = "DESCRIPCION :"
End Select
'Adoreg1.Refresh
Set DataGrid1.DataSource = Adoreg1
CarObj (cCod)
DataGrid1.SetFocus
End Sub

Private Sub CmdIng_Click()      ' Seleccionar
If Adoreg1.RecordCount > 0 Then
    cC = Adoreg1("tclave")
    cD = Adoreg1("tdescri")
    CmdSalir_Click
End If
End Sub

Private Sub CmdSalir_Click()    ' Salir
Unload Me
End Sub



Private Sub Form_Activate()
TxFiltro = ""
CmbOrden.ListIndex = 0

DataGrid1.SetFocus
End Sub

Private Sub Form_Load()
central Me                                 ' Centra el Formulario
'Adoreg1.ConnectionString = "Provider = Microsoft.Jet.OLEDB.3.51;Data Source= " & cRuta2 & ""

DataGrid1.ClearFields                       ' Limpia las Columnas
Set Adoreg1 = New ADODB.Recordset
Adoreg1.Open "SELECT Tclave,Tdescri FROM TABAYU WHERE TCOD = '" & cCod & "' ORDER BY TCLAVE", VGcnx, adOpenStatic
Set DataGrid1.DataSource = Adoreg1
'Adoreg1.Refresh
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
DataGrid1.Refresh
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
Case "65"
      Me.Caption = "TIPO DE ATENCION"
Case "70"
      Me.Caption = "BANCO"
Case "22"
      Me.Caption = "TIPO DE VENTA"
Case "23"
      Me.Caption = "TIPO DE PRECIOS"
Case "03"
      Me.Caption = "MONEDAS"
Case "28"
      Me.Caption = "ZONAS"
Case "27"
      Me.Caption = "VENDEDOR"
Case "05"
      Me.Caption = "UNIDAD DE MEDIDA"
Case "38"
      Me.Caption = "FAMILIA DE ARTICULOS"
Case "39"
      Me.Caption = "LINEA DE ARTICULOS"
Case "06"
      Me.Caption = "GRUPO DE ARTICULOS"
Case "07"
      Me.Caption = "CUENTAS CONTABLES"
Case "04"
      Me.Caption = "DOC. REFERENCIA"
Case "08"
      Me.Caption = "TIPO DE ARTICULOS"
Case "25"
      Me.Caption = "TARJETAS DE CREDITO"
End Select
End Sub
Private Sub TxFiltro_Change()           ' Filtro
Set Adoreg1 = New ADODB.Recordset
If Trim(TxFiltro) = "" Then
  
    Select Case nCom
    Case 0
        Adoreg1.Open "SELECT Tclave,Tdescri FROM TABAYU WHERE TCOD = '" & cCod & "' ORDER BY TCLAVE", VGcnx, adOpenStatic
    Case 1
        Adoreg1.Open "SELECT Tclave,Tdescri FROM TABAYU WHERE TCOD = '" & cCod & "' ORDER BY TDESCRI", VGcnx, adOpenStatic
    End Select
    Set DataGrid1.DataSource = Adoreg1
    'Adoreg1.Refresh
    CarObj (cCod)
    TxFiltro.SetFocus
End If

If Trim(TxFiltro) <> "" Then
    Select Case nCom
    Case 0
            Adoreg1.Open " Select Tclave,Tdescri from TABAYU WHERE MID(TCLAVE,1," & Len(Trim(TxFiltro)) & ") ='" & Trim(UCase(TxFiltro)) & "' AND TCOD = '" & cCod & "' ORDER BY TCLAVE", VGcnx, adOpenStatic
    Case 1
            Adoreg1.Open " Select Tclave,Tdescri from TABAYU WHERE Mid(TDESCRI, 1," & Len(Trim(TxFiltro)) & ")  ='" & Trim(UCase(TxFiltro)) & "' AND TCOD = '" & cCod & "' ORDER BY TDESCRI", VGcnx, adOpenStatic
    End Select
   ' Adoreg1.Refresh
    Set DataGrid1.DataSource = Adoreg1
    CarObj (cCod)
    If Adoreg1.RecordCount = 0 Then
        TxFiltro.SetFocus
    End If
End If
End Sub

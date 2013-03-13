VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmCalidadAvio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calidad de Avio"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DBGrid1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   3413
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
         BeginProperty Column00 
            ColumnWidth     =   1530.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1530.142
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6885
      Begin VB.TextBox TxFiltro 
         Height          =   300
         Left            =   1200
         MaxLength       =   45
         TabIndex        =   9
         Text            =   "TxFiltro"
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmCalidadAvio.frx":0000
         Left            =   5040
         List            =   "FrmCalidadAvio.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar  :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   255
         Left            =   4320
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   6855
      Begin VB.CommandButton command5 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   3600
         Picture         =   "FrmCalidadAvio.frx":0027
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   775
      End
      Begin VB.CommandButton Command7 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5760
         Picture         =   "FrmCalidadAvio.frx":0469
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2520
         Picture         =   "FrmCalidadAvio.frx":08AB
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   360
         Picture         =   "FrmCalidadAvio.frx":0CED
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   1440
         Picture         =   "FrmCalidadAvio.frx":112F
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   775
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   450
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Visible         =   0   'False
      Width           =   6840
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3945
         Picture         =   "FrmCalidadAvio.frx":1571
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1800
         Picture         =   "FrmCalidadAvio.frx":19B3
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   480
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2265
         MaxLength       =   1
         TabIndex        =   14
         Top             =   645
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2265
         MaxLength       =   45
         TabIndex        =   13
         Top             =   1125
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "C�digo:"
         Height          =   255
         Left            =   660
         TabIndex        =   16
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Descripci�n:"
         Height          =   255
         Left            =   660
         TabIndex        =   15
         Top             =   1125
         Width           =   975
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   450
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   210
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "FrmCalidadAvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim resp As String
Dim nTra As Integer
Dim cBase As String
Dim rs As New ADODB.Recordset
Dim VGDllGeneral As New dllgeneral.dll_general

Private Sub command5_Click()
    Dim cadena As String
    Dim cNomRepor  As String

cNomRepor = "CalidadAvio.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Calidad de Avios"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport + cNomRepor

                        
    CrystalReport1.Connect = VGcadenareport2
    CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
    
    CrystalReport1.formulas(0) = "emp ='" & VGparametros.RucEmpresa & "'"
    
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowState = crptMaximized
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
Else
    MsgBox "No existe el nombre del Reporte, verifique en Formatos", vbInformation, "Informaci�n"
    Exit Sub
End If

End Sub

Private Sub Command9_Click()
      DBGrid1.Visible = True
      'Command19.Visible = True
      Frame5.Visible = True
      Frame2.Visible = True
      Frame1.Visible = False
      Frame3.Visible = False
      DBGrid1.SetFocus
End Sub

Private Sub Text2_Change()
Dim I As Integer
Text1.text = UCase(Text1.text)
I = Len(Text1.text)
Text1.SelStart = I

End Sub


Private Sub TxFiltro_Change()
Dim nstr As String

'If rs.RecordCount > 0 Then
    If Trim(TxFiltro) <> "" Then
        If CmbOrden.ListIndex = 0 Then
            nstr = "CodMat like '" & Trim(UCase(TxFiltro)) & "%'"
        ElseIf CmbOrden.ListIndex = 1 Then
            nstr = "descripcion like '" & Trim(UCase(TxFiltro)) & "%'"
        End If
        nstr = "select * from calidad_a where " & nstr
    Else
       nstr = "select * from calidad_a"
    End If
'Else
'    nstr = "select * from Material"
'End If
Call Listado(nstr)
End Sub

Private Sub CmbOrden_Click()             ' Ordenar por
Dim nCom As Integer

nCom = CmbOrden.ListIndex

Select Case nCom
Case 0
    Data1.RecordSource = "Select * from calidad_a order by CodMat"
    
Case 1
    Data1.RecordSource = "Select * from calidad_a order by DESCRIPCIOn"
End Select
TxFiltro = ""
Data1.Refresh
If DBGrid1.Visible Then DBGrid1.SetFocus
End Sub

Private Sub Command1_Click()
    resp = "S"
    Limpiar
   Text1.Enabled = True
   'Command19.Visible = False
   DBGrid1.Visible = False
   Frame2.Visible = False
   Frame5.Visible = False
   Frame3.Caption = "Ingreso de Material de Avios"
   
   Frame1.Visible = True
   Frame3.Visible = True
    Text1.SetFocus
End Sub

Private Sub Command19_Click()
If rs.RecordCount > 0 Then
'    FrmArLineas.show 1
End If
End Sub
'Modificaci�n
Private Sub Command2_Click()
If rs.RecordCount > 0 Then
   Limpiar
    resp = "N"
    Frame3.Caption = "Modificaci�n de Material de Avios"
    DBGrid1.Visible = False
    'Command19.Visible = False
    Frame2.Visible = False
    Frame5.Visible = False
    Frame1.Visible = True
    Frame3.Visible = True

    Text1.text = rs.Fields("CodMat")
    Text1.Enabled = False
    
    If Not IsNull(rs.Fields("descripcion")) Then
      Text2.text = rs.Fields("descripcion")
    Else
      Text2.text = ""
    End If
    
   Text2.SetFocus
End If
End Sub
'Eliminaci�n
Private Sub Command3_Click()
On Error GoTo EliErr
Dim cSql1 As String
Dim CSQL2 As String, cSql3 As String
Dim cCodigo1 As String
Dim cSel1 As Recordset
Dim cCodigo As String

If rs.RecordCount > 0 Then
    

    If MsgBox("Seguro de Eliminar", vbQuestion + vbOKCancel, "Inventarios") = vbOK Then
        'If Existe(1, rs.Fields("CODIGO"), "MaeArt", "AMaterial", False) Then
        '    MsgBox "El Titulo de Tela no puede Eliminarse, porque esta registrada en Articulos"
        'Else
            nTra = 2
            cCodigo1 = Pos_Dato1(rs, "CodMat")
            nTra = 1
            nTra = 0
            
            Call Listado("select * from calidad_a")
        'End If
    End If
    DBGrid1.Refresh
    
    If DBGrid1.Visible Then DBGrid1.SetFocus
Else
    MsgBox "No existe ning�n registro para Eilminar", vbInformation, "Inventarios"
End If
Exit Sub

EliErr:
    MsgBox Err.Description
    'If nTra = 1 Then Vgcnx.RollbackTrans
End Sub
'Salir
Private Sub Command7_Click()
   Unload Me
 End Sub
'Grabar
Private Sub Command8_Click()
On Error GoTo GrabErr
Dim cFam As String

If resp = "S" Then
  If Trim(Text1) = "" Then
     MsgBox "Ingrese C�digo de Material de Avios", vbInformation, "Mensaje"
     Text1.SetFocus
     Exit Sub
  Else
       If VGDllGeneral.VerificaDatoExistente(VGCNx, "select * from calidad_a where CodMat='" & Trim(Text1) & "'") = 1 Then
          MsgBox "El c�digo de Material de Avios ya existe", vbInformation, "Mensaje"
          Text1.SetFocus
          Exit Sub
       End If
  End If
End If

  If Trim(Text2) = "" Then
     MsgBox "Ingrese Descripci�n de Material de Avios", vbExclamation, "Aviso"
     Text2.SetFocus
     Exit Sub
  End If
  
'  If Trim(Text3) <> "" Then
'        cBase = cRuta4
'        If UCase(Dir$(cBase)) = "BDCONTABILIDAD.MDB" Then
'            MsgBox "Ingrese Cuenta Contable", vbExclamation, "Aviso"
'        End If
'  End If
'    If Trim(Text4) <> "" Then
'        cBase = cRuta4
'        If UCase(Dir$(cBase)) = "BDCONTABILIDAD.MDB" Then
'            MsgBox "Ingrese Cuenta Contable", vbExclamation, "Aviso"
'        End If
'  End If
    If resp = "S" Then
        VGCNx.Execute "Insert Into calidad_a " & _
                          "(CodMat,descripcion)" & _
                          " VALUES(" & _
                          "'" & Trim(Text1.text) & "'," & _
                          "'" & SupCadSQL(Text2.text) & "')"
                          
        Limpiar
        Text1.SetFocus
    Else
        VGCNx.Execute "UPDATE calidad_a " & _
                          " SET descripcion='" & SupCadSQL(Text2.text) & "'" & _
                          " WHERE CodMat='" & Text1.text & "'"
                          
      DBGrid1.Visible = True
      'Command19.Visible = True
      Frame5.Visible = True
      Frame2.Visible = True
      Frame1.Visible = False
      Frame3.Visible = False
      DBGrid1.Visible = True
      DBGrid1.SetFocus
    End If
    Call Listado("SELECT * FROM calidad_a")
      
   
Exit Sub
'Resume
GrabErr:
    MsgBox Err.Description
End Sub

Sub Limpiar()
Text1 = ""
Text2 = ""
End Sub
Private Sub Form_Activate()
TxFiltro = ""
CmbOrden.ListIndex = 0
If DBGrid1.Visible And DBGrid1.Enabled Then DBGrid1.SetFocus
End Sub

Private Sub Form_Load()
central Me
Call Listado("Select * from calidad_a order by CodMat")
'Command19.Visible = True
End Sub

Sub Listado(wcad)
  Set DBGrid1.DataSource = Nothing
  Set rs = Nothing
  Set rs = New ADODB.Recordset
  rs.Open wcad, VGCNx, adOpenDynamic, adLockOptimistic
  
  Set DBGrid1.DataSource = rs
  With DBGrid1
      .Columns(0).Caption = "Codigo"
      .Columns(0).Width = 700
      .Columns(1).Caption = "Descripcion"
      .Columns(1).Width = 3800
      .MarqueeStyle = dbgHighlightRow
      .Refresh
  End With

End Sub


Private Sub Text1_GotFocus()
Enfoque Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim cFam As String

If KeyAscii = 13 Then
    If Trim(Text1) <> "" Then
       If Existe(1, Trim(Text1), "calidad_a", "CodMat", False) Then
          MsgBox "El c�digo de Material de Avios ya existe", vbInformation, "Mensaje"
          Text1 = "": Text1.SetFocus
          Exit Sub
       End If
    Else
          MsgBox "Ingrese c�digo de Material de Avios", vbInformation, "Mensaje"
          Text1 = "": Text1.SetFocus
    End If
    Text2.SetFocus
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub Text2_GotFocus()
Enfoque Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Trim(Text2) = "" Then
       MsgBox "Ingrese Descripcion de la Material de Avios ", vbInformation, "Mensaje"
       Text2 = "": Text2.SetFocus
    End If
    Command8.SetFocus
End If
End Sub




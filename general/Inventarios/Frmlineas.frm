VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form Frmlineas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Familias de Articulos"
   ClientHeight    =   4485
   ClientLeft      =   1950
   ClientTop       =   1755
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4481.552
   ScaleMode       =   0  'User
   ScaleWidth      =   7301.94
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   636
      Left            =   144
      TabIndex        =   20
      Top             =   96
      Width           =   6876
      Begin ctrlayuda_f.Ctr_Ayuda Ctrayu_familia 
         Height          =   348
         Left            =   1296
         TabIndex        =   21
         Top             =   240
         Width           =   5436
         _ExtentX        =   9578
         _ExtentY        =   609
         XcodMaxLongitud =   0
         xcodwith        =   300
         NomTabla        =   "familia"
         ListaCampos     =   "fam_codigo(1),fam_nombre(1)"
         XcodCampo       =   "fam_codigo"
         XListCampo      =   "fam_nombre"
         ListaCamposDescrip=   "Codigo,descripcion"
         ListaCamposText =   "fam_codigo,fam_nombre"
      End
      Begin VB.Label Label7 
         Caption         =   "Cod. "
         Height          =   396
         Left            =   192
         TabIndex        =   22
         Top             =   240
         Width           =   972
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   480
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   6885
      Begin VB.TextBox TxFiltro 
         Height          =   300
         Left            =   1200
         MaxLength       =   45
         TabIndex        =   16
         Text            =   "TxFiltro"
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "Frmlineas.frx":0000
         Left            =   5040
         List            =   "Frmlineas.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar  :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   165
      TabIndex        =   9
      Top             =   3360
      Width           =   6855
      Begin VB.CommandButton command5 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   3600
         Picture         =   "Frmlineas.frx":0027
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command7 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5760
         Picture         =   "Frmlineas.frx":0469
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2520
         Picture         =   "Frmlineas.frx":08AB
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   360
         Picture         =   "Frmlineas.frx":0CED
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   1440
         Picture         =   "Frmlineas.frx":112F
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   180
      TabIndex        =   13
      Top             =   3396
      Visible         =   0   'False
      Width           =   6840
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3945
         Picture         =   "Frmlineas.frx":1571
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1800
         Picture         =   "Frmlineas.frx":19B3
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   775
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
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   936
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDataGridLib.DataGrid DBGrid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   3625
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox Check1 
         Caption         =   "Consolida Factura"
         Height          =   495
         Left            =   600
         TabIndex        =   23
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2265
         MaxLength       =   8
         TabIndex        =   5
         Top             =   645
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2265
         MaxLength       =   45
         TabIndex        =   6
         Top             =   1125
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         Height          =   255
         Left            =   660
         TabIndex        =   12
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   1125
         Width           =   855
      End
   End
End
Attribute VB_Name = "Frmlineas"
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
    Dim CADENA As String
    Dim cNomRepor  As String

cNomRepor = "al_lineas.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Lineas de Articulos"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport + "\" + cNomRepor
                        
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
    MsgBox "No existe el nombre del Reporte, verifique en Formatos", vbInformation, "Información"
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

Private Sub Ctrayu_familia_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Call Listado("SELECT * FROM lineas where fam_codigo = '" & Ctrayu_familia.xclave & "'")
End Sub

Private Sub TxFiltro_Change()
Dim nstr As String

'If rs.RecordCount > 0 Then
    If Trim(TxFiltro) <> "" Then
        If CmbOrden.ListIndex = 0 Then
            nstr = " and lin_CODIGO like '" & Trim(UCase(TxFiltro)) & "%'"
        ElseIf CmbOrden.ListIndex = 1 Then
            nstr = " and lin_NOMBRE like '" & Trim(UCase(TxFiltro)) & "%'"
        End If
        nstr = "select * from lineas where fam_codigo = '" & Ctrayu_familia.xclave & "' " & nstr
    Else
       nstr = "select * from lineas where fam_codigo = '" & Ctrayu_familia.xclave & "'"
    End If
'End If
Call Listado(nstr)
End Sub

Private Sub CmbOrden_Click()             ' Ordenar por
Dim nCom As Integer

nCom = CmbOrden.ListIndex

Select Case nCom
Case 0
    Data1.RecordSource = "Select * from lineas where fam_codigo = '" & Ctrayu_familia.xclave & "' order by FAM_CODIGO"
    
Case 1
    Data1.RecordSource = "Select * from FAMILIA where fam_codigo = '" & Ctrayu_familia.xclave & "' order by FAM_NOMBRE"
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
   Frame3.Caption = "Ingreso de Familias"
   
   Frame1.Visible = True
   Frame3.Visible = True
    Text1.SetFocus
End Sub

Private Sub Command19_Click()
If rs.RecordCount > 0 Then
'    FrmArLineas.show 1
End If
End Sub
'Modificación
Private Sub Command2_Click()
If rs.RecordCount > 0 Then
   Limpiar
    resp = "N"
    Frame3.Caption = "Modificación de Familias"
    DBGrid1.Visible = False
    'Command19.Visible = False
    Frame2.Visible = False
    Frame5.Visible = False
    Frame1.Visible = True
    Frame3.Visible = True

    Text1.text = rs.Fields("lin_CODIGO")
    Text1.Enabled = False
    
    If Not IsNull(rs.Fields("lin_NOMBRE")) Then
      Text2.text = rs.Fields("lin_NOMBRE")
    Else
      Text2.text = ""
    End If
    If rs.Fields("lin_facturaconsolidada") Then
       Check1.Value = 1
     Else
       Check1.Value = 0
     End If
     
   Text2.SetFocus
End If
End Sub
'Eliminación
Private Sub Command3_Click()
On Error GoTo EliErr
Dim cSql1 As String
Dim CSQL2 As String, cSql3 As String
Dim cCodigo1 As String
Dim cSel1 As Recordset
Dim cCodigo As String

If rs.RecordCount > 0 Then
    CSQL2 = "Delete from GRUPO Where FAM_CODIGO= '" & rs.Fields("FAM_CODIGO") & "' and lin_CODIGO= '" & rs.Fields("lin_CODIGO") & "'"
    
    Dim cSqlA As String, cSelA As ADODB.Recordset
    cSqlA = "Select * GRUPO Where FAM_CODIGO= '" & rs.Fields("FAM_CODIGO") & "' and lin_CODIGO= '" & rs.Fields("lin_CODIGO") & "'"
    Set cSelA = New ADODB.Recordset
    cSelA.Open cSqlA, VGCNx, adOpenStatic
    If cSelA.RecordCount > 0 Then
       If MsgBox("La linea  seleccionada tiene registrada Grupos, al Eliminarla eliminará sus Grupos, desea Eliminarla de todas maneras", vbYesNo, "Eliminacion de Registro") = vbNo Then
          cSelA.Close: Exit Sub
       End If
    End If
    cSelA.Close
    If MsgBox("Seguro de Eliminar", vbQuestion + vbOKCancel, "Inventarios") = vbOK Then
        If Existe(1, rs.Fields("FAM_CODIGO"), "MaeArt", "AFAMILIA", False) And Existe(1, rs.Fields("LIN_CODIGO"), "MaeArt", "ALINEA", False) Then
            MsgBox "La Linea no puede Eliminarse, porque esta registrada en Articulos"
        Else
            nTra = 2
            cCodigo1 = Pos_Dato1(rs.Fields, "LIN_CODIGO")
            nTra = 1
            VGCNx.BeginTrans
            VGCNx.Execute CSQL2
            VGCNx.CommitTrans
            nTra = 0
            Call Listado("select * from lineas where fam_codigo = '" & Ctrayu_familia.xclave & "'")
        End If
    End If
    DBGrid1.Refresh
    
    If DBGrid1.Visible Then DBGrid1.SetFocus
Else
    MsgBox "No existe ningún registro para Eilminar", vbInformation, "Inventarios"
End If
Exit Sub

EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
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
     MsgBox "Ingrese Código de Linea ", vbInformation, "Mensaje"
     Text1.SetFocus
     Exit Sub
  Else
       If VGDllGeneral.VerificaDatoExistente(VGCNx, "select * from lineas where fam_codigo = '" & Ctrayu_familia.xclave & "' and  lin_CODIGO='" & Text1 & "'") = 1 Then
          MsgBox "El código de Linea ya existe", vbInformation, "Mensaje"
          Text1.SetFocus
          Exit Sub
       End If
  End If
End If

  If Trim(Text2) = "" Then
     MsgBox "Ingrese Descripción de lineas", vbExclamation, "Aviso"
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
        VGCNx.Execute "Insert Into LINEAs " & _
                          "(FAM_CODIGO,lin_codigo,lin_NOMBRE)" & _
                          " VALUES('" & Trim(Ctrayu_familia.xclave) & "'," & _
                          "'" & Text1.text & "'," & _
                          "'" & SupCadSQL(Text2.text) & "')"
    
        Limpiar
        Text1.SetFocus
    Else
        VGCNx.Execute "UPDATE lineas " & _
                          " SET lin_NOMBRE='" & SupCadSQL(Text2.text) & "', " & _
                          " lin_facturaconsolidada='" & Check1.Value & "' " & _
                          " WHERE fam_codigo = '" & Ctrayu_familia.xclave & "' and lin_CODIGO='" & Text1.text & "'"
      DBGrid1.Visible = True
      'Command19.Visible = True
      Frame5.Visible = True
      Frame2.Visible = True
      Frame1.Visible = False
      Frame3.Visible = False
      DBGrid1.SetFocus
    End If
    Call Listado("SELECT * FROM lineas where fam_codigo = '" & Ctrayu_familia.xclave & "'")
         
Exit Sub
GrabErr:
    MsgBox Err.Description
    Resume
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

'Dim sConexAux As ADODB.Connection
'Set sConexAux = New ADODB.Connection 'BD. Común
'sConexAux.CursorLocation = adUseClient
'sConexAux.ConnectionString = "Provider=SQLOLEDB.1;User ID='sa';password='administrador';Initial Catalog='FOX';Data Source='192.168.1.2'"
    '"Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGBUsuario & ";password='" & Trim(VGPassw) & "';Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGServer
'sConexAux.ConnectionString = sConexAux.ConnectionString & ";password='" & VGPassw & "'"
'sConexAux.Open

central Me
'sConexAux = vgcnx & ";password='" & VGPassw & "'"
'Call Ctrayu_familia.Conexion(sConexAux)
Call Ctrayu_familia.conexion(VGCNx)
'Command19.Visible = True
End Sub

Sub Listado(wcad)
  Set DBGrid1.DataSource = Nothing
  Set rs = Nothing
  
  Set rs = VGCNx.Execute(wcad)
  Set DBGrid1.DataSource = rs
  With DBGrid1
      .Columns(0).Caption = "Codigo Fam."
      .Columns(0).Width = 1000
      .Columns(1).Caption = "Cod. Linea"
      .Columns(1).Width = 1000
      .Columns(2).Caption = "Descripcion"
      .Columns(2).Width = 4500
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
       If Existe(1, Trim(Text1), "Lineas", "fam_codigo = '" & Ctrayu_familia.xclave & "' and lin_CODIGO", False) Then
          MsgBox "El código de Linea ya existe", vbInformation, "Mensaje"
          Text1 = "": Text1.SetFocus
          Exit Sub
       End If
    Else
          MsgBox "Ingrese código de Familia", vbInformation, "Mensaje"
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
       MsgBox "Ingrese Descripcion de Familia", vbInformation, "Mensaje"
       Text2 = "": Text2.SetFocus
    End If
End If
End Sub

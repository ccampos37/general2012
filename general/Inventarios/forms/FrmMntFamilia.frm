VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "textfer.ocx"
Begin VB.Form FrmMntFamilia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Familias de Articulos"
   ClientHeight    =   4800
   ClientLeft      =   1950
   ClientTop       =   1755
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleMode       =   0  'User
   ScaleWidth      =   7301.94
   ShowInTaskbar   =   0   'False
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
      Left            =   150
      TabIndex        =   16
      Top             =   120
      Width           =   6885
      Begin VB.TextBox TxFiltro 
         Height          =   300
         Left            =   1200
         MaxLength       =   45
         TabIndex        =   18
         Text            =   "TxFiltro"
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmMntFamilia.frx":0000
         Left            =   5040
         List            =   "FrmMntFamilia.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar  :"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   255
         Left            =   4320
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   165
      TabIndex        =   11
      Top             =   3720
      Width           =   6855
      Begin VB.CommandButton command5 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   3600
         Picture         =   "FrmMntFamilia.frx":0027
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
         Picture         =   "FrmMntFamilia.frx":0469
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2520
         Picture         =   "FrmMntFamilia.frx":08AB
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   360
         Picture         =   "FrmMntFamilia.frx":0CED
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   1440
         Picture         =   "FrmMntFamilia.frx":112F
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   180
      TabIndex        =   15
      Top             =   3750
      Visible         =   0   'False
      Width           =   6840
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3945
         Picture         =   "FrmMntFamilia.frx":1571
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1800
         Picture         =   "FrmMntFamilia.frx":19B3
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3135
      Left            =   450
      TabIndex        =   12
      Top             =   330
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2265
         MaxLength       =   8
         TabIndex        =   8
         Top             =   2745
         Width           =   2190
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2265
         MaxLength       =   8
         TabIndex        =   7
         Top             =   2340
         Width           =   2205
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
         Top             =   1005
         Width           =   3495
      End
      Begin TextFer.TxFer TxFcorrelativo 
         Height          =   345
         Left            =   2280
         TabIndex        =   27
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         BackColor       =   65535
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuGastos 
         Height          =   350
         Left            =   2280
         TabIndex        =   29
         Top             =   1800
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   609
         XcodMaxLongitud =   20
         xcodwith        =   1000
         NomTabla        =   "co_gastos"
         TituloAyuda     =   "Busqueda de Cuenta de Gastos"
         ListaCampos     =   "gastoscodigo(1),gastosdescripcion(1),gastosctrlcostos(1),cuentacodigo(1),tipoanaliticocodigo(1),habilitadodetraccion(1)"
         XcodCampo       =   "gastoscodigo"
         XListCampo      =   "gastosdescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "gastoscodigo,gastosdescripcion,gastosctrlcostos,cuentacodigo,tipoanaliticocodigo,habilitadodetraccion"
      End
      Begin VB.Label Label7 
         Caption         =   "Codigo Plan de gastos "
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   28
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Correltivo Codigo"
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   26
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "(Haber)"
         Height          =   240
         Left            =   4680
         TabIndex        =   24
         Top             =   2775
         Width           =   1110
      End
      Begin VB.Label Label5 
         Caption         =   "(Debe)"
         Height          =   240
         Left            =   4665
         TabIndex        =   23
         Top             =   2370
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "Cuenta Contable:"
         Height          =   270
         Left            =   675
         TabIndex        =   22
         Top             =   2760
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Contable:"
         Height          =   270
         Left            =   660
         TabIndex        =   21
         Top             =   2355
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         Height          =   255
         Left            =   300
         TabIndex        =   14
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   300
         TabIndex        =   13
         Top             =   1005
         Width           =   855
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
      Top             =   450
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDataGridLib.DataGrid DBGrid1 
      Height          =   2775
      Left            =   240
      TabIndex        =   25
      Top             =   840
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   4895
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
End
Attribute VB_Name = "FrmMntFamilia"
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

cNomRepor = "famarti.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Familia de Articulo"
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

Private Sub Text3_Change()
' Enfoque Text3
End Sub

Private Sub Text3_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
cBase = cRuta4
If UCase(Dir$(cBase)) = UCase(cNomBd4) Then
        Adodc2.Open "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION From Plan_Cuenta_Nacional", VGcnxCT, adOpenStatic
        frmReferencia.Conectar Adodc2, "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION From Plan_Cuenta_Nacional"
        frmReferencia.Label1.Caption = "Plan de Cuenta Nacional"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text3.text = (vGUtil(1))
        End If
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text3_DblClick
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cBase = cRuta4
   If Trim(Text3) <> "" Then
       If UCase(Dir$(cBase)) = UCase(cNomBd4) Then
        'MsgBox "Ingrese Cuenta Contable", vbInformation, "Información"
            If Existe(3, Text3, "PLAN_CUENTA_NACIONAL", "PLANCTA_CODIGO", False) = False Then
                    MsgBox "La Cuenta no existe", vbInformation, "Mensaje"
                    Text3.SetFocus: Exit Sub
             End If
        End If
    End If
   SendKeys "{tab}"
End If
End Sub

Private Sub Text4_Change()
Enfoque Text3
End Sub

Private Sub Text4_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
cBase = cRuta4
If UCase(Dir$(cBase)) = UCase(cNomBd4) Then
        Adodc2.Open "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION From Plan_Cuenta_Nacional", VGcnxCT, adOpenStatic
        frmReferencia.Conectar Adodc2, "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION From Plan_Cuenta_Nacional"
        frmReferencia.Label1.Caption = "Plan de Cuenta Nacional"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text4.text = (vGUtil(1))
        End If
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text4_DblClick
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cBase = cRuta4
   If Trim(Text4) <> "" Then
       If UCase(Dir$(cBase)) = UCase(cNomBd4) Then
        'MsgBox "Ingrese Cuenta Contable", vbInformation, "Información"
            If Existe(3, Text4, "PLAN_CUENTA_NACIONAL ", "PLANCTA_CODIGO", False) = False Then
                    MsgBox "La Cuenta no existe", vbInformation, "Mensaje"
                    Text3.SetFocus: Exit Sub
             End If
        End If
    End If
    Command8.SetFocus
End If
End Sub


Private Sub TxFiltro_Change()
Dim nstr As String

'If rs.RecordCount > 0 Then
    If Trim(TxFiltro) <> "" Then
        If CmbOrden.ListIndex = 0 Then
            nstr = "FAM_CODIGO like '" & Trim(UCase(TxFiltro)) & "%'"
        ElseIf CmbOrden.ListIndex = 1 Then
            nstr = "FAM_NOMBRE like '" & Trim(UCase(TxFiltro)) & "%'"
        End If
        nstr = "select * from familia where " & nstr
    Else
       nstr = "select * from familia"
    End If
'Else
'    nstr = "select * from familia"
'End If
Call Listado(nstr)
End Sub

Private Sub CmbOrden_Click()             ' Ordenar por
Dim nCom As Integer

nCom = CmbOrden.ListIndex

Select Case nCom
Case 0
   Call Listado("Select * from FAMILIA order by FAM_CODIGO")
   ' Data1.RecordSource = "Select * from FAMILIA order by FAM_CODIGO"
    
Case 1
    Call Listado("Select * from FAMILIA order by FAM_nombre")
'    Data1.RecordSource = "Select * from FAMILIA order by FAM_NOMBRE"
End Select
TxFiltro = ""
Data1.Refresh
DBGrid1.Refresh
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

    Text1.text = rs.Fields("FAM_CODIGO")
    Text1.Enabled = False
    
    If Not IsNull(rs.Fields("FAM_NOMBRE")) Then
      Text2.text = rs.Fields("FAM_NOMBRE")
    Else
      Text2.text = ""
    End If
    If Not IsNull(rs.Fields("correlativocodigo")) Then
      TxFcorrelativo.text = rs.Fields("correlativocodigo")
    Else
      TxFcorrelativo.text = 1
    End If
    If Not IsNull(rs.Fields("FAM_DEBE")) Then
      Text3.text = rs.Fields("FAM_DEBE")
    Else
      Text3.text = ""
    End If
    If Not IsNull(rs.Fields("FAM_HABER")) Then
      Text4.text = rs.Fields("FAM_HABER")
    Else
      Text4.text = ""
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
    cSql1 = "Delete from  LINEAS Where FAM_CODIGO= '" & rs.Fields("FAM_CODIGO") & "'"
    CSQL2 = "Delete from FAMILIA Where FAM_CODIGO= '" & rs.Fields("FAM_CODIGO") & "'"
 
    Dim cSqlA As String, cSelA As ADODB.Recordset
    
    cSqlA = "Select * FROM LINEAS WHERE FAM_CODIGO = '" & Trim(rs.Fields("FAM_CODIGO")) & "'"
    Set cSelA = New ADODB.Recordset
    cSelA.Open cSqlA, VGCNx, adOpenStatic
    If cSelA.RecordCount > 0 Then
       If MsgBox("La Familia seleccionada tiene registrada Lineas, al Eliminarla eliminará sus Lineas, desea Eliminarla de todas maneras", vbYesNo, "Eliminacion de Registro") = vbNo Then
          cSelA.Close: Exit Sub
       End If
    End If
    cSelA.Close
    

    If MsgBox("Seguro de Eliminar", vbQuestion + vbOKCancel, "Inventarios") = vbOK Then
        If Existe(1, rs.Fields("FAM_CODIGO"), "MaeArt", "AFAMILIA", False) Then
            MsgBox "La Familia no puede Eliminarse, porque esta registrada en Articulos"
        Else
            nTra = 2
            VGCNx.BeginTrans
            VGCNx.Execute CSQL2
            VGCNx.CommitTrans
            nTra = 0
            
            Call Listado("select * from familia")
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
     MsgBox "Ingrese Código de Familia ", vbInformation, "Mensaje"
     Text1.SetFocus
     Exit Sub
  Else
       If VGDllGeneral.VerificaDatoExistente(VGCNx, "select * from FAMILIA where FAM_CODIGO='" & Text1 & "'") = 1 Then
          MsgBox "El código de Familia ya existe", vbInformation, "Mensaje"
          Text1.SetFocus
          Exit Sub
       End If
  End If
End If

  If Trim(Text2) = "" Then
     MsgBox "Ingrese Descripción de Familia", vbExclamation, "Aviso"
     Text2.SetFocus
     Exit Sub
  End If
  If VGparametros.VGLongCodigo <> 0 And numero(TxFcorrelativo.text) = 0 Then
     MsgBox "Ingrese Correltivo mayor a CERO (0)", vbExclamation, "Aviso"
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
        VGCNx.Execute "Insert Into Familia " & _
                          "(FAM_CODIGO,FAM_NOMBRE,correlativocodigo,FAM_DEBE,FAM_HABER)" & _
                          " VALUES(" & _
                          "'" & Text1.text & "'," & _
                          "'" & SupCadSQL(Text2.text) & "'," & TxFcorrelativo.text & "," & _
                          "'" & IIf(Text3 <> "", Text3, " ") & "'," & _
                          "'" & IIf(Text4 <> "", Text4, " ") & "')"
    
        Limpiar
        Text1.SetFocus
    Else
        VGCNx.Execute "UPDATE Familia " & _
                          " SET FAM_NOMBRE='" & SupCadSQL(Text2.text) & "'," & _
                          " FAM_DEBE='" & IIf(Text3 <> "", Text3, " ") & "'," & _
                          " correlativocodigo=" & TxFcorrelativo.text & "," & _
                          " FAM_HABER='" & IIf(Text4 <> "", Text4, " ") & "'" & _
                          " WHERE FAM_CODIGO='" & Text1.text & "'"
      DBGrid1.Visible = True
      'Command19.Visible = True
      Frame5.Visible = True
      Frame2.Visible = True
      Frame1.Visible = False
      Frame3.Visible = False
      DBGrid1.SetFocus
    End If
    Call Listado("SELECT * FROM FAMILIA")
      
   
Exit Sub
GrabErr:
    MsgBox Err.Description
End Sub

Sub Limpiar()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub
Private Sub Form_Activate()
TxFiltro = ""
CmbOrden.ListIndex = 0
Call Ctr_AyuGastos.Conexion(VGCNx)
If DBGrid1.Visible And DBGrid1.Enabled Then DBGrid1.SetFocus
End Sub

Private Sub Form_Load()
central Me
Call Listado("Select * from FAMILIA order by FAM_CODIGO")
'Command19.Visible = True
End Sub

Sub Listado(wcad)
  Set DBGrid1.DataSource = Nothing
  Set rs = Nothing
  
  Set rs = VGCNx.Execute(wcad)
  Set DBGrid1.DataSource = rs
  With DBGrid1
      .Columns(0).Caption = "Codigo"
      .Columns(0).Width = 1000
      .Columns(1).Caption = "Descripcion"
      .Columns(1).Width = 3800
      .Columns(2).Caption = "Cuenta Contable"
      .Columns(2).Width = 1000
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
       If Existe(1, Trim(Text1), "FAMILIA", "FAM_CODIGO", False) Then
          MsgBox "El código de Familia ya existe", vbInformation, "Mensaje"
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
    Text3.SetFocus
End If
End Sub

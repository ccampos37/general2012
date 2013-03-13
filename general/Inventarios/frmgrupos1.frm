VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmgrupo 
   Caption         =   "Grupos"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form2"
   ScaleHeight     =   6975
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame frm_ingreso 
      Height          =   2055
      Left            =   480
      TabIndex        =   16
      Top             =   2910
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox Check1 
         Caption         =   "Consolida Factura"
         Height          =   495
         Left            =   600
         TabIndex        =   19
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2265
         MaxLength       =   8
         TabIndex        =   18
         Top             =   645
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2265
         MaxLength       =   45
         TabIndex        =   17
         Top             =   1125
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Código:"
         Height          =   255
         Left            =   660
         TabIndex        =   21
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   1125
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   420
      TabIndex        =   12
      Top             =   4860
      Visible         =   0   'False
      Width           =   6840
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3945
         Picture         =   "frmgrupos1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1800
         Picture         =   "frmgrupos1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   405
      TabIndex        =   6
      Top             =   4830
      Width           =   6855
      Begin VB.CommandButton command5 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   3600
         Picture         =   "frmgrupos1.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command7 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5760
         Picture         =   "frmgrupos1.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2520
         Picture         =   "frmgrupos1.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   360
         Picture         =   "frmgrupos1.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   1440
         Picture         =   "frmgrupos1.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame4 
      Height          =   870
      Left            =   510
      TabIndex        =   4
      Top             =   120
      Width           =   6876
      Begin ctrlayuda_f.Ctr_Ayuda Ctrayu_familia 
         Height          =   465
         Left            =   1290
         TabIndex        =   1
         Top             =   240
         Width           =   5420
         _ExtentX        =   9551
         _ExtentY        =   820
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
         Caption         =   "Cod. Familia"
         Height          =   396
         Left            =   192
         TabIndex        =   5
         Top             =   240
         Width           =   972
      End
   End
   Begin VB.Frame Frame6 
      Height          =   750
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   6876
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_ayulinea 
         Height          =   348
         Left            =   1296
         TabIndex        =   2
         Top             =   240
         Width           =   5436
         _ExtentX        =   9578
         _ExtentY        =   609
         XcodMaxLongitud =   0
         xcodwith        =   300
         NomTabla        =   "lineas"
         ListaCampos     =   "lin_codigo(1),lin_nombre(1)"
         XcodCampo       =   "lin_codigo"
         XListCampo      =   "lin_nombre"
         ListaCamposDescrip=   "Codigo,descripcion"
         ListaCamposText =   "lin_codigo,lin_nombre"
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Linea"
         Height          =   396
         Left            =   192
         TabIndex        =   3
         Top             =   240
         Width           =   972
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   720
      Top             =   4590
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid DBGrid1 
      Height          =   2895
      Left            =   360
      TabIndex        =   15
      Top             =   1950
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   5106
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
Attribute VB_Name = "frmgrupo"
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

cNomRepor = "al_grupos.RPT"
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
      frm_ingreso.Visible = False
      Frame2.Visible = True
      Frame3.Visible = False
      DBGrid1.SetFocus
End Sub

Private Sub Ctr_ayulinea_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Call Listado("SELECT * FROM grupo where fam_codigo = '" & Ctrayu_familia.xclave & "' and lin_codigo = '" & Ctr_ayulinea.xclave & "'")
    End Sub


Private Sub Command1_Click()
    resp = "S"
    Limpiar
   Text1.Enabled = True
   DBGrid1.Visible = False
   Frame2.Visible = False
   frm_ingreso.Visible = True
   frm_ingreso.Caption = "Ingreso de Grupos"
   
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
    Frame3.Caption = "Modificación de Grupos"
    DBGrid1.Visible = False
    Frame2.Visible = False
'   Frame5.Visible = False
'   Frame1.Visible = True
    Frame3.Visible = True
    frm_ingreso.Visible = True
    frm_ingreso.Visible = True
    Ctrayu_familia.Enabled = False
    Ctr_ayulinea.Enabled = False
    Text1.text = rs.Fields("gru_CODIGO")
    Text1.Enabled = False
    
    If Not IsNull(rs.Fields("gru_NOMBRE")) Then
      Text2.text = rs.Fields("gru_NOMBRE")
    Else
      Text2.text = ""
    End If
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
    CSQL2 = "Delete from GRUPO Where FAM_CODIGO= '" & rs.Fields("FAM_CODIGO") & "' and lin_CODIGO= '" & rs.Fields("LIN_CODIGO") & "' and gru_CODIGO= '" & rs.Fields("gru_CODIGO") & "'"
    If MsgBox("Seguro de Eliminar", vbQuestion + vbOKCancel, "Inventarios") = vbOK Then
            nTra = 2
            VGCNx.BeginTrans
            VGCNx.Execute CSQL2
            VGCNx.CommitTrans
    End If
    Call Listado("SELECT * FROM GRUPO where fam_codigo = '" & Ctrayu_familia.xclave & "' AND lin_codigo = '" & Ctr_ayulinea.xclave & "'")
    If DBGrid1.Visible Then DBGrid1.SetFocus
Else
    MsgBox "No existe ningún registro para Eilminar", vbInformation, "Inventarios"
End If
   DBGrid1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = False
    DBGrid1.SetFocus
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
Dim SQL As String

If resp = "S" Then
  If Trim(Text1) = "" Then
     MsgBox "Ingrese Código de Grupo ", vbInformation, "Mensaje"
     Text1.SetFocus
     Exit Sub
  Else
       If VGDllGeneral.VerificaDatoExistente(VGCNx, "select * from grupo where fam_codigo = '" & Ctrayu_familia.xclave & "' and  lin_CODIGO='" & Ctr_ayulinea.xclave & "' and  gru_CODIGO='" & Text1 & "' ") = 1 Then
          MsgBox "El código de Grupo ya existe", vbInformation, "Mensaje"
          Text1.SetFocus
          Exit Sub
       End If
  End If
End If

  If Trim(Text2) = "" Then
     MsgBox "Ingrese Descripción de Grupo", vbExclamation, "Aviso"
     Text2.SetFocus
     Exit Sub
  End If
  
    If resp = "S" Then
       SQL = "Insert Into Grupo (FAM_CODIGO,lin_codigo,gru_codigo,gru_NOMBRE) "
       SQL = SQL & " VALUES('" & Trim(Ctrayu_familia.xclave) & "',"
       SQL = SQL & "'" & Trim(Ctr_ayulinea.xclave) & "','" & Text1.text & "',"
       SQL = SQL & "'" & SupCadSQL(Text2.text) & "')"
    
        VGCNx.Execute (SQL)
    
        Limpiar
        Text1.SetFocus
    Else
       SQL = " UPDATE grupo SET gru_NOMBRE='" & SupCadSQL(Text2.text) & "'"
       SQL = SQL & " WHERE fam_codigo = '" & Ctrayu_familia.xclave & "'"
       SQL = SQL & " and lin_codigo = '" & Ctr_ayulinea.xclave & "' "
       SQL = SQL & " and gru_CODIGO='" & Trim(Text1.text) & "'"
       
       VGCNx.Execute (SQL)

      DBGrid1.Visible = True
      Frame2.Visible = True
      Frame3.Visible = False
      DBGrid1.SetFocus
    End If
    DBGrid1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = False
    DBGrid1.SetFocus
    frm_ingreso.Visible = False
    Ctr_ayulinea.Enabled = True
    Ctrayu_familia.Enabled = True
    Call Listado("SELECT * FROM GRUPO where fam_codigo = '" & Ctrayu_familia.xclave & "' AND lin_codigo = '" & Ctr_ayulinea.xclave & "'")
         
Exit Sub
GrabErr:
    MsgBox Err.Description
    Resume
End Sub

Sub Limpiar()
Text1 = ""
Text2 = ""
End Sub

Private Sub Ctrayu_familia_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim SQL As String
Dim rsql As New ADODB.Recordset
SQL = "SELECT * FROM lineas where fam_codigo = '" & Ctrayu_familia.xclave & "'"
Set rsql = VGCNx.Execute(SQL)
If rsql.RecordCount() > 0 Then
   Call Listado(SQL)
   Ctr_ayulinea.Enabled = True
   Ctr_ayulinea.filtro = "fam_codigo = '" & Ctrayu_familia.xclave & "'"
  Else
    MsgBox (" no existe detalle de linea ")
End If
End Sub

Private Sub Form_Activate()
If DBGrid1.Visible And DBGrid1.Enabled Then DBGrid1.SetFocus
Ctr_ayulinea.Enabled = False
Ctrayu_familia.Enabled = True
End Sub

Private Sub Form_Load()
central Me
Call Ctrayu_familia.Conexion(VGCNx)
Call Ctr_ayulinea.Conexion(VGCNx)
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
      .Columns(2).Caption = "Cod. grupo"
      .Columns(2).Width = 1000
      .Columns(3).Caption = "Descripcion"
      .Columns(3).Width = 4500
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
       If Existe(1, Trim(Text1), "Grupo", "fam_codigo ='" & Ctrayu_familia.xclave & "'and lin_codigo ='" & Ctr_ayulinea.xclave & "' and lin_CODIGO", False) Then
          MsgBox "El código de Grupo ya existe", vbInformation, "Mensaje"
          Text1 = "": Text1.SetFocus
          Exit Sub
       End If
    Else
          MsgBox "Ingrese código de Grupo", vbInformation, "Mensaje"
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



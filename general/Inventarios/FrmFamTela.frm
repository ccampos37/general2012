VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmFamTela 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Familia de Tela"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DBGrid1 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   4471
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6885
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmFamTela.frx":0000
         Left            =   5040
         List            =   "FrmFamTela.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxFiltro 
         Height          =   300
         Left            =   1200
         MaxLength       =   45
         TabIndex        =   2
         Text            =   "TxFiltro"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar  :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   735
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
      Left            =   570
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   330
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      Height          =   3135
      Left            =   600
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   20
         Top             =   2520
         Width           =   750
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2265
         MaxLength       =   45
         TabIndex        =   17
         Top             =   1125
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2265
         MaxLength       =   8
         TabIndex        =   16
         Top             =   645
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   18
         Top             =   1620
         Width           =   765
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   19
         Top             =   2025
         Width           =   750
      End
      Begin VB.Label Label5 
         Caption         =   "Raport :"
         Height          =   270
         Left            =   675
         TabIndex        =   25
         Top             =   2520
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción:"
         Height          =   255
         Left            =   660
         TabIndex        =   24
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         Height          =   255
         Left            =   660
         TabIndex        =   23
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Mescla :"
         Height          =   270
         Left            =   660
         TabIndex        =   22
         Top             =   1635
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Densidad:"
         Height          =   270
         Left            =   675
         TabIndex        =   21
         Top             =   2040
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   6855
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   1440
         Picture         =   "FrmFamTela.frx":0027
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   360
         Picture         =   "FrmFamTela.frx":0469
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2520
         Picture         =   "FrmFamTela.frx":08AB
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command7 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5760
         Picture         =   "FrmFamTela.frx":0CED
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton command5 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   3600
         Picture         =   "FrmFamTela.frx":112F
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   775
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   570
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   6840
      Begin VB.CommandButton Command8 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1800
         Picture         =   "FrmFamTela.frx":1571
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3945
         Picture         =   "FrmFamTela.frx":19B3
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   255
         Width           =   775
      End
   End
End
Attribute VB_Name = "FrmFamTela"
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

cNomRepor = "famartitela.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Familia de Tela"
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

Private Sub Text2_Change()
Dim I As Integer
Text1.text = UCase(Text1.text)
I = Len(Text1.text)
Text1.SelStart = I

End Sub

Private Sub Text3_Change()
' Enfoque Text3
Dim I As Integer
Text3.text = UCase(Text3.text)
I = Len(Text3.text)
Text3.SelStart = I
End Sub

Private Sub Text3_DblClick()
'Dim Adodc2 As adodb.Recordset
'Set Adodc2 = New adodb.Recordset
'cBase = cRuta4
'If UCase(Dir$(cBase)) = UCase(cNomBd4) Then
'        Adodc2.Open "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION From Plan_Cuenta_Nacional", VGcnxCT, adOpenStatic
'        frmReferencia.Conectar Adodc2, "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION From Plan_Cuenta_Nacional"
'        frmReferencia.Label1.Caption = "Plan de Cuenta Nacional"
'        frmReferencia.Show vbModal
'        Adodc2.Close
'        If vGUtil(1) <> "" Then
'                Text3.text = (vGUtil(1))
'        End If
'End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then Text3_DblClick
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
'   cBase = cRuta4
'   If Trim(Text3) <> "" Then
'       If UCase(Dir$(cBase)) = UCase(cNomBd4) Then
'        'MsgBox "Ingrese Cuenta Contable", vbInformation, "Información"
'            If Existe(3, Text3, "PLAN_CUENTA_NACIONAL", "PLANCTA_CODIGO", False) = False Then
'                    MsgBox "La Cuenta no existe", vbInformation, "Mensaje"
'                    Text3.SetFocus: Exit Sub
'             End If
'        End If
'    End If
   SendKeys "{tab}"
End If
End Sub

Private Sub Text4_Change()
Enfoque Text3
Dim I As Integer
Text4.text = UCase(Text4.text)
I = Len(Text4.text)
Text4.SelStart = I
End Sub

Private Sub Text4_DblClick()
'Dim Adodc2 As adodb.Recordset
'Set Adodc2 = New adodb.Recordset
'cBase = cRuta4
'If UCase(Dir$(cBase)) = UCase(cNomBd4) Then
'        Adodc2.Open "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION From Plan_Cuenta_Nacional", VGcnxCT, adOpenStatic
'        frmReferencia.Conectar Adodc2, "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION From Plan_Cuenta_Nacional"
'        frmReferencia.Label1.Caption = "Plan de Cuenta Nacional"
'        frmReferencia.Show vbModal
'        Adodc2.Close
'        If vGUtil(1) <> "" Then
'                Text4.text = (vGUtil(1))
'        End If
'End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then Text4_DblClick
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'   cBase = cRuta4
'   If Trim(Text4) <> "" Then
'       If UCase(Dir$(cBase)) = UCase(cNomBd4) Then
'        'MsgBox "Ingrese Cuenta Contable", vbInformation, "Información"
'            If Existe(3, Text4, "PLAN_CUENTA_NACIONAL ", "PLANCTA_CODIGO", False) = False Then
'                    MsgBox "La Cuenta no existe", vbInformation, "Mensaje"
'                    Text3.SetFocus: Exit Sub
'             End If
'        End If
'    End If
    Text5.SetFocus
End If
End Sub


Private Sub Text5_Change()
Dim I As Integer
Text5.text = UCase(Text5.text)
I = Len(Text5.text)
Text5.SelStart = I
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command8.SetFocus
End If
End Sub

Private Sub TxFiltro_Change()
Dim nstr As String

'If rs.RecordCount > 0 Then
    If Trim(TxFiltro) <> "" Then
        If CmbOrden.ListIndex = 0 Then
            nstr = "CODIGO like '" & Trim(UCase(TxFiltro)) & "%'"
        ElseIf CmbOrden.ListIndex = 1 Then
            nstr = "descripcio like '" & Trim(UCase(TxFiltro)) & "%'"
        End If
        nstr = "select * from familiatela where " & nstr
    Else
       nstr = "select * from familiatela"
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
    Data1.RecordSource = "Select * from FAMILIATELA order by CODIGO"
    
Case 1
    Data1.RecordSource = "Select * from FAMILIATELA order by DESCRIPCIO"
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
   Frame3.Caption = "Ingreso de Familias de Telas"
   
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
    Frame3.Caption = "Modificación de Familias de Telas"
    DBGrid1.Visible = False
    'Command19.Visible = False
    Frame2.Visible = False
    Frame5.Visible = False
    Frame1.Visible = True
    Frame3.Visible = True

    Text1.text = rs.Fields("CODIGO")
    Text1.Enabled = False
    
    If Not IsNull(rs.Fields("descripcio")) Then
      Text2.text = rs.Fields("descripcio")
    Else
      Text2.text = ""
    End If
    If Not IsNull(rs.Fields("mezcla")) Then
      Text3.text = rs.Fields("mezcla")
    Else
      Text3.text = ""
    End If
    If Not IsNull(rs.Fields("densidad")) Then
      Text4.text = rs.Fields("densidad")
    Else
      Text4.text = ""
    End If
    
    If Not IsNull(rs.Fields("raport")) Then
      Text5.text = rs.Fields("raport")
    Else
      Text5.text = ""
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
    'cSql1 = "Delete from  LINEAS Where FAM_CODIGO= '" & rs.Fields("FAM_CODIGO") & "'"
    'CSQL2 = "Delete from GRUPO Where FAM_CODIGO= '" & rs.Fields("FAM_CODIGO") & "'"
    
 
'    Dim cSqlA As String, cSelA As adodb.Recordset
'
'    cSqlA = "Select * FROM LINEAS WHERE FAM_CODIGO = '" & Trim(rs.Fields("FAM_CODIGO")) & "'"
'    Set cSelA = New adodb.Recordset
'    cSelA.Open cSqlA, Vgcnx, adOpenStatic
'    If cSelA.RecordCount > 0 Then
'       If MsgBox("La Familia seleccionada tiene registrada Lineas, al Eliminarla eliminará sus Lineas, desea Eliminarla de todas maneras", vbYesNo, "Eliminacion de Registro") = vbNo Then
'          cSelA.Close: Exit Sub
'       End If
'    End If
'    cSelA.Close
    

    If MsgBox("Seguro de Eliminar", vbQuestion + vbOKCancel, "Inventarios") = vbOK Then
        If Existe(1, rs.Fields("CODIGO"), "MaeArt", "AFAMILIA", False) Then
            MsgBox "La Familia de Tela no puede Eliminarse, porque esta registrada en Articulos"
        Else
            nTra = 2
            cCodigo1 = Pos_Dato1(rs, "CODIGO")
            nTra = 1
            'Vgcnx.BeginTrans
            'Vgcnx.Execute cSql1
            'Vgcnx.Execute CSQL2
            'Vgcnx.CommitTrans
            nTra = 0
            
            Call Listado("select * from familiatela")
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
     MsgBox "Ingrese Código de Familia de Tela", vbInformation, "Mensaje"
     Text1.SetFocus
     Exit Sub
  Else
       If VGDllGeneral.VerificaDatoExistente(VGCNx, "select * from FAMILIATELA where CODIGO='" & Trim(Text1) & "'") = 1 Then
          MsgBox "El código de Familia de tela ya existe", vbInformation, "Mensaje"
          Text1.SetFocus
          Exit Sub
       End If
  End If
End If

  If Trim(Text2) = "" Then
     MsgBox "Ingrese Descripción de Familia de Tela", vbExclamation, "Aviso"
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
        VGCNx.Execute "Insert Into Familiatela " & _
                          "(CODIGO,descripcio,mezcla,densidad,raport)" & _
                          " VALUES(" & _
                          "'" & Trim(Text1.text) & "'," & _
                          "'" & SupCadSQL(Text2.text) & "'," & _
                          "'" & IIf(Text3 <> "", Text3, "") & "'," & _
                          "'" & IIf(Text4 <> "", Text4, "") & "'," & _
                          "'" & IIf(Text5 <> "", Text5, "") & "')"
    
        Limpiar
        Text1.SetFocus
    Else
        VGCNx.Execute "UPDATE Familiatela " & _
                          " SET descripcio='" & SupCadSQL(Text2.text) & "'," & _
                          " mezcla='" & IIf(Text3 <> "", Text3, "") & "'," & _
                          " densidad='" & IIf(Text4 <> "", Text4, "") & "'," & _
                          " raport='" & IIf(Text5 <> "", Text5, "") & "'" & _
                          " WHERE CODIGO='" & Text1.text & "'"
                          
      DBGrid1.Visible = True
      'Command19.Visible = True
      Frame5.Visible = True
      Frame2.Visible = True
      Frame1.Visible = False
      Frame3.Visible = False
      DBGrid1.Visible = True
      DBGrid1.SetFocus
    End If
    Call Listado("SELECT * FROM FAMILIAtela")
      
   
Exit Sub
GrabErr:
    MsgBox Err.Description
End Sub

Sub Limpiar()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
End Sub
Private Sub Form_Activate()
TxFiltro = ""
CmbOrden.ListIndex = 0
If DBGrid1.Visible And DBGrid1.Enabled Then DBGrid1.SetFocus
End Sub

Private Sub Form_Load()
central Me
Call Listado("Select * from FAMILIAtela order by CODIGO")
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
      .Columns(2).Caption = "Mezcla"
      .Columns(2).Width = 800
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
       If Existe(1, Trim(Text1), "FAMILIATELA", "CODIGO", False) Then
          MsgBox "El código de Familia de Tela  ya existe", vbInformation, "Mensaje"
          Text1 = "": Text1.SetFocus
          Exit Sub
       End If
    Else
          MsgBox "Ingrese código de Familia de Tela", vbInformation, "Mensaje"
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
       MsgBox "Ingrese Descripcion de Familia de Tela ", vbInformation, "Mensaje"
       Text2 = "": Text2.SetFocus
    End If
    Text3.SetFocus
End If
End Sub


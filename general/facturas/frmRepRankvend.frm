VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRepRankvend 
   Caption         =   "Form2"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6930
   LinkTopic       =   "Form2"
   ScaleHeight     =   5130
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3240
      MaxLength       =   8
      TabIndex        =   8
      Top             =   2775
      Width           =   1305
   End
   Begin VB.ComboBox cmbBase 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2175
      Width           =   2880
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3480
      TabIndex        =   6
      Top             =   3855
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Top             =   3855
      Width           =   1215
   End
   Begin VB.CheckBox chk 
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   4
      Top             =   3255
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   495
      Width           =   2610
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   4155
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   615
      Visible         =   0   'False
      Width           =   1305
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   360
      Top             =   3015
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTHasta 
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1575
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97517569
      CurrentDate     =   37518
   End
   Begin MSComCtl2.DTPicker DTDesde 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   975
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97517569
      CurrentDate     =   37518
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   840
      TabIndex        =   14
      Top             =   2775
      Width           =   1995
   End
   Begin VB.Label lbl 
      Caption         =   "En Base a"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   13
      Top             =   2175
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   12
      Top             =   1095
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   11
      Top             =   1575
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Incluido IGV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   10
      Top             =   3255
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "Punto de Venta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   1320
      TabIndex        =   9
      Top             =   615
      Width           =   1695
   End
End
Attribute VB_Name = "frmRepRankvend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim TotalNeto As Double
Dim TotalBruto As Double
Dim d_porcentaje As Double
Dim d_monto As Double
Dim index_combo As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''
'Agregar:
Dim busca As New dll_apisgen.dll_apis
Dim adll As New dllgeneral.dll_general
''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmbBase_Click()
    If cmbBase.ListIndex <> index_combo Then
        lbl(0).Caption = cmbBase.Text
        txt(0) = ""
    End If
End Sub

Private Sub cmbBase_DropDown()
    index_combo = cmbBase.ListIndex
End Sub

Private Sub cmbBase_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
On Error GoTo Errores
 If DTDesde > DtHasta Then
     MsgBox "Fecha Desde debe ser mayor a Fecha Hasta", vbInformation, "AVISO"
     Exit Sub
 End If
                 
 If cmbBase.ListIndex = -1 Then
     MsgBox "Ingrese Parámetro En base a", vbInformation, "AVISO"
     Exit Sub
 End If
 
 If Not IsNumeric(txt(0)) Then
     MsgBox "Ingrese Números", vbInformation, "AVISO"
     Exit Sub
 End If
 
 If (txt(0) <= 0) Then
     MsgBox "Ingrese valores mayores a cero", vbInformation, "AVISO"
     Exit Sub
 End If
 
 Screen.MousePointer = 11
 
 Call Consulta_Reporte
Dim Param(9) As Variant
Dim formulas(8) As Variant
  
formulas(0) = "@Empresa='" & VGParametros.nomempresa & "'"
formulas(1) = "@ruc='" & VGParametros.RucEmpresa & "'"

If chk(0).Value = 1 Then
    formulas(2) = "Control_Visibles='N'"
Else
    formulas(2) = "Control_Visibles='S'"
End If
formulas(3) = "Desde='" & DTDesde & "'"
formulas(4) = "Hasta='" & DtHasta & "'"
formulas(5) = "EnBase='" & cmbBase.Text & "'"
formulas(6) = "Numero='" & txt(0) & "'"
If Combo1.ListIndex <> -1 Then
    formulas(7) = "PuntoVta='" & Combo1.Text & "'"
Else
    formulas(7) = "PuntoVta='TODOS'"
End If

Param(0) = TotalNeto
Param(1) = TotalBruto
Param(2) = d_porcentaje
Param(3) = d_monto
Param(4) = DTDesde
Param(5) = DtHasta
Param(6) = IIf(Trim(txt(1)) = "", "%", Trim(txt(1)))
Param(7) = VGParamSistem.BDEmpresa
Param(8) = VGParametros.empresacodigo

Call ImpresionRptProc("vt_Rankingvendedores.rpt", formulas, Param, , "Ranking de Vendedores")

' With oCrystalReport
'        .Reset
'        .ReportFileName = VGParamSistem.Rutareport & "RepvtRankingvendedores.rpt"
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''
'        .LogOnServer "pdssql.dll", _
'         busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", ""), _
'         busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", ""), _
'         busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", ""), _
'         busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "")
'        .Connect = _
'        "DSN=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "") & ";" & _
'        "DSQ=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "") & ";" & _
'        "UID=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "") & ";" & _
'        "PWD=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "")
'        '''''''''''''''''''''''''''''''''''''''''''''''''''
'        .DiscardSavedData = True
'        .Destination = crptToWindow
'        .WindowState = crptMaximized
'        .WindowShowPrintSetupBtn = True
'        .WindowShowExportBtn = True
'        .WindowShowZoomCtl = True
'        .WindowShowNavigationCtls = True
'        .WindowShowPrintBtn = True
'        .WindowTitle = "Ranking de Vendedores"
'        .formulas(0) = "Empresa='" & VGParametros.nomempresa & "'"
'        If chk(0).Value = 1 Then
'            .formulas(1) = "Control_Visibles='N'"
'        Else
'            .formulas(1) = "Control_Visibles='S'"
'        End If
'        .formulas(3) = "Desde='" & DTDesde & "'"
'        .formulas(4) = "Hasta='" & DTHasta & "'"
'        .formulas(5) = "EnBase='" & cmbBase.Text & "'"
'        .formulas(6) = "Numero='" & txt(0) & "'"
'        If Combo1.ListIndex <> -1 Then
'            .formulas(7) = "PuntoVta='" & Combo1.Text & "'"
'        Else
'            .formulas(7) = "PuntoVta='TODOS'"
'        End If
'        .StoredProcParam(0) = TotalNeto
'        .StoredProcParam(1) = TotalBruto
'        .StoredProcParam(2) = d_porcentaje
'        .StoredProcParam(3) = d_monto
'        .StoredProcParam(4) = DTDesde
'        .StoredProcParam(5) = DTHasta
'        .StoredProcParam(6) = IIf(Trim(txt(1)) = "", "%", Trim(txt(1)))
'        .StoredProcParam(7) = busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "BDDATOS", "")
'
'        .Action = 1
'
'  End With
  
Screen.MousePointer = 1

Exit Sub
Errores:
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
  Err = 0
  Exit Sub
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
          If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
          If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    MostrarFormVentas Me, "C2"
    Call adll.llenacombo(Combo1, "select puntovtacodigo,puntovtadescripcion from vt_puntoventa", VGCNx)
    DTDesde = Date
    DtHasta = Date
    Carga_Combo
End Sub

Private Sub Combo1_Click()
  If Combo1.ListCount > 0 Then
     txt(1) = adll.ComboDato(Combo1.Text)
  Else
     txt(1) = ""
  End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If cmbBase.ListIndex = -1 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
 txt(Index).Text = Format(txt(Index).Text, "###,##0.00")
End Sub

Private Function Carga_Combo()
   cmbBase.Clear
   cmbBase.AddItem ("Porcentaje de Ventas")
   cmbBase.AddItem ("Monto de Ventas")
End Function
Private Function Consulta_Reporte()

Dim SQL_TOTAL_NETO As String
Dim SQL_TOTAL_BRUTO_Sol As String
Dim SQL_TOTAL_BRUTO_Dol As String
Dim rs As New ADODB.Recordset
Dim codpuntoventa As String
TotalNeto = 0
TotalBruto = 0

 If Trim(txt(1)) = "" Then
    codpuntoventa = "%"
 Else
    codpuntoventa = Trim(txt(1))
 End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If cmbBase.ListIndex = 0 Then   ' En base a porcentaje de ventas
    
    'd_porcentaje = CDbl(CDbl(txt(0)) / 100)
    d_porcentaje = CDbl(txt(0))
    d_monto = 0
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ElseIf cmbBase.ListIndex = 1 Then   ' En base a monto de ventas
        
    d_porcentaje = 0
    d_monto = CDbl(txt(0))
            
End If
      
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'MONTO TOTAL DE VENTAS EN SOLES (NETO Y SIN IGV)

'SQL_TOTAL_NETO = _
'"SELECT TOTAL_NETO = isnull ( " & _
'"( " & _
'"SELECT SUM(IsNull(a.pedidototneto, 0)) As IMPORTE_SOLES " & _
'"FROM vt_pedido a " & _
'"JOIN vt_cargo b  ON (a.pedidonrofact = b.cargonumdoc  OR a.pedidonroboleta = b.cargonumdoc OR  a.pedidonrogiarem = b.cargonumdoc) " & _
'"WHERE a.pedidofechaanu IS NULL AND a.pedidomoneda = '01' " & _
'"AND a.puntovtacodigo LIKE ('" & codpuntoventa & "') " & _
'"AND b.cargoapefecemi BETWEEN '" & DTDesde & "' AND '" & DTHasta & "' " & _
'")   ,  0  ) " & _
'"+ " & _
'"isnull( " & _
'"(SELECT SUM(IsNull(e.pedidototneto, 0) * IsNull(s.tipocambioventa, 0)) As IMPORTE_DOLARES " & _
'"FROM vt_pedido e " & _
'"JOIN vt_cargo f ON (e.pedidonrofact = f.cargonumdoc  OR e.pedidonroboleta = f.cargonumdoc OR  e.pedidonrogiarem = f.cargonumdoc) " & _
'"JOIN ct_tipocambio s ON f.cargoapefecemi = s.tipocambiofecha " & _
'"WHERE e.pedidofechaanu IS NULL " & _
'"AND e.puntovtacodigo LIKE ('" & codpuntoventa & "') " & _
'"AND f.cargoapefecemi BETWEEN '" & DTDesde & "' AND '" & DTHasta & "' " & _
'"AND e.pedidomoneda = '02' " & _
'") " & _
'" , 0 ) "
'
'SQL_TOTAL_BRUTO = _
'"TOTAL_BRUTO = isnull ( " & _
'"(SELECT SUM(IsNull(y.pedidototneto, 0) - isnull(y.pedidototimpuesto,0) )  As IMP_SOLES " & _
'"FROM vt_pedido y " & _
'"JOIN vt_cargo x " & _
'"ON (y.pedidonrofact = x.cargonumdoc  OR y.pedidonroboleta = x.cargonumdoc OR  y.pedidonrogiarem = x.cargonumdoc) " & _
'"WHERE " & _
'"y.pedidofechaanu IS NULL " & _
'"AND y.puntovtacodigo LIKE ('" & codpuntoventa & "') " & _
'"AND x.cargoapefecemi BETWEEN '" & DTDesde & "' AND '" & DTHasta & "' " & _
'"AND y.pedidomoneda = '01' )   ,  0 )  " & _
'"+ " & _
'"isnull(  (SELECT  SUM( ( IsNull(q.pedidototneto, 0)- isnull(q.pedidototimpuesto,0) ) * IsNull(s.tipocambioventa, 0) ) " & _
'"As IMPORTE_DOLARES " & _
'"FROM vt_pedido q " & _
'"JOIN vt_cargo r " & _
'"ON (q.pedidonrofact = r.cargonumdoc  OR q.pedidonroboleta = r.cargonumdoc OR  q.pedidonrogiarem = r.cargonumdoc) " & _
'"JOIN ct_tipocambio s " & _
'"ON r.cargoapefecemi = s.tipocambiofecha " & _
'"WHERE " & _
'"q.pedidofechaanu IS NULL " & _
'"AND q.puntovtacodigo LIKE ('" & codpuntoventa & "') " & _
'"AND r.cargoapefecemi BETWEEN '" & DTDesde & "' AND '" & DTHasta & "' " & _
'"AND q.pedidomoneda = '02' ) " & _
'" ,  0 ) "
'
'
'Set rs = VGcnx.Execute(SQL_TOTAL_NETO & " , " & SQL_TOTAL_BRUTO)
'If rs(0) > 0 Then
'    TotalNeto = rs(0)
'    TotalBruto = rs(1)
'End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'MONTO TOTAL DE VENTAS EN SOLES (NETO Y SIN IGV) - INCLUYE CAMBIOS POR DOCUMENTOS Y PEDIDOS

SQL_TOTAL_NETO = _
"SELECT TOTAL_NETO = isnull ( " & _
"(  SELECT SUM ( IsNull( " & _
"   CASE  " & _
"   WHEN  b.documentotipo = 'A' THEN  a.pedidototneto*-1  " & _
"   WHEN  b.documentotipo = 'C' THEN  a.pedidototneto  " & _
"   ELSE  a.pedidototneto  " & _
"   END   " & _
" , 0 )  ) As IMPORTE_SOLES " & _
"FROM vt_pedido a  JOIN vt_documento b  ON a.pedidotipofac = b.documentocodigo Left join vt_modoventa a1 on a1.modovtacodigo=a.modovtacodigo " & _
"WHERE isnull(a1.modovtacanje,0)='0' and a.pedidofechaanu IS NULL AND a.pedidomoneda = '01' AND a.puntovtacodigo LIKE ('" & codpuntoventa & "') " & _
"AND a.pedidofechafact BETWEEN '" & DTDesde & "' AND '" & DtHasta & "' " & _
")   ,  0  )  +  " & _
"isnull( (SELECT SUM(IsNull( " & _
"   CASE  " & _
"   WHEN  f.documentotipo = 'A' THEN  e.pedidototneto*-1  " & _
"   WHEN  f.documentotipo = 'C' THEN  e.pedidototneto  " & _
"   ELSE  e.pedidototneto  " & _
"   END   " & _
" , 0) * IsNull(s.tipocambioventa, 0) ) As IMPORTE_DOLARES " & _
"FROM vt_pedido e JOIN vt_documento f ON e.pedidotipofac = f.documentocodigo Left join vt_modoventa b1 on b1.modovtacodigo=e.modovtacodigo " & _
"LEFT JOIN ct_tipocambio s ON e.pedidofechafact = s.tipocambiofecha " & _
"WHERE isnull(b1.modovtacanje,0)='0' and e.pedidofechaanu IS NULL AND e.puntovtacodigo LIKE ('" & codpuntoventa & "') " & _
"AND e.pedidofechafact BETWEEN '" & DTDesde & "' AND '" & DtHasta & "' AND e.pedidomoneda = '02' " & _
")   , 0 ) "

SQL_TOTAL_BRUTO_Sol = _
"TOTAL_BRUTO = isnull ( " & _
"(SELECT SUM( IsNull(" & _
"   CASE  " & _
"   WHEN  x.documentotipo = 'A' THEN  y.pedidototneto*-1  " & _
"   WHEN  x.documentotipo = 'C' THEN  y.pedidototneto  " & _
"   ELSE  y.pedidototneto  " & _
"   END   " & _
" , 0) - isnull(" & _
"   CASE  " & _
"   WHEN  x.documentotipo = 'A' THEN  y.pedidototimpuesto*-1  " & _
"   WHEN  x.documentotipo = 'C' THEN  y.pedidototimpuesto  " & _
"   ELSE  y.pedidototimpuesto  " & _
"   END   " & _
" ,0)  )  As IMP_SOLES " & _
"FROM vt_pedido y  " & _
"JOIN vt_documento x ON y.pedidotipofac = x.documentocodigo " & _
" Left join vt_modoventa a1 on a1.modovtacodigo=y.modovtacodigo " & _
"WHERE isnull(a1.modovtacanje,0)='0' and y.pedidofechaanu IS NULL AND y.puntovtacodigo LIKE ('" & codpuntoventa & "') " & _
"AND y.pedidofechafact BETWEEN '" & DTDesde & "' AND '" & DtHasta & "' " & _
"AND y.pedidomoneda = '01' )   ,  0 )  + "

SQL_TOTAL_BRUTO_Dol = _
"isnull(  (SELECT  SUM(  ( IsNull( " & _
"   CASE  " & _
"   WHEN  r.documentotipo = 'A' THEN  q.pedidototneto*-1  " & _
"   WHEN  r.documentotipo = 'C' THEN  q.pedidototneto  " & _
"   ELSE  q.pedidototneto  " & _
"   END   " & _
" , 0) - isnull( " & _
"   CASE  " & _
"   WHEN  r.documentotipo = 'A' THEN  q.pedidototimpuesto*-1  " & _
"   WHEN  r.documentotipo = 'C' THEN  q.pedidototimpuesto  " & _
"   ELSE  q.pedidototimpuesto  " & _
"   END   " & _
" ,0)  ) * IsNull(s.tipocambioventa, 0)  ) As IMPORTE_DOLARES " & _
"FROM vt_pedido q " & _
"JOIN vt_documento r ON q.pedidotipofac = r.documentocodigo " & _
" Left join vt_modoventa a1 on a1.modovtacodigo=q.modovtacodigo " & _
"LEFT JOIN ct_tipocambio s ON q.pedidofechafact = s.tipocambiofecha " & _
"WHERE isnull(a1.modovtacanje,0)='0' and q.pedidofechaanu IS NULL AND q.puntovtacodigo LIKE ('" & codpuntoventa & "') " & _
"AND q.pedidofechafact BETWEEN '" & DTDesde & "' AND '" & DtHasta & "' AND q.pedidomoneda = '02' ) " & _
" ,  0 ) "


Set rs = VGCNx.Execute(SQL_TOTAL_NETO & " , " & SQL_TOTAL_BRUTO_Sol & SQL_TOTAL_BRUTO_Dol)
If rs(0) > 0 Then
    TotalNeto = rs(0)
    TotalBruto = rs(1)
End If


End Function



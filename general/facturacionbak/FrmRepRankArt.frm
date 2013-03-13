VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmRepRankArt 
   Caption         =   "Ranking de Articulos"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Checktipo 
      Caption         =   "Incluye Intercompanias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   15
      Top             =   3240
      Width           =   2535
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
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
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
      Left            =   3675
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   1305
   End
   Begin MSComCtl2.DTPicker DTHasta 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1680
      Width           =   1845
      _ExtentX        =   3254
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
      Format          =   16842753
      CurrentDate     =   37518
   End
   Begin MSComCtl2.DTPicker DTDesde 
      Height          =   405
      Left            =   2760
      TabIndex        =   1
      Top             =   1080
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
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
      Format          =   16842753
      CurrentDate     =   37518
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   600
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Height          =   375
      Index           =   0
      Left            =   2760
      MaxLength       =   8
      TabIndex        =   4
      Top             =   2760
      Width           =   1305
   End
   Begin VB.ComboBox cmbBase 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
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
      Top             =   4080
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
      Top             =   4080
      Width           =   1215
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      XcodMaxLongitud =   3
      xcodwith        =   300
      NomTabla        =   "co_multiempresas"
      TituloAyuda     =   "Busqueda de Empresas"
      ListaCampos     =   "empresacodigo(1),empresadescripcion(1),agentederetencion(1)"
      XcodCampo       =   "empresacodigo"
      XListCampo      =   "empresadescripcion"
      ListaCamposDescrip=   "Codigo,Descripcion"
      ListaCamposText =   "empresacodigo,empresadescripcion,agentederetencion"
      Requerido       =   0   'False
   End
   Begin VB.Label Leempresa 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   14
      Top             =   180
      Width           =   915
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
      Left            =   840
      TabIndex        =   12
      Top             =   600
      Width           =   1695
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
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   2475
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
      Left            =   1320
      TabIndex        =   9
      Top             =   2280
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
      Left            =   1680
      TabIndex        =   8
      Top             =   1080
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
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
End
Attribute VB_Name = "FrmRepRankArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalVentas As Double
Dim d_porcentaje As Double
Dim d_cantidad As Double
Dim d_monto As Double
Dim index_combo As Integer
Dim orderby As Integer
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
Dim Param(9) As Variant
Dim formulas(7) As Variant
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

Param(0) = TotalVentas
Param(1) = d_cantidad
Param(2) = d_porcentaje
Param(3) = d_monto
Param(4) = DTDesde
Param(5) = DtHasta
Param(6) = IIf(Trim(txt(1)) = "", "%%", Trim(txt(1)))
Param(7) = VGCNx.DefaultDatabase
Param(8) = IIf(Ctr_Ayuempresa.xclave = "", "%%", Ctr_Ayuempresa.xclave)

If Ctr_Ayuempresa.xclave = "" Then
   formulas(0) = "@Empresa='T O D A S '"
Else
   formulas(0) = "@Empresa='" & Ctr_Ayuempresa.xnombre & "'"
End If
formulas(1) = "@ruc='" & VGParametros.RucEmpresa & "'"
formulas(2) = "Desde='" & DTDesde & "'"
formulas(3) = "Hasta='" & DtHasta & "'"
formulas(4) = "EnBase='" & cmbBase.Text & "'"
formulas(5) = "Numero='" & txt(0) & "'"
If Combo1.ListIndex <> -1 Then
    formulas(6) = "Puntovta='" & Combo1.Text & "'"
Else
    formulas(6) = "Puntovta='TODOS'"
End If

Call ImpresionRptProc("vt_RankingArticulo.rpt", formulas, Param, , "Ranking de Articulos")
 
Screen.MousePointer = 1

Exit Sub
Errores:
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
  Err = 0

End Sub
Private Sub Combo1_Click()
  If Combo1.ListCount > 0 Then
     txt(1) = adll.ComboDato(Combo1.Text)
  Else
     txt(1) = ""
  End If
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

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Load()
    MostrarForm Me, "C2"
    Call adll.llenacombo(Combo1, "select puntovtacodigo,puntovtadescripcion from vt_puntoventa", VGCNx)
    DTDesde = Date
    DtHasta = Date
    Call Ctr_Ayuempresa.conexion(VGCNx)
    Carga_Combo
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
   cmbBase.AddItem ("Cantidad de Articulos")
   cmbBase.AddItem ("Porcentaje de Ventas")
   cmbBase.AddItem ("Monto de Ventas")
End Function
Private Function Consulta_Reporte()

Dim SQLSoles As String
Dim SQLDolares As String
Dim rs As New ADODB.Recordset
Dim codpuntoventa As String
TotalVentas = 0

 If Trim(txt(1)) = "" Then
    codpuntoventa = "%"
 Else
    codpuntoventa = Trim(txt(1))
 End If

If cmbBase.ListIndex = 0 Then  ' En base a cantidad de articulos
    d_cantidad = CDbl(txt(0))
    d_porcentaje = 0
    d_monto = 0

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ElseIf cmbBase.ListIndex = 1 Then   ' En base a porcentaje de ventas
    d_cantidad = 0
    'd_porcentaje = CDbl(txt(0)) / 100
    d_porcentaje = CDbl(txt(0))
    d_monto = 0

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ElseIf cmbBase.ListIndex = 2 Then   ' En base a monto de ventas
    d_cantidad = 0
    d_porcentaje = 0
    d_monto = CDbl(txt(0))
    
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MONTO TOTAL DE VENTAS EN SOLES

'SQL = _
'"SELECT TOTAL_VENTAS =   isnull ( " & _
'"(SELECT SUM(IsNull(z.detpedmontoimpto, 0)) As IMPORTE_SOLES " & _
'"FROM  vt_detallepedido z " & _
'"JOIN vt_pedido y " & _
'"ON z.pedidonumero = y.pedidonumero " & _
'"JOIN vt_cargo X " & _
'"ON (y.pedidonrofact = x.cargonumdoc  OR y.pedidonroboleta = x.cargonumdoc OR  y.pedidonrogiarem = x.cargonumdoc) " & _
'"WHERE " & _
'"y.pedidofechaanu IS NULL AND y.puntovtacodigo LIKE ('" & codpuntoventa & "') " & _
'"AND x.cargoapefecemi BETWEEN '" & DTDesde & "' AND '" & DTHasta & "' " & _
'"AND y.pedidomoneda = '01' )   ,  0 ) " & _
'"+ " & _
'"isnull ( (SELECT SUM(IsNull(p.detpedmontoimpto, 0) * IsNull(s.tipocambioventa, 0)) As IMPORTE_DOLARES " & _
'"FROM  vt_detallepedido p " & _
'"JOIN vt_pedido q " & _
'"ON p.pedidonumero = q.pedidonumero " & _
'"JOIN vt_cargo r " & _
'"ON (q.pedidonrofact = r.cargonumdoc  OR q.pedidonroboleta = r.cargonumdoc OR  q.pedidonrogiarem = r.cargonumdoc) " & _
'"JOIN ct_tipocambio s " & _
'"ON r.cargoapefecemi = s.tipocambiofecha " & _
'"WHERE " & _
'"q.pedidofechaanu IS NULL AND q.puntovtacodigo LIKE ('" & codpuntoventa & "') " & _
'"AND r.cargoapefecemi BETWEEN '" & DTDesde & "' AND '" & DTHasta & "' " & _
'"AND q.pedidomoneda = '02' )   , 0 ) "
'
'Set rs = VGcnx.Execute(SQL)
'If rs(0) > 0 Then
'    TotalVentas = rs(0)
'End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MONTO TOTAL DE VENTAS EN SOLES  - INCLUYE CAMBIOS POR DOCUMENTOS Y PEDIDOS

SQLSoles = _
"SELECT TOTAL_VENTAS =   isnull ( " & _
"(SELECT SUM(IsNull( " & _
"   CASE  " & _
"   WHEN  x.documentotipo = 'A' THEN  z.detpedmontoimpto*-1  " & _
"   WHEN  x.documentotipo = 'C' THEN  z.detpedmontoimpto  " & _
"   ELSE  z.detpedmontoimpto  " & _
"   END   " & _
" ,0)  ) As IMPORTE_SOLES " & _
"FROM  vt_detallepedido z " & _
"JOIN vt_pedido y " & _
"ON z.pedidonumero = y.pedidonumero " & _
"JOIN vt_documento x " & _
"ON y.pedidotipofac = x.documentocodigo " & _
" Left JOIN vt_modoventa q " & _
"ON q.modovtacodigo = y.modovtacodigo " & _
"WHERE isnull(q.modovtacanje,0)='0' and " & _
"y.pedidofechaanu IS NULL AND y.puntovtacodigo LIKE ('" & codpuntoventa & "') " & _
"AND y.pedidofechafact BETWEEN '" & DTDesde & "' AND '" & DtHasta & "' " & _
"AND y.pedidomoneda = '01' and z.empresacodigo='" & VGParametros.empresacodigo & "' )   ,  0 ) "

SQLDolares = _
" + " & _
"isnull ( (SELECT SUM( IsNull( " & _
"   CASE  " & _
"   WHEN  r.documentotipo = 'A' THEN  p.detpedmontoimpto*-1  " & _
"   WHEN  r.documentotipo = 'C' THEN  p.detpedmontoimpto  " & _
"   ELSE  p.detpedmontoimpto  " & _
"   END   " & _
" , 0) * IsNull(s.tipocambioventa, 0)  ) As IMPORTE_DOLARES " & _
"FROM  vt_detallepedido p " & _
"JOIN vt_pedido q " & _
"ON p.pedidonumero = q.pedidonumero " & _
"JOIN vt_documento r " & _
"ON q.pedidotipofac = r.documentocodigo " & _
"LEFT JOIN ct_tipocambio s " & _
"ON q.pedidofechafact = s.tipocambiofecha " & _
"LEFT JOIN vt_modoventa a1 " & _
"ON a1.modovtacodigo = q.modovtacodigo " & _
"WHERE isnull(a1.modovtacanje,0)='0' and " & _
"q.pedidofechaanu IS NULL AND q.puntovtacodigo LIKE ('" & codpuntoventa & "') " & _
"AND q.pedidofechafact BETWEEN '" & DTDesde & "' AND '" & DtHasta & "' " & _
"AND q.pedidomoneda = '02'  and p.empresacodigo='" & VGParametros.empresacodigo & "' )   , 0 ) "
            
Set rs = VGCNx.Execute(SQLSoles & SQLDolares)
If rs(0) > 0 Then
    TotalVentas = rs(0)
End If


End Function


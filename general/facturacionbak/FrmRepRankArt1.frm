VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmRepRankArt 
   Caption         =   "Ranking de Articulos"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   240
      Top             =   2280
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
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   8
      Top             =   1680
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
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1200
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
      Left            =   2640
      TabIndex        =   1
      Top             =   2280
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
      Left            =   1200
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTHasta 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   60882945
      CurrentDate     =   37489
   End
   Begin MSComCtl2.DTPicker DTDesde 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   60882945
      CurrentDate     =   37489
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
      TabIndex        =   9
      Top             =   1680
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
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "Fecha Desde"
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
      Left            =   1200
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "Fecha Hasta"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmRepRankArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CadCodProd As String
Dim TotalVentas As Double
Dim index_combo As Integer

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

 If DTDesde > DTHasta Then
     MsgBox "Fecha Desde debe ser mayor a Fecha Hasta", vbInformation, "AVISO"
     Exit Sub
 End If
                 
 If cmbBase.ListIndex = -1 Or txt(0) = "" Then
     MsgBox "Ingrese Parámetro En base a", vbInformation, "AVISO"
     Exit Sub
 End If
 
  Call Consulta_Reporte
  'MDIPrincipal.oCrystalReport.Connect = CadenaRep
  oCrystalReport.ReportFileName = RutaRepProc & "RepRankArt.rpt"
  oCrystalReport.Destination = crptToWindow
  oCrystalReport.WindowState = crptMaximized
  oCrystalReport.DiscardSavedData = True
  'MDIPrincipal.oCrystalReport.StoredProcParam(0) = DTDesde
  'MDIPrincipal.oCrystalReport.StoredProcParam(1) = DTHasta
  'MDIPrincipal.oCrystalReport.StoredProcParam(2) = IIf(cmbBase.ListIndex = 0, CDbl(txt(0)), 0)
  'MDIPrincipal.oCrystalReport.StoredProcParam(3) = IIf(cmbBase.ListIndex = 1, CDbl(txt(0)) / 100, 0)
  'MDIPrincipal.oCrystalReport.StoredProcParam(4) = IIf(cmbBase.ListIndex = 2, CDbl(txt(0)), 0)
  oCrystalReport.StoredProcParam(0) = CadCodProd
  oCrystalReport.StoredProcParam(1) = TotalVentas
  oCrystalReport.Action = 1
  
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub


Private Sub DTDesde_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DTHasta_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Form_Load()
    MostrarForm Me, "C2"
    DTDesde = Date
    DTHasta = Date
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
 txt(Index).Text = Format(txt(Index).Text, "#0.00")
End Sub

Private Function Carga_Combo()
   cmbBase.Clear
   'cmbBase.AddItem ("Monto de Ventas")
   'cmbBase.AddItem ("Porcentaje de Ventas")
   cmbBase.AddItem ("Cantidad de Articulos")
End Function
Private Function Consulta_Reporte()

Dim sql As String
Dim RS As New ADODB.Recordset
Dim PORC_ACUM As Double

PORC_ACUM = 0
TotalVentas = 0
CadCodProd = " "

If cmbBase.ListIndex = 0 Then  ' En base a cantidad de articulos
  sql = "SELECT TOP " & CInt(txt(0)) & " " & _
       "a.productocodigo as CODIGO_PRODUCTO," & _
       "b.productodescripcion as PRODUCTO," & _
       "SUM(isnull(a.detpedcantentreg,0)) as CANTIDAD," & _
       "SUM(Isnull(a.detpedmontoimpto, 0)) As IMPORTE " & _
       "FROM " & _
       "vt_detallepedido a " & _
       "Join " & _
       "vt_producto b " & _
       "ON " & _
       "a.productocodigo = b.productocodigo " & _
       "Join " & _
       "vt_pedido c " & _
       "ON " & _
       "a.pedidonumero = c.pedidonumero " & _
       "WHERE " & _
       "c.pedidofecha between '" & DTDesde & "' AND '" & DTHasta & "' " & _
       "AND isnull(c.pedidofechaanu,'') = '' " & _
       "GROUP BY " & _
       "a.productocodigo , b.productodescripcion " & _
       "ORDER BY " & _
       "CANTIDAD Desc"
       
        'CADENA DE CODIGOS DE PRODUCTO - CANTIDAD DE ARTICULOS
        Set RS = cn.Execute(sql)
        If RS.RecordCount > 0 Then
            RS.MoveFirst
            Do While Not RS.EOF
               'CadCodProd = CadCodProd & "'" & RS(0) & "',"
                CadCodProd = CadCodProd & RS(0) & ","
                RS.MoveNext
            Loop
            CadCodProd = Left(CadCodProd, Len(RTrim(CadCodProd)) - 1)
        Else
            CadCodProd = " "
        End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ElseIf cmbBase.ListIndex = 1 Then   ' En base a porcentaje de ventas
    'sql = "SELECT TOP " & CInt(txt(0)) & " PERCENT WITH TIES " &
    sql = "SELECT " & _
        "a.productocodigo as CODIGO_PRODUCTO," & _
        "b.productodescripcion as PRODUCTO," & _
        "SUM(isnull(a.detpedcantentreg,0)) as CANTIDAD," & _
        "SUM(isnull(a.detpedmontoimpto,0)) as IMPORTE," & _
        "PORCENTAJE_VTAS =  SUM(isnull(a.detpedmontoimpto,0))/" & _
        "   (SELECT SUM (pedidototneto) " & _
        "    FROM vt_pedido WHERE isnull(pedidofechaanu,'') = '') " & _
        "FROM " & _
        "vt_detallepedido a " & _
        "JOIN " & _
        "vt_producto b " & _
        "ON " & _
        "a.productocodigo = b.productocodigo " & _
        "JOIN " & _
        "vt_pedido c " & _
        "ON " & _
        "a.pedidonumero = c.pedidonumero " & _
        "WHERE " & _
        "c.pedidofecha BETWEEN '" & DTDesde & "' AND '" & DTHasta & "' " & _
        "AND isnull(c.pedidofechaanu,'') = '' " & _
        "GROUP BY  " & _
        "a.productocodigo , b.productodescripcion " & _
        "ORDER BY " & _
        "PORCENTAJE_VTAS Desc "
        
        'CADENA DE CODIGOS DE PRODUCTO - PORCENTAJE DE VENTAS
    Set RS = cn.Execute(sql)
    If RS.RecordCount > 0 Then
        RS.MoveFirst
        Do While Not RS.EOF
            PORC_ACUM = PORC_ACUM + RS(4)
            If PORC_ACUM <= txt(0) / 100 Then
                CadCodProd = CadCodProd & RS(0) & ","
            'Else
            '    Exit Do
            End If
            RS.MoveNext
        Loop
        If CadCodProd <> " " Then
            CadCodProd = Left(CadCodProd, Len(RTrim(CadCodProd)) - 1)
        End If
    Else
        CadCodProd = ""
    End If
    
End If

'MONTO TOTAL DE VENTAS
sql = "SELECT SUM (pedidototneto) " & _
      "FROM vt_pedido WHERE isnull(pedidofechaanu,'') = '' "
Set RS = cn.Execute(sql)
If RS(0) > 0 Then
    TotalVentas = RS(0)
End If

End Function

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmRepArtxProveedor 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   825
      Left            =   315
      TabIndex        =   32
      Top             =   1215
      Width           =   6045
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_ayuAlmacen 
         Height          =   315
         Left            =   990
         TabIndex        =   34
         Top             =   270
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         XcodMaxLongitud =   2
         xcodwith        =   300
         NomTabla        =   "tabalm"
         TituloAyuda     =   "Busqueda de Almacenes"
         ListaCampos     =   "TAALMA(1),TADESCRI(1),almacenvalorizado(1)"
         XcodCampo       =   "taalma"
         XListCampo      =   "tadescri"
         ListaCamposDescrip=   "Código,Descripción,Valorizacion"
         ListaCamposText =   "TAALMA,TADESCRI,almacenvalorizado"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Almacen"
         Height          =   195
         Left            =   135
         TabIndex        =   33
         Top             =   315
         Width           =   615
      End
   End
   Begin VB.Frame FrameProveedor 
      Height          =   1170
      Left            =   225
      TabIndex        =   21
      Top             =   2100
      Width           =   6165
      Begin VB.CheckBox Check2 
         Caption         =   "todos"
         Height          =   255
         Left            =   5040
         TabIndex        =   28
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "todos"
         Height          =   255
         Left            =   5040
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuProveedor1 
         Height          =   315
         Left            =   1080
         TabIndex        =   22
         Top             =   360
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   556
         XcodMaxLongitud =   11
         xcodwith        =   1200
         NomTabla        =   "cp_proveedor"
         TituloAyuda     =   "Busqueda de Proveedor"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),clienteruc(1),clientetelefono(1),proveedorcontribuyente(2)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "clientecodigo,clienterazonsocial,clienteruc,clientetelefono,proveedorcontribuyente"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAnalitico 
         Height          =   315
         Left            =   1080
         TabIndex        =   29
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         XcodMaxLongitud =   11
         xcodwith        =   900
         TituloAyuda     =   "Busqueda de Centro de Costos"
         ListaCampos     =   "entidadcodigo(1),entidadrazonsocial(1)"
         XcodCampo       =   "entidadcodigo"
         XListCampo      =   "entidadrazonsocial"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "entidadcodigo,entidadrazonsocial"
         Requerido       =   0   'False
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   390
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Entidad"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   765
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Height          =   990
      Left            =   300
      TabIndex        =   18
      Top             =   180
      Width           =   6105
      Begin VB.OptionButton OptEntidad 
         Caption         =   "Entidades"
         Height          =   495
         Left            =   1800
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptProveedor 
         Caption         =   "Proveedores"
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5040
         Picture         =   "FrmRepArtxProveedor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   195
         Width           =   775
      End
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   4020
         Picture         =   "FrmRepArtxProveedor.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   195
         Width           =   780
      End
   End
   Begin VB.Frame Reporte 
      Height          =   570
      Left            =   300
      TabIndex        =   15
      Top             =   5610
      Width           =   6045
      Begin VB.OptionButton Optmensual 
         Caption         =   "Mensualizado"
         Height          =   255
         Left            =   4440
         TabIndex        =   26
         Top             =   200
         Width           =   1455
      End
      Begin VB.OptionButton OptDetallado 
         Caption         =   "Detallado"
         Height          =   270
         Left            =   2640
         TabIndex        =   17
         Top             =   200
         Width           =   1065
      End
      Begin VB.OptionButton OptResumido 
         Caption         =   "Resumido"
         Height          =   315
         Left            =   960
         TabIndex        =   16
         Top             =   200
         Value           =   -1  'True
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1440
      Left            =   240
      TabIndex        =   5
      Top             =   3285
      Width           =   6105
      Begin VB.TextBox TxtMov1 
         Height          =   300
         Left            =   2010
         TabIndex        =   11
         Top             =   615
         Width           =   800
      End
      Begin VB.TextBox TxtMov2 
         Height          =   300
         Left            =   3630
         TabIndex        =   10
         Top             =   615
         Width           =   800
      End
      Begin VB.TextBox txtMov3 
         Height          =   300
         Left            =   5085
         TabIndex        =   9
         Top             =   615
         Width           =   800
      End
      Begin VB.TextBox txtDev1 
         Height          =   300
         Left            =   2010
         TabIndex        =   8
         Top             =   1020
         Width           =   800
      End
      Begin VB.TextBox txtDev2 
         Height          =   300
         Left            =   3645
         TabIndex        =   7
         Top             =   1020
         Width           =   800
      End
      Begin VB.TextBox txtDev3 
         Height          =   300
         Left            =   5085
         TabIndex        =   6
         Top             =   1020
         Width           =   800
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuMoneda 
         Height          =   315
         Left            =   2040
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         XcodMaxLongitud =   2
         NomTabla        =   "gr_moneda"
         TituloAyuda     =   "Busqueda de Moneda"
         ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
         XcodCampo       =   "monedacodigo"
         XListCampo      =   "monedadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "monedacodigo,monedadescripcion"
      End
      Begin VB.Label Label4 
         Caption         =   "Moneda Reporte"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Mov. de Compras :"
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   690
         Width           =   1560
      End
      Begin VB.Label Label5 
         Caption         =   "Dev. Compras :"
         Height          =   285
         Left            =   255
         TabIndex        =   12
         Top             =   1020
         Width           =   1320
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Fecha"
      Height          =   810
      Left            =   300
      TabIndex        =   0
      Top             =   4770
      Width           =   6000
      Begin MSComCtl2.DTPicker txtHFec 
         Height          =   300
         Left            =   4620
         TabIndex        =   1
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   37445
      End
      Begin MSComCtl2.DTPicker txtDFec 
         Height          =   270
         Left            =   1620
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   476
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   37445
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   330
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   3720
         TabIndex        =   3
         Top             =   345
         Width           =   420
      End
   End
End
Attribute VB_Name = "FrmRepArtxProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Mensaje As String
Dim strsql As String
Const activo = &H80000005
Const inactivo = &H8000000F

Private Sub CmbMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       TxtMov1.SetFocus
    End If
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       txtDFec.SetFocus
    End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Ctr_AyuProveedor1.Enabled = True
 Else
   Ctr_AyuProveedor1.Enabled = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   Ctr_AyuAnalitico.Enabled = True
 Else
   Ctr_AyuAnalitico.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
'    MDIPrincipal.cancelado = True
    Unload Me
End Sub

Private Sub cmdImp_Click()
Dim vrep As String
Dim aparam(8) As Variant
Dim aform(6) As Variant
Dim Reporte As String
Dim filtro As String
On Error GoTo Err
    
If Ctr_ayuAlmacen.xclave = "" Then
   Mensaje = "Ingrese Codigo de Almacen "
   MsgBox Mensaje, vbExclamation, "Error"
  Ctr_ayuAlmacen.SetFocus
  Exit Sub
End If
If Check1.Value = 0 And Ctr_AyuProveedor1.xclave = "" Then
            Mensaje = "Ingrese Codigo de Proveedor "
            MsgBox Mensaje, vbExclamation, "Error"
            Ctr_AyuProveedor1.SetFocus
            Exit Sub
End If
If Check2.Value = 0 And Ctr_AyuAnalitico.xclave = "" Then
            Mensaje = "Ingrese Codigo de Entidad "
            MsgBox Mensaje, vbExclamation, "Error"
            Ctr_AyuAnalitico.SetFocus
            Exit Sub
End If
   
If txtDFec > txtHFec Then
        Mensaje = "Fecha Desde no puede ser posterior a Fecha Hasta"
        MsgBox Mensaje, vbExclamation, "Error"
        txtDFec.SetFocus
        Exit Sub
End If
If OptProveedor.Value = True Then
   If OptDetallado.Value = True Then
      Reporte = "al_comprasxproveedor.rpt"
    ElseIf OptResumido.Value = True Then
           Reporte = "al_comprasxproveedorResumen.rpt"
         Else
           Reporte = "al_comprasxproveedorMensual.rpt"
   End If
Else
   If OptDetallado.Value = True Then
      Reporte = "al_comprasxentidad.rpt"
    ElseIf OptResumido.Value = True Then
           Reporte = "al_comprasxentidadResumen.rpt"
         Else
           Reporte = "al_comprasxentidadMensual.rpt"
   End If
End If

    aform(0) = "xMoneda='" & Ctr_AyuMoneda.xclave & "'"
    aform(1) = "xDprov='" & Ctr_AyuProveedor1.xclave & "'"
    aform(2) = "xHProv='" & Ctr_AyuAnalitico.xnombre & "'"
    aform(3) = "xDFecha='" & txtDFec.Value & "'"
    aform(4) = "xHFecha='" & txtHFec.Value & "'"
    aform(5) = "xAlmacen='" & Ctr_ayuAlmacen.xnombre & "'"
    
    aparam(0) = VGCNx.DefaultDatabase
    If Ctr_AyuProveedor1.xclave <> "" Then
       aparam(1) = Ctr_AyuProveedor1.xclave
     Else
       aparam(1) = "%%"
    End If
    If Check2.Value = 0 Then
       aparam(2) = Ctr_AyuAnalitico.xclave
     Else
       aparam(2) = "%%"
    End If
    aparam(3) = txtDFec.Value
    aparam(4) = txtHFec.Value
    If Ctr_AyuMoneda.xclave <> "" Then
       aparam(5) = Ctr_AyuMoneda.xclave
     Else
       aparam(5) = "%%"
    End If
    Call filtrotransa
    aparam(6) = VGComputer + "_compras"
    aparam(7) = Ctr_ayuAlmacen.xclave
    Call ImpresionRptProc(Reporte, aform, aparam, , " Compras x proveedor " & Reporte)
Exit Sub
Err:
  MsgBox Err.Description
  Exit Sub
  Resume

End Sub
Private Sub filtrotransa()
Dim rsql As New ADODB.Recordset
SQL = VGComputer + "_compras"
If ExisteElem(0, VGCNx, SQL) Then Set rsql = VGCNx.Execute(" drop table " & SQL & "")
   Set rsql = VGCNx.Execute("create table " & SQL & " ( transa nvarchar(3) null )")

If TxtMov1.text <> "" Then Set rsql = VGCNx.Execute("insert " & SQL & " ( transa) values ('I" & TxtMov1.text & "')")
If TxtMov2.text <> "" Then Set rsql = VGCNx.Execute("insert " & SQL & " ( transa) values ('I" & TxtMov2.text & "')")
If txtMov3.text <> "" Then Set rsql = VGCNx.Execute("insert " & SQL & " ( transa) values ('I" & txtMov3.text & "')")
If txtDev1.text <> "" Then Set rsql = VGCNx.Execute("insert " & SQL & " ( transa) values ('S" & txtDev1.text & "')")
If txtDev2.text <> "" Then Set rsql = VGCNx.Execute("insert " & SQL & " ( transa) values ('S" & txtDev2.text & "')")
If txtDev3.text <> "" Then Set rsql = VGCNx.Execute("insert " & SQL & " ( transa) values ('S" & txtDev3.text & "')")

End Sub


Private Sub Form_Load()
    txtDFec = VGParamSistem.FechaTrabajo
    txtHFec = VGParamSistem.FechaTrabajo
    Call Ctr_AyuProveedor1.Conexion(VGCNx)
    Call Ctr_AyuAnalitico.Conexion(VGCNx)
    Call Ctr_AyuMoneda.Conexion(VGCNx)
    Call Ctr_ayuAlmacen.Conexion(VGCNx)
    Check1.Value = 1
    Check2.Value = 1
End Sub

Private Sub OptDetallado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    End If
End Sub

Private Sub OptResumido_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        SendKeys ("{TAB}")
    End If
End Sub

Private Sub txtDev1_DblClick()
Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT TT_CODMOV,TT_DESCRI FROM TABTRANSA WHERE TT_TIPMOV='S'"
    Adodc2.Open strsql, VGCNx, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Transacciones"
    frmReferencia.inicio
    frmReferencia.Show vbModal
    Rem MVV Adodc2.Close
    Set Adodc2 = Nothing
    
    If vGUtil(1) <> "" Then
         txtDev1.text = vGUtil(1)
         txtDev2.SetFocus
    End If

End Sub

Private Sub txtDev1_GotFocus()
  Enfoque txtDev1
End Sub

Private Sub txtDev1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then txtDev1_DblClick
End Sub

Private Sub txtDev1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtDev1 = Trim(txtDev1)
        If txtDev1 <> "" Then
            If Not Existe(1, txtDev1, "Tabtransa", "TT_CODMOV", False) Then
                Mensaje = "El Código de la Transacción  no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                Enfoque txtDev1
            End If
        End If
        txtDev2.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtDFec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidFecha(txtDFec) Then
            Mensaje = "Fecha No Válida"
            MsgBox Mensaje, vbExclamation, "Error"
            txtDFec.SetFocus
        Else
            txtHFec.SetFocus
        End If
    End If
End Sub


Private Sub txtDFec_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
        SendKeys ("{TAB}")
    End If
End Sub



Private Sub txtHFec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidFecha(txtHFec) Then
            Mensaje = "Fecha No Válida"
            MsgBox Mensaje, vbExclamation, "Error"
            txtHFec.SetFocus
        Else
            cmdImp.SetFocus
        End If
    End If
End Sub

Private Sub txtHFec_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
        SendKeys ("{TAB}")
    End If
End Sub

Private Sub TxtMov1_DblClick()
    Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT TT_CODMOV,TT_DESCRI FROM TABTRANSA "
    Adodc2.Open strsql, VGCNx, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Transacciones"
    frmReferencia.inicio
    frmReferencia.Show vbModal
    Rem MVV Adodc2.Close
    Set Adodc2 = Nothing
    
    If vGUtil(1) <> "" Then
        TxtMov1.text = vGUtil(1)
        TxtMov2.SetFocus
    End If
End Sub

Private Sub TxtMov1_GotFocus()
  Enfoque TxtMov1
End Sub

Private Sub TxtMov1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then TxtMov1_DblClick
End Sub

Private Sub TxtMov1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        TxtMov1 = Trim(TxtMov1)
        If TxtMov1 <> "" Then
            If Not Existe(1, TxtMov1, "Tabtransa", "TT_CODMOV", False) Then
                Mensaje = "El Código de la Transacción  no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                Enfoque TxtMov1
            End If
        End If
        TxtMov2.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub TxtMov2_DblClick()
Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT TT_CODMOV,TT_DESCRI FROM TABTRANSA "
    Adodc2.Open strsql, VGCNx, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Transacciones"
    frmReferencia.inicio
    frmReferencia.Show vbModal
    Rem MVV Adodc2.Close
    Set Adodc2 = Nothing
    
    If vGUtil(1) <> "" Then
        TxtMov2.text = vGUtil(1)
        txtMov3.SetFocus
    End If
End Sub

Private Sub TxtMov2_GotFocus()
Enfoque TxtMov2
End Sub

Private Sub TxtMov2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 112 Then TxtMov2_DblClick
End Sub

Private Sub TxtMov2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        TxtMov2 = Trim(TxtMov2)
        If TxtMov2 <> "" Then
            If Not Existe(1, TxtMov2, "Tabtransa", "TT_CODMOV", False) Then
                Mensaje = "El Código de la Transacción  no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                Enfoque TxtMov2
            End If
        End If
        txtMov3.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtMov3_DblClick()
Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT TT_CODMOV,TT_DESCRI FROM TABTRANSA "
    Adodc2.Open strsql, VGCNx, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Transacciones"
    frmReferencia.inicio
    frmReferencia.Show vbModal
    Rem MVV Adodc2.Close
    Set Adodc2 = Nothing
    
    If vGUtil(1) <> "" Then
        txtMov3.text = vGUtil(1)
         txtDev1.SetFocus
    End If

End Sub

Private Sub txtMov3_GotFocus()
  Enfoque txtMov3
End Sub

Private Sub txtMov3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 112 Then txtMov3_DblClick
End Sub

Private Sub txtMov3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtMov3 = Trim(txtMov3)
        If txtMov3 <> "" Then
            If Not Existe(1, txtMov3, "Tabtransa", "TT_CODMOV", False) Then
                Mensaje = "El Código de la Transacción  no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                Enfoque txtMov3
            End If
        End If
        txtDev1.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtdev2_DblClick()
Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT TT_CODMOV,TT_DESCRI FROM TABTRANSA WHERE TT_TIPMOV='S'"
    Adodc2.Open strsql, VGCNx, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Transacciones"
    frmReferencia.inicio
    frmReferencia.Show vbModal
    Rem MVV Adodc2.Close
    Set Adodc2 = Nothing
    
    If vGUtil(1) <> "" Then
         txtDev2.text = vGUtil(1)
         txtDev3.SetFocus
    End If

End Sub

Private Sub txtdev2_GotFocus()
  Enfoque txtDev2
End Sub

Private Sub txtdev2_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then txtdev2_DblClick
End Sub

Private Sub txtdev2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtDev2 = Trim(txtDev2)
        If txtDev2 <> "" Then
            If Not Existe(1, txtDev2, "Tabtransa", "TT_CODMOV", False) Then
                Mensaje = "El Código de la Transacción  no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                Enfoque txtDev2
            End If
        End If
        txtDev3.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

'**
Private Sub txtdev3_DblClick()
Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT TT_CODMOV,TT_DESCRI FROM TABTRANSA WHERE TT_TIPMOV='S'"
    Adodc2.Open strsql, VGCNx, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Transacciones"
    frmReferencia.inicio
    frmReferencia.Show vbModal
    Rem MVV Adodc2.Close
    Set Adodc2 = Nothing
    
    If vGUtil(1) <> "" Then
         txtDev3.text = vGUtil(1)
         txtDFec.SetFocus
    End If

End Sub

Private Sub txtdev3_GotFocus()
  Enfoque txtDev3
End Sub

Private Sub txtdev3_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then txtdev3_DblClick
End Sub

Private Sub txtdev3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtDev3 = Trim(txtDev3)
        If txtDev3 <> "" Then
            If Not Existe(1, txtDev3, "Tabtransa", "TT_CODMOV", False) Then
                Mensaje = "El Código de la Transacción  no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                Enfoque txtDev3
            End If
        End If
        txtDFec.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub



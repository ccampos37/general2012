VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepIngAlmCompMeses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadísticas de ingresos por Compras - Comparativo por meses"
   ClientHeight    =   4980
   ClientLeft      =   3225
   ClientTop       =   2385
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5760
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4920
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Caption         =   "Fecha"
      Height          =   1245
      Left            =   60
      TabIndex        =   19
      Top             =   2700
      Width           =   3060
      Begin MSComCtl2.DTPicker txtHFec 
         Height          =   300
         Left            =   930
         TabIndex        =   4
         Top             =   735
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MMMM 'del' yyyy"
         Format          =   59375619
         CurrentDate     =   37445
      End
      Begin MSComCtl2.DTPicker txtDFec 
         Height          =   270
         Left            =   945
         TabIndex        =   3
         Top             =   285
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   476
         _Version        =   393216
         CustomFormat    =   "MMMM 'del' yyyy"
         Format          =   59375619
         CurrentDate     =   37445
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   765
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   330
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1440
      Left            =   60
      TabIndex        =   16
      Top             =   1230
      Width           =   5625
      Begin VB.TextBox txtDev3 
         Height          =   300
         Left            =   4365
         TabIndex        =   28
         Top             =   1020
         Width           =   800
      End
      Begin VB.TextBox txtDev2 
         Height          =   300
         Left            =   3165
         TabIndex        =   27
         Top             =   1020
         Width           =   800
      End
      Begin VB.TextBox txtDev1 
         Height          =   300
         Left            =   2010
         TabIndex        =   26
         Top             =   1020
         Width           =   800
      End
      Begin VB.TextBox txtMov3 
         Height          =   300
         Left            =   4365
         TabIndex        =   24
         Top             =   615
         Width           =   800
      End
      Begin VB.TextBox TxtMov2 
         Height          =   300
         Left            =   3150
         TabIndex        =   23
         Top             =   615
         Width           =   800
      End
      Begin VB.TextBox TxtMov1 
         Height          =   300
         Left            =   2010
         TabIndex        =   22
         Top             =   615
         Width           =   800
      End
      Begin VB.ComboBox cmbMoneda 
         Height          =   315
         ItemData        =   "frmRepIngAlmCompMeses.frx":0000
         Left            =   2010
         List            =   "frmRepIngAlmCompMeses.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "Dev. Compras :"
         Height          =   285
         Left            =   255
         TabIndex        =   25
         Top             =   1020
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Mov. de Compras :"
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   690
         Width           =   1560
      End
      Begin VB.Label Label4 
         Caption         =   "Moneda "
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   315
         Width           =   795
      End
   End
   Begin VB.Frame Reporte 
      Caption         =   "Reporte por"
      Height          =   1245
      Left            =   3180
      TabIndex        =   15
      Top             =   2700
      Width           =   2505
      Begin VB.OptionButton Optambos 
         Caption         =   "Ambos"
         Height          =   270
         Left            =   270
         TabIndex        =   29
         Top             =   915
         Width           =   1785
      End
      Begin VB.OptionButton OptCantidad 
         Caption         =   "Cantidad"
         Height          =   315
         Left            =   270
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1710
      End
      Begin VB.OptionButton Optimporte 
         Caption         =   "Importe"
         Height          =   270
         Left            =   270
         TabIndex        =   6
         Top             =   585
         Width           =   1785
      End
   End
   Begin VB.Frame Frame2 
      Height          =   990
      Left            =   60
      TabIndex        =   14
      Top             =   3930
      Width           =   5625
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   1980
         Picture         =   "frmRepIngAlmCompMeses.frx":0016
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   195
         Width           =   780
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   2865
         Picture         =   "frmRepIngAlmCompMeses.frx":0458
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   195
         Width           =   775
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Artículo"
      Height          =   1170
      Left            =   45
      TabIndex        =   9
      Top             =   45
      Width           =   5625
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   0
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   0
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.TextBox txtDArt 
         Height          =   285
         Left            =   960
         MaxLength       =   8
         TabIndex        =   0
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox txtHArt 
         Height          =   285
         Left            =   960
         MaxLength       =   8
         TabIndex        =   1
         Top             =   705
         Width           =   975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   765
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   390
         Width           =   465
      End
      Begin VB.Label lblDArt 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2040
         TabIndex        =   11
         Top             =   300
         Width           =   3375
      End
      Begin VB.Label lblHArt 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2040
         TabIndex        =   10
         Top             =   705
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmRepIngAlmCompMeses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Mensaje As String
Dim strsql As String
Const activo = &H80000005
Const inactivo = &H8000000F

Private Sub cmbMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       TxtMov1.SetFocus
    End If
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       txtDFec.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
'    MDIPrincipal.cancelado = True
    Unload Me
End Sub

Private Sub cmdImp_Click()
Dim vrep As String
'On Error GoTo err
    
If txtDArt.text = "" Then
     MsgBox "Debe seleccionar codigo de artículo inicial", vbInformation, "sistema de Compras"
     txtDArt.SetFocus
  Exit Sub
  End If
  
  If txtHArt.text = "" Then
     MsgBox "Debe seleccionar codigo de artículo final", vbInformation, "sistema de Compras"
     txtHArt.SetFocus
  Exit Sub
  End If
  
  If txtDArt.text > txtHArt.text Then
     MsgBox "El artículo inicial no debe ser mayor al artículo final", vbInformation, "sistema de Compras"
     txtDArt.SetFocus
  Exit Sub
  End If
  
    If txtDArt <> "" And lblDArt = "" Then
        If Not Existe(1, txtDArt, "maeart", "acodigo", False) Then
            Mensaje = "El Código de Artículo no existe"
            MsgBox Mensaje, vbExclamation, "Error"
            txtDArt.SetFocus
            Exit Sub
        Else
            lblDArt = Devolver_Dato(1, txtDArt, "maeart", "acodigo", False, "adescri")
        End If
    End If
    If txtHArt <> "" And lblHArt = "" Then
        If Not Existe(1, txtHArt, "maeart", "acodigo", False) Then
            Mensaje = "El Código de Artículo no existe"
            MsgBox Mensaje, vbExclamation, "Error"
            txtHArt.SetFocus
            Exit Sub
        Else
            lblHArt = Devolver_Dato(1, txtHArt, "maeart", "acodigo", False, "adescri")
        End If
    End If
    If txtDArt = "" Then
        If txtHArt <> "" Then
            Mensaje = "No ha especificado Desde Artículo"
            MsgBox Mensaje, vbExclamation, "Error"
            txtDArt.SetFocus
            Exit Sub
        End If
    Else
        If txtHArt = "" Then
            Mensaje = "No ha especificado Hasta Artículo"
            MsgBox Mensaje, vbExclamation, "Error"
            txtHArt.SetFocus
            Exit Sub
        End If
    End If
    If txtDArt <> "" Then
        If txtDArt > txtHArt Then
            Mensaje = "Desde Artículo y Hasta Artículo no forman un rango válido"
            MsgBox Mensaje, vbExclamation, "Error"
            txtDArt.SetFocus
            Exit Sub
        End If
    End If
    
  '  If Not ValidFecha(txtDFec) Then
  '      mensaje = "Fecha No Válida"
  '      MsgBox mensaje, vbExclamation, "Error"
  '      txtDFec.SetFocus
  '  End If
  '  If Not ValidFecha(txtHFec) Then
  '      mensaje = "Fecha No Válida"
  ''      MsgBox mensaje, vbExclamation, "Error"
   '     txtDFec.SetFocus
   ' End If
    
    If txtDFec > txtHFec Then
        Mensaje = "Fecha Desde no puede ser posterior a Fecha Hasta"
        MsgBox Mensaje, vbExclamation, "Error"
        txtDFec.SetFocus
        Exit Sub
    End If

 Screen.MousePointer = 11
 
 If Genera_Reporte Then
 
   If OptCantidad.Value = True Then
       vrep = "\comp0038.rpt"
   Else
     If Optimporte.Value = True Then
       vrep = "\comp0039.rpt"
     Else
        vrep = "\comp0040.rpt"
    End If
   End If
   
    CrystalReport1.Reset
    Rem MVV Data1.DatabaseName = strsql

    CrystalReport1.ReportFileName = cRutP & vrep
    Rem MVV Ubi_Tab CrystalReport1
    Detalle CrystalReport1
    CrystalReport1.Formulas(0) = "Empresa='" & UCase(VGNemp) & "'"
    CrystalReport1.Formulas(1) = "Hora='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "xMoneda='" & cmbMoneda.text & "'"
    CrystalReport1.Formulas(3) = "Art1='" & txtDArt.text & "'"
    CrystalReport1.Formulas(4) = "Art2='" & txtHArt.text & "'"
    CrystalReport1.WindowTitle = vrep & " -- Control de Inventarios"
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Action = 1
 End If
 Screen.MousePointer = vbDefault
Exit Sub
err:
  MsgBox err.Description
  Exit Sub
  Resume
End Sub

Private Sub Form_Load()
    CentForm Me
    txtDFec = VG_FecTrab
    txtHFec = VG_FecTrab
    cmbMoneda.ListIndex = 0
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

Private Sub txtDArt_Change()
    lblDArt = ""
End Sub

Private Sub txtDArt_DblClick()
    Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT acodigo,adescri,aunidad FROM maeart WHERE NOT(AFSTOCK='N' AND AFSERIE='N' AND AFLOTE='N') "
    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Artículos"
    frmReferencia.Inicio
    frmReferencia.Show vbModal
    Rem MVV Adodc2.Close
    Set Adodc2 = Nothing
    
    If vGUtil(1) <> "" Then
        txtDArt = vGUtil(1)
        lblDArt = vGUtil(2)
        txtHArt.SetFocus
    End If
End Sub

Private Sub txtDArt_GotFocus()
    Enfoque txtDArt
End Sub

Private Sub txtDArt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then txtDArt_DblClick
End Sub

Private Sub txtDArt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDArt = Trim(txtDArt)
        If txtDArt <> "" Then
            If Not Existe(1, txtDArt, "maeart", "acodigo", False) Then
                Mensaje = "El Código de Artículo no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                Enfoque txtDArt
            Else
                lblDArt = Devolver_Dato(1, txtDArt, "maeart", "acodigo", False, _
                    "adescri")
            End If
        End If
        txtHArt.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtDev1_DblClick()
Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT TT_CODMOV,TT_DESCRI FROM TABTRANSA WHERE TT_TIPMOV='S'"
    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Transacciones"
    frmReferencia.Inicio
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

Private Sub txtHArt_Change()
    lblHArt = ""
End Sub

Private Sub txtHArt_DblClick()
    Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT acodigo,adescri,aunidad FROM maeart WHERE NOT(AFSTOCK='N' AND AFSERIE='N' AND AFLOTE='N') "
    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Artículos"
    frmReferencia.Inicio
    frmReferencia.Show vbModal
    Rem MVV Adodc2.Close
    Set Adodc2 = Nothing
    
    If vGUtil(1) <> "" Then
        txtHArt = vGUtil(1)
        lblHArt = vGUtil(2)
        cmbMoneda.SetFocus
    End If
End Sub

Private Sub txtHArt_GotFocus()
    Enfoque txtHArt
End Sub

Private Sub txtHArt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then txtHArt_DblClick
End Sub

Private Sub txtHArt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtHArt = Trim(txtHArt)
        If txtHArt <> "" Then
            If Not Existe(1, txtHArt, "maeart", "acodigo", False) Then
                Mensaje = "El Código de Artículo no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                Enfoque txtHArt
            Else
                lblHArt = Devolver_Dato(1, txtHArt, "maeart", "acodigo", False, _
                    "adescri")
            End If
        End If
        cmbMoneda.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
    
    strsql = "SELECT TT_CODMOV,TT_DESCRI FROM TABTRANSA WHERE TT_TIPMOV='I'"
    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Transacciones"
    frmReferencia.Inicio
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
    
    strsql = "SELECT TT_CODMOV,TT_DESCRI FROM TABTRANSA WHERE TT_TIPMOV='I'"
    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Transacciones"
    frmReferencia.Inicio
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
    
    strsql = "SELECT TT_CODMOV,TT_DESCRI FROM TABTRANSA WHERE TT_TIPMOV='I'"
    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Transacciones"
    frmReferencia.Inicio
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
    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Transacciones"
    frmReferencia.Inicio
    frmReferencia.Show vbModal
    Rem mvv Adodc2.Close
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
    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Transacciones"
    frmReferencia.Inicio
    frmReferencia.Show vbModal
    Rem mvv Adodc2.Close
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

Private Function Genera_Reporte() As Boolean
Dim cAnno1 As String, cAnno2 As String
Dim cMes1 As String, cMes2 As String
Dim NI As Integer, cC As String
Dim X As Integer
Genera_Reporte = True

cAnno1 = Format(Year(txtDFec), "0000")
cAnno2 = Format(Year(txtHFec), "0000")

If Val(cAnno1) < Val(cAnno2) Then
    MsgBox "El Año Final debe ser mayor o igual al Año Inicial "
    Genera_Reporte = False
    Exit Function
End If

cMes1 = Format(Month(txtDFec), "00")
cMes2 = Format(Month(txtHFec), "00")

If (cMes1 < cMes2) And (cAnno1 < cAnno2) Then
    MsgBox "El Mes Inicial debe ser menor al Mes Final"
    Genera_Reporte = False
    Exit Function
ElseIf (Val(cMes2) - Val(cMes1)) < 0 Then
    MsgBox "Los Meses deben estar en el rango de los 12 meses"
    Genera_Reporte = False
    Exit Function
End If

If funcExisteTabla(cConexCom, "COMP_MEN") Then cConexCom.Execute "DROP TABLE COMP_MEN"
        strsql = " SELECT ACODIGO AS CMCODIGO, ADESCRI AS CMDESCRI, AUNIDAD AS CMUNI,"
        strsql = strsql & " SPACE(100) AS CMPROVEEDOR ,0.00 AS CMCANT01, 0.00 AS CMCANT02, 0.00 AS CMCANT03,"
        strsql = strsql & " 0.00 AS CMCANT04,0.00 AS CMCANT05,0.00 AS CMCANT06,0.00 AS CMCANT07,"
        strsql = strsql & "0.00 AS CMCANT08,0.00 AS CMCANT09,0.00 AS CMCANT10,0.00 AS CMCANT11,0.00 AS CMCANT12,"
        strsql = strsql & " 0.00 AS CMIMP01,0.00 AS CMIMP02,0.00 AS CMIMP03,0.00 AS CMIMP04,0.00 AS CMIMP05,"
        strsql = strsql & " 0.00 AS CMIMP06,0.00 AS CMIMP07,0.00 AS CMIMP08,0.00 AS CMIMP09,0.00 AS CMIMP10,"
        Rem MVV strsql = strsql & "0.00 AS CMIMP11,0.00 AS CMIMP12  INTO COMP_MEN  in '" & cConexCom.Properties(7) & "' FROM MAEART"
        strsql = strsql & "0.00 AS CMIMP11,0.00 AS CMIMP12  INTO COMP_MEN FROM MAEART"
        strsql = strsql & " WHERE NOT(AFSTOCK='N' AND AFSERIE='N' AND AFLOTE='N') AND "
        strsql = strsql & " ACODIGO>='" & txtDArt.text & "' AND ACODIGO<='" & txtHArt.text & "' ORDER BY ACODIGO "
        cConexCom.Execute strsql
        
  For NI = Val(cMes1) To Val(cMes2)
       cC = Format(NI, "00")
       If funcExisteTabla(cConexCom, "IMPORTES") Then cConexCom.Execute "DROP TABLE IMPORTES"
       strsql = "SELECT DECODIGO, " & _
       " sum((CASE C.CATD WHEN 'NS' THEN DECANTID*-1 ELSE DECANTID END)) AS CANTTOTAL, " & _
       " sum((CASE c.cacodmon WHEN 'ME' THEN (CASE C.CATD WHEN 'NS' THEN decantid*-1*deprecio " & _
       " ELSE decantid*deprecio END) ELSE " & _
       " (CASE WHEN catipcam<>0 THEN (CASE C.CATD WHEN 'NS' THEN decantid*-1*(deprecio/catipcam) " & _
       " ELSE decantid*(deprecio/catipcam) END) ELSE (CASE C.CATD WHEN 'NS' THEN decantid*-1*(deprecio/1) " & _
       " ELSE decantid*(deprecio/1) END) END) END)) AS TOTALME, " & _
       " sum((CASE c.cacodmon WHEN 'MN' THEN (CASE C.CATD WHEN 'NS' THEN decantid*-1*deprecio " & _
       " ELSE decantid*deprecio END) ELSE " & _
       " (CASE WHEN catipcam<>0 THEN (CASE C.CATD WHEN 'NS' THEN decantid*-1*(deprecio * catipcam) " & _
       " ELSE decantid*(deprecio*catipcam) END) ELSE (CASE C.CATD WHEN 'NS' THEN decantid*-1*(deprecio *1) " & _
       " ELSE decantid*(deprecio*1) END) END) END)) AS TOTALMN"
       
       strsql = strsql & " INTO IMPORTES "
       strsql = strsql & " FROM movalmdet AS d INNER JOIN movalmcab AS c"
       strsql = strsql & " ON (c.catd=d.detd) AND (c.caalma=d.dealma) AND (c.canumdoc=d.denumdoc)"
       strsql = strsql & " Where Month(C.CaFecDoc) =" & NI & " and YEAR(C.CAFECDOC)='" & Trim(cAnno1) & "' AND"
       strsql = strsql & " (C.CATD= 'NI' OR C.CATD='NS') AND "
       strsql = strsql & " (C.CACODMOV='" & TxtMov1.text & "' OR C.CACODMOV='" & TxtMov2.text & "' OR C.CACODMOV='" & txtMov3.text & "' "
       strsql = strsql & " OR C.CACODMOV='" & txtDev1.text & "' OR C.CACODMOV='" & txtDev2.text & "' OR  C.CACODMOV='" & txtDev3.text & "' )"
       strsql = strsql & " AND D.DECODIGO>='" & txtDArt.text & "' AND D.DECODIGO<='" & txtHArt.text & "'  "
       strsql = strsql & " GROUP BY decodigo"
       
'       strsql = "SELECT DECODIGO, sum(iif(C.CATD='NS',DECANTID*-1,DECANTID)) AS CANTTOTAL,"
'       strsql = strsql & " sum(iif(c.cacodmon='MN', IIF(C.CATD='NS',decantid*deprecio*-1,decantid*deprecio),"
'       strsql = strsql & " IIF(C.CATD='NS', decantid*deprecio*catipcam*-1,decantid*deprecio*catipcam))) AS TOTALMN,"
'       strsql = strsql & "sum(iif(c.cacodmon='ME', IIF(C.CATD='NS',decantid*-1*deprecio,decantid*deprecio),"
'       strsql = strsql & " iif(catipcam<>0,IIF(C.CATD='NS',decantid*-1*(deprecio/catipcam),decantid*(deprecio/catipcam)),"
'       strsql = strsql & " IIF(C.CATD='NS',decantid*-1*(deprecio/1),decantid*(deprecio/1))))) AS TOTALME "
'       strsql = strsql & " INTO IMPORTES in '" & cConexCom.Properties(7) & "' "
'       strsql = strsql & " FROM movalmdet AS d INNER JOIN movalmcab AS c"
'       strsql = strsql & " ON (c.catd=d.detd) AND (c.caalma=d.dealma) AND (c.canumdoc=d.denumdoc)"
'       strsql = strsql & " Where Month(C.CaFecDoc) =" & NI & " and YEAR(C.CAFECDOC)='" & Trim(cAnno1) & "' AND"
'       strsql = strsql & " (C.CATD= 'NI' OR C.CATD='NS') AND "
'       strsql = strsql & " (C.CACODMOV='" & TxtMov1.text & "' OR C.CACODMOV='" & TxtMov2.text & "' OR C.CACODMOV='" & txtMov3.text & "' "
'       strsql = strsql & " OR C.CACODMOV='" & txtDev1.text & "' OR C.CACODMOV='" & txtDev2.text & "' OR  C.CACODMOV='" & txtDev3.text & "' )"
'       strsql = strsql & " AND D.DECODIGO>='" & txtDArt.text & "' AND D.DECODIGO<='" & txtHArt.text & "'  "
'       strsql = strsql & " GROUP BY decodigo"
       cConexCom.Execute strsql
     
       Dim t As Variant
       t = Timer
       While Timer < (t + 5)
        DoEvents
       Wend
       strsql = "UPDATE  COMP_MEN SET COMP_MEN.CMCANT" & cC & " =CAST(IMPORTES.CANTTOTAL AS float), " & _
            " COMP_MEN.CMIMP" & cC & " = CAST(" & IIf(cmbMoneda.text = "MN", "IMPORTES.TOTALMN", "IMPORTES.TOTALME") & _
            " AS float) FROM IMPORTES WHERE COMP_MEN.CMCODIGO=IMPORTES.DECODIGO"
       cConexCom.Execute strsql
   Next
 End Function


VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmGenerasaldosini 
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCuenta 
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   635
      XcodMaxLongitud =   0
      xcodwith        =   800
      NomTabla        =   "ct_cuenta"
      ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
      XcodCampo       =   "cuentacodigo"
      XListCampo      =   "cuentadescripcion"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "cuentacodigo,cuentadescripcion"
      Requerido       =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Cuenta de Resultados para el actual ejercicio"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "FrmGenerasaldosini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Not ValidaIngreso Then Exit Sub
    Call GenSaldos
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Ctr_AyudaCuenta.conexion(VGCNx): Ctr_AyudaCuenta.Filtro = "empresacodigo ='" & VGParametros.empresacodigo & "'"
End Sub
Private Function ValidaIngreso() As Boolean
ValidaIngreso = False
    If Ctr_AyudaCuenta.xclave = "" Then
        MsgBox "Debe Ingresar un Cuenta de resultados de ejercicio", vbExclamation
        Ctr_AyudaCuenta.SetFocus
        Exit Function
    End If
ValidaIngreso = True
End Function
Private Sub GenSaldos()
  On Error GoTo SaldosIniciales
    Screen.MousePointer = 11
    VGCNx.BeginTrans
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_GenerarSaldoIniciales_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@annoact") = VGParamSistem.Anoproceso
        .Parameters("@annoant") = VGParamSistem.Anoproceso - 1
        .Parameters("@ultnivel") = VG_aNIVELES(VGnumnivelescuenta - 1)
        .Parameters("@CuentaResEjer") = Ctr_AyudaCuenta.xclave
        .Execute
    End With
    VGCNx.CommitTrans
    
    'Recalcular Saldos
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_recalacum_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
          .Parameters("@base") = VGParamSistem.BDEmpresa
          .Parameters("@empresa") = VGParametros.empresacodigo
          .Parameters("@anno") = VGParamSistem.Anoproceso
          .Parameters("@mespro") = "01"
          .Parameters("@user") = VGParamSistem.Usuario
    End With
    VGCommandoSP.Execute
    
    Screen.MousePointer = 1
    MsgBox "Se Generaron Saldos Iniciales del año " & VGParamSistem.Anoproceso & " Satisfactoriamente ", vbInformation
    Unload Me
    Exit Sub
SaldosIniciales:
    Screen.MousePointer = 1
    VGCNx.RollbackTrans
    MsgBox "ERROR : No se actualizarón los Saldos Iniciales del año " & VGParamSistem.Anoproceso & " " & Chr(13) & err.Description, vbExclamation
End Sub




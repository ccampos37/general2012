VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmAjusteDiferxPend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuste de Difer x Cambio fin Mes"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   3660
   Begin VB.Frame Frame1 
      Height          =   2025
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   3495
      Begin VB.ComboBox CmbTcambio 
         Height          =   315
         ItemData        =   "FrmAjusDiferxPend.frx":0000
         Left            =   1095
         List            =   "FrmAjusDiferxPend.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1515
         Width           =   2250
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_SubAsiento 
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   1065
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
         XcodMaxLongitud =   4
         xcodwith        =   450
         NomTabla        =   "ct_subasiento"
         TituloAyuda     =   "Busqueda de  SubAsiento"
         ListaCampos     =   "subasientocodigo(1),subasientodescripcion(1),monedacodigo(1),subasientoglosa(1),subasientorepitedoc(2)"
         XcodCampo       =   "subasientocodigo"
         XListCampo      =   "subasientodescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion,Moneda"
         ListaCamposText =   "subasientocodigo,subasientodescripcion,monedacodigo,subasientoglosa,subasientorepitedoc"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Asiento 
         Height          =   300
         Left            =   165
         TabIndex        =   0
         Top             =   450
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   529
         XcodMaxLongitud =   3
         xcodwith        =   150
         NomTabla        =   "ct_asiento"
         TituloAyuda     =   "Busqueda de Asiento"
         ListaCampos     =   "asientocodigo(1), asientodescripcion(1),flaggrabado(2),controlnref(2),nemotecref(1)"
         XcodCampo       =   "asientocodigo"
         XListCampo      =   "asientodescripcion"
         ListaCamposDescrip=   "Codigo,Descripción,OperGraba"
         ListaCamposText =   "asientocodigo,asientodescripcion,flaggrabado,controlnref,nemotecref"
         Requerido       =   0   'False
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "T/Cambio :"
         Height          =   195
         Left            =   165
         TabIndex        =   8
         Top             =   1575
         Width           =   795
      End
      Begin VB.Label lbSubAsiento 
         BackStyle       =   0  'Transparent
         Caption         =   "Subasiento :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   855
         Width           =   1590
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Asiento :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   225
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1830
      TabIndex        =   4
      Top             =   2160
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   705
      TabIndex        =   3
      Top             =   2160
      Width           =   1125
   End
End
Attribute VB_Name = "FrmAjusteDiferxPend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Not ValidaIngreso Then Exit Sub
    Call GenAjuste
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
    CtrAyu_Asiento.SetFocus
End Sub

Private Sub Form_Load()
     Width = 3750
     Height = 2985
     Call CtrAyu_Asiento.conexion(VGCNx): CtrAyu_Asiento.Filtro = "asientocodigo <>'00'"
     Call CtrAyu_SubAsiento.conexion(VGCNx): CtrAyu_SubAsiento.Filtro = "subasientocodigo <>'00'"
     CmbTcambio.ListIndex = 1
End Sub
Private Sub CtrAyu_Asiento_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    CtrAyu_SubAsiento.Filtro = "asientocodigo='" & Trim$(CtrAyu_Asiento.xclave) & "'"
    CtrAyu_SubAsiento.xclave = "": CtrAyu_SubAsiento.xnombre = ""
End Sub
Private Function ValidaIngreso() As Boolean
ValidaIngreso = False
    If CtrAyu_Asiento.xclave = "" Then
        MsgBox "Debe Ingresar un Codigo de Asiento", vbExclamation
        CtrAyu_Asiento.SetFocus
        Exit Function
    End If
    If CtrAyu_SubAsiento.xclave = "" Then
        MsgBox "Debe Ingresar un Codigo de Sub - Asiento", vbExclamation
        CtrAyu_SubAsiento.SetFocus
        Exit Function
    End If
ValidaIngreso = True
End Function
Private Sub GenAjuste()
On Error GoTo Ajuste
    VGGeneral.BeginTrans
    Screen.MousePointer = 11
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_AjusPend_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@Servidor") = VGParamSistem.Servidor
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@Ano") = VGParamSistem.Anoproceso
        .Parameters("@mes") = Format(CInt(VGParamSistem.Mesproceso), "00")
        .Parameters("@FormCamb") = Left(CmbTcambio.Text, 2)
        .Parameters("@asiento") = Trim$(CtrAyu_Asiento.xclave)
        .Parameters("@subasiento") = Trim$(CtrAyu_SubAsiento.xclave)
        .Parameters("@User") = VGParamSistem.Usuario
        Set rs = .Execute
    End With
    VGGeneral.CommitTrans
    If rs.State = 0 Then
        MsgBox "El Ajuste de Diferencia de Cambio de Documentos " & Chr(13) & _
               "Pendientes se Realizo Satisfactoriamente ", vbInformation
      Else
         MsgBox "El Ajuste de Diferencia de Cambio de Documentos " & Chr(13) & _
               "Pendientes No se Realizará porque no Encontro " & Chr(13) & _
               "Ni un Documento a Ajustar ", vbExclamation
        
    End If
    Screen.MousePointer = 1
    Exit Sub
Ajuste:
    VGGeneral.RollbackTrans
    Screen.MousePointer = 1
    MsgBox "No se genero el Ajuste " & Chr(13) & _
           err.Description, vbExclamation
End Sub

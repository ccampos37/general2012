VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmAjusDiferxDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuste de Difer x Cambio por documento"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4230
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1770
      TabIndex        =   0
      Top             =   270
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   39845889
      CurrentDate     =   39833
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2010
      TabIndex        =   3
      Top             =   1725
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   885
      TabIndex        =   2
      Top             =   1725
      Width           =   1125
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCCosto 
      Height          =   300
      Left            =   420
      TabIndex        =   1
      Top             =   1110
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   529
      XcodMaxLongitud =   10
      xcodwith        =   400
      NomTabla        =   "ct_centrocosto"
      TituloAyuda     =   "Busqueda Centro Costo"
      ListaCampos     =   "centrocostocodigo(1),centrocostodescripcion(2)"
      XcodCampo       =   "centrocostocodigo"
      XListCampo      =   "centrocostodescripcion"
      ListaCamposDescrip=   "centrocostocodigo,centrocostodescripcion"
      ListaCamposText =   "centrocostocodigo,centrocostodescripcion"
   End
   Begin VB.Label Label2 
      Caption         =   "Centro Costo - ajuste por pérdida:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   420
      TabIndex        =   5
      Top             =   840
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha ajuste:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   450
      TabIndex        =   4
      Top             =   300
      Width           =   1245
   End
End
Attribute VB_Name = "FrmAjusDiferxDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Not ValidaIngreso Then Exit Sub
    Call GenAjuste
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Ctr_AyudaCCosto.conexion(VGCNx): Ctr_AyudaCCosto.Filtro = "empresacodigo ='" & VGParametros.empresacodigo & "'"
    DTPicker1.Value = Fecha(2, VGParamSistem.FechaTrabajo)
'    DTPicker1.SetFocus
End Sub
Private Function ValidaIngreso() As Boolean
Dim tipocambio As Integer
tipocambio = 0
ValidaIngreso = False
    If Ctr_AyudaCCosto.xclave = "" Then
        MsgBox "Debe Ingresar un Centro de Costo", vbExclamation
        Ctr_AyudaCCosto.SetFocus
        Exit Function
    End If
tipocambio = XRecuperaTipoCambio(DTPicker1, Compra, VGcnxCT)
If tipocambio = 0 Then
   MsgBox "Fecha de Ajuste no tiene tipo de cambio", vbExclamation
        DTPicker1.SetFocus
        Exit Function
    End If
ValidaIngreso = True
End Function
Private Sub GenAjuste()
On Error GoTo Ajuste
    VGGeneral.BeginTrans
    Screen.MousePointer = 11
    
    EliminaAjustes
'  Mayoriza
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "Ct_AjusDifCambxDoc_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@Ano") = Year(DTPicker1.Value)
        .Parameters("@Mes") = Month(DTPicker1.Value)
        .Parameters("@Asiento") = "055"
        .Parameters("@SubAsiento") = "0099"
        .Parameters("@AjusteDebe") = Trim(VGParametros.sistemactaajustedeb)
        .Parameters("@AjusteHaber") = Trim(VGParametros.sistemactaajustehab)
        .Parameters("@Fecha") = DTPicker1.Value
        .Parameters("@Usuario") = VGParamSistem.Usuario
        .Parameters("@NombrePC") = VGcomputer
        .Parameters("@TipoCambio1") = XRecuperaTipoCambio(DTPicker1.Value, 1, VGcnxCT)
        .Parameters("@TipoCambio2") = XRecuperaTipoCambio(DTPicker1.Value, 2, VGcnxCT)
        .Parameters("@CCosto") = Ctr_AyudaCCosto.xclave
        Set rs = .Execute
    End With
    
    Mayoriza
    
    VGGeneral.CommitTrans
    If rs.State = 0 Then
        MsgBox "El Ajuste de Diferencia de Cambio de Documentos " & Chr(13) & _
               "Cancelados se Realizo Satisfactoriamente ", vbInformation
      Else
         MsgBox "El Ajuste de Diferencia de Cambio de Documentos " & Chr(13) & _
               "Cancelados No se Realizará porque no Encontro " & Chr(13) & _
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
Private Sub EliminaAjustes()
    Dim rs1 As ADODB.Recordset
    Dim SQL As String
    Set rs1 = New ADODB.Recordset
    
    'Eliminando asientos de ajuste
    'Ajuste Ganancia
    SQL = "Delete From ct_cabcomprob" & Year(DTPicker1.Value) & " Where empresacodigo='" & VGParametros.empresacodigo & "' " & _
        "And cabcomprobmes=" & Month(DTPicker1.Value) & _
        " AND cabcomprobnumero='" & Format(Month(DTPicker1.Value), "00") & "05500001' " & _
        " AND subasientocodigo='0099' AND asientocodigo='055'"
    Set rs1 = VGCNx.Execute(SQL)
    Set rs1 = Nothing
    'Ajuste Perdida
    SQL = "Delete From ct_cabcomprob" & Year(DTPicker1.Value) & " Where empresacodigo='" & VGParametros.empresacodigo & "' " & _
        "And cabcomprobmes=" & Month(DTPicker1.Value) & _
        " AND cabcomprobnumero='" & Format(Month(DTPicker1.Value), "00") & "05500002' " & _
        " AND subasientocodigo='0099' AND asientocodigo='055'"
    Set rs1 = VGCNx.Execute(SQL)
    Set rs1 = Nothing
End Sub
Private Sub Mayoriza()
Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_mayoriza_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@anno") = Year(DTPicker1.Value)
        .Parameters("@mespro") = Month(DTPicker1.Value)
        .Parameters("@user") = VGParamSistem.Usuario
        .Execute
    End With
End Sub


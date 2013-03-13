VERSION 5.00
Begin VB.Form FrmRepCOCompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Compras"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5670
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      Begin VB.ComboBox CmbTipo 
         Height          =   315
         ItemData        =   "FrmRepCOCompras.frx":0000
         Left            =   2370
         List            =   "FrmRepCOCompras.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   255
         Width           =   2700
      End
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmRepCOCompras.frx":0053
         Left            =   1245
         List            =   "FrmRepCOCompras.frx":005D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   645
         Width           =   3825
      End
      Begin VB.ComboBox Cmbempresa 
         Height          =   315
         ItemData        =   "FrmRepCOCompras.frx":00A3
         Left            =   1230
         List            =   "FrmRepCOCompras.frx":00AA
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   3825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Registro de Compras :"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   315
         Width           =   2145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenar por :"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Empresar :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1155
         Width           =   750
      End
   End
   Begin VB.CommandButton axBCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   2700
      TabIndex        =   1
      Top             =   1860
      Width           =   1275
   End
   Begin VB.CommandButton axbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1350
      TabIndex        =   0
      Top             =   1860
      Width           =   1275
   End
End
Attribute VB_Name = "FrmRepCOCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSparCompras As ADODB.Recordset

Private Sub axBAceptar_Click()
    Call imprimir
End Sub

Private Sub axBCancelar_Click()
    Unload Me
End Sub

Private Sub CmbTipo_Click()
    If CmbTipo.ListIndex = 2 Then
        Label2.Enabled = False
        CmbOrden.Enabled = False
      Else
        Label2.Enabled = True
        CmbOrden.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Call cargaempresa(1)
    CmbTipo.ListIndex = 0
    CmbOrden.ListIndex = 0
End Sub
Public Sub imprimir()
Dim arrform(2) As Variant, arrparm(6) As Variant
Dim NombreRep As String, Cadorden As String
Dim mon As String
 arrparm(0) = VGParamSistem.BDEmpresa
 arrparm(1) = VGParamSistem.BDEmpresaCT
 arrparm(2) = VGParamSistem.Servidor
 arrparm(3) = "%%"
 If VGParametros.sistemamultiempresas Then
    If Left(Cmbempresa.Text, 2) = "00" Then
       arrparm(3) = "%%"
     Else
       arrparm(3) = Left(Cmbempresa.Text, 2)
    End If
End If
 arrparm(4) = VGParamSistem.Anoproceso
 arrparm(5) = RTrim(VGParamSistem.Mesproceso)
 arrform(0) = "Xmes=" & CInt(Trim(VGParamSistem.Mesproceso))
 arrform(1) = "Xano=" & CInt(Trim(VGParamSistem.Anoproceso))
 Select Case CmbTipo.ListIndex
        Case 0
            NombreRep = "rptCORegistroCompras.rpt"
        Case 1
            NombreRep = "rptCORegistroComprasTD.rpt"
        Case 2
            NombreRep = "rptCORegistroComprasPRO.rpt"
    End Select
    Cadorden = ""
   Call ImpresionRptProc(NombreRep, arrform, arrparm, Cadorden, "Registro de Compras ")
    Exit Sub
ImprimeRegCompras:
MsgBox err.Description
End Sub
Public Sub cargaempresa(op As Integer)
Dim REG1 As New ADODB.Recordset
Set REG1 = New ADODB.Recordset
REG1.Open "Select * from co_multiempresas order by empresacodigo ", VGCNx, adOpenStatic
 If REG1.EOF Then Exit Sub
 If REG1.BOF Then Exit Sub
 If VGParametros.sistemamultiempresas = True Then
 Do While Not REG1.EOF
    If Not IsNull(REG1.Fields("empresadescripcion")) And REG1.Fields("empresacodigo") <> "00" Then
        Cmbempresa.AddItem REG1.Fields("empresacodigo") + "-" + REG1.Fields("empresadescripcion")
    End If
      REG1.MoveNext
      If REG1.EOF Then Exit Do
 Loop
    REG1.MoveFirst
 End If
 Cmbempresa.ListIndex = 0
End Sub

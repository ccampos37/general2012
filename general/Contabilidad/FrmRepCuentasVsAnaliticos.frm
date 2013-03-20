VERSION 5.00
Begin VB.Form FrmRepCuentasVsAnaliticos 
   Caption         =   "Ctas. Contables. vs Ctas. Analisis"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   4500
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbNivel 
      Height          =   315
      Left            =   1650
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   390
      Width           =   1920
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   900
      TabIndex        =   1
      Top             =   1050
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   2250
      TabIndex        =   0
      Top             =   1050
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nivel :"
      Height          =   195
      Index           =   1
      Left            =   1020
      TabIndex        =   3
      Top             =   480
      Width           =   450
   End
End
Attribute VB_Name = "FrmRepCuentasVsAnaliticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private Sub cmdAceptar_Click()
    Screen.MousePointer = 11
    Call imprimir
    Screen.MousePointer = 1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CargaNivel
    cmbNivel.Text = "1"
End Sub
Private Sub CargaNivel()
    Dim i As Integer
    For i = 1 To VGnumnivelescuenta
        cmbNivel.AddItem Format(i, "0")
    Next
    cmbNivel.ListIndex = 0
End Sub

Public Sub imprimir()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(2) As Variant, arrparm(7) As Variant
Dim mon As String

     '@Base,@empresa, @Anno, @computer, @Nivel, @tipo
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = VGParamSistem.Mesproceso
    arrparm(4) = VGcomputer
    arrparm(5) = LongitudNivel(CInt(cmbNivel.Text))
    arrparm(6) = "1"
    
    Call ImpresionRptProc("ct_CuentasVsAnaliticos.rpt", arrform, arrparm)
End Sub
Private Function LongitudNivel(ByVal nivel As Integer) As Integer
Dim rs As ADODB.Recordset
    LongitudNivel = 2
    Set rs = New ADODB.Recordset
    Set rs = VGCNx.Execute("SELECT sistemaconfiguracuenta FROM ct_sistema")
    If Not rs.EOF Then
        LongitudNivel = Mid$(rs!sistemaconfiguracuenta, nivel * 2 - 1, 1)
        '2*4*8
    End If
End Function


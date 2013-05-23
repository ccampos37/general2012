VERSION 5.00
Begin VB.Form FrmLibroRegistrodeCompras 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   1155
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   5175
      Begin VB.ComboBox CmbTipo 
         Height          =   315
         ItemData        =   "FrmLibroregistrodeCompras.frx":0000
         Left            =   2370
         List            =   "FrmLibroregistrodeCompras.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   255
         Width           =   2700
      End
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmLibroregistrodeCompras.frx":001F
         Left            =   1245
         List            =   "FrmLibroregistrodeCompras.frx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   645
         Width           =   3825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Tipo de Registro de Compras :"
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   315
         Width           =   2145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ordenar por :"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   720
         Width           =   930
      End
   End
   Begin VB.CommandButton axbAceptar 
      Caption         =   "&Aceptar"
      Height          =   480
      Left            =   1785
      TabIndex        =   1
      Top             =   1620
      Width           =   1275
   End
   Begin VB.CommandButton axBCancelar 
      Caption         =   "&Cancelar"
      Height          =   480
      Left            =   3135
      TabIndex        =   0
      Top             =   1620
      Width           =   1275
   End
End
Attribute VB_Name = "FrmLibroRegistrodeCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
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
    Set RSparCompras = New ADODB.Recordset
    CmbTipo.ListIndex = 0
    CmbOrden.ListIndex = 0
End Sub

Public Sub imprimir()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(1) As Variant
Dim arrparm(9) As Variant
Dim NombreRep As String, CadOrden As String
Dim mon As String
     '@BASE, @ANNO, @MES, @ASIENTOSPLAN, @CTASPLANCOMP, @CTASIGV
    Set RSparCompras = New ADODB.Recordset
    SQL = "select * from ct_paramlibaux where paramlibauxtipo='CO'"
    Set RSparCompras = VGCNx.Execute(SQL)
    If RSparCompras.RecordCount = 0 Then
        MsgBox "No existen parametros para el registros de compras"
        Exit Sub
    End If
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = Trim$(VGParamSistem.Mesproceso)
    arrparm(4) = RSparCompras!paramlibauxasiento
    arrparm(5) = RSparCompras!paramlibauxcuenta
    arrparm(6) = RSparCompras!paramlibauxigv
    arrparm(7) = RSparCompras!paramlibauxies
    arrparm(8) = RSparCompras!paramlibauxirenta
    
    arrform(0) = "periodo='" & DesMes(Trim$(VGParamSistem.Mesproceso)) & "'"
    NombreRep = "ct_LibroRegistroComprasSunat.rpt"
    CadOrden = ""
    CadOrden = "+{ct_registrocompras_rpt.numauxiliar},"
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Registro de Compras ")
End Sub

Private Function ArmaCriterio(cad As String, car As String, Optional campocrit As String) As String
Dim pos As Integer, cadaux As String, criterio As String
Dim valor As String
    criterio = ""
    Do While True
        pos = InStr(1, cad, car, vbTextCompare)
        If pos = 0 Then Exit Do
        If campocrit = "" Or Trim$(car) = "," Then
            valor = "''" & Left(cad, pos - 1) & "''"
          Else
            valor = "''" & Left(cad, pos) & "''"
        End If
        cad = Right(cad, (Len(cad) - pos))
        If campocrit <> "" Then
            criterio = criterio & campocrit & " like " & valor & " or "
          Else
            criterio = criterio & valor & car
        End If
    Loop
    If campocrit <> "" Then
        ArmaCriterio = Left(criterio, Len(criterio) - 3)
      Else
        ArmaCriterio = Left(criterio, Len(criterio) - 1)
    End If
End Function


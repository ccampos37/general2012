VERSION 5.00
Begin VB.Form frmRepVentas 
   Caption         =   "Registro de Ventas"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   1905
   ScaleWidth      =   5520
   Begin VB.CommandButton axBCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   2655
      TabIndex        =   6
      Top             =   1245
      Width           =   1215
   End
   Begin VB.CommandButton axbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1350
      TabIndex        =   5
      Top             =   1245
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.ComboBox CmbTipo 
         Height          =   315
         ItemData        =   "FrmRepVentas.frx":0000
         Left            =   2370
         List            =   "FrmRepVentas.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   255
         Width           =   2700
      End
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmRepVentas.frx":0051
         Left            =   1245
         List            =   "FrmRepVentas.frx":005E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   630
         Width           =   3825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Registro de Compras :"
         Height          =   195
         Left            =   105
         TabIndex        =   4
         Top             =   315
         Width           =   2145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenar por :"
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   720
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmRepVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSparVentas As ADODB.Recordset

Private Sub axBAceptar_Click()
    Call imprimir
End Sub

Private Sub axBCancelar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Height = 2070: Width = 5355
    Set RSparVentas = New ADODB.Recordset
    CmbTipo.ListIndex = 0
    CmbOrden.ListIndex = 0
End Sub

Public Sub imprimir()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(1) As Variant, arrparm(10) As Variant
Dim NombreRep As String, CadOrden As String
Dim mon As String
     '@BASE, @ANNO, @MES, @ASIENTOSPLAN, @CTASPLANCOMP, @CTASIGV
    Set RSparVentas = New ADODB.Recordset
    RSparVentas.Open "select * from ct_paramlibaux where paramlibauxtipo='VT'", VGCNx, adOpenKeyset, adLockReadOnly
    If RSparVentas.RecordCount = 0 Then
        MsgBox "No existen parametros para el registros de Ventas"
        Exit Sub
    End If
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = Trim$(VGParamSistem.Mesproceso)
    arrparm(4) = RSparVentas!paramlibauxasiento
    arrparm(5) = RSparVentas!paramlibauxcuenta
    arrparm(6) = RSparVentas!paramlibauxigv
    arrparm(7) = RSparVentas!paramlibauxies
    arrparm(8) = RSparVentas!paramlibauxirenta
    arrparm(9) = "74%"
    Select Case CmbTipo.ListIndex
        Case 0
            NombreRep = "rptRegistroVentas.Rpt"
        Case 1
            NombreRep = "rptRegistroVentas.Rpt"
           'NombreRep = "rptRegistroVentasprov.Rpt"
        Case 2
            NombreRep = "rptRegistroVentasLibros.rpt"
    End Select
    CadOrden = ""
    
    arrform(0) = "periodo='" & VGvardllgen.DesMes(Trim$(VGParamSistem.Mesproceso)) & "'"
    
    If CmbTipo.ListIndex < 2 Then
       Select Case CmbOrden.ListIndex
           Case 0
               CadOrden = "+{ct_registroventas_rpt.detcomprobnumdocumento},"
           Case 1
               CadOrden = "+{ct_registroVentas_rpt.cabcomprobnumero},"
           Case 2
               CadOrden = "+{ct_registroVentas_rpt.detcomprobfechaemision},"
       End Select
    End If
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Registro de Ventas")
End Sub

Private Function ArmaCriterio(cad As String, car As String, Optional campocrit As String) As String
Dim pos As Integer, cadaux As String, criterio As String
Dim Valor As String
    criterio = ""
    Do While True
        pos = InStr(1, cad, car, vbTextCompare)
        If pos = 0 Then Exit Do
        If campocrit = "" Or Trim$(car) = "," Then
            Valor = "''" & Left(cad, pos - 1) & "''"
          Else
            Valor = "''" & Left(cad, pos) & "''"
        End If
        cad = Right(cad, (Len(cad) - pos))
        If campocrit <> "" Then
            criterio = criterio & campocrit & " like " & Valor & " or "
          Else
            criterio = criterio & Valor & car
        End If
    Loop
    If campocrit <> "" Then
        ArmaCriterio = Left(criterio, Len(criterio) - 3)
      Else
        ArmaCriterio = Left(criterio, Len(criterio) - 1)
    End If
End Function

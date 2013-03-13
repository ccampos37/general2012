VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmRepHonorarios 
   Caption         =   "Reporte de Honoraios"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2325
   ScaleWidth      =   6180
   Begin VB.CommandButton axbCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3142
      TabIndex        =   4
      Top             =   1725
      Width           =   1275
   End
   Begin VB.CommandButton axbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1717
      TabIndex        =   3
      Top             =   1725
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "frmRepHonorarios.frx":0000
         Left            =   1155
         List            =   "frmRepHonorarios.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   150
         Width           =   3825
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   330
         Left            =   1140
         TabIndex        =   6
         Top             =   675
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   582
         XcodMaxLongitud =   0
         xcodwith        =   1300
         NomTabla        =   "v_analiticoentidad"
         ListaCampos     =   "analiticocodigo(1),entidadrazonsocial(1)"
         XcodCampo       =   "analiticocodigo"
         XListCampo      =   "entidadrazonsocial"
         ListaCamposDescrip=   "Código,Razon_Social"
         ListaCamposText =   "analiticocodigo,entidadrazonsocial"
         Requerido       =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ordenar por :"
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   240
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmRepHonorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSparHonorario As ADODB.Recordset

Private Sub axBAceptar_Click()
  If ValidaParametros = True Then
    Call imprimir
  End If
End Sub

Private Sub axBCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Height = 2730: Width = 6300
    Set RSparHonorario = New ADODB.Recordset
    CmbOrden.ListIndex = 0
    Ctr_Ayuda2.conexion VGCNx
End Sub

Public Sub imprimir()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(2) As Variant, arrparm(9) As Variant
Dim NombreRep As String, CadOrden As String
Dim mon As String
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = Trim$(VGParamSistem.Mesproceso)
    arrparm(4) = RSparHonorario!paramlibauxasiento
    arrparm(5) = RSparHonorario!paramlibauxcuenta
    arrparm(6) = RSparHonorario!paramlibauxies
    arrparm(7) = RSparHonorario!paramlibauxirenta
    arrparm(8) = IIf(Ctr_Ayuda2.xclave = Empty, "%%", Trim$(Ctr_Ayuda2.xclave))
    arrform(0) = "@TituloReporte='" & "Reporte de Honorarios" & "'"
    arrform(1) = "@mes='MES DE " & VGvardllgen.DesMes(VGParamSistem.Mesproceso) & " - " & VGParamSistem.Anoproceso & "'"
    NombreRep = "rptHonorarios.rpt"
    CadOrden = Empty
    Select Case CmbOrden.ListIndex
        Case 0
            CadOrden = "+{ct_registroHonorarios_rpt.cabcomprobnumero},"
        Case 1
            CadOrden = "+{ct_registroHonorarios_rpt.detcomprobfechaemision},"
        Case 2
            CadOrden = "+{ct_registroHonorarios_rpt.detcomprobnumdocumento},"
        Case 3
            CadOrden = "+{ct_registroHonorarios_rpt.entidadrazonsocial},+{ct_registroHonorarios_rpt.detcomprobnumdocumento},"
    End Select
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Reporte de Honorarios")
End Sub
Public Sub CapturaParm(ByRef Asientos As String, CtasComp As String, CtasIGV As String)
'SET @ASIENTOSPLAN=' (''064'') '
'SET @CTASPLANCOMP=' (A.cuentacodigo like ''94%'' or A.cuentacodigo like ''33%'') '
'SET @CTASIGV1=' (A.cuentacodigo like ''401174'') '*/
'SET @CTASIGV2=' (A.cuentacodigo like ''403140'') '*/
'paramlibauxasiento --> Asientos de Honorarios
'paramlibauxcuenta  --> Cuentas de Honorarios
'paramlibauxigv --> cuentas de Impuesto Renta e IES
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

Function ValidaParametros() As Boolean
    Set VGvardllgen = New dll_general
    Set RSparHonorario = New ADODB.Recordset
    RSparHonorario.Open "select * from ct_paramlibaux where paramlibauxtipo='HO'", VGCNx, adOpenKeyset, adLockReadOnly
    If RSparHonorario.RecordCount = 0 Then
        MsgBox "No existen parametros para el Reporte de Honorarios", vbInformation, Caption
        ValidaParametros = False
        Exit Function
    End If
    
    If RSparHonorario!paramlibauxirenta = Empty Or RSparHonorario!paramlibauxies = Empty Then
        MsgBox "Faltan los Parámetros de Impuesto a la Renta ó IES", vbInformation, Caption
        ValidaParametros = False
        Exit Function
    End If
    
    If VGvardllgen.ESNULO(RSparHonorario!paramlibauxirenta, 0) = 0 Or VGvardllgen.ESNULO(RSparHonorario!paramlibauxies, 0) = 0 Then
        MsgBox "Faltan los Parámetros de Impuesto a la Renta ó IES", vbInformation, Caption
        ValidaParametros = False
        Exit Function
    End If

    ValidaParametros = True

End Function

VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmRepCajaBancos 
   Caption         =   "Reporte de Caja y Bancos"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2370
   ScaleWidth      =   6270
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   15
      TabIndex        =   2
      Top             =   0
      Width           =   6225
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   360
         Left            =   855
         TabIndex        =   4
         Top             =   465
         Width           =   5280
         _ExtentX        =   9313
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
         Caption         =   "Cuenta:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   585
         Width           =   1095
      End
   End
   Begin VB.CommandButton axbAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1717
      TabIndex        =   1
      Top             =   1725
      Width           =   1275
   End
   Begin VB.CommandButton axbCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3142
      TabIndex        =   0
      Top             =   1725
      Width           =   1275
   End
End
Attribute VB_Name = "frmRepCajaBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim RSparVentas As ADODB.Recordset

Private Sub axBAceptar_Click()
    Call imprimir
End Sub

Private Sub axBCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Height = 2625: Width = 6255
    Set RSparVentas = New ADODB.Recordset
    Ctr_Ayuda1.conexion VGCNx
    Ctr_Ayuda1.Filtro = "cuentacodigo like '10%' AND empresacodigo ='" & VGParametros.empresacodigo & "' "
End Sub

Public Sub imprimir()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(3) As Variant, arrparm(7) As Variant
Dim NombreRep As String, CadOrden As String
Dim mon As String
    Set VGvardllgen = New dll_general
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = Format(VGParamSistem.Mesproceso - 1, "0#")
    arrparm(4) = Format(VGParamSistem.Mesproceso, "0#")
    arrparm(5) = IIf(Ctr_Ayuda1.xclave = Empty, "10%", Trim$(Ctr_Ayuda1.xclave) & "%%")
    arrparm(6) = "%%"
    
    Set VGvardllgen = New dllgeneral.dll_general
    arrform(0) = "@TituloReporte='" & "Libro Caja Bancos " & "'"   'Ctr_Ayuda1.xclave
    arrform(1) = "@Mes='" & VGvardllgen.DesMes(VGParamSistem.Mesproceso) & "'"
    arrform(2) = "@anno='" & VGParamSistem.Anoproceso & "'"
    
    Call ImpresionRptProc("ct_LibroCajaBancos.rpt", arrform, arrparm)

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


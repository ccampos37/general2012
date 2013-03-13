VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmImpCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Cajero"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker DtHasta 
      Height          =   285
      Left            =   1590
      TabIndex        =   1
      Top             =   750
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   503
      _Version        =   393216
      Format          =   94502913
      CurrentDate     =   39675
   End
   Begin MSComCtl2.DTPicker DtDesde 
      Height          =   285
      Left            =   1590
      TabIndex        =   0
      Top             =   270
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   503
      _Version        =   393216
      Format          =   94502913
      CurrentDate     =   39675
   End
   Begin VB.CommandButton CmdImp 
      Caption         =   "Im&primir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1020
      TabIndex        =   2
      Top             =   1290
      Width           =   1605
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   750
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Desde :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   3
      Top             =   300
      Width           =   1245
   End
End
Attribute VB_Name = "FrmImpCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdImp_Click()
Dim arrform(3) As Variant
Dim arrparam(4) As Variant

If DtDesde > DtHasta Then
    MsgBox "Fecha Desde debe ser mayor a Fecha Hasta", vbInformation, "AVISO"
    Exit Sub
End If
    arrform(0) = "Desde='" & DtDesde & "'"
    arrform(1) = "empresa='" & VGCNx.DefaultDatabase & "'"
    arrform(2) = "Hasta='" & DtHasta & "'"
    
    arrparam(0) = VGCNx.DefaultDatabase
    arrparam(1) = DtDesde
    arrparam(2) = DtHasta
    arrparam(3) = VGParametros.cajerocodigo
    
    Call ImpresionRptProc("RptCaja.rpt", arrform, arrparam, "", "Registro de Ventas")

End Sub

Public Sub ImpresionRptProc(cNombreReporte As String, PFormulas(), Param(), Optional ORDEN As String, Optional Titulo As String)
Dim I As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .WindowTitle = Titulo
        Call PropCrystal(MDIPrincipal.CryRptProc)
        .ReportFileName = VGParamSistem.Rutareport
        If Right(VGParamSistem.Rutareport, 1) <> "\" Then
           .ReportFileName = VGParamSistem.Rutareport & "\"
        End If
        .ReportFileName = .ReportFileName & VGParamSistem.carpetareportes
        If Right(.ReportFileName, 1) <> "\" Then
        .ReportFileName = .ReportFileName & "\"
        End If
        .ReportFileName = .ReportFileName & cNombreReporte
        If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGcadenareport2
         End If
        
        .formulas(0) = "Desde='" & DtDesde & "'"
        .formulas(1) = "empresa='" & Trim(VGParametros.nomempresa) & "'"
        .formulas(2) = "Hasta='" & DtHasta & "'"

        If UBound(PFormulas) > 0 Then
            For I = 0 To UBound(PFormulas) - 1
                .formulas(2 + I) = PFormulas(I)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For I = 0 To UBound(Param) - 1
                .StoredProcParam(I) = Param(I)
            Next
        End If
        If ORDEN <> "" Then Call CrystOrden(MDIPrincipal.CryRptProc, ORDEN)
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub

Private Sub CrystOrden(ByRef cry As CrystalReport, cad As String)
Dim pos As Integer, cadaux As String, I As Integer
Dim valor As String
    Do While True
        pos = InStr(1, cad, ",", vbTextCompare)
        I = 0
        If pos = 0 Then Exit Do
        valor = Left(cad, pos - 1)
        cry.SortFields(I) = valor
        I = I + 1
        cad = Right(cad, (Len(cad) - pos))
    Loop
End Sub

Private Sub DtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub DtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub Form_Load()
DtDesde.Value = Date
DtHasta.Value = Date
End Sub



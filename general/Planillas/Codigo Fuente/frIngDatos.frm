VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frIngDatos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingreso de Datos"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2528
      TabIndex        =   8
      Top             =   2145
      Width           =   1365
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   788
      TabIndex        =   7
      Top             =   2145
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "Otros Datos del Trabajador"
      Height          =   1845
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4440
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   360
         Left            =   4050
         TabIndex        =   11
         Top             =   1020
         Visible         =   0   'False
         Width           =   240
      End
      Begin AplisetControlText.Aplitext tipoT 
         Height          =   315
         Left            =   990
         TabIndex        =   5
         Top             =   1020
         Visible         =   0   'False
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   556
         MaxLength       =   30
         Text            =   ""
      End
      Begin VB.OptionButton tipoBNo 
         Caption         =   "No"
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   1050
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.OptionButton tipoBSi 
         Caption         =   "Si"
         Height          =   285
         Left            =   990
         TabIndex        =   3
         Top             =   1050
         Visible         =   0   'False
         Width           =   630
      End
      Begin MSComCtl2.DTPicker tipoF 
         Height          =   330
         Left            =   990
         TabIndex        =   2
         Top             =   1027
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   36788
      End
      Begin AplisetControlText.Aplitext tipoN 
         Height          =   300
         Left            =   975
         TabIndex        =   1
         Top             =   1042
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         MaxLength       =   10
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xConcepto 
         Height          =   315
         Left            =   975
         TabIndex        =   0
         Top             =   480
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   556
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   1110
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Concepto:"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   525
         Width           =   735
      End
   End
End
Attribute VB_Name = "frIngDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMACEPTAR_CLICK()
    Select Case cmAceptar.Tag
        Case "T"
            If tipoT.Text = "" Then
                MsgBox "Falta registrar un valor de tipo alfanúmerico (Texto)", vbInformation
                tipoT.SetFocus
                Exit Sub
            End If
            DBSYSTEM.Execute "UPDATE TRABAJADORES SET " & xConcepto.Tag & "='" & tipoT.Text & "' WHERE CODTRAB='" & Trim(frTrab.xCodTrab.Text) & "'"
        Case "N"
            DBSYSTEM.Execute "UPDATE TRABAJADORES SET " & xConcepto.Tag & "=" & tipoN.Text & " WHERE CODTRAB='" & Trim(frTrab.xCodTrab.Text) & "'"
        Case "F"
            DBSYSTEM.Execute "UPDATE TRABAJADORES SET " & xConcepto.Tag & "=" & DateSQL(tipoF.Value) & " WHERE CODTRAB='" & Trim(frTrab.xCodTrab.Text) & "'"
        Case "B"
            DBSYSTEM.Execute "UPDATE TRABAJADORES SET " & xConcepto.Tag & "=" & IIf(tipoBSi.Value, -1, 0) & " WHERE CODTRAB='" & Trim(frTrab.xCodTrab.Text) & "'"
        Case Else
            Beep
    End Select
    Unload Me
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me



'Dim RSCUENTA As New ADODB.Recordset
'Set RSCUENTA = New ADODB.Recordset
'RSCUENTA.Open "SELECT PLANCTA_CODIGO,PLANCTA_DESCRIPCION FROM PLAN_CUENTA_NACIONAL WHERE PLANCTA_NIVEL=" & REGSISTEMA.scNivelCta, CONECTARDBSQL(xEmp.Text & "BDCONTABILIDAD"), adOpenKeyset, adLockReadOnly
'frmComun.CONECTAR RSCUENTA
'frmComun.Show 1
'If VGUTIL(1) <> "" Then
' xcuenta.Text = VGUTIL(1)
'End If
 
    
    
End Sub

Private Sub Command1_Click()
Dim RSCUENTA As New ADODB.Recordset
Set RSCUENTA = New ADODB.Recordset
RSCUENTA.Open "SELECT PLANCTA_CODIGO,PLANCTA_DESCRIPCION FROM PLAN_CUENTA_NACIONAL WHERE PLANCTA_NIVEL=" & REGSISTEMA.scNivelCta, CONECTARDBSQL(DevuelveValor("SELECT CODEMP FROM CFGASIENTOS", DBSYSTEM) & "BDCONTABILIDAD"), adOpenKeyset, adLockReadOnly
frmComun.CONECTAR RSCUENTA
frmComun.Show 1
If VGUTIL(1) <> "" Then
 tipoT.Text = VGUTIL(1)
End If

End Sub

Private Sub XCONCEPTO_DBLCLICK()
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "DATATRAB", DBSYSTEM, adOpenStatic, adLockReadOnly
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        cmAceptar.Enabled = True
        tipoT.Visible = False
        tipoN.Visible = False
        tipoF.Visible = False
        tipoBSi.Visible = False
        tipoBNo.Visible = False
        Select Case RSAUX!TIPODATA
            Case "T"
                tipoT.Visible = True
            Case "N"
                tipoN.Visible = True
                tipoN.Text = "0"
            Case "F"
                tipoF.Visible = True
                tipoF.Value = Date
            Case "B"
                tipoBSi.Visible = True
                tipoBNo.Visible = True
            Case Else
                Set RSAUX = Nothing
                Exit Sub
        End Select
        xConcepto.Text = VGUTIL(2)
        xConcepto.Tag = VGUTIL(1)
        cmAceptar.Tag = RSAUX!TIPODATA
    Else
        cmAceptar.Enabled = False
    End If
    Set RSAUX = Nothing
End Sub


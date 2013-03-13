VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frUbigeo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ubicación Geográfica (INEI)"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   Icon            =   "frUbigeo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3023
      TabIndex        =   5
      Top             =   2430
      Width           =   1305
   End
   Begin VB.CommandButton cmAcepta 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1403
      TabIndex        =   4
      Top             =   2430
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar UBIGEO"
      Height          =   2175
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5565
      Begin AplisetControlText.Aplitext xDist 
         Height          =   315
         Left            =   1620
         TabIndex        =   8
         Top             =   1380
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         Locked          =   -1  'True
         Text            =   "*** Seleccione"
      End
      Begin AplisetControlText.Aplitext xProv 
         Height          =   315
         Left            =   1620
         TabIndex        =   7
         Top             =   900
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         Locked          =   -1  'True
         Text            =   "*** Seleccione"
      End
      Begin AplisetControlText.Aplitext xDep 
         Height          =   315
         Left            =   1620
         TabIndex        =   6
         Top             =   420
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         Locked          =   -1  'True
         Text            =   "*** Seleccione"
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Distrito"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1410
         Width           =   480
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   945
         Width           =   660
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frUbigeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSDEPS As New ADODB.Recordset
Dim RSPROV As New ADODB.Recordset
Dim RSDIST As New ADODB.Recordset

Private Sub cmAcepta_Click()
    If xDist.Tag = "" Then
        MsgBox "NO HA TERMINADO DE SELECCIONAR EL UBIGEO", vbCritical
        Exit Sub
    End If
    SaveSetting App.CompanyName, "UBIGEO", "DEPARTAMENTO", xDep.Text
    SaveSetting App.CompanyName, "UBIGEO", "PROVINCIA", xProv.Text
    SaveSetting App.CompanyName, "UBIGEO", "DISTRITO", xDist.Text
    SaveSetting App.CompanyName, "UBIGEO", "CODDEPARTAMENTO", xDep.Tag
    SaveSetting App.CompanyName, "UBIGEO", "CODPROVINCIA", xProv.Tag
    SaveSetting App.CompanyName, "UBIGEO", "CODDISTRITO", xDist.Tag
    vpCodTmp = xDist.Tag
    vpTrasPrm = xDist.Text
    Unload Me
End Sub

Private Sub CMCANCELA_Click()
    Unload Me
End Sub

Private Sub FORM_ACTIVATE()
    If xDist.Text = "" Then xDep.SetFocus Else xDist.SetFocus
End Sub

Private Sub FORM_Load()
    If vpCodTmp = "" Then
        'QUIERE DECIR QUE ES POR PRIMERA VEZ
    End If
    'SE CARGA EL RECORDSET DE DEPARTAMENTOS
    RSDEPS.Open "UBIDEP", DbSystem, adOpenStatic
    xDep.Text = GetSetting(App.CompanyName, "UBIGEO", "DEPARTAMENTO", "")
    xDist.Text = GetSetting(App.CompanyName, "UBIGEO", "DISTRITO", "")
    xProv.Text = GetSetting(App.CompanyName, "UBIGEO", "PROVINCIA", "")
    xDep.Tag = GetSetting(App.CompanyName, "UBIGEO", "CODDEPARTAMENTO", "")
    xDist.Tag = GetSetting(App.CompanyName, "UBIGEO", "CODDISTRITO", "")
    xProv.Tag = GetSetting(App.CompanyName, "UBIGEO", "CODPROVINCIA", "")
End Sub

Private Sub FORM_UnLoad(CANCEL As Integer)
    RSDEPS.Close
    Set RSDEPS = Nothing
    Set RSPROV = Nothing
    Set RSDIST = Nothing
End Sub

Private Sub XDEP_DblClick()
    frmComun.CONECTAR RSDEPS
    frmComun.Show 1
    If vgUtil(1) <> "" Then
        xDep.Text = vgUtil(1) & " : " & vgUtil(2)
        xDep.Tag = vgUtil(1)
        xProv.Tag = ""
        xProv.Text = ""
        xDist.Tag = ""
        xDist.Text = ""
    End If
End Sub

Private Sub XDIST_DblClick()
    If xProv.Tag = "" Then
        MsgBox "DEBE SELECCIONAR PRIMERO LA PROVINCIA", vbCritical
        Exit Sub
    End If
    Dim CADSTR As String
    CADSTR = "SELECT * FROM UBIDIST WHERE CODIGO LIKE '" & xProv.Tag & "%' ORDER BY NOMBRE"
    RSDIST.Open CADSTR, DbSystem, adOpenStatic
    frmComun.CONECTAR RSDIST
    frmComun.Show 1
    If vgUtil(1) <> "" Then
        xDist.Text = vgUtil(1) & " : " & vgUtil(2)
        xDist.Tag = vgUtil(1)
    End If
    RSDIST.Close
End Sub

Private Sub XPROV_DblClick()
    If xDep.Tag = "" Then
        MsgBox "DEBE SELECCIONAR PRIMERO UN DEPARTAMENTO", vbCritical
        Exit Sub
    End If
    Dim CADSTR As String
    CADSTR = "SELECT * FROM UBIPROV WHERE CODIGO LIKE '" & xDep.Tag & "%' ORDER BY NOMBRE"
    RSPROV.Open CADSTR, DbSystem, adOpenStatic
    frmComun.CONECTAR RSPROV
    frmComun.Show 1
    If vgUtil(1) <> "" Then
        xProv.Text = vgUtil(1) & " : " & vgUtil(2)
        xProv.Tag = vgUtil(1)
        xDist.Text = ""
        xDist.Tag = ""
    End If
    RSPROV.Close
End Sub


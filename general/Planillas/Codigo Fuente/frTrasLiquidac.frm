VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frTrasLiquidac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso a Planillas"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "frTrasLiquidac.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   2861
      TabIndex        =   7
      Top             =   2970
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1053
      TabIndex        =   6
      Top             =   2970
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione el Periodo a donde traspasar"
      Height          =   1755
      Left            =   105
      TabIndex        =   0
      Top             =   1110
      Width           =   4995
      Begin AplisetControlText.Aplitext xGratif 
         Height          =   285
         Left            =   105
         TabIndex        =   5
         Top             =   1320
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xVacac 
         Height          =   285
         Left            =   105
         TabIndex        =   3
         Top             =   645
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Gratificaciones"
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   1050
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vacaciones"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"frTrasLiquidac.frx":08CA
      Height          =   825
      Left            =   900
      TabIndex        =   1
      Top             =   180
      Width           =   3960
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "frTrasLiquidac.frx":09A9
      Top             =   225
      Width           =   480
   End
End
Attribute VB_Name = "frTrasLiquidac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    RegLiquida.CronoVac = xVacac.Tag
    RegLiquida.CronoGrat = xGratif.Tag
    RegLiquida.Cancel = False
    Unload Me
End Sub

Private Sub Command2_Click()
    RegLiquida.Cancel = True
    Unload Me
End Sub

Private Sub Form_Load()
    xVacac.Tag = 0
    xGratif.Tag = 0
End Sub

Private Sub xGratif_Click()
    If Not RegLiquida.ActGrati Then Exit Sub
    Dim RsMeses As New ADODB.Recordset
    RsMeses.Open "SELECT CODIGO, NOMBRE FROM NOMBOL WHERE MES IN (SELECT MESACTIVO FROM MESESACT)", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RsMeses.RecordCount = 0 Then
        MsgBox "No se han encontrado meses en actividad", vbCritical
        Set RsMeses = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RsMeses
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xGratif.Text = RsMeses!NOMBRE
        xGratif.Tag = RsMeses!CODIGO
    End If
    Set RsMeses = Nothing
End Sub

Private Sub xVacac_DblClick()
    If Not RegLiquida.ActVaca Then Exit Sub
    Dim RsMeses As New ADODB.Recordset
    RsMeses.Open "SELECT CODIGO, NOMBRE FROM NOMBOL WHERE MES IN (SELECT MESACTIVO FROM MESESACT)", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RsMeses.RecordCount = 0 Then
        MsgBox "No se han encontrado meses en actividad", vbCritical
        Set RsMeses = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RsMeses
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xVacac.Text = RsMeses!NOMBRE
        xVacac.Tag = RsMeses!CODIGO
    End If
    Set RsMeses = Nothing
End Sub

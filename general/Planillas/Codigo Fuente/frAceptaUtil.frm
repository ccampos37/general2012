VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frAceptaUtil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aceptar Planilla de Utilidad"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frAceptaUtil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   2577
      TabIndex        =   7
      Top             =   2475
      Width           =   1215
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   889
      TabIndex        =   6
      Top             =   2475
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Aceptar Planilla"
      Height          =   1590
      Left            =   120
      TabIndex        =   0
      Top             =   765
      Width           =   4455
      Begin AplisetControlText.Aplitext xPeriodo 
         Height          =   330
         Left            =   1605
         TabIndex        =   8
         Top             =   780
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   582
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Periodo de Pago"
         Height          =   195
         Left            =   210
         TabIndex        =   1
         Top             =   795
         Width           =   1185
      End
      Begin VB.Label xMonto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   2610
         TabIndex        =   3
         Top             =   330
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Participacion a distribuir"
         Height          =   195
         Left            =   225
         TabIndex        =   2
         Top             =   390
         Width           =   2070
      End
   End
   Begin VB.Label xPlanilla 
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   345
      Width           =   4455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Planilla de Utilidad"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1290
   End
End
Attribute VB_Name = "frAceptaUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CMACEPTAR_CLICK()
    If xPERIODO.Tag = "" Then
        MsgBox "Seleccione un PERIODO para aceptar la Planilla Utilidad", vbExclamation
        xPERIODO.SetFocus
        Exit Sub
    End If
    If MsgBox("Esta seguro de aceptar los valores para la planilla de utilidad estos valores, serán transferidos a los movimientos del PERIODO que Ud. haya seleccionado " & xPlanilla.Caption, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    DbSystem.Execute "UPDATE UTIL SET CERRADO=1,NOMBOL=" & xPERIODO.Tag & " WHERE CODIGO=" & vpTrasPrm
    Unload Me
End Sub

Private Sub cmCancelar_Click()
    Unload Me
End Sub

Private Sub FORM_LOAD()
    xPlanilla.Caption = DevuelveValor("SELECT NOMBRE FROM UTIL WHERE CODIGO=" & vpTrasPrm, DbSystem)
    xMonto.Caption = Format(DevuelveValor("SELECT PARTDIST FROM UTIL WHERE CODIGO=" & vpTrasPrm, DbSystem), "###,###,##0.00 ")
End Sub

Private Sub XPERIODO_DBLCLICK()
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT NOMBOL.CODIGO, NOMBOL.NOMBRE FROM NOMBOL, MESESACT WHERE NOMBOL.MES=MESESACT.MESACTIVO", DbSystem, adOpenStatic, adLockReadOnly
    If RSAUX.RecordCount = 0 Or RSAUX.EOF Then
        MsgBox "No se encontrado meses activos", vbInformation
        cmAceptar.Enabled = False
        Set RSAUX = Nothing
        Exit Sub
    End If
    frmComun.Conectar RSAUX
    frmComun.Show 1
    If vgUtil(1) <> "" Then
        xPERIODO.Tag = vgUtil(1)
        xPERIODO.Text = vgUtil(2)
    End If
    Set RSAUX = Nothing
End Sub


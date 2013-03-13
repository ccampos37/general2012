VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRestSalAnt 
   Caption         =   "Restaurar Cierre Anteriores"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   Icon            =   "frmRestCierre.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3180
      Left            =   84
      TabIndex        =   2
      Top             =   90
      Width           =   4296
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1260
         TabIndex        =   5
         Top             =   2700
         Width           =   1932
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MMMM 'del' yyyy"
         Format          =   16318467
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   495
         Picture         =   "frmRestCierre.frx":08CA
         Top             =   435
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   $"frmRestCierre.frx":25C4
         Height          =   870
         Left            =   315
         TabIndex        =   6
         Top             =   1170
         Width           =   3360
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccione el mes de donde va realizar el cambio de valorización."
         Height          =   495
         Left            =   285
         TabIndex        =   4
         Top             =   2025
         Width           =   3105
      End
      Begin VB.Label Label1 
         Caption         =   "Después de realizar esta opción debe valorizar de nuevo todos los meses afectados. "
         Height          =   810
         Left            =   1320
         TabIndex        =   3
         Top             =   330
         Width           =   2325
      End
   End
   Begin VB.CommandButton Command21 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   660
      Left            =   2304
      Picture         =   "frmRestCierre.frx":2658
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3324
      Width           =   696
   End
   Begin VB.CommandButton Command20 
      Caption         =   "&Aceptar"
      Height          =   660
      Left            =   1584
      Picture         =   "frmRestCierre.frx":2A9A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3324
      Width           =   696
   End
End
Attribute VB_Name = "frmRestSalAnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command20_Click()
Dim nMes As Integer
Dim Rsql1 As String
On Error GoTo Err
nMes = Month(DTPicker1)

If MsgBox("Se va a levantar el cierre del mes de : " & MonthName(nMes) & " a la fecha Actual ", vbInformation, "Aviso") Then
   'Rsql1 = "Update MovAlmCab set CACIERRE =  FALSE " & _
                   " where  CAALMA = '" & VGAlma & "'   AND MONTH(CAFECDOC) >=" & nMes & " AND YEAR(CAFECDOC) = " & Year(DTPicker1)
   Rsql1 = "Update MovAlmCab set CACIERRE =  FALSE " & _
                   " where  CAALMA = '" & VGAlma & "' AND FORMAT( YEAR(CAFECDOC),'0000' )+ FORMAT( MONTH(CAFECDOC),'00' ) >='" & Format(DTPicker1.Year, "0000") + Format(DTPicker1.Month, "00") & "'"
   
   
   VGcnx.Execute Rsql1
   
   Rsql1 = "DELETE FROM AL_CIERRESMENSUALES where  CIERRALMA = '" & VGAlma & "'   AND CIERRMES>='" & Format(DTPicker1.Year, "0000") & Format(DTPicker1.Month, "00") & "'"
   VGcnx.Execute Rsql1
     
   MsgBox "Se realizó sastifactoriamente", vbInformation, "Aviso"
End If
Exit Sub
Err:
   MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub Command21_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  DTPicker1 = Date
End Sub

'rSql = "SELECT TOP 1 CAFECDOC From MovAlmCab  WHERE CACIERRE = false AND CAALMA = '" & VGAlma & "'  AND MONTH(CAFECDOC) <= " & nMes & " ORDER BY CAFECDOC asc"
'Set Rs2 = New ADODB.Recordset
'Rs2.Open rSql, Vgcnx, adOpenStatic
'If Not Rs2.EOF Then
'    Rs2.MoveFirst
'    nMes = Month(Rs2(0))
'    If nMes = 13 Then
'         nMes = 1
'    End If
'    MsgBox "Se va a levantar el cierre del  mes : " & MonthName(nMes), vbInformation, "Aviso"
'End If

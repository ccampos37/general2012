VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frmDiaEspecial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Goce de Vacaciones Pendientes"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "frmDiaEspecial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   2745
      TabIndex        =   11
      Top             =   5145
      Width           =   1305
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1155
      TabIndex        =   10
      Top             =   5145
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Definición"
      Height          =   4095
      Left            =   150
      TabIndex        =   5
      Top             =   915
      Width           =   4875
      Begin AplisetControlText.Aplitext xDescripcion 
         Height          =   300
         Left            =   240
         TabIndex        =   17
         Top             =   3600
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   529
         MaxLength       =   50
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xTrab 
         Height          =   300
         Left            =   240
         TabIndex        =   2
         Top             =   675
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   529
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xDias 
         Height          =   315
         Left            =   2670
         TabIndex        =   9
         Top             =   2910
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         MaxLength       =   2
         Text            =   "1"
         TipoDato        =   "N"
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   330
         Left            =   2670
         TabIndex        =   7
         Top             =   2490
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         Format          =   25034753
         CurrentDate     =   36811
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   330
         Left            =   2670
         TabIndex        =   4
         Top             =   2055
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         Format          =   25034753
         CurrentDate     =   36811
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   3330
         Width           =   840
      End
      Begin VB.Label xPendiente 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   2670
         TabIndex        =   15
         Top             =   1410
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Días Pendientes"
         Height          =   195
         Left            =   2655
         TabIndex        =   14
         Top             =   1155
         Width           =   1185
      End
      Begin VB.Label xPeriodo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Periodo Vacacional"
         Height          =   315
         Left            =   255
         TabIndex        =   13
         Top             =   1410
         Width           =   2220
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1140
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Trabajador"
         Height          =   195
         Left            =   225
         TabIndex        =   0
         Top             =   420
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total &Días"
         Height          =   195
         Left            =   915
         TabIndex        =   8
         Top             =   2970
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de &Regreso"
         Height          =   195
         Left            =   915
         TabIndex        =   6
         Top             =   2550
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de &Salida"
         Height          =   195
         Left            =   915
         TabIndex        =   3
         Top             =   2100
         Width           =   1155
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDiaEspecial.frx":0ABA
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   810
      TabIndex        =   1
      Top             =   120
      Width           =   4275
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmDiaEspecial.frx":0B6A
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   5265
   End
End
Attribute VB_Name = "frmDiaEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CMACEPTAR_CLICK()
    If Val(xDias.Text) > Val(xPendiente.Caption) Then
        MsgBox "Los dias no pueden ser mayores a los especificados como pendientes, los cuales pueden ser como máximo: " & xPendiente.Caption, vbInformation
        xDias.Text = xPendiente.Caption
        Exit Sub
    End If
    If xDias.Text = 0 Then
        MsgBox "Los dias no pueden especificarse en cero", vbInformation
        xDias.SetFocus
        Exit Sub
    End If
    DBSYSTEM.Execute "UPDATE HISTOVAC SET DIAS=DIAS-" & xDias.Text & " WHERE CODIGO=" & xPendiente.Tag
    DBSYSTEM.Execute "INSERT INTO DIASGOCE (CODIGO,FECHAINI,FECHAFIN,DIAS,DESCRIPCION) VALUES (" & xPendiente.Tag & "," & DateSQL(xFechaIni.Value) & "," & DateSQL(xFechaFin.Value) & "," & xDias.Text & ",'" & xDescripcion.Text & "')"
    Unload Me
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub FORM_LOAD()
    xFechaIni.Value = Date
    xFechaFin.Value = Date
End Sub

Private Sub XDIAS_CHANGE()
    xFechaFin.Value = DateAdd("D", Val(xDias.Text) - 1, xFechaIni.Value)
End Sub

Private Sub XFECHAFIN_CHANGE()
    xDias.Text = DateDiff("D", xFechaIni.Value, xFechaFin.Value) + 1
End Sub

Private Sub XFECHAINI_CHANGE()
    'XDIAS.TEXT = DATEDIFF("D", XFECHAINI.VALUE, XFECHAFIN.VALUE) + 1
    xFechaFin.Value = DateAdd("D", Val(xDias.Text) - 1, xFechaIni.Value)
End Sub


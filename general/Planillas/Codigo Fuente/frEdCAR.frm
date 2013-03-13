VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frEdCAR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edición de Centros de Alto Riesgo"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   Icon            =   "frEdCAR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   2280
      TabIndex        =   12
      Top             =   2550
      Width           =   1125
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   840
      TabIndex        =   11
      Top             =   2550
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Centros de Alto Riesgo"
      Height          =   2235
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   4035
      Begin AplisetControlText.Aplitext txRUC 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1830
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         MaxLength       =   11
         Text            =   ""
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext txCorrela 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   1500
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   5
         Text            =   ""
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext txTasa 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   1140
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   8
         Text            =   ""
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext txDescrip 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   780
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         MaxLength       =   35
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext txCodigo 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   420
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         MaxLength       =   6
         Text            =   ""
         TipoCodigo      =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C."
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1860
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Correlativo"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1500
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tasa (en %)"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1140
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   780
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frEdCAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMACEPTAR_CLICK()
    On Error GoTo ERRORES
    If txCodigo.Text = "" Then
        MsgBox "DEBE INGRESAR UN CÓDIGO VÁLIDO"
        Exit Sub
    End If
    If txDescrip.Text = "" Then
        MsgBox "DEBE INGRESAR UNA DESCRIPCIÓN/NOMBRE DEL CENTRO DE ALTO RIESGO VÁLIDA"
        Exit Sub
    End If
    If txRUC.Text = "" Then
        MsgBox "DEBE INGRESAR UN NÚMERO DE R.U.C.", vbCritical
        Exit Sub
    End If
    If Not Validar_RUC(txRUC.Text) Then Exit Sub
    If VPTAREA = "NUEVO" Then
        DBSYSTEM.Execute ("INSERT INTO CENTROSAR (CODCAR,NOMBRE,TASA,CORRELATIVO,RUC) SELECT '" & txCodigo.Text & "','" & txDescrip.Text & "'," & ESNULO(txTasa.Text, 0) & "," & ESNULO(txCorrela.Text, 0) & ",'" & txRUC.Text & "'")
    Else
        DBSYSTEM.Execute "UPDATE CENTROSAR SET NOMBRE='" & txDescrip.Text & "', TASA= " & txTasa.Text & ",CORRELATIVO=" & txCorrela.Text & ", RUC='" & "" & txRUC.Text & "' WHERE CODCAR='" & txCodigo.Text & "'"
    End If
    Unload Me
    Exit Sub
ERRORES:
    MsgBox "CÓDIGO DUPLICADO", vbCritical
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub FORM_LOAD()
    If VPTAREA = "EDITAR" Then
        txCodigo.Text = frCAR.lvCAR.SelectedItem.Text
        txDescrip.Text = frCAR.lvCAR.SelectedItem.SubItems(1)
        txTasa.Text = frCAR.lvCAR.SelectedItem.SubItems(2)
        txCorrela.Text = frCAR.lvCAR.SelectedItem.SubItems(3)
        txRUC.Text = frCAR.lvCAR.SelectedItem.SubItems(4)
        txCodigo.Locked = True
    End If
End Sub

Private Sub TXRUC_LOSTFOCUS()
    If txRUC.Text = "" Then Exit Sub
    I = Validar_RUC(txRUC.Text)
End Sub


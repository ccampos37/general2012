VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmLiskits 
   Caption         =   "Listado de Kits"
   ClientHeight    =   4005
   ClientLeft      =   1365
   ClientTop       =   1545
   ClientWidth     =   3645
   LinkTopic       =   "Form2"
   ScaleHeight     =   4005
   ScaleWidth      =   3645
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   165
      TabIndex        =   6
      Top             =   1215
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox TxA1 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxA2 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Listar Todo"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Código Inicial  :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Código Final   :"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   180
      TabIndex        =   3
      Top             =   2670
      Width           =   3255
      Begin VB.CommandButton CmdA 
         Caption         =   "&Aceptar"
         Height          =   675
         Left            =   600
         Picture         =   "frmLiskits.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdS 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   1800
         Picture         =   "frmLiskits.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   165
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.OptionButton Option1 
         Caption         =   "Por Código "
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Descripción"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   2055
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   2880
         Top             =   3000
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
End
Attribute VB_Name = "frmLiskits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Check1_Click()
'If Check1.Value = 1 Then
'    TxC1.Enabled = False
'    TxC2.Enabled = False
'Else
'    TxC1.Enabled = True
'    TxC2.Enabled = True
'End If
'End Sub
'
'Private Sub TxA1_DblClick()
'Static Adodc2 As ADODB.Recordset
'Set Adodc2 = New ADODB.Recordset
'    Select Case Index
'     Case 0:
'
'         Adodc2.Open "SELECT CCODCLI,CNOMCLI FROM MAECLI", Vgcnx, adOpenStatic, adLockOptimistic
'         frmReferencia.conectar Adodc2, "SELECT CCODCLI,CNOMCLI FROM MAECLI"
'         frmReferencia.Label1.Caption = "Maestro de Clientes"
'         frmReferencia.show vbmodal
'         Adodc2.Close
'         If vGUtil(1) <> "" Then
'           TxA1.text = (vGUtil(1))
'           'Text1(1) = VGUTIL(2)
'         End If
'   End Select
'
'End Sub
'
'Private Sub TxA1_GotFocus()
'Enfoque TxA1
'End Sub
'
'Private Sub TxA1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then TxA1_DblClick
'End Sub
'
'Private Sub TxA1_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'
'    If Trim(TxA1) <> "" Then
'          If Existe(1, TxA1, "kits", "Codkit", False) = False Then
'            MsgBox "Codigo de Cliente no existe", vbInformation, mensaje1
'            TxA1.SetFocus: Exit Sub
'          End If
'    End If
'    TxA2.SetFocus: Exit Sub
'
'Else
'    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End If
'End Sub
'
'Private Sub TxA2_DblClick()
'Static Adodc2 As ADODB.Recordset
'Set Adodc2 = New ADODB.Recordset
'    Select Case Index
'     Case 0:
'
'         Adodc2.Open "SELECT acodigo,adescri FROM acodigo", Vgcnx, adOpenStatic, adLockOptimistic
'         frmReferencia.conectar Adodc2, "SELECT CCODCLI,CNOMCLI FROM MAECLI"
'         frmReferencia.Label1.Caption = "Maestro de Clientes"
'         frmReferencia.show vbmodal
'         Adodc2.Close
'         If vGUtil(1) <> "" Then
'           TxA2.text = (vGUtil(1))
'           'Text1(1) = VGUTIL(2)
'         End If
'   End Select
'End Sub
'
'Private Sub TxA2_GotFocus()
'Enfoque TxA2
'End Sub
'
'Private Sub TxA2_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then TxA2_DblClick
'End Sub
'
'Private Sub TxA2_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Trim(TxA2) <> "" Then
'          If Existe(1, TxA2, "MAECLI", "CCODCLI", False) = False Then
'            MsgBox "Codigo de Cliente no existe", vbInformation, mensaje1
'            TxA2.SetFocus: Exit Sub
'          End If
'    End If
'    CmdA.SetFocus: Exit Sub
'
'Else
'    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End If
'End Sub
'
'Private Sub TxC1_DblClick()
'Static Adodc2 As ADODB.Recordset
'Set Adodc2 = New ADODB.Recordset
'    Select Case Index
'     Case 0:
'        Adodc2.Open "SELECT CNUMRUC,CNOMCLI FROM MAECLI", Vgcnx, adOpenStatic, adLockOptimistic
'        frmReferencia.conectar Adodc2, "SELECT CNUMRUC,CNOMCLI FROM MAECLI"
'        frmReferencia.Label1.Caption = "RUC de Clientes"
'        frmReferencia.show vbmodal
'        Adodc2.Close
'         If vGUtil(1) <> "" Then
'           TxC1.text = (vGUtil(1))
'           'Text1(1) = VGUTIL(2)
'         End If
'   End Select
'End Sub
'
'Private Sub TxC1_GotFocus()
'Enfoque TxC1
'End Sub
'
'Private Sub TxC1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then TxC1_DblClick
'End Sub
'
'Private Sub TxC1_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Trim(TxC1) <> "" Then
'          If Existe(1, TxC1, "MAECLI", "CNUMRUC", False) = False Then
'            MsgBox "RUC de Cliente no existe", vbInformation, mensaje1
'            TxC1.SetFocus: Exit Sub
'          End If
'    End If
'    TxC2.SetFocus: Exit Sub
'
'Else
'    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End If
'End Sub
'
'Private Sub TxC2_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then TxC2_DblClick
'End Sub
'Private Sub TxC2_DblClick()
'Static Adodc2 As ADODB.Recordset
'Set Adodc2 = New ADODB.Recordset
'    Select Case Index
'     Case 0:
'        Adodc2.Open "SELECT CNUMRUC,CNOMCLI FROM MAECLI", Vgcnx, adOpenStatic, adLockOptimistic
'        frmReferencia.conectar Adodc2, "SELECT CNUMRUC,CNOMCLI FROM MAECLI"
'        frmReferencia.Label1.Caption = "RUC de Clientes"
'        frmReferencia.show vbmodal
'        Adodc2.Close
'         If vGUtil(1) <> "" Then
'           TxC2.text = (vGUtil(1))
'           'Text1(1) = VGUTIL(2)
'         End If
'   End Select
'End Sub
'
'Private Sub TxC2_GotFocus()
'Enfoque TxC2
'End Sub
'
'Private Sub Limpiar()
'TxC1 = ""
'TxC2 = ""
'End Sub


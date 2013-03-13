VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmKit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Articulos"
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   4935
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3315
         Left            =   120
         TabIndex        =   11
         Top             =   900
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   5847
         _Version        =   393216
         SelectionMode   =   1
      End
      Begin VB.Label Label6 
         Caption         =   "<< Selecione con la tecla TAB o un click del Mouse >>"
         Height          =   255
         Left            =   195
         TabIndex        =   12
         Top             =   270
         Width           =   4455
      End
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Agrupar"
      Height          =   4335
      Left            =   5040
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton CmdBuscarArt 
         Caption         =   "..."
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid Salida 
         Height          =   3015
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin VB.Label Label2 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   2775
      End
   End
   Begin VB.CommandButton CmdEnviar 
      Caption         =   "&Enviar"
      Height          =   375
      Left            =   3705
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmKit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As Recordset
Dim Adoreg1 As ADODB.Recordset


Private Sub CmdBuscarArt_Click()
  VGForm1 = 13
  FormAyuArt1.Show 1
End Sub

Private Sub CmdEnviar_Click()
  Dim i As Integer
    If FG1.Rows = 1 Then Exit Sub
    Salida.Rows = 1
    For i = 0 To FG1.Rows - 1
      If FG1.TextMatrix(i, 0) = "*" Then
         Salida.AddItem (FG1.TextMatrix(i, 1) & vbTab & FG1.TextMatrix(i, 2) & vbTab & FG1.TextMatrix(i, 3) & vbTab & FG1.TextMatrix(i, 4) & vbTab & FG1.TextMatrix(i, 5) & vbTab & FG1.TextMatrix(i, 6))
      End If
    Next i

End Sub

Private Sub CmdGrabar_Click()
  grabar
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim RSQL As String
 RSQL = "select  p.ACODIGO, p.ADESCRI,p.AUNIDAD,p.AFAMILIA" & _
        "from MaeArt p LEFT JOIN StkArt n  ON  p.ACODIGO = n.STCODIGO" & _
        "group by p.ACODIGO ,p.ADESCRI,p.AUNIDAD,p.AFAMILIA  " '
 'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
 
 Set rs = cConexCom.Execute(RSQL)
 While Not rs.EOF
        FG1.AddItem (" " & vbTab & rs(0) & vbTab & rs(1) & vbTab & rs(2) & vbTab & rs(3))
        rs.MoveNext
 Wend
End Sub

Private Sub grabar()
   Dim i As Integer
   Dim CANT As Double
   Dim monto As Double
   Dim RSQL As String
   Set Adoreg1 = New ADODB.Recordset
   On Error GoTo Err
  
   RSQL = "select * from kits"
   Adoreg1.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
   For i = 1 To Salida.Rows - 1
                Adoreg1.AddNew
                Adoreg1("kit_codigo") = Text1
                Adoreg1("kit_acodigo") = Salida.TextMatrix(i, 0)
                Adoreg1("kit_cantidad") = Salida.TextMatrix(i, 2)
                Adoreg1.UpdateBatch
                Adoreg1.Requery
                CANT = Salida.TextMatrix(i, 3) + CANT
                monto = monto + Salida.TextMatrix(i, 4) * Salida.TextMatrix(i, 3)
  Next i
  ' actualiza el precio promedio
  
  If monto <> 0 Then
     'promedio = CANT / monto
     'cantidad = CANT
     'rsql = "Update stkart set stkdis =" & CANT & " where   stk='" & nroguia & "' and  acodigo='" & Text1 & "' and  codigo ='" & Salida.TextMatrix(i, 0) & "'and  giasa ='" & Salida.TextMatrix(i, 5) & "'"
     cConexCom.Execute RSQL
     'InicializaSalida
  Else
   MsgBox "El monto no puede ser igual a cero ", vbExclamation, "Grabar"
  End If
  Set Adoreg1 = Nothing
  Exit Sub
Err:
 monto = Err.Number
  MsgBox Err.Description
  'MsgBox err.Number
  
End Sub

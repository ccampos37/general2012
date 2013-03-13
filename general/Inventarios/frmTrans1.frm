VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTransMabel 
   Caption         =   "Enviar de IASA a Mabel"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9660
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   5025
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Enviar"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Agrupar"
      Height          =   4335
      Left            =   5040
      TabIndex        =   14
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton CmdBuscarArt 
         Caption         =   "..."
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid Salida 
         Height          =   3015
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Guia"
      Height          =   4335
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   4935
      Begin VB.TextBox TxtNroGuia 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3015
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36686
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Nro. Guia"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "<< Selecione con la tecla TAB o un click del Mouse >>"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Ubicacion"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmTransMabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim promedio As Double
Dim cantidad As Double

Private Sub CmdBuscarArt_Click()
  VGForm1 = 12
  FormAyuArt1.Show 1
End Sub

Private Sub CmdGrabar_Click()
Dim i As Integer, f As Integer
 If Trim(Text1) = "" Then
     MsgBox "Antes de Grabar ,debe asociar", vbExclamation, "Guia"
     Exit Sub
 End If
 If Salida.Rows = 1 Then Exit Sub
 grabar
 
otro:
 f = FG1.Rows - 1
  For i = 0 To f
   If FG1.TextMatrix(i, 0) = "*" Then
          If FG1.Rows > 2 Then
              FG1.RemoveItem i
              GoTo otro
         Else
             FG1.Clear
             FG1.Rows = 1
             FG1.Row = 0
             inicializaFG1
            Exit For
       End If
     End If
     'MsgBox f & " " & i
  Next i

 
  
 ' Wend
  Salida.Clear
  Salida.Rows = 1
  InicializaSalida
 Text1 = ""
 Label3 = ""
End Sub


Private Sub Command3_Click()
 On Error GoTo cErrorAbrir
 CommonDialog1.CancelError = True
 CommonDialog1.Filter = "Dos (*.dbf)|*.dbf|Todos los Archivos (*.*)|*.*"
 CommonDialog1.FilterIndex = 1
 CommonDialog1.Action = 1
 Text11 = CommonDialog1.FileName
SalirAbrir:
 Text11.SetFocus
 Exit Sub
cErrorAbrir:
 Resume SalirAbrir
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
   Dim i As Integer
   If FG1.Rows = 1 Then Exit Sub
  Salida.Rows = 1
  For i = 0 To FG1.Rows - 1
     If FG1.TextMatrix(i, 0) = "*" Then
        Salida.AddItem (FG1.TextMatrix(i, 1) & vbTab & FG1.TextMatrix(i, 2) & vbTab & FG1.TextMatrix(i, 3) & vbTab & FG1.TextMatrix(i, 4) & vbTab & FG1.TextMatrix(i, 5))
     End If
  Next i
End Sub

Private Sub Form_Load()
  Data3.DatabaseName = cRuta2
  Data3.RecordSource = "Select * from stkart"
  central Me
  FG1.Clear
  Salida.Clear
  inicializaFG1
  InicializaSalida
  limpiaTexto
End Sub

Private Sub Text11_DblClick()
 If InStr(Trim(Text11), "dbf") Then
      TxtNroGuia.SetFocus
  End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
  Dim db As Database
  Dim rs As Recordset
  Dim rsql As String
   If KeyAscii = 13 And InStr(Trim(Text11), "gmabel.dbf") Then
         rsql = "select  * from mabel "
         Set db = Workspaces(0).OpenDatabase(cRuta2)           '     "C:\WINDOWS
         Set rs = db.OpenRecordset(rsql, dbOpenSnapshot)
         If rs.EOF Then
              MsgBox "No hay registro  en la guia", vbCritical
              End
         End If
         FG1.Rows = 1
         rs.MoveFirst
         While Not rs.EOF
               FG1.AddItem (" " & vbTab & rs(0) & vbTab & rs(1) & vbTab & rs(2) & vbTab & Format(rs(3), "####0.00") & vbTab & Format(rs(4), "####0.000"))
               rs.MoveNext
         Wend
         TxtNroGuia.SetFocus
   End If
End Sub


Private Sub FG1_Click()
  If FG1.TextMatrix(FG1.Row, 0) = "*" Then
     FG1.TextMatrix(FG1.Row, 0) = " "
  Else
     FG1.TextMatrix(FG1.Row, 0) = "*"
  End If
End Sub

Private Sub FG1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        FG1_Click
   End If
End Sub

Private Sub grabar()
   Dim i As Integer
   Dim CANT As Double
   Dim monto As Double
 
  For i = 1 To Salida.Rows - 1
          CANT = Salida.TextMatrix(i, 3) + CANT
         monto = monto + Salida.TextMatrix(i, 4)
  Next i

  If monto <> 0 Then
     promedio = CANT / monto
     cantidad = CANT
     grabastk
     InicializaSalida
  Else
   MsgBox "El monto no puede ser igual a cero ", vbExclamation, "Grabar"
  End If
  
End Sub

Private Sub inicializaFG1()
  FG1.FormatString = "  ^ Seleccion | Codigo Art.|   Descripcion |  Unidad|Cantidad | Precio Total. "
  FG1.Row = 0
  FG1.Cols = 6
  FG1.ColWidth(0) = 800
  FG1.ColWidth(1) = 1200
  FG1.ColWidth(2) = 2000
  FG1.ColWidth(3) = 890
  FG1.ColWidth(4) = 1200
  FG1.ColWidth(5) = 1200
End Sub

Private Sub InicializaSalida()
 Salida.FormatString = "    Codigo Art.|   Descripcion |  Unidad|Cantidad | Precio Total. "
  Salida.Row = 0
  Salida.Cols = 5

  Salida.ColWidth(0) = 1200
  Salida.ColWidth(1) = 2600
  Salida.ColWidth(2) = 890
  Salida.ColWidth(3) = 1200
  Salida.ColWidth(4) = 1200

End Sub

Private Sub limpiaTexto()
   Text1 = ""
   Text11 = ""
   TxtNroGuia = ""
   Label3 = ""
  
End Sub


Public Sub grabastk()
   Dim criterio As String
   Dim cadena As String
   Dim auxdisp As Double
   criterio = " STCODIGO = " & Chr$(34) + Text1 + Chr$(34)
   criterio = criterio + "and  STALMA = " & Chr$(34) + VGAlma + Chr$(34)
   Data3.Recordset.FindFirst criterio
   If Not Data3.Recordset.NoMatch Then
     
     Data3.Recordset.Edit
     auxdisp = Data3.Recordset("STSKDIS")
   Else
      Data3.Recordset.AddNew
       Data3.Recordset("STALMA") = VGAlma   '"01"
       Data3.Recordset("STCODIGO") = Text1
  End If
     If Data3.Recordset("STKPREPRO") <> 0 Then  'no se registrado algun precio
        Data3.Recordset("STKPREPRO") = (promedio * cantidad + auxdisp * Data3.Recordset("STKPREPRO")) / (cantidad + auxdisp)
         Data3.Recordset("STSKDIS") = cantidad + auxdisp
     Else
      Data3.Recordset("STKPREPRO") = promedio
      Data3.Recordset("STKPREULT") = promedio
      Data3.Recordset("STKFECULT") = DTPicker1
       Data3.Recordset("STSKDIS") = cantidad
     End If
     If IsNull(Data3.Recordset("stkfecult")) Then
         ' Data3.Recordset("stkfecult") = " "
     End If
   
     
   Data3.Recordset.Update
   Data3.Refresh
End Sub

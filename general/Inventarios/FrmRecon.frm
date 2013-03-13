VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRecon 
   Caption         =   "Generación de Toma de Inventario"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7770
   LinkTopic       =   "Form2"
   ScaleHeight     =   5895
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox t 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   930
      TabIndex        =   5
      Top             =   5070
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   285
      TabIndex        =   4
      Top             =   30
      Width           =   7215
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   5610
         TabIndex        =   0
         Top             =   300
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   36756
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   3945
         Left            =   240
         TabIndex        =   1
         Top             =   750
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   6959
         _Version        =   393216
         Cols            =   3
      End
      Begin VB.Label Label2 
         Caption         =   "Ingrese las Cantidades Físicas"
         Height          =   300
         Left            =   315
         TabIndex        =   7
         Top             =   390
         Width           =   2340
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha  :"
         Height          =   255
         Left            =   4620
         TabIndex        =   6
         Top             =   330
         Width           =   795
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Grabar"
      Height          =   675
      Left            =   2838
      Picture         =   "FrmRecon.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5085
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   4158
      Picture         =   "FrmRecon.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5085
      Width           =   775
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   300
      Top             =   5235
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
End
Attribute VB_Name = "FrmRecon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As Integer
Dim Condi As String
Dim Valor As String
Dim corre As String

Private Sub CmdAceptar_Click()
Dim Update As String
Dim I As Integer
Condi = "": Valor = ""
Flex.Col = 0
For I = 1 To Con
    Flex.Row = I
    Flex.Col = 1
    Flex.Col = 4
    Update = "Update INVFIS set ICANTFIS=" & Valor & " Where ICODART='" & Condi & "' "
    cConexCom.Execute Update
Next
CmdAceptar.Enabled = False
CmdSalir.SetFocus
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim I As Integer
Dim Remi As Integer

If KeyCode = 13 Then
  Flex.Clear
  Flex.Row = 0
  Remi = 1
  For I = 1 To Con - 2
    Flex.RemoveItem (Remi)
  Next
  Con = 1
  CargaFlex
'  DTPicker1.SetFocus
End If
End Sub

Private Sub CargaFlex()
Dim rSql As String
Dim Adodc1 As New ADODB.Recordset

Flex.Rows = 1
Flex.Cols = 5
Flex.FormatString = " |   Código |   Descripción    | Cant. Stock | Cant. Física "
Flex.ColWidth(0) = 600
Flex.ColWidth(1) = 800
Flex.ColWidth(2) = 2800
Flex.ColWidth(3) = 1200
Flex.ColWidth(4) = 1200

rSql = "Select ICODART,ADESCRI,ICANTFIS "
rSql = rSql & "from HISTO_INV  A,INVFIS B, MAEART C where A.hCORRELA=B.iCORRELA and "
rSql = rSql & "B.ICODART=C.ACODIGO and hfecha=#" & Format(DTPicker1, "mm/dd/yyyy") & "# "
rSql = rSql & "Order by ICodArt"

Adodc1.Open rSql, cConexCom, adOpenDynamic, adLockOptimistic

If Adodc1.RecordCount = 0 Then
    MsgBox "No hay artículos disponibles en el almacén", vbInformation, "Aviso"
    Adodc1.Close
    Exit Sub
End If
  
  Con = Adodc1.RecordCount
  Adodc1.MoveFirst
'  Flex.Visible = False
  cCod = "": nStock = 0
  Do While Not Adodc1.EOF
     cCod = Adodc1(0)
     If Flex.RowIsVisible(Flex.Row) = True Then
        Flex.AddItem (" " & vbTab & Adodc1(0) & vbTab & Adodc1(1) & vbTab & Adodc1(2))
     End If
     If Flex.Row = 1 And Flex.text = "" Then
        If Flex.RowIsVisible(1) = False Then
          For I = 1 To 3
              Flex.Col = I
              Flex.text = Adodc1(I - 1)
          Next
        End If
     End If
     Adodc1.MoveNext
     If Adodc1.EOF Then Exit Do
  Loop
  Flex.Visible = True
  
  t.Visible = False
   Flex.Col = 4
   Flex.Row = 1
Adodc1.Close
End Sub

Private Sub flex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
     Flex.TextMatrix(Flex.Row, Flex.Col) = ""
End If
End Sub

Private Sub flex_KeyPress(KeyAscii As Integer)
If Flex.Col = 0 Or Flex.Col = 1 Or Flex.Col = 2 Or Flex.Col = 3 Then Exit Sub
If KeyAscii = 12 Then

End If
alinear
t.Visible = True
t.SetFocus
If KeyAscii <> 13 And KeyAscii <> 27 And KeyAscii <> 9 And KeyAscii <> 8 Then
   t.text = t.text & Chr(KeyAscii)
End If
t.SelStart = Len(t.text)
t.SelLength = 0
End Sub

Private Sub Form_Activate()
CargaFlex
Flex.SetFocus
End Sub

Private Sub Form_Load()
central Me
DTPicker1 = Date
t.ZOrder (0)
End Sub

Sub alinear()
t.Width = Flex.CellWidth
t.Left = Flex.CellLeft + Flex.Left + Frame1.Left
t.Top = Flex.CellTop + Flex.Top + Frame1.Top
t.Height = Flex.CellHeight
End Sub

Private Sub flex_RowColChange()
If Flex.Col = 0 And Flex.text <> "" Then
    Flex.Col = 4
End If
 If Flex.Col = 1 And Flex.text <> "" Then
    Condi = Flex.text: Flex.Col = 4
End If
 If Flex.Col = 2 And Flex.text <> "" Then
    Flex.Col = 4
End If
 If Flex.Col = 3 And Flex.text <> "" Then
    Flex.Col = 4
End If
 If Flex.Col = 4 And Flex.text <> "" Then
    Valor = Flex.text
End If


'If Flex.Row = 1 Or Flex.Row = 2 Then
'        If Flex.Col <> 1 Then
'             Flex.Col = 1
'        End If
'Else
'   If Flex.Row = 3 And Flex.Col = 1 Then
'      Exit Sub
'   End If
'   If Flex.TextMatrix(Flex.Row - 1, 1) <> "" And Flex.TextMatrix(Flex.Row - 1, 2) = "" And Flex.Row <> 3 Then
'        MsgBox "Ingrese el nombre del banco", vbExclamation, "Error"
'         Flex.Row = Flex.Row - 1
'         Flex.Col = 2
'   ElseIf Flex.TextMatrix(Flex.Row - 1, 2) <> "" And Flex.TextMatrix(Flex.Row - 1, 3) = "" Then
'        MsgBox "Ingrese el nro de Tarjeta o Cuenta", vbExclamation, "Error"
'         Flex.Row = Flex.Row - 1
'         Flex.Col = 3
'  ElseIf Flex.TextMatrix(Flex.Row - 1, 3) <> "" And Flex.TextMatrix(Flex.Row - 1, 4) = "" And Flex.Row > 5 Then
'        MsgBox "Ingrese el Tipo de Tarjeta ", vbExclamation, "Error"
'         Flex.Row = Flex.Row - 1
'         Flex.Col = 4
''   ElseIf flex.TextMatrix(flex.Row + 1, 3) <> "" And flex.TextMatrix(flex.Row + 1, 4) = "" And flex.Row = 5 Then
''        MsgBox "Ingrese el Tipo de Tarjeta ", vbExclamation, "Error"
''         flex.Row = flex.Row + 1
''         flex.Col = 4
'  ElseIf Flex.Row = 3 Or Flex.Row = 4 Then
'             If Flex.Col = 4 Then
'                  Flex.Col = 1
'             End If
'  End If
' End If
 
End Sub

Private Sub t_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case Is = 13
             If Flex.Col = 4 And Len(t) <> 0 Then
                 If Not IsNumeric(t) Then
                     MsgBox "Ingrese un dato válido ", vbExclamation, "Error"
                     t = ""
                     Exit Sub
                 Else
                   t = Format(t, "##0.00")
                 End If
             End If
             Flex.text = t.text
             t.Visible = False
             t.text = ""
             If Flex.Col = 4 And Flex.Row < Con Then
'                     If Flex.Row <> 6 Then
'                         Flex.Row = Flex.Row + 1
'                     End If
'                    Flex.Col = 1
'             ElseIf Flex.Col = 3 And (Flex.Row = 3 Or Flex.Col = 4) Then
                    Flex.Row = Flex.Row + 1
                    Flex.Col = 4
                    Flex.SetFocus
             Else
              CmdAceptar.SetFocus
             End If
      Case Is = 27
                  t.Visible = False
                  t.text = ""

   End Select
   If Flex.Col = 1 Then
      
   End If
End Sub



VERSION 5.00
Begin VB.Form formIngSerie 
   Caption         =   "Ingreso de Serie"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   1470
      TabIndex        =   2
      Top             =   3675
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3405
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3495
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   375
         MaxLength       =   20
         TabIndex        =   0
         Top             =   375
         Width           =   2775
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   360
         TabIndex        =   5
         Top             =   792
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   2760
      TabIndex        =   3
      Top             =   3675
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   195
      TabIndex        =   1
      Top             =   3675
      Width           =   1215
   End
End
Attribute VB_Name = "formIngSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adodc1 As ADODB.Recordset
Attribute adodc1.VB_VarHelpID = -1
Public almacen As String

'Procedure adodc1_willupdate()
'
'End
Private Sub Command1_Click()
frmTraIng.Flex1.TextMatrix(frmTraIng.Flex1.Row, 5) = Format(List1.ListCount, "###,##0.00")
  If Val(frmTraIng.Flex1.TextMatrix(frmTraIng.Flex1.Row, 5)) < List1.ListCount Then
    MsgBox "Se ha excedido en el ingreso : " & List1.ListCount - frmTraIng.nro_serie, vbInformation, "Aviso"
    Exit Sub
  End If
  If Val(frmTraIng.Flex1.TextMatrix(frmTraIng.Flex1.Row, 5)) > List1.ListCount Then
    MsgBox "Falta ingresar " & Val(frmTraIng.Flex1.TextMatrix(frmTraIng.Flex1.Row, 5)) - List1.ListCount, vbInformation, "Aviso"
    Exit Sub
  End If
  grabar
  List1.Clear
  Command2_Click
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Command3_Click()
  If List1.ListIndex <> -1 Then List1.RemoveItem List1.ListIndex
End Sub

Private Sub Form_Load()
  List1.Clear
  Set adodc1 = New ADODB.Recordset
  CargaLista (VGcod)
  'Adodc1.ActiveConnection = "Provider = Microsoft.Jet.OLEDB.3.51;Data Source= " & RUTA & NAMEBD & ""
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim codigo As String
Dim criterio As String
  If KeyAscii = 13 Then
    If Val(frmTraIng.Flex1.TextMatrix(frmTraIng.Flex1.Row, 4)) = 0 Then
       MsgBox "La Cantidad Recibida No puede ser 0 "
       Unload Me
       Exit Sub
    End If
    If Text1 <> "" Then
         If Val(frmTraIng.Flex1.TextMatrix(frmTraIng.Flex1.Row, 4)) <= List1.ListCount Then
            MsgBox "No puede ingresar mas series", vbInformation, "Aviso"
            Exit Sub
         End If
         If List1.ListCount <> 0 Then
            List1.ListIndex = 0
            While List1.ListIndex <> (List1.ListCount - 1)
              If List1.text = Trim(Text1) Then Exit Sub
              List1.ListIndex = List1.ListIndex + 1
            Wend
         End If
         If List1.text = Trim(Text1) Then Exit Sub
         criterio = "select * from stkseri  where STSSERIE= '" & Text1.text & "' and  STSCODIGO = '" & VGcod & " ' "
         If adodc1.State = 1 Then
         End If
         adodc1.Open criterio, VGcnx, adOpenStatic
         If adodc1.RecordCount = 0 Then
            List1.AddItem Trim(Text1)
            Enfoque Text1
         Else
            MsgBox "El codigo de las serie  existe", vbOKOnly + vbExclamation, "Aviso"
            Text1.SetFocus
         End If
         adodc1.Close
   Else
         SendKeys "{tab}"
         KeyAscii = 0
   End If
  End If
End Sub

Sub grabar()
Dim I As Integer
On Error GoTo Err
' ReDim arrayserie(List1.ListCount)'Este codigo es para ingreso automatico
' i = List1.ListIndex
' For i = 0 To List1.ListCount - 1
'   arrayserie(i) = List1.text
'   List1.ListIndex = i
' Next i
'
' Exit Sub

If Trim(almacen) = "" Then
   MsgBox "Seleccione primero el Almacen, para poder registrar las Series  ", vbInformation, "Aviso...!"
   Exit Sub
End If

'If List1.ListIndex = -1 Then Exit Sub
Dim criterio  As String
  Set adodc1 = New ADODB.Recordset
  With adodc1
     If List1.ListCount <> 0 Then List1.ListIndex = 0
     'RMM***********************************************************************08/08/2001
     'Vgcnx.Execute "delete from art_serie"
     VGcnx.Execute "delete from art_serie where acodigo='" & VGcod & "'"
     '***********************************************************************
     criterio = "select * from art_serie "
     adodc1.Open criterio, VGcnx, adOpenDynamic, adLockBatchOptimistic
     
     While List1.ListIndex <> (List1.ListCount - 1)
         .AddNew
         .Fields("alma") = almacen
         .Fields("acodigo") = VGcod
         .Fields("serie") = List1.text
         .UpdateBatch
          List1.ListIndex = List1.ListIndex + 1
     Wend
     If List1.text <> "" Then
        .AddNew
        .Fields("alma") = almacen
        .Fields("acodigo") = VGcod        ' form.text1  qwue tiene el codigo
        .Fields("serie") = List1.text
        .UpdateBatch
     End If
  End With
Exit Sub
Err:
  MsgBox Err.Description, vbInformation, "Aviso"
  
End Sub

Sub CargaLista(ByVal arCod As String)
Dim rs As New ADODB.Recordset
rs.Open "select * from art_serie where alma='" & almacen & "' and acodigo='" & arCod & "'", VGcnx, adOpenStatic, adLockBatchOptimistic
List1.Clear
Do While Not rs.EOF
   List1.AddItem rs!Serie
   rs.MoveNext
Loop
rs.Close
End Sub

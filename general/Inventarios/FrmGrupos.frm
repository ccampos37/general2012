VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmArGrupos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos de Articulos"
   ClientHeight    =   4800
   ClientLeft      =   1935
   ClientTop       =   2145
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleMode       =   0  'User
   ScaleWidth      =   7392.835
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   6855
      Begin VB.CommandButton command5 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   4200
         Picture         =   "FrmGrupos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command7 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5400
         Picture         =   "FrmGrupos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   3000
         Picture         =   "FrmGrupos.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   600
         Picture         =   "FrmGrupos.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   1800
         Picture         =   "FrmGrupos.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3960
         Picture         =   "FrmGrupos.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1800
         Picture         =   "FrmGrupos.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   775
      End
   End
End
Attribute VB_Name = "FrmArGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim resp As String
Dim Lin As String
Dim DesFam As String
Dim DesLin As String
Dim Fam As String
Dim nTra As Integer

Private Sub command5_Click()
Dim CTIME As String
If Data1.Recordset.RecordCount > 0 Then
  With FrmArFam
        CTIME = Format(Time, "hh:mm:ss")
        .CrystalReport1.WindowTitle = "Inv038 -- Control de Inventarios"
        .CrystalReport1.ReportFileName = cRutP & "inv038.Rpt"
        Call Ubi_Tab(.CrystalReport1)
        .CrystalReport1.WindowShowPrintBtn = True
        .CrystalReport1.WindowShowRefreshBtn = True
        .CrystalReport1.WindowShowSearchBtn = True
        .CrystalReport1.WindowShowPrintSetupBtn = True
        .CrystalReport1.formulas(0) = "Hora = '" & CTIME & "'"
        .CrystalReport1.formulas(1) = "Empresa = '" & Mid(VGNemp, 1, 20) & "'"
        .CrystalReport1.formulas(2) = "Familia = '" & Mid(DesFam, 1, 17) & "'"
        .CrystalReport1.formulas(3) = "Linea = '" & Mid(DesLin, 1, 17) & "'"
        .CrystalReport1.SelectionFormula = "{GRUPO.FAM_CODIGO} =  '" & Trim(Fam) & "' AND {GRUPO.LIN_CODIGO} =  '" & Trim(Lin) & "'"
        .CrystalReport1.WindowTop = 100
        .CrystalReport1.WindowLeft = 150
        .CrystalReport1.DiscardSavedData = True
        .CrystalReport1.Destination = crptToWindow
        If .CrystalReport1.Status <> 2 Then .CrystalReport1.Action = 1
  End With
 End If
End Sub

Private Sub Command9_Click()
DBGrid1.Visible = True
'Command19.Visible = True
Frame5.Visible = True
Frame2.Visible = True
Frame1.Visible = False
Frame3.Visible = False
DBGrid1.SetFocus
End Sub

Private Sub TxFiltro_Change()
'If Data1.Recordset.RecordCount > 0 Then
    If Trim(TxFiltro) <> "" Then
        Data1.Recordset.MoveFirst
        
        If CmbOrden.ListIndex = 0 Then
            Data1.Recordset.FindFirst "GRU_CODIGO like '" & Trim(UCase(TxFiltro)) & "*'"
        ElseIf CmbOrden.ListIndex = 1 Then
            Data1.Recordset.FindFirst "GRU_NOMBRE like '" & Trim(UCase(TxFiltro)) & "*'"
        End If
        If Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
        
    End If
'End If
End Sub

Private Sub CmbOrden_Click()             ' Ordenar por
Dim nCom As Integer

nCom = CmbOrden.ListIndex

Select Case nCom
Case 0
    Data1.RecordSource = "Select * from GRUPO where  FAM_CODIGO='" & Fam & "' AND LIN_CODIGO='" & Lin & "' order by GRU_CODIGO"
Case 1
    Data1.RecordSource = "Select * from GRUPO where  FAM_CODIGO='" & Fam & "' AND LIN_CODIGO='" & Lin & "' order by GRU_NOMBRE"
End Select
TxFiltro = ""
Data1.Refresh
If DBGrid1.Visible Then DBGrid1.SetFocus
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyBack Then
    If Len(TxFiltro) - 1 > 0 Then
        TxFiltro = Left(TxFiltro, Len(TxFiltro) - 1)
    Else
        TxFiltro = ""
    End If
    KeyAscii = 0
ElseIf KeyAscii <> 13 Then
        TxFiltro = TxFiltro & Chr(KeyAscii)
End If
End Sub

Private Sub Command1_Click()
resp = "S"
Limpiar
Text1.Enabled = True

DBGrid1.Visible = False
Frame2.Visible = False
Frame5.Visible = False
Frame3.Caption = "Ingreso de Grupos"

Frame1.Visible = True
Frame3.Visible = True
Text1.SetFocus
End Sub

Private Sub Command2_Click()
If Data1.Recordset.RecordCount > 0 Then
    Limpiar
    resp = "N"
    Frame3.Caption = "Modificación de Grupos"
    DBGrid1.Visible = False
    Frame2.Visible = False
    Frame5.Visible = False
    Frame1.Visible = True
    Frame3.Visible = True

    Text1.text = Data1.Recordset("GRU_CODIGO")
    Text1.Enabled = False
    
    If Not IsNull(Data1.Recordset("GRU_NOMBRE")) Then
          Text2.text = Data1.Recordset("GRU_NOMBRE")
    Else
          Text2.text = ""
    End If
   Text2.SetFocus
End If
End Sub

Private Sub Command3_Click()
On Error GoTo EliErr
Dim cSql1 As String
Dim CSQL2 As String, cSql3 As String
Dim cCodigo1 As String
Dim cSel1 As Recordset
Dim cCodigo As String

If Data1.Recordset.RecordCount > 0 Then
    If MsgBox("Seguro de Eliminar", vbQuestion + vbOKCancel, mensaje1) = vbOK Then
            nTra = 2
            cCodigo1 = Pos_Dato1(Data1.Recordset, "GRU_CODIGO")
            Data1.Refresh
            If cCodigo1 <> "" Then
                Data1.Recordset.FindFirst "GRU_CODIGO='" & cCodigo1 & "'"
            End If
    End If
    DBGrid1.Refresh
    
    If DBGrid1.Visible Then DBGrid1.SetFocus
Else
    MsgBox "No existe ningún registro para Eliminar", vbInformation, mensaje1
End If
Exit Sub
EliErr:
    MsgBox Err.Description
    'If nTra = 1 Then cConexCom.RollbackTrans
End Sub

Private Sub Command7_Click()
Unload Me
 End Sub

Private Sub Command8_Click()
On Error GoTo GrabErr
Dim cFam As String
If resp = "S" Then
  If Text1 = "" Then
         MsgBox "Ingrese Código de Grupo ", vbInformation, "Mensaje"
         Text1.SetFocus
        Exit Sub
  Else
       If Existe(1, Trim(Text1), "GRUPO", "GRU_CODIGO", False, Fam, "FAM_CODIGO", Lin, "LIN_CODIGO") Then
              MsgBox "El código de Grupo ya existe", vbInformation, "Mensaje"
              Text1.SetFocus
               Exit Sub
       End If
  End If
End If

If Text2 = "" Then
     MsgBox "Ingrese Descripción de Grupo", vbExclamation, "Aviso"
     Text2.SetFocus
     Exit Sub
End If

If resp = "S" Then
       Data1.Recordset.AddNew
Else
       Data1.Recordset.Edit
End If
Data1.Recordset("FAM_CODIGO") = Fam
Data1.Recordset("LIN_CODIGO") = Lin
Data1.Recordset("GRU_CODIGO") = Text1.text
If Not IsNull(Text2.text) Then
       Data1.Recordset("GRU_NOMBRE") = Text2.text
Else
        Data1.Recordset("GRU_NOMBRE") = " "
End If
Data1.Recordset.Update
Data1.Refresh
DBGrid1.Refresh
   
Data1.Recordset.FindFirst "GRU_CODIGO = '" & Text1.text & "'"
   
If resp = "S" Then
    Limpiar
    Text1.SetFocus
Else
     'Label1.Visible = True
     DBGrid1.Visible = True
     'Command19.Visible = True
      Frame5.Visible = True
      Frame2.Visible = True
      Frame1.Visible = False
      Frame3.Visible = False
      DBGrid1.SetFocus
End If
Exit Sub
GrabErr:
    MsgBox Err.Description
    'If nTra = 1 Then cConexCom.RollbackTrans
End Sub

Sub Limpiar()
Text1 = ""
Text2 = ""
End Sub

Private Sub Form_Activate()
TxFiltro = ""
'CmbOrden.ListIndex = 0
'If DBGrid1.Visible Then DBGrid1.SetFocus
End Sub

Private Sub Form_Load()
central Me
'Init_ControlDBGrid DBGrid1

'Fam = FrmArLineas.Fam
'DesFam = FrmArFam.Data1.Recordset("FAM_NOMBRE")
'Lin = FrmArLineas.Data1.Recordset("LIN_CODIGO")
'DesLin = FrmArLineas.Data1.Recordset("LIN_NOMBRE")

Me.Caption = "Grupos de la Linea :  " & Mid(DesLin, 1, 15) & Space(2) & "Familia : " & Mid(DesFam, 1, 15)
End Sub

Private Sub Text1_GotFocus()
Enfoque Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim cFam As String

If KeyAscii = 13 Then
    If Trim(Text1) <> "" Then
           If Existe(1, Trim(Text1), "GRUPO", "GRU_CODIGO", False, Fam, "FAM_CODIGO", Lin, "LIN_CODIGO") Then
                  MsgBox "El código de Grupo ya existe", vbInformation, "Mensaje"
                  Text1 = "": Text1.SetFocus
                    Exit Sub
            End If
    Else
            MsgBox "Ingrese código de Grupo", vbInformation, "Mensaje"
            Text1 = "": Text1.SetFocus
    End If
    Text2.SetFocus
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub Text2_GotFocus()
Enfoque Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Text2) = "" Then
       MsgBox "Ingrese Descripcion de Grupo", vbInformation, "Mensaje"
       Text2 = "": Text2.SetFocus
    End If
    Command8.SetFocus
End If
End Sub

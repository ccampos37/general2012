VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FrmArEquival 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unidades Equivalentes"
   ClientHeight    =   4725
   ClientLeft      =   3195
   ClientTop       =   2235
   ClientWidth     =   6360
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleMode       =   0  'User
   ScaleWidth      =   6423.283
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   5655
      Begin VB.CommandButton command5 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   3480
         Picture         =   "FrmArEquival.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command7 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4560
         Picture         =   "FrmArEquival.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   240
         Picture         =   "FrmArEquival.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2400
         Picture         =   "FrmArEquival.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   1320
         Picture         =   "FrmArEquival.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   5895
      Begin VB.TextBox TxFiltro 
         Height          =   300
         Left            =   960
         MaxLength       =   45
         TabIndex        =   11
         Text            =   "TxFiltro"
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmArEquival.frx":154A
         Left            =   4200
         List            =   "FrmArEquival.frx":1551
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar  :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2715
      Left            =   210
      TabIndex        =   1
      Top             =   750
      Visible         =   0   'False
      Width           =   5925
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "OOOOOO"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2160
         MaxLength       =   25
         TabIndex        =   2
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Factor :"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   1680
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid DBGrid1 
      Height          =   2565
      Left            =   210
      TabIndex        =   20
      Top             =   780
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   4524
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmArEquival"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim resp As String
Dim rsk As New ADODB.Recordset
Dim Uni As String
Dim pdato As String
Dim pdata As String * 20

Public Property Let bdato(pdata)
    pdato = pdata
End Property

Public Property Let bdata(pdata)
    pdata = pdata
End Property




Private Sub command5_Click()
Dim CTIME As String

If rsk.RecordCount > 0 Then
    CTIME = Format(Time, "hh:mm:ss")
    With FrmMntUnidMedida
            .CrystalReport1.WindowTitle = "Inv036 -- Sistema de Inventarios "
            .CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv036.Rpt"
            Call Ubi_Tab(.CrystalReport1)
            .CrystalReport1.formulas(0) = "Hora = '" & CTIME & "'"
            .CrystalReport1.formulas(1) = "Empresa = '" & Mid(VGParametros.RucEmpresa, 1, 20) & "'"
            .CrystalReport1.formulas(2) = "Familia = ''"     '& Mid(FrmArUnidades.Data1.Recordset("UM_NOMBRE"), 1, 17) & "'"
            .CrystalReport1.SelectionFormula = "{TABEQUI.EQUNIPRI} =  '" & Trim(Uni) & "'"
            .CrystalReport1.WindowShowPrintBtn = True
            .CrystalReport1.WindowShowRefreshBtn = True
            .CrystalReport1.WindowShowSearchBtn = True
            .CrystalReport1.WindowShowPrintSetupBtn = True
            .CrystalReport1.DiscardSavedData = True
            .CrystalReport1.Destination = crptToWindow
            If .CrystalReport1.Status <> 2 Then .CrystalReport1.Action = 1
    End With
 End If
End Sub

Private Sub Command9_Click()
DbGrid1.Visible = True
'Command19.Visible = True
Frame5.Visible = True
Frame2.Visible = True
Frame1.Visible = False
Frame3.Visible = False
DbGrid1.SetFocus
End Sub

Private Sub TxFiltro_Change()
Dim ncondi As String

'If Data1.Recordset.RecordCount > 0 Then
    If Trim(TxFiltro) <> "" Then
        If cmbOrden.ListIndex = 0 Then
            ncondi = "EQUNIEQUI like '" & Trim(UCase(TxFiltro)) & "%'"
            ncondi = "select * from tabequi where EQUNIPRI='" & pdato & "' and " & ncondi
        Else
            ncondi = "select * from tabequi where EQUNIPRI='" & pdato & "'"
        End If
    Else
        ncondi = "Select * from TABEQUI where EQUNIPRI='" & pdato & "' order by EQUNIEQUI"
    End If
    Call Listado(ncondi)
'End If
End Sub


Sub Listado(wcad)
  Set DbGrid1.DataSource = Nothing
  Set rsk = Nothing
  
  Set rsk = VGCNx.Execute(wcad)
  Set DbGrid1.DataSource = rsk
  With DbGrid1
      .Columns(0).Caption = "Codigo"
      .Columns(0).Width = 1000
      .Columns(1).Caption = "Descripcion"
      .Columns(1).Width = 3800
      .MarqueeStyle = dbgHighlightRow
      .Refresh
  End With

End Sub


Private Sub CmbOrden_Click()             ' Ordenar por
Dim nCom As Integer

nCom = cmbOrden.ListIndex

Select Case nCom
Case 0
    'Data1.RecordSource = "Select * from TABEQUI where EQUNIPRI='" & FrmArUnidades.Data1.Recordset("UM_ABREV") & "' order by EQUNIEQUI"
    Call Listado("Select * from TABEQUI where EQUNIPRI='" & pdato & "' order by EQUNIEQUI")
End Select
TxFiltro = ""
'Data1.Refresh
If DbGrid1.Visible Then DbGrid1.SetFocus
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
DbGrid1.Visible = False
Frame2.Visible = False
Frame5.Visible = False
Frame3.Caption = "Ingreso de Unidades Equivalentes"

Frame1.Visible = True
Frame3.Visible = True
Text1.SetFocus
End Sub

Private Sub Command2_Click()
If rsk.RecordCount > 0 Then
    Limpiar
    resp = "N"
    Frame3.Caption = "Modificación de Unidades Equivalentes"
    
    DbGrid1.Visible = False
    'Command19.Visible = False
    Frame2.Visible = False
    Frame5.Visible = False
    Frame1.Visible = True
    Frame3.Visible = True

    Text1.text = rsk.Fields("EQUNIEQUI")
    Text1.Enabled = False
    
    If Not IsNull(rsk.Fields("EQCANTEQUI")) Then
            Text2.text = rsk.Fields("EQCANTEQUI")
    Else
            Text2.text = "0"
    End If
    If Devolver_Dato(1, Text1, "TABUNIMED", "UM_ABREV", False, "UM_NOMBRE") <> "" Then
            Label1 = Devolver_Dato(1, Text1, "TABUNIMED", "UM_ABREV", False, "UM_NOMBRE")
    Else
            Label1 = ""
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

If rsk.RecordCount > 0 Then

    If MsgBox("Seguro de Eliminar", vbQuestion + vbOKCancel, mensaje1) = vbOK Then
             cCodigo1 = DbGrid1.Columns(1).text
             VGCNx.Execute "Delete From TABEQUI where EQUNIPRI='" & pdato & "' and EQUNIEQUI='" & cCodigo1 & "'"
    End If
    Call Listado("select * from TABEQUI where EQUNIPRI='" & pdato & "'")
Else
    MsgBox "No existe ningún registro para Eilminar", vbInformation, mensaje1
End If
Exit Sub
EliErr:
    MsgBox Err.Description
    'If nTra = 1 Then Vgcnx.RollbackTrans
End Sub

Private Sub Command7_Click()
Unload Me
 End Sub

Private Sub Command8_Click()
On Error GoTo GrabErr
Dim cUni As String

If resp = "S" Then
    If Text1 = "" Then
         MsgBox "Ingrese Código de Unidad Equivalente ", vbInformation, "Mensaje"
        Text1.SetFocus
        Exit Sub
    Else
        If Existe(1, Trim(Text1), "TABEQUI", "EQUNIEQUI", False, Uni, "EQUNIPRI") Then
            MsgBox "El código de Equivalencias ya existe", vbInformation, "Mensaje"
            Text1.SetFocus
            Exit Sub
        End If
    End If
End If

If Text2 = "" Then Text2 = "0"

If Text2 = "0" Then
     MsgBox "Ingrese Factor de Conversión", vbExclamation, "Aviso"
     Text2.SetFocus
     Exit Sub
End If
  
If resp = "S" Then
    VGCNx.Execute "INSERT INTO tabequi " & _
                      "(EQUNIPRI,EQUNIEQUI,EQCANTEQUI)" & _
                      " VALUES(" & _
                      "'" & pdato & "'," & _
                      "'" & Text1.text & "'," & _
                      Val(IIf(Len(Trim(Text2.text)) = 0, 0, Text2)) & ")"

    Call Listado("select * from TABEQUI where EQUNIPRI='" & pdato & "'")
    Limpiar
    Text1.SetFocus

Else
    VGCNx.Execute "UPDATE tabequi " & _
                      " SET EQUNIEQUI ='" & Text1.text & "'," & _
                      " EQCANTEQUI =" & Val(IIf(Len(Trim(Text2.text)) = 0, 0, Text2))

    Call Listado("select * from TABEQUI where EQUNIPRI='" & pdato & "'")
    
    DbGrid1.Visible = True
    Frame5.Visible = True
    Frame2.Visible = True
    Frame1.Visible = False
    Frame3.Visible = False
    DbGrid1.SetFocus

End If
Exit Sub
GrabErr:
    MsgBox Err.Description
    'If nTra = 1 Then Vgcnx.RollbackTrans
End Sub

Sub Limpiar()
Text1 = ""
Text2 = "0"
Label1 = ""
End Sub

Private Sub Form_Activate()
TxFiltro = ""
cmbOrden.ListIndex = 0
If DbGrid1.Visible Then DbGrid1.SetFocus
End Sub

Private Sub Form_Load()
central Me
'Init_ControlDBGrid DBGrid1

Me.Caption = "Equivalencias de la Unid. de Medida :  " & pdata
Call Listado("select * from TABEQUI where EQUNIPRI='" & pdato & "'")

End Sub

Private Sub Text1_GotFocus()
Enfoque Text1
End Sub

Private Sub Text1_DblClick()
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset

Adodc2.Open "SELECT UM_ABREV,UM_NOMBRE FROM TABUNIMED WHERE UM_ESTADO='A'", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "SELECT UM_ABREV,UM_NOMBRE FROM TABUNIMED WHERE UM_ESTADO='A'"
frmReferencia.Label1.Caption = "Unidades de Medida"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
    Text1.text = (vGUtil(1))
    Label1 = (Mid(vGUtil(2), 1, 15))
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Text1_DblClick
Else
    If KeyCode = 46 Then Label1 = ""
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim cFam As String
If KeyAscii = 13 Then
    If Trim(Text1) <> "" Then
    
       If Devolver_Dato(1, Text1, "TABUNIMED", "UM_ABREV", False, "UM_NOMBRE") <> "" Then
          Label1 = Devolver_Dato(1, Text1, "TABUNIMED", "UM_ABREV", False, "UM_NOMBRE")
       Else
          MsgBox "Unidad de Medida no existe", vbInformation, "Mensaje"
          Label1 = ""
          Text1.SetFocus: Exit Sub
       End If
        
       If Existe(1, Trim(Text1), "TABEQUI", "EQUNIEQUI", False, Uni, "EQUNIPRI") Then
          MsgBox "El código de Equivalencia ya existe", vbInformation, "Mensaje"
          Text1.SetFocus
          Exit Sub
       End If
    Else
          MsgBox "Ingrese Código de Equivalencia", vbInformation, "Mensaje"
          Text1.SetFocus: Exit Sub
    End If
    Text2.SetFocus
Else
    'If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub Text2_GotFocus()
Enfoque Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim I As Integer
If KeyAscii = 13 Then
    If Trim(Text2) = "" Then Text2 = "0"
    If Trim(Text2) = "0" Then
       MsgBox "Ingrese Factor de Conversión", vbInformation, "Mensaje"
       Text2.SetFocus: Exit Sub
    End If
    Command8.SetFocus
Else
  If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And Chr$(KeyAscii) <> "." And KeyAscii <> 8 Then
     KeyAscii = 0
  Else
     If Chr$(KeyAscii) = "." Then
        For I = 1 To Len(Text1)
            If Mid(Text1, I, 1) = "." Then KeyAscii = 0: Exit Sub
        Next
     End If
  End If
End If
End Sub

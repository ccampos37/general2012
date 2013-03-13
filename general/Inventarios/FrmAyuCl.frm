VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FrmAyuCliente 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6930
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "S&alir"
      CausesValidation=   0   'False
      Height          =   675
      Left            =   3675
      Picture         =   "FrmAyuCl.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3930
      Width           =   775
   End
   Begin VB.CommandButton CmdIng 
      Caption         =   "&Selec."
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   675
      Left            =   1980
      Picture         =   "FrmAyuCl.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3930
      Width           =   775
   End
   Begin VB.Frame Frame5 
      Height          =   630
      Left            =   150
      TabIndex        =   0
      Top             =   0
      Width           =   6600
      Begin VB.TextBox TxFiltro 
         Height          =   300
         Left            =   930
         TabIndex        =   2
         Text            =   "TxFiltro"
         Top             =   210
         Width           =   2655
      End
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmAyuCl.frx":0884
         Left            =   4365
         List            =   "FrmAyuCl.frx":088E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   195
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar :"
         Height          =   255
         Left            =   195
         TabIndex        =   4
         Top             =   255
         Width           =   870
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   255
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmAyuCl.frx":08AC
      Height          =   3195
      Left            =   165
      TabIndex        =   5
      Top             =   690
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5636
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ForeColor       =   -2147483635
      HeadLines       =   1
      RowHeight       =   17
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "CLIENTECODIGO"
         Caption         =   " CODIGO"
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
         DataField       =   "CLIENTERAZONSOCIAL"
         Caption         =   "              RAZON SOCIAL"
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
      BeginProperty Column02 
         DataField       =   "CLIENTERUC"
         Caption         =   "   R.U.C."
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
         MarqueeStyle    =   4
         ScrollBars      =   2
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   3344.882
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1395.213
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAyuCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCod As String, cNom As String
Public cDir As String, cRuc As String
Public nPago As String, nDes As String
Public nCli As Integer, nT As Integer
Public guiasTerceros As Integer
Dim adodc1 As ADODB.Recordset
Dim nCom As Integer, nCursor As Integer
Private Sub CmbOrden_Click()            ' Ordenar por
nCom = CmbOrden.ListIndex
Set adodc1 = New ADODB.Recordset
Select Case nCom
Case 0
    ' adodc1.Open "Select Clientecodigo as ccodcli,clienterazonsocial as cnomcli,clienteruc as cnumruc,clientedireccion as cdircli," & _
    '        "clientetipopersona as ctipvta,'0' as NPORDES FROM vt_cliente ORDER BY clientecodigo", VGCNx, adOpenDynamic, adLockOptimistic
    
    'adodc1.Open "SELECT B.CLIENTECODIGO,B.CLIENTERAZONSOCIAL,A.CLIENTERUC,CLIENTEDIRECCION FROM V_ALMACENYVENTAS A " _
    '& " INNER JOIN VT_CLIENTE B ON A.CLIENTERUC=B.CLIENTERUC WHERE A.PUNTOVTACODIGO='" & VGparametros.puntovta & "' AND A.EMPRESACODIGO='" & VGparametros.empresacodigo & "' group by B.CLIENTECODIGO,B.CLIENTERAZONSOCIAL,A.CLIENTERUC,CLIENTEDIRECCION ", VGCNx, adOpenDynamic, adLockOptimistic
    adodc1.Open "SELECT CLIENTECODIGO,CLIENTERAZONSOCIAL,CLIENTEruc,clientedireccion, clienteguiasterceros from VT_CLIENTE", VGCNx, adOpenDynamic, adLockOptimistic

Case 1
    'adodc1.Open "Select Clientecodigo as ccodcli,clienterazonsocial as cnomcli,clienteruc as cnumruc,clientedireccion as cdircli," & _
    '"clientetipopersona as ctipvta,'0' as NPORDES FROM vt_cliente ORDER BY clienterazonsocial", VGCNx, adOpenDynamic, adLockOptimistic

    adodc1.Open "Select Clientecodigo,clienterazonsocial,clienteruc,CLIENTEDIRECCION , clienteguiasterceros " & _
    " FROM vt_cliente ORDER BY clienterazonsocial", VGCNx, adOpenDynamic, adLockOptimistic

Case 2
    'Adodc1.Open "Select CCODCLI,CNOMCLI,CNUMRUC,CDIRCLI,CTIPVTA,NPORDES FROM MAECLI ORDER BY CNUMRUC", Vgcnx, adOpenStatic
    adodc1.Open "Select Clientecodigo as ccodcli,clienterazonsocial as cnomcli,clienteruc as cnumruc,clientedireccion as cdircli, clienteguiasterceros, " & _
            "clientetipopersona as ctipvta,'0' as NPORDES FROM vt_cliente ORDER BY clienteruc", VGCNx, adOpenDynamic, adLockOptimistic
End Select
Set DataGrid1.DataSource = adodc1
DataGrid1.SetFocus
TxFiltro = ""
End Sub

Private Sub CmdIng_Click()
If adodc1.RecordCount <> 0 Then
    cCod = adodc1("CLIENTECODIGO")
    cNom = IIf(IsNull(adodc1("CLIENTERAZONSOCIAL")), "", adodc1("CLIENTERAZONSOCIAL"))
    cRuc = IIf(IsNull(adodc1("CLIENTERUC")), "", adodc1("CLIENTERUC"))
'    cRuc = IIf(IsNull(adodc1("CLIENTEcodigo")), "", adodc1("CLIENTEcodigo"))
    cDir = IIf(IsNull(adodc1("CLIENTEDIRECCION")), "", adodc1("CLIENTEDIRECCION"))
    guiasTerceros = IIf(IsNull(adodc1("clienteguiasTerceros")), 0, adodc1("CLIENTEguiasterceros"))
    nCli = 1
    CmdSalir_Click
End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DataGrid1_DblClick()
CmdIng_Click
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

Private Sub Form_Activate()
CmbOrden.ListIndex = 1
End Sub

Private Sub Form_Load()
AlinearAyuda Me              ' Centrar Formulario
Init_ControlDataGrid DataGrid1
Set adodc1 = New ADODB.Recordset

'adodc1.Open "SELECT B.CLIENTECODIGO,B.CLIENTERAZONSOCIAL,A.CLIENTERUC,CLIENTEDIRECCION FROM V_ALMACENYVENTAS A " _
'& " INNER JOIN VT_CLIENTE B ON A.CLIENTERUC=B.CLIENTERUC WHERE A.PUNTOVTACODIGO='" & VGparametros.puntovta & "' AND A.EMPRESACODIGO='" & VGparametros.empresacodigo & "' group by B.CLIENTECODIGO,B.CLIENTERAZONSOCIAL,A.CLIENTERUC,CLIENTEDIRECCION ", VGCNx, adOpenDynamic, adLockOptimistic
Set adodc1 = VGCNx.Execute("SELECT CLIENTECODIGO,CLIENTERAZONSOCIAL,clienteruc,CLIENTEDIRECCION ,clienteGuiasTerceros from VT_CLIENTE")

'If adodc1.RecordCount > 0 Then
    Set DataGrid1.DataSource = adodc1
    DataGrid1.Refresh
    nCli = 0: TxFiltro = "": cCod = "": cNom = "": cRuc = "": cDir = "": guiasTerceros = 0
'Else
'    MsgBox "gdfgdf"
'    Unload Me
'End If

End Sub

Private Sub TxFiltro_Change()
Dim rrsql As New ADODB.Recordset
    If Trim(TxFiltro) <> "" Then
    SQL = "SELECT CLIENTECODIGO,CLIENTERAZONSOCIAL,clienteruc,CLIENTEDIRECCION ,clienteGuiasTerceros from VT_CLIENTE"
    If CmbOrden.ListIndex = 0 Then
          Set adodc1 = VGCNx.Execute(SQL & " where clientecodigo LIKE '" & Trim(UCase(TxFiltro)) & "%'")
        ElseIf CmbOrden.ListIndex = 1 Then
          Set adodc1 = VGCNx.Execute(SQL & " where Clienterazonsocial LIKE ('%" & Trim(UCase(TxFiltro)) & "%')")
        ElseIf CmbOrden.ListIndex = 2 Then
          Set adodc1 = VGCNx.Execute(SQL & " where CNUMRUC LIKE '" & Trim(UCase(TxFiltro)) & "%'")
        End If
    End If
   Set DataGrid1.DataSource = adodc1
    DataGrid1.Refresh
End Sub

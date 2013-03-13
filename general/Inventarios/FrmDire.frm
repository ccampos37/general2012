VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmDire 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direcciones del Cliente"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6150
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmDire.frx":0000
      Height          =   1365
      Left            =   345
      TabIndex        =   0
      Top             =   195
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2408
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "CCODCLI"
         Caption         =   "  CODIGO"
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
         DataField       =   "CDIRCLI"
         Caption         =   "         DIRECCION DEL CLIENTE"
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
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   3764.977
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1440
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox Txdirec 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   795
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         Height          =   270
         Left            =   360
         TabIndex        =   8
         Top             =   405
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Dirección :"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   825
         Width           =   1575
      End
      Begin VB.Label LbCliente 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   420
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdSalir2 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   3330
      Picture         =   "FrmDire.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1710
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   675
      Left            =   1890
      Picture         =   "FrmDire.frx":0457
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1710
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton CmdIng 
      Caption         =   "&Ingreso"
      Height          =   675
      Left            =   1050
      Picture         =   "FrmDire.frx":0899
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1710
      Width           =   775
   End
   Begin VB.CommandButton CmdEli 
      Caption         =   "&Eliminar"
      Height          =   675
      Left            =   2610
      Picture         =   "FrmDire.frx":0CDB
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1710
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   4170
      Picture         =   "FrmDire.frx":111D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1710
      Width           =   775
   End
End
Attribute VB_Name = "FrmDire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cCliente As String
Dim adodc1 As ADODB.Recordset
Dim cSql1 As String
Dim cRec As ADODB.Recordset
Dim cCod As String
Dim cDir As String
Dim nTra As Integer

Private Sub CmdEli_Click()
On Error GoTo EliErr
If adodc1.RecordCount > 0 Then
    cCod = adodc1("CCODCLI")
    cDir = adodc1("CDIRCLI")
    If MsgBox("Seguro de Eliminar la Dirección", vbQuestion + vbOKCancel, "Mensaje") = vbOK Then
        cSql1 = "Delete From DIRE_CLIENTE Where CCODCLI = '" & Trim(cCod) & "' and CDIRCLI = '" & Trim(cDir) & "'"
        nTra = 1
        VGcnx.BeginTrans
        VGcnx.Execute cSql1
        VGcnx.CommitTrans
        nTra = 0
        adodc1.Requery
        DataGrid1.Refresh
    End If
End If
Exit Sub
EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGcnx.RollbackTrans
End Sub

Private Sub CmdGrabar_Click()
On Error GoTo GrabErr

cSql1 = "Select * from DIRE_CLIENTE Where CCODCLI = '" & Trim(LbCliente) & "' and CDIRCLI = '" & Trim(Txdirec) & "'"
Set cRec = New ADODB.Recordset
cRec.Open cSql1, VGcnx, adOpenStatic, adLockOptimistic
If cRec.RecordCount > 0 Then
    MsgBox "La dirección se encuentra registrada", vbInformation, "Mensaje"
    cRec.Close: Txdirec.SetFocus
    Exit Sub
End If
cRec.Close
cSql1 = "Insert Into DIRE_CLIENTE (CCODCLI,CDIRCLI) Values ('" & Trim(LbCliente) & "','" & Trim(Txdirec) & "')"
nTra = 1
VGcnx.BeginTrans
VGcnx.Execute cSql1
VGcnx.CommitTrans
nTra = 0
adodc1.Requery
CmdSalir2_Click
Exit Sub
GrabErr:
    MsgBox Err.Description
    If nTra = 1 Then VGcnx.RollbackTrans
End Sub

Private Sub CmdIng_Click()
OculObj2 (False)
OculObj (True)
LbCliente = cCliente
Txdirec = ""
If Txdirec.Enabled And Txdirec.Visible Then
    Txdirec.SetFocus
End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSalir2_Click()
OculObj (False)
OculObj2 (True)
DataGrid1.SetFocus
End Sub

Private Sub Form_Load()
central Me             ' Centrar Formulario
Init_ControlDataGrid DataGrid1
Set adodc1 = New ADODB.Recordset
adodc1.Open "Select * from DIRE_CLIENTE where CCODCLI = '" & cCliente & "'", VGcnx, adOpenStatic
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh
Me.Caption = Caption & "  " & "-" & "  " & cCliente
OculObj (False)
End Sub

Private Sub OculObj(nTipo As Boolean)  ' Cliente y dirección
Frame1.Visible = nTipo
CmdGrabar.Visible = nTipo
CmdSalir2.Visible = nTipo
End Sub

Private Sub OculObj2(nTipo As Boolean)
DataGrid1.Visible = nTipo
CmdIng.Visible = nTipo
CmdEli.Visible = nTipo
CmdSalir.Visible = nTipo
End Sub

Private Sub Txdirec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

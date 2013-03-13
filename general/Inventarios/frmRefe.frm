VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmref 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Referencias"
   ClientHeight    =   4905
   ClientLeft      =   2055
   ClientTop       =   2400
   ClientWidth     =   4845
   Icon            =   "frmRefe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   420
      Left            =   975
      TabIndex        =   4
      Top             =   4020
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   420
      Left            =   2535
      TabIndex        =   3
      Top             =   4020
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   150
      MaxLength       =   25
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3135
      Left            =   150
      TabIndex        =   1
      Top             =   750
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5530
      _Version        =   393216
      HeadLines       =   2
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enterprise Solutions S.A."
      ForeColor       =   &H80000010&
      Height          =   195
      Index           =   3
      Left            =   2640
      TabIndex        =   6
      Top             =   4635
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enterprise Solutions S.A."
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   2
      Left            =   2625
      TabIndex        =   5
      Top             =   4620
      Width           =   1740
   End
   Begin VB.Label Label2 
      Caption         =   "Buscar:"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   90
      Width           =   1095
   End
End
Attribute VB_Name = "frmref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Campos(1 To 3) As String
'Dim Conexion As String
Dim csql As String
Dim Adodc3 As New ADODB.Recordset

Private Sub cmdCancel_Click()
 vGUtil(1) = ""
 vGUtil(2) = ""
 Unload Me
End Sub

Private Sub cmdOK_Click()
 If Adodc3.RecordCount <> 0 Then
  If Not IsNull(DataGrid1.Bookmark) Then
   Adodc3.Bookmark = DataGrid1.Bookmark
  End If
  vGUtil(1) = Adodc3.Fields(0)
  vGUtil(2) = Adodc3.Fields(1)
 Else
  vGUtil(1) = ""
  vGUtil(2) = ""
 End If
 Unload Me
End Sub


Private Sub DataGrid1_DblClick()
 cmdOK_Click
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo Err1
    Adodc3.Sort = DataGrid1.Columns(ColIndex).DataField
    DataGrid1.Tag = DataGrid1.Columns(ColIndex).DataField
    'Set DataGrid1.DataSource = Adodc3
    DataGrid1.Refresh
    Text1 = ""
    With DataGrid1
        .Columns(0).Caption = "Código"
        .Columns(0).Width = 1000
        .Columns(1).Caption = "Descripción"
        .Columns(1).Width = 3000
    End With
    Exit Sub
Err1:
    Resume Next
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then
 If Len(Text1) - 1 > 0 Then
  Text1 = Left(Text1, Len(Text1) - 1)
 Else
  Text1 = ""
 End If
 KeyAscii = 0
ElseIf KeyAscii <> 13 Then
 Text1 = Text1 & Chr(KeyAscii)
End If
End Sub

Private Sub Form_Activate()
DataGrid1.SetFocus
End Sub

Private Sub Form_Load()
 Init_ControlDataGrid DataGrid1
 Set DataGrid1.DataSource = Adodc3
 With DataGrid1
        .Columns(0).Caption = "Código"
        .Columns(0).Width = 1000
        .Columns(1).Caption = "Descripción"
        .Columns(1).Width = 3000
 End With
 DataGrid1_HeadClick 0
End Sub

Public Sub Conectar(AD As ADODB.Recordset, Optional ByVal Sq As String, Optional strCampo1 As String = "*", Optional strCampo2 As String = "*")
    Set Adodc3 = AD
End Sub

Private Sub Label2_Click(Index As Integer)
    MsgBox "Producto desarrollado por Enterprise Solutions S.A. para uso exclusivo de sus clientes, de acuerdo a un contrato establecido entre ambos", vbInformation
End Sub

Private Sub Text1_Change()
    On Error Resume Next
 Dim C As String
 Dim Ant As Integer
 If Adodc3.RecordCount <> 0 Then
  Ant = Adodc3.Bookmark
  Adodc3.AbsolutePosition = 1
  If Adodc3.Fields(DataGrid1.Tag).Type = adDate Then Exit Sub
  If Text1 <> "" Then
    C = Adodc3.Fields(DataGrid1.Tag).name & " LIKE " & "'" & UCase(Trim(Text1)) & "*'"
    Adodc3.Find C
    If Adodc3.EOF Then
    Adodc3.AbsolutePosition = Ant
   End If
  End If
 End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case Chr(KeyAscii)
        Case "'", ",": KeyAscii = 0
    End Select
End Sub

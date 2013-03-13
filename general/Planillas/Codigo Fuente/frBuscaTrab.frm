VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frmBuscaTrab 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3975
   ClientLeft      =   2025
   ClientTop       =   2085
   ClientWidth     =   5400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5400
   Begin AplisetControlText.Aplitext Text2 
      Height          =   285
      Left            =   1605
      TabIndex        =   1
      Top             =   0
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   503
      Text            =   ""
   End
   Begin AplisetControlText.Aplitext Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3660
      Left            =   15
      TabIndex        =   2
      Top             =   300
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   6456
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3750.236
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBuscaTrab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CAMPOS(1 To 3) As String
Dim Conexion As String
Dim CSQL As String
Dim ADODC3 As ADODB.Recordset

Private Sub CMDCANCEL_Click()
 vgUtil(1) = ""
 vgUtil(2) = ""
 Unload Me
End Sub

Private Sub CMDOK_Click()
 If ADODC3.RecordCount <> 0 Then
  If Not IsNull(DataGrid1.Bookmark) Then
   ADODC3.Bookmark = DataGrid1.Bookmark
  End If
  vgUtil(1) = ADODC3.Fields(0)
  vgUtil(2) = ADODC3.Fields(1)
 Else
  vgUtil(1) = ""
  vgUtil(2) = ""
 End If
 Unload Me
End Sub

Private Sub COMBO1_Click(ByVal INDICE As Byte)
    On Error GoTo ERR1
    If ADODC3.State = 1 Then
        ADODC3.Close
    End If
    Select Case INDICE
    Case 0:
        ADODC3.Open CSQL & " ORDER BY " & CAMPOS(1), Conexion, adOpenDynamic, adLockOptimistic
        Set DataGrid1.DataSource = ADODC3
        DataGrid1.Refresh
        Text1 = ""
    Case 1:
        ADODC3.Open CSQL & " ORDER BY " & CAMPOS(2), Conexion, adOpenDynamic, adLockOptimistic
        Set DataGrid1.DataSource = ADODC3
        DataGrid1.Refresh
        Text1 = ""
    End Select
    With DataGrid1
        .Columns(0).DataField = ADODC3.Fields(0).Name
        .Columns(1).DataField = ADODC3.Fields(1).Name
        .Columns(0).Caption = "CÓDIGO"
        '.COLUMNS(0).WIDTH = 1000
        .Columns(1).Caption = "DESCRIPCIÓN"
        '.COLUMNS(1).WIDTH = 3000
    End With
    Exit Sub
ERR1:
    Resume Next
End Sub

Private Sub TEXT2_CHANGE()
 Dim C As String
 Dim ANT As Integer
 If ADODC3.RecordCount <> 0 Then
  ANT = ADODC3.Bookmark
  ADODC3.AbsolutePosition = 1
  If ADODC3.Fields(1).Type = adDate Then Exit Sub
  If Text2.Text <> "" Then
    C = ADODC3.Fields(1).Name & " LIKE " & "'" & UCase(Trim(Text2.Text)) & "*'"
   ADODC3.FIND C
   If ADODC3.EOF Then
    ADODC3.AbsolutePosition = ANT
   End If
  End If
 End If
End Sub

Private Sub TEXT2_GOTFOCUS()
    COMBO1_Click (1)
End Sub

Private Sub DATAGRID1_DblClick()
 CMDOK_Click
End Sub

Private Sub DATAGRID1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CMDOK_Click
    End If
    If KeyAscii = 27 Then
        CMDCANCEL_Click
    End If
If KeyAscii = vbKeyBack Then
 If Len(Text1) - 1 > 0 Then
  Text1 = Left(Text1, Len(Text1) - 1)
 Else
  Text1 = ""
 End If
 KeyAscii = 0
ElseIf KeyAscii <> 13 Then
 Text1.Text = Text1.Text & Chr(KeyAscii)
End If
End Sub

Private Sub FORM_ACTIVATE()
DataGrid1.SetFocus
End Sub

Private Sub FORM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CMDOK_Click
    End If
    If KeyAscii = 27 Then
        CMDCANCEL_Click
    End If
End Sub

Private Sub FORM_Load()
 'ME.LEFT = FORMPRINCIPAL.SCALEWIDTH - ME.WIDTH
 'ME.TOP = FORMPRINCIPAL.HEIGHT - FORMPRINCIPAL.SCALEHEIGHT
 'ALINEARAYUDA ME
 'INIT_CONTROLDATAGRID DATAGRID1
 Set DataGrid1.DataSource = ADODC3
 With DataGrid1
        .Columns(0).Caption = "CÓDIGO"
        '.COLUMNS(0).WIDTH = 1000
        .Columns(1).Caption = "DESCRIPCIÓN"
        '.COLUMNS(1).WIDTH = 3000
 End With
End Sub

Public Sub CONECTAR(AD As ADODB.Recordset, Optional ByVal SQ As String, Optional STRCAMPO1 As String = "*", Optional STRCAMPO2 As String = "*")
 If IsMissing(SQ) Or SQ = "" Then
    If InStr(AD.Source, "SELECT") = 0 Then 'SI ES SOLO UNA TABLA
        SQ = "SELECT * FROM " & AD.Source
    Else
        SQ = UCase(AD.Source)
        If InStr(SQ, "ORDER BY") > 0 Then
            SQ = Left(SQ, InStr(SQ, "ORDER BY") - 1)
        End If
    End If
 End If
 SQ = UCase(AD.Source)
 If InStr(SQ, "ORDER BY") > 0 Then
    SQ = Left(SQ, InStr(SQ, "ORDER BY") - 1)
 End If
Set ADODC3 = AD
 Conexion = ADODC3.ActiveConnection
 If STRCAMPO1 = "*" Then CAMPOS(1) = Trim(ADODC3.Fields(0).Name) Else CAMPOS(1) = STRCAMPO1
 If STRCAMPO2 = "*" Then CAMPOS(2) = Trim(ADODC3.Fields(1).Name) Else CAMPOS(2) = STRCAMPO2
 CSQL = SQ
 COMBO1_Click (0)
End Sub

Private Sub TEXT1_CHANGE()
 Dim C As String
 Dim ANT As Integer
 If ADODC3.RecordCount <> 0 Then
  ANT = ADODC3.Bookmark
  ADODC3.AbsolutePosition = 1
  If ADODC3.Fields(0).Type = adDate Then Exit Sub
  If Text1.Text <> "" Then
    C = ADODC3.Fields(0).Name & " LIKE " & "'" & UCase(Trim(Text1.Text)) & "*'"
   ADODC3.FIND C
   If ADODC3.EOF Then
    ADODC3.AbsolutePosition = ANT
   End If
  End If
 End If
End Sub

Private Sub Text1_GotFocus()
    COMBO1_Click (0)
End Sub


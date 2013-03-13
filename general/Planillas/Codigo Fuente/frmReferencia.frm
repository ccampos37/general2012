VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmComun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Referencias"
   ClientHeight    =   4905
   ClientLeft      =   2055
   ClientTop       =   2400
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4845
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   420
      Left            =   975
      TabIndex        =   6
      Top             =   4020
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   420
      Left            =   2535
      TabIndex        =   5
      Top             =   4020
      Width           =   1320
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmReferencia.frx":0000
      Left            =   2835
      List            =   "frmReferencia.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   345
      Width           =   1890
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   150
      MaxLength       =   25
      TabIndex        =   0
      Top             =   360
      Width           =   2535
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
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4410
      Picture         =   "frmReferencia.frx":0023
      Top             =   4560
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Criterio:"
      Height          =   255
      Index           =   1
      Left            =   2835
      TabIndex        =   4
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Buscar:"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   90
      Width           =   1095
   End
End
Attribute VB_Name = "frmComun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CAMPOS(1 To 3) As String
'Dim Conexion As String
Dim CSQL As String
Dim ADODC3 As ADODB.Recordset
Dim flCarga As Boolean

Private Sub CMDCANCEL_Click()
 VGUTIL(1) = ""
 VGUTIL(2) = ""
 Unload Me
End Sub

Private Sub CMDOK_Click()
If C = 0 Then
 If ADODC3.RecordCount <> 0 Then
  If Not IsNull(DataGrid1.Bookmark) Then
   ADODC3.Bookmark = DataGrid1.Bookmark
  End If
  VGUTIL(1) = ADODC3.Fields(0)
  VGUTIL(2) = ADODC3.Fields(1)
 Else
  VGUTIL(1) = ""
  VGUTIL(2) = ""
 End If
 Unload Me
Else
 Dim RSDETALLE As ADODB.Recordset
    Set RSDETALLE = New ADODB.Recordset
    'Obtener el concepto para el detalle de adelanto
        DBAUXCOM.Execute "INSERT INTO ##TMPDETALLE(NOMBRE,TIPO) VALUES('" & ADODC3!Codigo & "'," & ADODC3!TIPO & ")" & ""
        RSDETALLE.Open "SELECT NOMBRE FROM ##TMPDETALLE", DBAUXCOM, adOpenKeyset, adLockOptimistic
    Set frAutoAd.DtgDetalle.DataSource = RSDETALLE
End If
End Sub

Private Sub Combo1_Click()
    If Not flCarga Then Exit Sub
   On Error GoTo Err1
    ADODC3.Sort = CAMPOS(Combo1.ListIndex + 1)
'    Select Case Combo1.ListIndex
'    Case 0:
'        ADODC3.Open ADODC3.Source & " ORDER BY " & CAMPOS(1), Conexion, adOpenDynamic, adLockOptimistic
'    Case 1:
'        ADODC3.Open ADODC3.Source & " ORDER BY " & CAMPOS(2), Conexion, adOpenDynamic, adLockOptimistic
'    End Select
    Set DataGrid1.DataSource = ADODC3
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

Private Sub DATAGRID1_DblClick()
 CMDOK_Click
End Sub

Private Sub DataGrid1_HeadClick(ByVal COLINDEX As Integer)
    Combo1.ListIndex = COLINDEX
End Sub

Private Sub DATAGRID1_KeyPress(KeyAscii As Integer)
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
 'Me.Left = FormPrincipal.ScaleWidth - Me.Width
 'Me.Top = FormPrincipal.Height - FormPrincipal.ScaleHeight
 'AlinearAyuda Me
    flCarga = False
    Init_ControlDataGrid DataGrid1
    Set DataGrid1.DataSource = ADODC3
    With DataGrid1
           .Columns(0).Caption = "Código"
           .Columns(0).Width = 1000
           .Columns(1).Caption = "Descripción"
           .Columns(1).Width = 3000
    End With
    Combo1.ListIndex = 0
    flCarga = True
End Sub

Public Sub CONECTAR(AD As ADODB.Recordset, Optional ByVal SQ As String, Optional STRCAMPO1 As String = "*", Optional STRCAMPO2 As String = "*")
     If IsMissing(SQ) Or SQ = "" Then
        If InStr(AD.Source, "SELECT") = 0 Then 'Si es solo una tabla
            SQ = "SELECT * FROM " & AD.Source
        Else
            SQ = UCase(AD.Source)
            If InStr(SQ, "ORDER BY") > 0 Then
                SQ = Left(SQ, InStr(SQ, "ORDER BY") - 1)
            End If
        End If
     End If
     SQ = UCase(SQ)
     If InStr(SQ, "ORDER BY") > 0 Then
        SQ = Left(SQ, InStr(SQ, "ORDER BY") - 1)
     End If
     Set ADODC3 = New ADODB.Recordset
     'ADODC3.CursorLocation = adUseServer
     Set ADODC3 = AD
'     Conexion = Adodc3.ActiveConnection
     If STRCAMPO1 = "*" Then CAMPOS(1) = Trim(ADODC3.Fields(0).Name) Else CAMPOS(1) = STRCAMPO1
     If STRCAMPO2 = "*" Then CAMPOS(2) = Trim(ADODC3.Fields(1).Name) Else CAMPOS(2) = STRCAMPO2
     CSQL = SQ
End Sub
Private Sub TEXT1_CHANGE()
On Error GoTo handler
 Dim C As String
 Dim ANT As Integer
 If ADODC3.RecordCount <> 0 Then
  ANT = ADODC3.Bookmark
  ADODC3.AbsolutePosition = 1
  If ADODC3.Fields(Combo1.ListIndex).Type = adDate Then Exit Sub
  If Text1 <> "" Then
    C = ADODC3.Fields(Combo1.ListIndex).Name & " LIKE " & "'" & UCase(Trim(Text1)) & "*'"
    ADODC3.FIND C
    If ADODC3.EOF Then
    ADODC3.AbsolutePosition = ANT
   End If
  End If
 End If
 
 Exit Sub
handler:
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case Chr(KeyAscii)
        Case "'", ",": KeyAscii = 0
    End Select
End Sub

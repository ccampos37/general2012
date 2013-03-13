VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmReferencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Referencias"
   ClientHeight    =   5376
   ClientLeft      =   2052
   ClientTop       =   2400
   ClientWidth     =   6888
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5376
   ScaleWidth      =   6888
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   675
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4470
      Width           =   775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4470
      Width           =   775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1050
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   1080
      Width           =   4725
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2925
      Left            =   120
      TabIndex        =   1
      Top             =   1470
      Width           =   6615
      _ExtentX        =   11663
      _ExtentY        =   5165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
            ColumnWidth     =   1272.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3539.906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Criterio:"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   5
      Top             =   690
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Buscar:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   450
      TabIndex        =   3
      Top             =   120
      Width           =   5880
   End
End
Attribute VB_Name = "frmReferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Campos(1 To 3) As String
Dim Conexion As String
Dim csql As String
Dim Adodc3 As ADODB.Recordset

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
  vGUtil(1) = IIf(IsNull(Adodc3.Fields(0)), "", Adodc3.Fields(0))
  vGUtil(2) = IIf(IsNull(Adodc3.Fields(1)), "", Adodc3.Fields(1))
 Else
  vGUtil(1) = ""
  vGUtil(2) = ""
 End If
 Unload Me
End Sub

Private Sub Combo1_Click()

If Adodc3.State = 1 Then
 Adodc3.Close
End If

Select Case Combo1.ListIndex
 Case 0:
        Adodc3.Open csql & " ORDER BY " & Campos(1), Conexion, adOpenDynamic, adLockOptimistic
        Set DataGrid1.DataSource = Adodc3
        DataGrid1.Refresh
        Text1 = ""
 Case 1:
        Adodc3.Open csql & " ORDER BY " & Campos(2), Conexion, adOpenDynamic, adLockOptimistic
        Set DataGrid1.DataSource = Adodc3
        DataGrid1.Refresh
        Text1 = ""
 Case 2:
        Adodc3.Open csql & " ORDER BY " & Campos(4), Conexion, adOpenDynamic, adLockOptimistic
        Set DataGrid1.DataSource = Adodc3
        DataGrid1.Refresh
        Text1 = ""
End Select
With DataGrid1
        .Columns(0).Caption = "Código"
        .Columns(0).Width = 1500
        .Columns(1).Caption = "Descripción"
        .Columns(1).Width = 3000
 End With
End Sub

Private Sub DataGrid1_DblClick()
 cmdOK_Click
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
 'Me.Left = FrmPrincipal.ScaleWidth - Me.Width
 'Me.Top = FrmPrincipal.Height - FrmPrincipal.ScaleHeight
 'AlinearAyuda Me
' Init_ControlDataGrid DataGrid1
' Set DataGrid1.DataSource = Adodc3
' With DataGrid1
'        .Columns(0).Caption = "Código"
'        .Columns(0).Width = 1500
'        .Columns(1).Caption = "Descripción"
'        .Columns(1).Width = 3000
' End With
'Combo1.ListIndex = 0
End Sub

Public Sub Conectar(AD As ADODB.Recordset, Sq As String)
 Set Adodc3 = AD
 Conexion = Adodc3.ActiveConnection
 Campos(1) = Trim(Adodc3.Fields(0).Name)
 Campos(2) = Trim(Adodc3.Fields(1).Name)
 csql = Sq
End Sub

Private Sub Text1_Change()
 Dim C As String
 Dim Ant As Integer
 If Adodc3.RecordCount <> 0 Then
  Ant = Adodc3.Bookmark
  Adodc3.AbsolutePosition = 1
  If Text1 <> "" Then
   If Combo1.ListIndex = 0 Then
    C = Campos(1) & " LIKE " & "'" & UCase(Trim(Text1)) & "*'"
   ElseIf Combo1.ListIndex = 2 Then
    C = Campos(4) & " LIKE " & "'" & UCase(Trim(Text1)) & "*'"
   Else
    C = Campos(2) & " LIKE " & "'" & UCase(Trim(Text1)) & "*'"
   End If
   If Trim(Text1) = "" Then Exit Sub
   Adodc3.Find C
   If Adodc3.EOF Then
    Adodc3.AbsolutePosition = Ant
   End If
  End If
 End If
End Sub

Sub Inicio()
    Set DataGrid1.DataSource = Adodc3
    
        With DataGrid1
        If Campos(1) = "CNUMRUC" Then
            CmbOrden.List(0) = "RUC"
            .Columns(0).Caption = "RUC"
        Else
            .Columns(0).Caption = "Código"
        End If
        'Asigna etiquetas al DBGrid
        .Columns(0).Width = 1000
        .Columns(1).Caption = "Descripción"
        .Columns(1).Width = 3000
    End With
    If Combo1.ListIndex = 0 Then
        Combo1_Click
    Else
'        Combo1.ListIndex = 0
    End If
End Sub


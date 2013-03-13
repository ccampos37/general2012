VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmComun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmComun.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   4755
   Begin VB.Frame Frame1 
      Caption         =   "Búsqueda de Registros"
      Height          =   930
      Left            =   150
      TabIndex        =   3
      Top             =   75
      Width           =   4500
      Begin VB.ComboBox xCampo 
         Height          =   315
         ItemData        =   "frmComun.frx":0442
         Left            =   2460
         List            =   "frmComun.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   1935
      End
      Begin VB.TextBox xValor 
         Height          =   315
         Left            =   90
         TabIndex        =   5
         Top             =   540
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   315
         Width           =   360
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3615
      Left            =   165
      TabIndex        =   0
      Top             =   1065
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   6376
      _Version        =   393216
      BackColor       =   12648447
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
      Caption         =   "Selección de Registros"
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   675
      Left            =   2612
      Picture         =   "frmComun.frx":0465
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4740
      Width           =   775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   675
      Left            =   1367
      Picture         =   "frmComun.frx":08A7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4740
      Width           =   775
   End
End
Attribute VB_Name = "frmComun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Adoreg As ADODB.Recordset

Private Sub cmdCancel_Click()
 vgUtil(1) = ""
 vgUtil(2) = ""
 Unload Me
End Sub

Private Sub cmdOK_Click()
If Not IsNull(DataGrid1.Bookmark) Then
 Adoreg.Bookmark = DataGrid1.Bookmark
End If
 If Adoreg.RecordCount <> 0 Then
  vgUtil(1) = Adoreg.Fields(0)
  vgUtil(2) = Adoreg.Fields(1)
 Else
  vgUtil(1) = ""
  vgUtil(2) = ""
 End If
 Unload Me
End Sub

Private Sub DataGrid1_DblClick()
 cmdOK_Click
End Sub

Private Sub Form_Load()
 Me.Left = MDIMain.ScaleWidth - Me.Width
 Me.Top = MDIMain.Height - MDIMain.ScaleHeight
 Init_ControlDataGrid DataGrid1
 VarTemp = ""
 Set DataGrid1.DataSource = Adoreg
 DataGrid1.Columns(0).Caption = "Código"
 DataGrid1.Columns(1).Caption = "Descripción"
 xCampo.ListIndex = 0
End Sub

Public Sub Conectar(AD As ADODB.Recordset)
 Set Adoreg = AD
End Sub

Private Sub xValor_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrBusq
    If KeyAscii = 13 Then 'Si presionó enter
        Dim PosAnt As Long
        If Adoreg.RecordCount = 0 Then
            Exit Sub
        End If
        PosAnt = Adoreg.Bookmark
        Adoreg.MoveFirst
        Adoreg.Find xCampo.Text & "='" & xValor.Text & "'"
        If Adoreg.EOF Then
            Beep
            Adoreg.Bookmark = PosAnt
        End If
    End If
    Exit Sub
ErrBusq:
    Exit Sub
End Sub

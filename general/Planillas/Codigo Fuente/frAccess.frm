VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frAccess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrador de Tablas"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   Icon            =   "frAccess.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpenRecordset 
      Caption         =   "Abrir Recordset >>"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   4275
      Width           =   1740
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   3525
      Top             =   2670
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Abrir Base de Datos Access"
      Filter          =   "Base de Datos Access |*.mdb"
   End
   Begin VB.CommandButton cmdRunSQL 
      Caption         =   "Ejecutar SQL"
      Height          =   510
      Left            =   6420
      TabIndex        =   8
      Top             =   5265
      Width           =   945
   End
   Begin VB.TextBox xSQL 
      Height          =   990
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   4785
      Width           =   6240
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   3585
      Left            =   2565
      TabIndex        =   6
      Top             =   1110
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6324
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      Caption         =   "Analizador de Recordsets"
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
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   105
      TabIndex        =   4
      Top             =   1365
      Width           =   2340
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información de la Base de Datos"
      Height          =   885
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7260
      Begin VB.CommandButton cmAbrir 
         Height          =   330
         Left            =   6645
         Picture         =   "frAccess.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   345
         Width           =   390
      End
      Begin VB.Label xBase 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1425
         TabIndex        =   2
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Base de Datos"
         Height          =   195
         Left            =   165
         TabIndex        =   1
         Top             =   390
         Width           =   1050
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desarrollado por Daniel Yafac Baquedano"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   11
      Top             =   5850
      Width           =   2985
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desarrollado por Daniel Yafac Baquedano: danielyafac@hotmail.com"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   5865
      Width           =   4890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tablas en la base de datos"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   1125
      Width           =   1920
   End
End
Attribute VB_Name = "frAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CN1 As New ADODB.Connection
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Private Sub CMABRIR_Click()
    Dim RSEMPRESA As New ADODB.Recordset
    RSEMPRESA.Open "SELECT DIRALMACEN,NOMBRE FROM EMPRESAS", DBSTARPLAN
    frmComun.CONECTAR RSEMPRESA
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xBase.Caption = "BD :" & VGUTIL(1) & " Nombre:" & VGUTIL(2)
        xBase.Tag = VGUTIL(1)
        Call OPENDATA
    End If
End Sub
Public Sub OPENDATA()
Dim CAD As String
On Error GoTo ERR
    Set CN1 = New ADODB.Connection
    CAD = VGUTIL(1)
    Set CN1 = CONECTARDBSQL(CAD)
    
    Set RS2 = Nothing
    Set RS2 = CN1.OpenSchema(adSchemaTables)
    List1.Clear
    Do While Not RS2.EOF
        List1.AddItem RS2!TABLE_NAME
        RS2.MoveNext
    Loop
    If List1.ListCount = 0 Then
        MsgBox "No se ha encontrado tablas en esta Base de Datos", vbCritical
        cmdOpenRecordset.Enabled = False
    Else
        List1.ListIndex = 0
        cmdOpenRecordset.Enabled = True
    End If
Exit Sub
ERR:
    MsgBox "Error no se pudo abrir la Base de Datos ,falta actualizar"
End Sub

Private Sub CMDOPENRECORDSET_Click()
    On Error GoTo ERRNando
    If List1.ListIndex = -1 Then Exit Sub
    Set RS1 = Nothing
    RS1.Open List1.Text, CN1, adOpenStatic, adLockReadOnly
    Set xData.DataSource = RS1
    Exit Sub
ERRNando:
    Exit Sub
End Sub

Private Sub CMDRUNSQL_Click()
    Dim X As Long
    On Error GoTo ERRSQL
    If xSQL.Text = "" Then Exit Sub
    CN1.Execute xSQL.Text, X
    MsgBox "REGISTROS AFECTADOS: " & X
ERRSQL:
    Exit Sub
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RS1 = Nothing
    Set RS2 = Nothing
    Set CN1 = Nothing
End Sub

Private Sub XDATA_HEADCLICK(ByVal COLINDEX As Integer)
    RS1.Sort = xData.Columns(COLINDEX).Caption
End Sub


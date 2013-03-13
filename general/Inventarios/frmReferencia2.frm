VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmReferencia1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Referencias"
   ClientHeight    =   5025
   ClientLeft      =   2055
   ClientTop       =   2400
   ClientWidth     =   4845
   ControlBox      =   0   'False
   Icon            =   "frmReferencia2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   4845
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   675
      Left            =   1200
      Picture         =   "frmReferencia2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   3120
      Picture         =   "frmReferencia2.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   775
   End
   Begin VB.ComboBox cmbOrden 
      Height          =   315
      ItemData        =   "frmReferencia2.frx":114E
      Left            =   3405
      List            =   "frmReferencia2.frx":115B
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtFiltro 
      Height          =   285
      Left            =   165
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   150
      TabIndex        =   1
      Top             =   1440
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4683
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1275.024
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
      Left            =   3390
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Buscar:"
      Height          =   255
      Index           =   0
      Left            =   165
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblTit 
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
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   4065
   End
End
Attribute VB_Name = "frmReferencia1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Campos(1 To 3) As String

Dim Conexion As String
Dim csql As String
Dim Adodc3 As ADODB.Recordset
Dim k As Integer

Private Sub cmdCancel_Click()
    k = 0
    vGUtil(1) = ""
    vGUtil(2) = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If Adodc3.RecordCount <> 0 Then
        If Not IsNull(DataGrid1.Bookmark) Then
            Adodc3.Bookmark = DataGrid1.Bookmark
        End If
        If Adodc3.Fields.count >= 1 Then vGUtil(1) = Adodc3(0)
        If Adodc3.Fields.count >= 2 Then vGUtil(2) = Adodc3(1)
        If Adodc3.Fields.count >= 3 Then vGUtil(3) = Adodc3(2)
        If Adodc3.Fields.count >= 4 Then vGUtil(4) = Adodc3(3)
    Else
        vGUtil(1) = ""
        vGUtil(2) = ""
        vGUtil(3) = ""
        vGUtil(4) = ""
    End If
    k = 0
    Unload Me
End Sub

Private Sub CmbOrden_Click()
Dim ConexionAux As String
    k = k + 1
    
    If Adodc3.State = 1 Then
        Adodc3.Close
    End If
    'xx = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=FOX;Data Source=192.168.1.2;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=REUNIONES;Use Encryption for Data=False;Tag with column collation when possible=False"
    'xx = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID='sa';password='administrador';Initial Catalog='FOX';Data Source='192.168.1.2'"
    'Conexion
    Select Case CmbOrden.ListIndex
        Case 0
            Adodc3.Open csql & " ORDER BY " & Campos(2), VGCNx, adOpenDynamic, _
            adLockOptimistic
            Set DataGrid1.DataSource = Adodc3
            DataGrid1.Refresh
            txtFiltro = ""
        Case 1
            Adodc3.Open csql & " ORDER BY " & Campos(3), ConexionAux, adOpenDynamic, _
                adLockOptimistic
            Set DataGrid1.DataSource = Adodc3
            DataGrid1.Refresh
            txtFiltro = ""
        Case 2
            Adodc3.Open csql & " ORDER BY " & Campos(4), ConexionAux, adOpenDynamic, _
                adLockOptimistic
            Set DataGrid1.DataSource = Adodc3
            DataGrid1.Refresh
            txtFiltro = ""
    End Select
    
    With DataGrid1
        If Campos(1) = "CNUMRUC" Then
            CmbOrden.List(0) = "RUC"
            .Columns(0).Caption = "RUC"
        Else
            .Columns(0).Caption = "Código"
        End If
        .Columns(0).Width = 1300
        .Columns(1).Caption = "Fecha"
        .Columns(1).Width = 1500
        .Columns(2).Caption = "Estado"
        .Columns(2).Width = 2000
    End With
    
 '   If k > 1 Then DataGrid1.SetFocus
End Sub

Private Sub DataGrid1_DblClick()
    cmdOK_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then
        If Len(txtFiltro) - 1 > 0 Then
            txtFiltro = Left(txtFiltro, Len(txtFiltro) - 1)
        Else
            txtFiltro = ""
        End If
        KeyAscii = 0
    ElseIf KeyAscii <> 13 Then
        txtFiltro = txtFiltro & Chr(KeyAscii)
    End If
End Sub

Private Sub Form_Activate()
'    Inicio
    If DataGrid1.Enabled And DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub Form_Load()
    central Me
    
    Init_ControlDataGrid DataGrid1
End Sub

Public Sub Conectar(AD As ADODB.Recordset, Sq As String)
    Set Adodc3 = AD
    Conexion = Adodc3.ActiveConnection
    Campos(1) = Trim(Adodc3.Fields(0).Name)
    Campos(2) = Trim(Adodc3.Fields(1).Name)
    Campos(3) = Trim(Adodc3.Fields(2).Name)
    csql = Sq
End Sub

Private Sub txtFiltro_Change()
    Dim AdodcX As ADODB.Recordset
    Dim PriFila As Integer, nCursor
    Dim Origen As String
    Dim vFiltro As String
    
    Origen = Me.ActiveControl.Name
    vFiltro = Trim(UCase(txtFiltro))
    
    Set AdodcX = New ADODB.Recordset
    Set AdodcX = Adodc3.Clone
    
    If AdodcX.RecordCount > 0 Then
        If vFiltro <> "" Then
            nCursor = AdodcX.Bookmark
            AdodcX.MoveFirst
            
            AdodcX.Find Campos(CmbOrden.ListIndex + 1) & " LIKE '" & vFiltro & "*'"
            If AdodcX.EOF Then
                AdodcX.Bookmark = nCursor
            Else
                DataGrid1.SetFocus
                PriFila = DataGrid1.FirstRow
                Adodc3.Bookmark = AdodcX.Bookmark
                If PriFila <> DataGrid1.FirstRow Then DataGrid1.SetFocus
                If Origen = "txtFiltro" Then txtFiltro.SetFocus
            End If
        End If
    End If
    Set AdodcX = Nothing
End Sub

Private Sub txtFiltro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Sub inicio()
    Set DataGrid1.DataSource = Adodc3
    
        With DataGrid1
        If Campos(1) = "CNUMRUC" Then
            CmbOrden.List(0) = "RUC"
            .Columns(0).Caption = "RUC"
        Else
            .Columns(0).Caption = "Código"
        End If
        'Asigna etiquetas al DBGrid
        .Columns(0).Width = 1300
        .Columns(1).Caption = "Fecha"
        .Columns(1).Width = 1500
        .Columns(2).Caption = "Estado"
        .Columns(2).Width = 2000
    End With
    If CmbOrden.ListIndex = 0 Then
        CmbOrden_Click
    Else
        CmbOrden.ListIndex = 0
    End If
End Sub

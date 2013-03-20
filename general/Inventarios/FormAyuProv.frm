VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FormAyuProv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proveedores"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   ControlBox      =   0   'False
   Icon            =   "FormAyuProv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4080
      Picture         =   "FormAyuProv.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   735
      Left            =   2280
      Picture         =   "FormAyuProv.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6855
      Begin MSDataGridLib.DataGrid DBGrid1 
         Height          =   2655
         Left            =   210
         TabIndex        =   7
         Top             =   810
         Width           =   6495
         _ExtentX        =   11456
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
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1395
         TabIndex        =   0
         Top             =   360
         Width           =   2265
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormAyuProv.frx":114E
         Left            =   4560
         List            =   "FormAyuProv.frx":115B
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "FormAyuProv.frx":117A
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Filtro"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Indice"
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "FormAyuProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsprove As New ADODB.Recordset

Private Sub Combo1_Click()
Text1 = ""
Label1 = Combo1.text
If Combo1.ListIndex = 0 Then
'        Data1.RecordSource = "Select * from Maeprov order by prvccodigo"
    Call Listado("Select clientecodigo,clienterazonsocial,clientedireccion,clienteruc from cp_proveedor order by clientecodigo")
ElseIf Combo1.ListIndex = 1 Then
'        Data1.RecordSource = "Select * from Maeprov order by PRVCNOMBRE"
    Call Listado("Select clientecodigo,clienterazonsocial,clientedireccion,clienteruc from cp_proveedor order by clienterazonsocial")
ElseIf Combo1.ListIndex = 2 Then
'        Data1.RecordSource = "Select * from Maeprov order by PRVCRUC"
    Call Listado("Select clientecodigo,clienterazonsocial,clientedireccion,clienteruc from cp_proveedor order by clienteRUC")
End If
'Data1.Refresh
End Sub


Sub Listado(wcad)
  Set DbGrid1.DataSource = Nothing
  Set rsprove = Nothing
  
  Set rsprove = VGCNx.Execute(wcad)
  Set DbGrid1.DataSource = rsprove
  With DbGrid1
      .Columns(0).Caption = "Codigo"
      .Columns(0).Width = 1000
      .Columns(1).Caption = "Descripcion"
      .Columns(1).Width = 3800
      .MarqueeStyle = dbgHighlightRow
      .Refresh
  End With

End Sub


Private Sub Command1_Click()
If rsprove.RecordCount > 0 Then
    If VGForm1 = 21 Then
        'frmGuiaMabel.TxTProveedor.text = DBGrid1.Columns(0).text
    ElseIf VGForm1 = 11 Then
       frmPrueba.TxtProveedor.text = DbGrid1.Columns(0).text
    ElseIf VGForm1 = 12 Then
        FrmRegistro.TxtProveedor.text = DbGrid1.Columns(0).text
        FrmRegistro.Label13.Caption = DbGrid1.Columns(1).text
    ElseIf VGForm1 = 13 Then ' nuevo
    '    FrmMntMovimientos.TxtProveedor.text = DBGrid1.Columns(0).text
    '    FrmMntMovimientos.Label13.Caption = DBGrid1.Columns(1).text
    
    End If
End If
  Unload Me
End Sub

Private Sub Command8_Click()
   Unload Me
End Sub
Private Sub DBGrid1_DblClick()
   Command1_Click
End Sub

Private Sub Form_Activate()
  Text1.SetFocus
End Sub

Private Sub Form_Load()
   AlinearAyuda Me
'   Data1.DatabaseName = cRuta2
'   Data1.RecordSource = "select * from maeprov order by prvccodigo "
    Call Listado("select clientecodigo,clienterazonsocial,clientedireccion from cp_proveedor order by clientecodigo ")
'    PRVCCODIGO  PRVCNOMBRE                                         PRVCDIRECC                                         PRVCLOCALI      PRVCPAISAC      PRVCTELEF1      PRVCFAXACR      PRVCTIPOAC PRVCGIROAC PRVCREPRES                               PRVCCARREP           PRVCTELREP      PRVDFECCRE                                             PRVCUSER
'  Init_ControlDBGrid DBGrid1
  Combo1.ListIndex = 0
  Label1 = Combo1.text
  
End Sub

Private Sub Text1_Change()
   Dim ncar As String
   Dim criterio As String
   ncar = Str$(Len(Text1.text))   'REVISAR PARA TODOS
   If Combo1.ListIndex = 0 Then
     criterio = "Left(clientecodigo," & ncar & ") = '" & Text1.text & "'"
   ElseIf Combo1.ListIndex = 1 Then
     criterio = "Left(clienterazonsocial," & ncar & ") = '" & Text1.text & "'"
   Else
     criterio = "Left(clienteruc," & ncar & ") = '" & Text1.text & "'"
   End If
   'Data1.Recordset.FindFirst criterio
   Call Listado("select clientecodigo,clienterazonsocial,clientedireccion,clienteruc from cp_proveedor where " & criterio & " order by clientecodigo ")
   
End Sub

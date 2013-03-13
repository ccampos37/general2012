VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frmayuunidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unidades"
   ClientHeight    =   4770
   ClientLeft      =   3600
   ClientTop       =   1365
   ClientWidth     =   5085
   Icon            =   "FrmAyuUnidades.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5085
   Begin VB.Frame Frame2 
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
      TabIndex        =   3
      Top             =   0
      Width           =   4932
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   2790
      End
      Begin MSDataGridLib.DataGrid DbGrid1 
         Height          =   2412
         Left            =   216
         TabIndex        =   4
         Top             =   1020
         Width           =   4548
         _ExtentX        =   8043
         _ExtentY        =   4260
         _Version        =   393216
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "EQUNIPRI"
            Caption         =   "Uni. Ref."
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
            DataField       =   "EQUNIEQUI"
            Caption         =   "Uni.Med."
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
            DataField       =   "EQCANTEQUI"
            Caption         =   "Factor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   3588
      Width           =   4932
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   675
         Left            =   720
         Picture         =   "FrmAyuUnidades.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command8 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   2520
         Picture         =   "FrmAyuUnidades.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   775
      End
   End
End
Attribute VB_Name = "Frmayuunidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rsayuda As New ADODB.Recordset

Private Sub Command1_Click()
Select Case VGForm1
   Case 1            'desde   1 al  case 3 estuvo con comentario
     FormCreacion.Text2.text = Rsayuda.Fields(1)
   Case 2 'varform = "FrmCreacionSin"
     FrmCreacionSin.Text3.text = Rsayuda.Fields(0)
   Case 3 'articulos
      FormArticulos.lblUnidad.Caption = Rsayuda.Fields(1)
 '  Case 4  'UNIDADES
 '    FrmArUniMed.Text4 = Data2.Recordset.Fields("UM_ABREV")  ' VGform 4
      
  End Select
  '
  VGabrev = Rsayuda.Fields(0)
 ' FrmArUniMed.Label7 = Rsayuda.Fields("UM_NOMBRE")
  Unload Me
End Sub

Private Sub Command8_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
DBGrid1.SetFocus
End Sub

Private Sub Form_Load()
  Frmayuunidades.Caption = "Unidades de referencia"
  Call Listado("SELECT * FROM TABequi where equniequi= '" & VGabrev & "'   ")
  AlinearAyuda Me
End Sub


Sub Listado(wcad)
  Set DBGrid1.DataSource = Nothing
  Set Rsayuda = Nothing
  
  Set Rsayuda = VGCNx.Execute(wcad)
  Set DBGrid1.DataSource = Rsayuda
  With DBGrid1
      .Columns(0).Caption = "Codigo Ref."
      .Columns(0).Width = 1000
      .Columns(1).Caption = "Codigo Unidad"
      .Columns(1).Width = 1000
      .Columns(2).Caption = "Descripcion"
      .Columns(2).Width = 3800
      .MarqueeStyle = dbgHighlightRow
      .Refresh
  End With

End Sub




Private Sub Text1_Change()
  Dim ncar As String
  ncar = Str$(Len(Text1))
  criterio = "Left(TCLAVE," & ncar & ") = '" & Text1 & "'"
  
  Call Listado("SELECT * FROM TABAYU where TCOD= '" & codayu & "' AND " & criterio)
End Sub



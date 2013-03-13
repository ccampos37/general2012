VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FormAyuda 
   Caption         =   "Form de ayuda"
   ClientHeight    =   4665
   ClientLeft      =   75
   ClientTop       =   1950
   ClientWidth     =   9600
   Icon            =   "FormAyuda.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   9600
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
      Left            =   240
      TabIndex        =   2
      Top             =   135
      Width           =   9135
      Begin MSDataGridLib.DataGrid DbGrid1 
         Height          =   2415
         Left            =   210
         TabIndex        =   5
         Top             =   1020
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   4260
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
         TabIndex        =   3
         Top             =   480
         Width           =   2790
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   735
      Left            =   1920
      Picture         =   "FormAyuda.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   930
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   5085
      Picture         =   "FormAyuda.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   930
   End
End
Attribute VB_Name = "FormAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rsayuda As New ADODB.Recordset
Private Sub Command1_Click()
 If Rsayuda.RecordCount > 0 Then
    FrmRegistro.Text9.text = Rsayuda.Fields(1)
    frmCenCos.Text3.text = Rsayuda.Fields(1)
    FrmDesKits.Text9.text = Rsayuda.Fields(1)
 End If
 Unload Me
End Sub

Private Sub Command8_Click()
 Unload Me
End Sub

Private Sub DBGrid1_Click()
  'Command1_Click
End Sub

Private Sub Form_Load()
  codayu = "12"
  FormAyuda.Caption = "Autorizado"

  Call Listado("SELECT * FROM TABAYU where TCOD= '" & codayu & "'   ")
  AlinearAyuda Me
End Sub


Sub Listado(wcad)
  Set DbGrid1.DataSource = Nothing
  Set Rsayuda = Nothing
  
  Set Rsayuda = VGCNx.Execute(wcad)
  Set DbGrid1.DataSource = Rsayuda
  With DbGrid1
      .Columns(0).Caption = "Codigo"
      .Columns(0).Width = 1000
      .Columns(1).Caption = "Descripcion"
      .Columns(1).Width = 3800
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

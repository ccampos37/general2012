VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmValidacionSeries 
   Caption         =   "Validación de las Series del Registro de Ventas"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   690
      Left            =   105
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3600
      Width           =   5940
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2955
      Left            =   105
      TabIndex        =   0
      Top             =   450
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   5212
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
End
Attribute VB_Name = "frmValidacionSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim rs As ADODB.Recordset
  Dim rsM As ADODB.Recordset
  Dim SQL As String
  Dim serie As String
  Dim numeroini As Long
  Dim contador As Long
    
  'Ejecutar el Store del Registro de Ventas: Generar un Temporal
    
  Set rs = New ADODB.Recordset
  SQL = "select cabcomprobnumero,serie=left(detcomprobnumdocumento,3),"
  SQL = SQL & "numero=substring(detcomprobnumdocumento,5,8) from regventas "
  SQL = SQL & "order by 2,3"
  
  Set rs = VGCNx.Execute(SQL)
  
  Set DataGrid1.DataSource = rs
  rs.MoveFirst
  If Not rs.EOF And Not rs.BOF Then
     
    Do Until rs.EOF
       serie = rs(1)
       'Print serie
       numeroini = CLng(rs(2))
       Do Until rs.EOF And serie = rs(1)
          If numeroini <> rs(2) Then
              Text1.Text = Text1.Text & rs(2) & " / "
              rs.MoveNext
              numeroini = CLng(rs(2))
          End If

       Loop
    Loop
     
  End If

End Sub

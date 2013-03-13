VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmAyuLote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Lote"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "FrmAyuLote.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DBGrid1 
      Height          =   3015
      Left            =   210
      TabIndex        =   2
      Top             =   120
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   5318
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
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmAyuLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim adoreg As ADODB.Recordset
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()

If rs.RecordCount > 0 Then
   FrmDistribucion_1.Textlote = rs("STSLOTE")
   FrmDistribucion_1.txtcanti(3) = IIf(Not IsNull(rs("STSLKDIS")), rs("STSLKDIS"), 0)
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
central Me                                 ' Centra el Formulario
'Data1.DatabaseName = cRuta2
'Data1.RecordSource = "SELECT STSLOTE,STSFECVEN,STSFECFAB,STSLKDIS FROM STKLOTE WHERE  STSALMA = '" & VGAlma & "' and STSCODIGO = '" & VGcod & "' AND STSLKDIS<> 0 ORDER BY STSFECVEN"
'Data1.Refresh
Call Listado("SELECT STSLOTE,STSFECVEN,STSFECFAB,STSLKDIS FROM STKLOTE WHERE  STSALMA = '" & VGAlma & "' and STSCODIGO = '" & VGcod & "' AND STSLKDIS<> 0 ORDER BY STSFECVEN")
'CarObj                                ' Objetos

End Sub



Sub Listado(wcad)
  Set DBGrid1.DataSource = Nothing
  Set rs = Nothing
  
  Set rs = VGcnx.Execute(wcad)
  Set DBGrid1.DataSource = rs
  With DBGrid1
      .Columns(0).Caption = "Nro Lote"
      .Columns(0).Width = 2000
      .Columns(1).Caption = "Fec/Vcto"
      .Columns(1).Width = 1300
      .Columns(2).Caption = "StsFecFab"
      .Columns(2).Width = 1300
      .Columns(3).Caption = "Stslkdis"
      .Columns(3).Width = 1300
      .MarqueeStyle = dbgHighlightRow
      .Refresh
  End With

End Sub

Private Sub CarObj()        ' Carga Objetos
 DBGrid1.Columns(0).Locked = True
 DBGrid1.Columns(0).WrapText = True
' DataGrid1.Columns(0).Alignment = dbgCenter
 
 DBGrid1.Columns(0).Caption = "   CODIGO"
 DBGrid1.Columns(1).Caption = "  FECHA VCTO"
 DBGrid1.Columns(2).Caption = "  FECHA FAB"
 DBGrid1.Columns(3).Caption = "  STOCK"

End Sub



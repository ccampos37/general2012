VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmPrueba 
   Caption         =   "Consulta Proveedor - Articulo"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   Icon            =   "frmPrueba.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   3855
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6800
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8805
      Begin VB.CommandButton CmdConsultar 
         Caption         =   "&Consultar"
         Height          =   315
         Left            =   6600
         TabIndex        =   10
         Top             =   585
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   4680
         MaxLength       =   20
         TabIndex        =   9
         Top             =   600
         Width           =   1590
      End
      Begin VB.TextBox TxtProveedor 
         Height          =   285
         Left            =   4680
         MaxLength       =   11
         TabIndex        =   8
         Top             =   225
         Width           =   1590
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   49741825
         CurrentDate     =   36699
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   49741825
         CurrentDate     =   36699
      End
      Begin VB.Label Label4 
         Caption         =   "Articulo"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta la Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Desde la Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6780
      TabIndex        =   0
      Top             =   5385
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\inventarios.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "maeart"
      Top             =   5370
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsql As New ADODB.Recordset

Private Sub CmdConsultar_Click()
    DoSql
    
End Sub

Private Sub Command1_Click()
  Unload Me
End Sub


Private Sub Form_Load()
On Error GoTo FormLoad_Err
  central Me
  DTPicker1 = Date
  DTPicker1 = DateAdd("m", -1, Date)
  DTPicker2 = Date

  DoSql

  Data1.Refresh
FormLoad_Exit:
Exit Sub
FormLoad_Err:
   Exit Sub
End Sub

Sub DoInitialSettings()
  
End Sub

Sub DoSql()
    Dim mysql As String
    If Trim(TxtProveedor) = "" And Trim(Text1) <> "" Then
       mysql = "select cacodpro +'   '+canompro as Provedor,decodigo +' '+adescri as Articulo,decantid as Cantidad ,cafecdoc as Fecha,canumdoc as Numdoc from  maeart,movalmcab,movalmdet where denumdoc= canumdoc  and dealma = caalma and detd =catd and decodigo='" & Text1 & "' and  acodigo = decodigo and  cafecdoc  between '" & DTPicker1.Value & "' and '" & DTPicker2.Value & "'"
    ElseIf Trim(TxtProveedor) <> "" And Trim(Text1) = "" Then
      mysql = "select cacodpro +'   '+canompro as Provedor,decodigo +' '+adescri as Articulo,decantid as Cantidad,cafecdoc as Fecha,canumdoc as Numdoc from  maeart,movalmcab,movalmdet where denumdoc= canumdoc  and dealma = caalma and detd =catd and cacodpro ='" & TxtProveedor & "' and  acodigo = decodigo and  cafecdoc  between '" & DTPicker1.Value & "' and '" & DTPicker2.Value & "'"
    ElseIf Trim(TxtProveedor) = "" And Trim(Text1) = "" Then
      mysql = "select cacodpro +'  '+canompro as Provedor,decodigo +' '+adescri as Articulo,decantid as Cantidad ,cafecdoc as Fecha,canumdoc as Numdoc  from  maeart,movalmcab,movalmdet where denumdoc= canumdoc  and dealma = caalma and detd =catd and acodigo = decodigo  and  cafecdoc  between '" & DTPicker1.Value & "' and '" & DTPicker2.Value & "'"
    End If
    Set rsql = VGCNx.Execute(mysql)
    TDBGrid1.DataSource = rsql
    TDBGrid1.Refresh
End Sub

Sub DoSort()

End Sub

Private Sub Text1_DblClick()
   VGForm1 = 14
   FormAyuArt1.Show 1
   If Trim(Text1) <> "" Then
       TxtProveedor = ""
       CmdConsultar.SetFocus
  End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then
    Text1_DblClick
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Not codigo(Text1) Then
          CmdConsultar.SetFocus
       End If
    Else
       TxtProveedor = ""
    End If
End Sub

'**********     PROVEEDOR ****************
Private Sub TxtProveedor_DblClick()
 VGForm1 = 11
  FormAyuProv.Show 1
  If Trim(TxtProveedor) <> "" Then
      Text1 = ""
      CmdConsultar.SetFocus
  End If
End Sub

Private Sub TxtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
    TxtProveedor_DblClick
   End If
   End Sub

Private Sub TxtProveedor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And TxtProveedor <> "" Then
           TxtProveedor = Trim(TxtProveedor)
           If prove(TxtProveedor) <> "" Then
              CmdConsultar.SetFocus
           End If
   Else
            Text1 = ""
   End If
End Sub

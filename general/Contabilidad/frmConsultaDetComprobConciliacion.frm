VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{2B12169D-6738-11D2-BF5B-00A024982E5B}#19.0#0"; "AXBUTTON.OCX"
Begin VB.Form frmConsultaDetComprobConciliacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle del Comprobante  Contable Nº "
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin TrueOleDBGrid70.TDBGrid TDBG_ConsultaDetalle 
      Height          =   2415
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4260
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Inaf"
      Columns(0).DataField=   "plantillaasientoinafecto"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Item"
      Columns(1).DataField=   "detcomprobitem"
      Columns(1).DataWidth=   5
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Op"
      Columns(2).DataField=   "operacioncodigo"
      Columns(2).DataWidth=   2
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Cod. analitico"
      Columns(3).DataField=   "analiticocodigo"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Cuenta"
      Columns(4).DataField=   "cuentacodigo"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "T/D"
      Columns(5).DataField=   "documentocodigo"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Nº Documento"
      Columns(6).DataField=   "detcomprobnumdocumento"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "ID"
      Columns(7).DataField=   "indicador"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Monto Soles"
      Columns(8).DataField=   "montosol"
      Columns(8).NumberFormat=   "###,###,###,###.00"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Monto Dolares"
      Columns(9).DataField=   "montouss"
      Columns(9).NumberFormat=   "###,###,###,###.00"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   4
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Auto"
      Columns(10).DataField=   "detcomprobauto"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1032"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=953"
      Splits(0)._ColumnProps(9)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=258"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=582"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=503"
      Splits(0)._ColumnProps(15)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=260"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=2037"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1958"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=260"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=2487"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2408"
      Splits(0)._ColumnProps(26)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=260"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=794"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=714"
      Splits(0)._ColumnProps(32)=   "Column(5).AllowSizing=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._ColStyle=260"
      Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(35)=   "Column(6).Width=3201"
      Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=3122"
      Splits(0)._ColumnProps(38)=   "Column(6)._ColStyle=260"
      Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(40)=   "Column(7).Width=609"
      Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=529"
      Splits(0)._ColumnProps(43)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(44)=   "Column(7)._ColStyle=260"
      Splits(0)._ColumnProps(45)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(46)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(47)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(49)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(50)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(51)=   "Column(9).Width=2752"
      Splits(0)._ColumnProps(52)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(9)._WidthInPix=2672"
      Splits(0)._ColumnProps(54)=   "Column(9).AllowSizing=0"
      Splits(0)._ColumnProps(55)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(56)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(57)=   "Column(10).Width=1402"
      Splits(0)._ColumnProps(58)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(10)._WidthInPix=1323"
      Splits(0)._ColumnProps(60)=   "Column(10).AllowSizing=0"
      Splits(0)._ColumnProps(61)=   "Column(10)._ColStyle=513"
      Splits(0)._ColumnProps(62)=   "Column(10).Order=11"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      RowDividerStyle =   6
      Caption         =   "Detalle del Comprobante Seleccionado"
      MultipleLines   =   0
      CellTips        =   2
      CellTipsWidth   =   0
      MultiSelect     =   2
      DataView        =   1
      AnimateWindow   =   2
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   1140.095
      ViewColumnWidth =   9764.788
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HC0C0C0&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H344A87&"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=47,.parent=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=56,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=48,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=49,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=50,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=52,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=51,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=53,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=54,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=55,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=57,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=58,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=16,.parent=47,.alignment=2,.bgcolor=&HFCEDE4&"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=48"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=49"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=51"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=47,.alignment=1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=48,.alignment=0"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=49"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=51"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=66,.parent=47"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=48,.alignment=0"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=49"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=51"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=47"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=48,.alignment=0"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=49"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=51"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=74,.parent=47"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=48,.alignment=0"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=49"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=51"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=130,.parent=47"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=127,.parent=48,.alignment=0"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=128,.parent=49"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=129,.parent=51"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=134,.parent=47"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=131,.parent=48,.alignment=0"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=132,.parent=49"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=133,.parent=51"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=138,.parent=47"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=135,.parent=48,.alignment=0"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=136,.parent=49"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=137,.parent=51"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=154,.parent=47,.alignment=1,.bgcolor=&H80000014&"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=151,.parent=48,.alignment=2"
      _StyleDefs(70)  =   ":id=151,.bgcolor=&H8000000F&"
      _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=152,.parent=49"
      _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=153,.parent=51,.bgcolor=&H80000018&"
      _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=158,.parent=47,.alignment=1,.bgcolor=&H80000014&"
      _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=155,.parent=48,.alignment=2"
      _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=156,.parent=49"
      _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=157,.parent=51,.bgcolor=&H80000018&"
      _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=162,.parent=47,.alignment=2"
      _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=159,.parent=48,.alignment=2"
      _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=160,.parent=49"
      _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=161,.parent=51"
      _StyleDefs(81)  =   "Named:id=33:Normal"
      _StyleDefs(82)  =   ":id=33,.parent=0"
      _StyleDefs(83)  =   "Named:id=34:Heading"
      _StyleDefs(84)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(85)  =   ":id=34,.wraptext=-1"
      _StyleDefs(86)  =   "Named:id=35:Footing"
      _StyleDefs(87)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(88)  =   "Named:id=36:Selected"
      _StyleDefs(89)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(90)  =   "Named:id=37:Caption"
      _StyleDefs(91)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(92)  =   "Named:id=38:HighlightRow"
      _StyleDefs(93)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(94)  =   "Named:id=39:EvenRow"
      _StyleDefs(95)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(96)  =   "Named:id=40:OddRow"
      _StyleDefs(97)  =   ":id=40,.parent=33"
      _StyleDefs(98)  =   "Named:id=41:RecordSelector"
      _StyleDefs(99)  =   ":id=41,.parent=34"
      _StyleDefs(100) =   "Named:id=42:FilterBar"
      _StyleDefs(101) =   ":id=42,.parent=33"
   End
   Begin axButtonControl.axButton axButton1 
      Height          =   390
      Left            =   4605
      TabIndex        =   1
      Top             =   2565
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Salir"
      MaskColor       =   -2147483633
      Style           =   1
   End
End
Attribute VB_Name = "frmConsultaDetComprobConciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_NumeroComprobante As String

Private Sub Form_Load()
    Call EjecutarConsultaDetalle(m_NumeroComprobante)
    Me.Caption = Me.Caption & " " & m_NumeroComprobante
End Sub

Private Sub EjecutarConsultaDetalle(xParam As String)
  Dim rsdetalle As New ADODB.Recordset
  Dim sqlcad As String
  sqlcad = "SELECT plantillaasientoinafecto,detcomprobitem,operacioncodigo,analiticocodigo,cuentacodigo,"
  sqlcad = sqlcad & "documentocodigo,detcomprobnumdocumento,indicador= case when detcomprobdebe>0 then 'D' else 'H' end,"
  sqlcad = sqlcad & "montosol=case when detcomprobdebe>0 then detcomprobdebe else detcomprobhaber end,"
  sqlcad = sqlcad & "montouss=case when detcomprobussdebe>0 then detcomprobussdebe else detcomprobusshaber end,detcomprobauto,monedacodigo "
  sqlcad = sqlcad & "FROM " & VGParamSistem.TablaDetcomprob & " WHERE cabcomprobnumero='" & xParam & " '"
  
  Set rsdetalle = New ADODB.Recordset
  Set rsdetalle = VGcnx.Execute(sqlcad)
  Set TDBG_ConsultaDetalle.DataSource = rsdetalle

End Sub

Property Let NumeroComprobante(valor As String)
  m_NumeroComprobante = valor
End Property

Private Sub axButton1_Click()
    Unload Me
End Sub

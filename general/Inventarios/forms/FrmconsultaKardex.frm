VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmConsultakardex 
   Caption         =   "Form2"
   ClientHeight    =   7620
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   ScaleHeight     =   7620
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1455
      Left            =   11400
      TabIndex        =   13
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2566
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   5655
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9975
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Fecha"
      Columns(0).DataField=   "cafecdoc"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Numero"
      Columns(1).DataField=   "canumdoc"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nro.transf"
      Columns(2).DataField=   "canrotransf"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Tipo"
      Columns(3).DataField=   "catipmov"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Mov"
      Columns(4).DataField=   "cacodmov"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Transaccion"
      Columns(5).DataField=   "transacciondescripcion"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "cantidad"
      Columns(6).DataField=   "cant1"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Saldo"
      Columns(7).DataField=   "saldo"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1852"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1773"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2064"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1984"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1879"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1799"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=556"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=476"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=635"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=556"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=3122"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=3043"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=1032"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=953"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=1482"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1402"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=69,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=71,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=74,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=67"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=68"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=69"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=71"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=86,.parent=67"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=68"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=69"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=71"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=90,.parent=67"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=87,.parent=68"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=88,.parent=69"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=89,.parent=71"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=94,.parent=67"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=91,.parent=68"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=92,.parent=69"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=93,.parent=71"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=98,.parent=67"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=95,.parent=68"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=96,.parent=69"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=97,.parent=71"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=102,.parent=67"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=99,.parent=68"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=100,.parent=69"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=101,.parent=71"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=106,.parent=67"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=103,.parent=68"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=104,.parent=69"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=105,.parent=71"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=110,.parent=67"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=107,.parent=68"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=108,.parent=69"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=109,.parent=71"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51773441
      CurrentDate     =   41093
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAlm1 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   661
      XcodMaxLongitud =   0
      xcodwith        =   100
      NomTabla        =   "tabalm"
      TituloAyuda     =   "Almacenes"
      ListaCampos     =   "TAALMA(1),TADESCRI(1),empresacodigo(1)"
      XcodCampo       =   "TAALMA"
      XListCampo      =   "TADESCRI"
      ListaCamposDescrip=   "Codigo,Descripcion,empresa"
      ListaCamposText =   "TAALMA,TADESCRI,empresacodigo"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAlm2 
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   1200
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   661
      XcodMaxLongitud =   0
      xcodwith        =   100
      NomTabla        =   "tabalm"
      TituloAyuda     =   "Almacenes"
      ListaCampos     =   "TAALMA(1),TADESCRI(1),empresacodigo(1)"
      XcodCampo       =   "TAALMA"
      XListCampo      =   "TADESCRI"
      ListaCamposDescrip=   "Codigo,Descripcion,empresa"
      ListaCamposText =   "TAALMA,TADESCRI,empresacodigo"
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuCodigo 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   661
      XcodMaxLongitud =   0
      xcodwith        =   1000
      NomTabla        =   "maeart"
      TituloAyuda     =   "Almacenes"
      ListaCampos     =   "acodigo(1), adescri(1)"
      XcodCampo       =   "acodigo"
      XListCampo      =   "adescri"
      ListaCamposDescrip=   "Codigo,Descripcion,empresa"
      ListaCamposText =   "acodigo, adescri"
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51773441
      CurrentDate     =   41093
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuPtoVta 
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   840
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   661
      XcodMaxLongitud =   0
      xcodwith        =   300
      NomTabla        =   "vt_puntoventa"
      TituloAyuda     =   "Almacenes"
      ListaCampos     =   "puntovtacodigo(1),puntovtadescripcion(1)"
      XcodCampo       =   "puntovtacodigo"
      XListCampo      =   "puntovtadescripcion"
      ListaCamposDescrip=   "Codigo,Descripcion"
      ListaCamposText =   "puntovtacodigo,puntovtadescripcion"
      Requerido       =   0   'False
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
      Height          =   5655
      Left            =   8400
      TabIndex        =   17
      Top             =   1680
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9975
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Fecha"
      Columns(0).DataField=   "cafecdoc"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Numero"
      Columns(1).DataField=   "canumdoc"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nro.transf"
      Columns(2).DataField=   "canrotransf"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Tipo"
      Columns(3).DataField=   "catipmov"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Mov"
      Columns(4).DataField=   "cacodmov"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Transaccion"
      Columns(5).DataField=   "transacciondescripcion"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "cantidad"
      Columns(6).DataField=   "cant1"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Saldo"
      Columns(7).DataField=   "saldo"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1852"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1773"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2064"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1984"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1879"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1799"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=556"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=476"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=635"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=556"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=3122"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=3043"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=1032"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=953"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=1482"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1402"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=69,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=71,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=74,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=75,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=67"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=68"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=69"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=71"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=86,.parent=67"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=68"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=69"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=71"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=90,.parent=67"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=87,.parent=68"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=88,.parent=69"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=89,.parent=71"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=94,.parent=67"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=91,.parent=68"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=92,.parent=69"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=93,.parent=71"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=98,.parent=67"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=95,.parent=68"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=96,.parent=69"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=97,.parent=71"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=102,.parent=67"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=99,.parent=68"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=100,.parent=69"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=101,.parent=71"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=106,.parent=67"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=103,.parent=68"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=104,.parent=69"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=105,.parent=71"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=110,.parent=67"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=107,.parent=68"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=108,.parent=69"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=109,.parent=71"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   960
      TabIndex        =   18
      Top             =   2880
      Width           =   12015
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   3135
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5530
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
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   3135
         Left            =   7800
         TabIndex        =   20
         Top             =   840
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5530
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
      Begin VB.Label Label8 
         Caption         =   "Detalle Transferencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9600
         TabIndex        =   22
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Detalle Documentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Hasta :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Desde :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   15
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Punto de venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LblSaldo2 
      Height          =   495
      Left            =   7800
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label LblSaldo1 
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Alm.destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Alm. Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "FrmConsultakardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rsql1 As New ADODB.Recordset
Attribute Rsql1.VB_VarHelpID = -1
Dim rsql2 As New ADODB.Recordset
Attribute rsql2.VB_VarHelpID = -1
Dim alm1 As Integer
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Ctr_AyuAlm1_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Call mostrar1
End Sub
Private Sub Ctr_AyuAlm2_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Call Mostrar2
End Sub
Private Sub Ctr_AyuCodigo_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Call saldoalmacenes
Call mostrar1
If Ctr_AyuAlm2.xclave <> "" Then Call Mostrar2
End Sub
Private Sub Ctr_AyuPtoVta_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Call saldoalmacenes
End Sub



Private Sub Form_Load()
DTPicker1.Value = "01/01/2010"
DTPicker2.Value = Date
Ctr_AyuAlm1.conexion VGCNx
Ctr_AyuAlm2.conexion VGCNx
Ctr_AyuCodigo.conexion VGCNx
Ctr_AyuPtoVta.conexion VGCNx
End Sub
Private Sub mostrar1()
Dim saldo As Double
Dim RSQL As New ADODB.Recordset
If Ctr_AyuCodigo.xclave = "" Then Exit Sub
SQL = " select cafecdoc,catipmov,transacciondescripcion=left(transacciondescripcion,20),cacodmov,canumdoc,cant1=decantid, saldo=0, "
SQL = SQL & " canrotransf from v_kardex "
SQL = SQL & " where almacenempresa='" & VGparametros.empresacodigo & "' and decodigo='" & Ctr_AyuCodigo.xclave & "'"
SQL = SQL & " and  cafecdoc >='" & DTPicker1.Value & "' and cafecdoc<='" & DTPicker2.Value & "'"
If Ctr_AyuPtoVta.xclave <> "" Then SQL = SQL & " and puntovtacodigo='" & Ctr_AyuPtoVta.xclave & "'"
If Ctr_AyuAlm1.xclave <> "" Then SQL = SQL & " and dealma='" & Ctr_AyuAlm1.xclave & "'"
SQL = SQL & " and decantid <> 0  order by cafecdoc,catipmov "

Set RSQL = VGCNx.Execute(SQL)
Call carga1

If RSQL.RecordCount = 0 Then Set TDBGrid1.DataSource = Rsql1
If RSQL.RecordCount = 0 Then TDBGrid1.Refresh
If RSQL.RecordCount = 0 Then Exit Sub
RSQL.MoveFirst
saldo = 0
Do While Not RSQL.EOF
   If RSQL!catipmov = "I" Then
      saldo = saldo + RSQL!cant1
    Else
      saldo = saldo - RSQL!cant1
   End If
      Rsql1.AddNew
      Rsql1!CAFECDOC = Format(RSQL!CAFECDOC, "dd/mm/yyyy")
      Rsql1!catipmov = RSQL!catipmov
      Rsql1!CANUMDOC = RSQL!CANUMDOC
      Rsql1!caNROtransf = ESNULO(RSQL!caNROtransf, "")
      Rsql1!transacciondescripcion = RSQL!transacciondescripcion
      Rsql1!cacodmov = RSQL!cacodmov
      Rsql1!cant1 = RSQL!cant1
      Rsql1!saldo = saldo
      Rsql1.Update
      If Rsql1!saldo < 0 Then
       End If
   RSQL.MoveNext
Loop
Set TDBGrid1.DataSource = Rsql1
TDBGrid1.Refresh
End Sub
Private Sub carga1()
   Set Rsql1 = Nothing
   Call Rsql1.Fields.Append("cafecdoc", adChar, 10)
   Call Rsql1.Fields.Append("catipmov", adChar, 1)
   Call Rsql1.Fields.Append("transacciondescripcion", adChar, 20)
   Call Rsql1.Fields.Append("cacodmov", adChar, 2)
   Call Rsql1.Fields.Append("canumdoc", adChar, 11)
   Call Rsql1.Fields.Append("canrotransf", adChar, 11)
   Call Rsql1.Fields.Append("cant1", adDouble)
   Call Rsql1.Fields.Append("saldo", adDouble)
   Rsql1.Open
End Sub
Private Sub Mostrar2()
Dim saldo As Double
Dim RSQL As New ADODB.Recordset
If Ctr_AyuCodigo.xclave = "" Then Exit Sub
SQL = " select cafecdoc,catipmov,transacciondescripcion=left(transacciondescripcion,20),cacodmov,canumdoc,cant1=decantid, saldo=0 , "
SQL = SQL & " canrotransf from v_kardex "
SQL = SQL & " where almacenempresa='" & VGparametros.empresacodigo & "' and decodigo='" & Ctr_AyuCodigo.xclave & "'"
SQL = SQL & " and  cafecdoc >='" & DTPicker1.Value & "' and cafecdoc<='" & DTPicker2.Value & "'"
If Ctr_AyuAlm2.xclave <> "" Then SQL = SQL & " and dealma='" & Ctr_AyuAlm2.xclave & "'"
SQL = SQL & " and decantid <> 0 order by cafecdoc,catipmov "
Set RSQL = VGCNx.Execute(SQL)
Call Carga2

If RSQL.RecordCount = 0 Then Set TDBGrid2.DataSource = rsql2
If RSQL.RecordCount = 0 Then TDBGrid2.Refresh
If RSQL.RecordCount = 0 Then Exit Sub
RSQL.MoveFirst
saldo = 0
Do While Not RSQL.EOF
   If RSQL!catipmov = "I" Then
      saldo = saldo + RSQL!cant1
    Else
      saldo = saldo - RSQL!cant1
   End If
   rsql2.AddNew
   rsql2!CAFECDOC = Format(RSQL!CAFECDOC, "dd/mm/yyyy")
   rsql2!catipmov = RSQL!catipmov
   rsql2!cacodmov = RSQL!cacodmov
   rsql2!CANUMDOC = RSQL!CANUMDOC
   rsql2!caNROtransf = ESNULO(RSQL!caNROtransf, "")
   rsql2!transacciondescripcion = RSQL!transacciondescripcion
   rsql2!cant1 = RSQL!cant1
   rsql2!saldo = saldo
   rsql2.Update
   RSQL.MoveNext
Loop
Set TDBGrid2.DataSource = rsql2
TDBGrid2.Refresh
End Sub

Private Sub Carga2()
   Set rsql2 = Nothing
   Call rsql2.Fields.Append("cafecdoc", adChar, 10)
   Call rsql2.Fields.Append("catipmov", adChar, 1)
   Call rsql2.Fields.Append("transacciondescripcion", adChar, 20)
   Call rsql2.Fields.Append("cacodmov", adChar, 2)
   Call rsql2.Fields.Append("canumdoc", adChar, 11)
   Call rsql2.Fields.Append("canrotransf", adChar, 11)
   Call rsql2.Fields.Append("cant1", adDouble)
   Call rsql2.Fields.Append("saldo", adDouble)
   rsql2.Open
End Sub
Private Sub saldoalmacenes()
Dim ralmacen As New ADODB.Recordset
SQL = " select codigo=stalma, descripcion=tadescri, saldo=isnull(stskdis,0) from tabalm a "
SQL = SQL & " inner join stkart b on a.taalma=b.stalma where empresacodigo='" & VGparametros.empresacodigo & "'"
SQL = SQL & " and stcodigo='" & Ctr_AyuCodigo.xclave & "' and isnull(stskdis,0) <> 0 "
If Ctr_AyuPtoVta.xclave <> "" Then
   SQL = SQL & " and a.puntovtacodigo='" & Ctr_AyuPtoVta.xclave & "'"
End If
Set ralmacen = VGCNx.Execute(SQL)
Set DataGrid1.DataSource = ralmacen
End Sub


VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAyuda 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9060
   Begin VB.Frame Frame3 
      Height          =   4365
      Left            =   150
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   8685
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   2475
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   4366
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
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=194,.bold=0,.fontsize=825,.italic=0"
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
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   4
         Left            =   4590
         TabIndex        =   24
         Top             =   3330
         Width           =   795
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3300
         MaxLength       =   2
         TabIndex        =   23
         Top             =   3360
         Width           =   465
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1050
         TabIndex        =   22
         Top             =   3360
         Width           =   1155
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancela"
         Height          =   375
         Left            =   7530
         TabIndex        =   28
         Top             =   3840
         Width           =   945
      End
      Begin VB.CommandButton cElige 
         Caption         =   "&Acepta"
         Height          =   375
         Left            =   6510
         TabIndex        =   27
         Top             =   3840
         Width           =   945
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   8100
         TabIndex        =   26
         Top             =   3270
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   6270
         TabIndex        =   25
         Top             =   3330
         Width           =   885
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   90
         TabIndex        =   16
         Top             =   120
         Width           =   8505
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   2190
            TabIndex        =   19
            Top             =   150
            Width           =   6195
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   810
            MaxLength       =   2
            TabIndex        =   18
            Top             =   150
            Width           =   435
         End
         Begin VB.Label Label2 
            Caption         =   "Producto"
            Height          =   225
            Index           =   1
            Left            =   1410
            TabIndex        =   20
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Almacen"
            Height          =   225
            Index           =   4
            Left            =   90
            TabIndex        =   17
            Top             =   180
            Width           =   735
         End
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota Ingreso"
         Height          =   285
         Index           =   1
         Left            =   3870
         TabIndex        =   33
         Top             =   3870
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "Nota Ingreso"
         Height          =   225
         Index           =   1
         Left            =   2820
         TabIndex        =   32
         Top             =   3930
         Width           =   945
      End
      Begin VB.Label Label5 
         Caption         =   "Nota Salida"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   31
         Top             =   3960
         Width           =   945
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota Salida"
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   30
         Top             =   3900
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
         Height          =   225
         Index           =   5
         Left            =   5550
         TabIndex        =   21
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Almacen Origen"
         Height          =   405
         Index           =   3
         Left            =   2460
         TabIndex        =   15
         Top             =   3300
         Width           =   705
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   30
         X2              =   8580
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   30
         X2              =   8610
         Y1              =   3735
         Y2              =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Producto"
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   14
         Top             =   3390
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "Almacen Destino"
         Height          =   435
         Left            =   7320
         TabIndex        =   13
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Cant.Ref"
         Height          =   225
         Index           =   0
         Left            =   3870
         TabIndex        =   12
         Top             =   3360
         Width           =   735
      End
   End
   Begin VB.CommandButton cTraslado 
      Caption         =   "&Traslado"
      Height          =   435
      Left            =   210
      TabIndex        =   10
      Top             =   5550
      Width           =   915
   End
   Begin TrueOleDBGrid70.TDBGrid DGrid1 
      Height          =   3855
      Left            =   180
      TabIndex        =   2
      Top             =   1590
      Width           =   8595
      _ExtentX        =   15161
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
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2752"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=192,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Arial"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Arial"
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
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   9
      Top             =   6144
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   180
      TabIndex        =   7
      Top             =   -30
      Width           =   8565
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Elija su Tipo de Busqueda e Ingrese su Dato a buscar"
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
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   6105
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   450
         Picture         =   "FrmAyuda.frx":0000
         Top             =   210
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   180
      TabIndex        =   6
      Top             =   870
      Width           =   8595
      Begin VB.CommandButton cBusca 
         BackColor       =   &H0000C0C0&
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   7170
         TabIndex        =   1
         Top             =   210
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   270
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2040
         TabIndex        =   0
         Top             =   270
         Width           =   4965
      End
   End
   Begin VB.CommandButton cCerrar 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Cerrar"
      Height          =   435
      Left            =   7620
      TabIndex        =   4
      Top             =   5535
      Width           =   1170
   End
   Begin VB.CommandButton cAcepto 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Acepta"
      Height          =   435
      Left            =   6360
      TabIndex        =   3
      Top             =   5550
      Width           =   1170
   End
End
Attribute VB_Name = "FrmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xtipo As Integer
Dim AOrden, ACondi As String
Dim DLLGENERALAYUDA As New dllgeneral.dll_general
Dim vcon As New ADODB.Connection
Dim xtabla, xCampos, xOrden, xCondi As String
Dim xdata, xdato As String
Dim aflag As Integer
Dim nlongi(1) As Integer
Dim nfiltra() As String
Private Sub cAcepto_Click()
  Call DGrid1_KeyDown(13, 0)
End Sub

Private Sub cBusca_Click()
    Call Text1_KeyPress(13)
End Sub

Private Sub cCancela_Click()
  Frame3.Visible = False
  Set TDBGrid1.DataSource = Nothing
End Sub
Private Sub cCerrar_Click()
  Dim acmd As New ADODB.Command
  Dim J As Integer
  nAyuda = "": nDetalle = ""
  If xdata = "2" And aflag = 1 And Len(Trim(Label4(0))) > 0 And Len(Trim(Text2(2))) > 0 Then
     For J = 1 To 2
        If J = 1 Then
            VGCNx.Execute "UPDATE movalmcab " & _
                     " set catipmov='A'" & _
                     " where catd='NS' and canumdoc='" & Trim(Label4(0)) & "' and caalma='" & Text2(1) & "'"
        Else
            VGCNx.Execute "UPDATE movalmcab " & _
                     " set catipmov='A'" & _
                     " where catd='NI' and canumdoc='" & Trim(Label4(1)) & "' and caalma='" & Text2(3) & "'"
        End If
        Set acmd.ActiveConnection = VGgeneral
        acmd.CommandType = adCmdStoredProc
        acmd.CommandTimeout = 0
        acmd.CommandText = "vt_actualizoalma_pro"
        acmd.Prepared = True
        With acmd
            .Parameters("@basedes") = CStr(VGCNx.DefaultDatabase)
            If J = 1 Then
              .Parameters("@almacen") = Trim(Text2(1))
            Else
              .Parameters("@almacen") = Trim(Text2(3))
            End If
            .Parameters("@tipo") = "2"
            .Parameters("@articulo") = Trim(Text2(0))
            .Parameters("@cantidad") = Trim(Text2(2))
        End With
        acmd.Execute
    Next J
  End If
  Unload Me
  Set acmd = Nothing
End Sub

Private Sub cElige_Click()
    Dim J As Integer
    Dim nsql As String
    Dim ltipo As String
    Dim lzona As String
    Dim xserie As String * 3
    Dim xfactu As Double  'String * 8
    Dim xtipofac As String * 2
    Dim ndato As String
    Dim nflag As Integer
    
    Dim acmd As New ADODB.Command
    Dim asql As New ADODB.Recordset
    Dim arbusca As New ADODB.Recordset
    Dim wCabe(40)
    
    If Len(Trim(Text2(0))) = 0 Or Len(Trim(Text2(1))) = 0 Or Len(Trim(Text2(3))) = 0 Or Len(Trim(Text2(2))) = 0 Then
      MsgBox "Debe seleccionar el articulo..Verifique!!", vbInformation, "AVISO"
      Exit Sub
    End If
    
    On Error GoTo nerror
    Text2(0).Text = Trim(TDBGrid1.Columns(0).Text)
    Text2(1).Text = Text4    'almacen origen
    Text2(3).Text = FrmPedidoVentanilla.Ctr_Ayuda3.xclave  'almacen destino

    If xdata = "2" And aflag = 1 And Len(Trim(Text2(0))) = 0 Then
      Frame3.Visible = False
      Exit Sub
    End If
    
    If TDBGrid1.ApproxCount > 0 Then
       If Val(TDBGrid1.Columns(3).Text) < Val(Text2(2)) Then
           MsgBox "La Cantidad solicitada excede al stock disponible..Verifique!!", vbInformation, "AVISO"
           Exit Sub
       End If
    End If
   
    Text3 = Trim(TDBGrid1.Columns(1).Text)
    
    '******** CABECERA DE MOVIMIENTO *****************
    For J = 1 To 29
        wCabe(J) = ""
    Next J
    Label4(0) = "": Label4(1) = ""
    
    If DLLGENERALAYUDA.VerificaDatoExistente(VGCNx, "select * from stkart where stalma='" & Text2(3) & "' and stcodigo='" & Text2(0) & "'") = 0 Then
        VGCNx.Execute "insert into stkart " & _
                        "(stalma,stcodigo,stskdis)" & _
                        " Values ('" & Text2(3) & "','" & Text2(0) & "',0)"
    End If
    Set asql = VGCNx.Execute("select * from  num_documentos where ctncodigo='TR'")
    If asql.RecordCount > 0 Then
        ndato = Right("000000000" & Trim(CStr(asql!ctnnumero)), 11)                    'nro pedido"
    Else
       MsgBox " No existe documentos de transacciones...Verifique!!", vbInformation, "AVISO"
       asql.Close
       Set asql = Nothing
       Exit Sub
    End If
    asql.Close
    Set asql = Nothing

    VGCNx.Execute "update num_documentos " & _
                    " set ctnnumero=ctnnumero+1 " & _
                    " where ctncodigo='TR'"

    
    For J = 1 To 2
        wCabe(1) = g_ptoventa                        'Pto Venta
        Set asql = Nothing
        If J = 1 Then
            ' de Almacen origen
           Set asql = VGCNx.Execute("select * from tabalm where taalma='" & Text2(1).Text & "'")
           If asql.RecordCount > 0 Then
               wCabe(2) = Right("0000000000000" & Trim(CStr(asql!tanumsal)), 11)                       'nro pedido"
           End If
           asql.Close
           Set asql = Nothing
           VGCNx.Execute "update tabalm " & _
                           " set tanumsal=tanumsal+1 " & _
                           " where taalma='" & Text2(1) & "'"
                           
           Label4(0) = wCabe(2)
        Else
            ' al almacen destino
           Set asql = VGCNx.Execute("select * from tabalm where taalma='" & Text2(3) & "'")
           If asql.RecordCount > 0 Then
               wCabe(2) = Right("0000000000000" & Trim(CStr(asql!tanument)), 11)                       'nro pedido"
           End If
           asql.Close
           Set asql = Nothing
           VGCNx.Execute "update tabalm " & _
                           " set tanument=tanument+1 " & _
                           " where taalma='" & Text2(3) & "'"
           Label4(1) = wCabe(2)
        End If
        wCabe(3) = ndato                      'nro factura
        wCabe(4) = "TR"                      'nro boleta
        wCabe(5) = ""                      'nro guia
        wCabe(6) = 0                       'dscto gral
        wCabe(7) = 0                       'dscto promocional
        wCabe(8) = 0                       'dscto especial
        wCabe(9) = g_TipoSol               'moneda
        wCabe(10) = 0                      'tipo de cambio
        wCabe(11) = 0                      'lista de precios
        wCabe(12) = ""                'mensajes
        wCabe(13) = ""                     'modo de venta
        wCabe(14) = FrmPedido.MBox(10)               'fecha de atencion
        wCabe(15) = "00"                   'forma de pago
        wCabe(16) = ""                     'cliente
        wCabe(17) = ""                     'vendedor
        wCabe(18) = 0                      'comision
        If J = 1 Then
           wCabe(19) = Text2(1)           'almacen
        Else
           wCabe(19) = Text2(3)           'almacen
        End If
        wCabe(20) = 0                     'otros gastos
        wCabe(21) = 0                     'nota pedido
        wCabe(22) = 0                     'orden de compra
        wCabe(23) = 0                     'autorizacion
        wCabe(24) = 0                     'dias pago
        wCabe(25) = 0                     'Total Cantidad
        wCabe(26) = 0                     'Total Bruto
        wCabe(27) = 0                     'total fletes --T.D.
        wCabe(28) = 0                     'Total Igv
        wCabe(29) = 0         'Neto a Facturar
        wCabe(30) = ""             'entrega pedido
        wCabe(31) = ""                    'nombre cliente
        wCabe(32) = ""                    'direccion
        wCabe(33) = ""                    'ruc
        wCabe(34) = FrmPedido.MBox(10)                           'fechafactura
        wCabe(35) = 0                     'Total Descuentos Globales
        wCabe(36) = 0                    'Total Descuentos Cliente
        wCabe(37) = 0                  'Total Descuentos Oficina
        wCabe(38) = 0                       'Total Descuentos Item
        wCabe(39) = 0                      'Total Descuentos Linea
        wCabe(40) = 0                      'Total Descuentos x Promocion
        
        Set acmd.ActiveConnection = VGgeneral
        acmd.CommandType = adCmdStoredProc
        acmd.CommandText = "vt_ingresoalma_pro"
        acmd.CommandTimeout = 0
        acmd.Prepared = True
        With acmd
            .Parameters("@base") = VGCNx.DefaultDatabase
            .Parameters("@tabla") = "movalmcab"
            If J = 1 Then
              .Parameters("@tipo") = "2"
            Else
              .Parameters("@tipo") = "3"
            End If
            .Parameters("@puntovta") = wCabe(1)
            .Parameters("@numero") = wCabe(2)
            .Parameters("@factura") = wCabe(3)
            .Parameters("@boleta") = wCabe(4)
            .Parameters("@guia") = wCabe(5)
            .Parameters("@dsctoglobal") = wCabe(6)
            .Parameters("@dsctoppago") = wCabe(7)
            .Parameters("@dsctovtaofi") = wCabe(8)
            .Parameters("@moneda") = IIf(wCabe(9) = g_TipoSol, "S", "D")
            .Parameters("@tipocambio") = wCabe(10)
            .Parameters("@listaprecio") = wCabe(11)
            .Parameters("@mensaje") = wCabe(12)
            .Parameters("@modoventa") = wCabe(13)
            .Parameters("@fecha") = wCabe(14)
            .Parameters("@formapago") = wCabe(15)
            .Parameters("@cliente") = wCabe(16)
            .Parameters("@vendedor") = wCabe(17)
            .Parameters("@porcomision") = wCabe(18)
            .Parameters("@almacen") = wCabe(19)
            .Parameters("@totalotros") = wCabe(20)
            .Parameters("@notaped") = wCabe(21)
            .Parameters("@ordencompra") = wCabe(22)
            .Parameters("@autoriza") = wCabe(23)
            .Parameters("@diaspago") = wCabe(24)
            .Parameters("@totalitem") = wCabe(25)
            .Parameters("@totalbruto") = wCabe(26)
            .Parameters("@totalflete") = wCabe(27)
            .Parameters("@totalimpuesto") = wCabe(28)
            .Parameters("@totalneto") = wCabe(29)
            .Parameters("@usuario") = g_usuario
            .Parameters("@fechaactual") = Date
            .Parameters("@totaldsctoxlinea") = wCabe(39)
            .Parameters("@montodsctoppago") = 0
            .Parameters("@entregapedido") = wCabe(30)
            .Parameters("@razon") = wCabe(31)
            .Parameters("@direccion") = wCabe(32)
            .Parameters("@ruc") = wCabe(33)
            .Parameters("@fechafactura") = wCabe(34)
            .Parameters("@TDGlobal") = wCabe(35)
            .Parameters("@TDCliente") = wCabe(36)
            .Parameters("@TDOficina") = wCabe(37)
            .Parameters("@TDItem") = wCabe(38)
            .Parameters("@TDPromo") = wCabe(40)
            .Parameters("@empresa") = VGParametros.empresacodigo
        End With
        acmd.Execute
        Set acmd = Nothing
        DoEvents
          
       '** Actualizamos detalle
    
        Set acmd.ActiveConnection = VGgeneral
        acmd.CommandType = adCmdStoredProc
        acmd.CommandTimeout = 0
        acmd.CommandText = "vt_ingresodetallealma_pro"
        acmd.Prepared = True
        With acmd
            .Parameters("@base") = VGCNx.DefaultDatabase
            .Parameters("@tabla") = "movalmdet" ' nsql
            If J = 1 Then
              .Parameters("@tipo") = "2"
            Else
              .Parameters("@tipo") = "3"
            End If
            .Parameters("@item") = 1
            .Parameters("@numero") = wCabe(2)
            .Parameters("@producto") = Trim(Text2(0))   'Trim(MBox2(1).Text)
            .Parameters("@unidad") = ""
            .Parameters("@cantidad") = Trim(Text2(2))   'Trim(txtcanti(1).Text)
            .Parameters("@preciopacto") = 0
            .Parameters("@dsctoxitem") = 0
            .Parameters("@importebruto") = 0
            .Parameters("@porcomision") = 0
            .Parameters("@mdsctoitem") = 0
            .Parameters("@mdsctoxlinea") = 0
            .Parameters("@mdsctoxprom") = 0
            .Parameters("@mimpor") = 0
            .Parameters("@unidadref") = Trim(Text2(4))   'rtxtcanti(1)
             If J = 1 Then
               .Parameters("@almacen") = Trim(Text2(1))
             Else
               .Parameters("@almacen") = Trim(Text2(3))
             End If
        End With
        acmd.Execute
        Set acmd = Nothing
                    
        Set acmd.ActiveConnection = VGgeneral
        acmd.CommandType = adCmdStoredProc
        acmd.CommandTimeout = 0
        acmd.CommandText = "vt_actualizoalma_pro"
        acmd.Prepared = True
        With acmd
            .Parameters("@basedes") = CStr(VGCNx.DefaultDatabase)
            .Parameters("@almacen") = wCabe(19)
            If J = 1 Then
              .Parameters("@tipo") = "1"
            Else
              .Parameters("@tipo") = "2"
            End If
            .Parameters("@articulo") = Trim(Text2(0))   'Trim(MBox2(1).Text)
            .Parameters("@cantidad") = Trim(Text2(2))   'txtcanti(1)
        End With
        acmd.Execute
        Set acmd = Nothing
    Next J
    
'    Set acmd.ActiveConnection = VGcnx
'    acmd.CommandText = "al_actualizaproducto_pro"
'    acmd.CommandType = adCmdStoredProc
'    acmd.Prepared = True
'    With acmd
'        .Parameters("@baseini") = VGcnx.DefaultDatabase
'        .Parameters("@basefin") = VGcnx.DefaultDatabase
'        .Parameters("@almacen") = Trim(Text2(3))
'        .Parameters("@articulo") = Trim(Text2(0))
'        .Parameters("@tipo") = "1"
'    End With
'    acmd.Execute
'    Set acmd = Nothing
     VGCNx.Execute "Delete From vt_producto where productocodigo='" & Trim(Text2(0)) & "' and almacencodigo='" & Trim(Text2(3)) & "'"
     VGCNx.Execute "Delete From listapre1 where productocodigo='" & Trim(Text2(0)) & "' and almacencodigo='" & Trim(Text2(3)) & "'"
     
     VGCNx.Execute "INSERT INTO VT_PRODUCTO " & _
                "( productocodigo,productodescripcion,productodescrcorta," & _
                " grupovtacodigo,productofamiliacodigo,productocategoriacodigo, " & _
                " productotipo,unidadcodigo,productoporcimpto,productoestunidreferencia," & _
                " unidadreferencial,unidadfactorconv,productoprecvta," & _
                " monedacodigo,USUARIOCODIGO,FECHAACT,almacencodigo)" & _
                "  SELECT DISTINCT " & _
                "ACODIGO as productocodigo,ADESCRI as productodescripcion," & _
                "left(ADESCRI,30) as productodescrcorta,AGRUPO as grupovtacodigo," & _
                "AFAMILIA as productofamiliacodigo,0 as productocategoriacodigo," & _
                " ATIPO as productotipo,AUNIDAD as unidadcodigo," & _
                "0 as productoporcimpto,0 as productoestunidreferencia," & _
                "'R' as unidadreferencial,0 as unidadfactorconv,0 AS productoprecvta," & _
                " CASE ACODMON WHEN 'MN' THEN '01' ELSE '02' END  as monedacodigo," & _
                "'CAMTEX' as usuariocodigo," & _
                "GETDATE() as fechaact," & _
                "'" & Trim(Text2(3)) & "' as almacencodigo " & _
                "  From [" & Trim(VGCNx.DefaultDatabase) & "].DBO.MAEART Where acodigo ='" & Trim(Text2(0)) & "'"
        
     VGCNx.Execute "INSERT INTO listapre1 " & _
                "( productocodigo,productodescripcion,productodescrcorta," & _
                " grupovtacodigo,productofamiliacodigo,productocategoriacodigo, " & _
                " productotipo,unidadcodigo,productoporcimpto,productoestunidreferencia," & _
                " unidadreferencial,unidadfactorconv,productoprecvta," & _
                " monedacodigo)" & _
                "  SELECT DISTINCT " & _
                "ACODIGO as productocodigo,ADESCRI as productodescripcion," & _
                "left(ADESCRI,30) as productodescrcorta,AGRUPO as grupovtacodigo," & _
                "AFAMILIA as productofamiliacodigo,0 as productocategoriacodigo," & _
                " ATIPO as productotipo,AUNIDAD as unidadcodigo," & _
                "0 as productoporcimpto,0 as productoestunidreferencia," & _
                "'R' as unidadreferencial,0 as unidadfactorconv,0 AS productoprecvta," & _
                " CASE ACODMON WHEN 'MN' THEN '01' ELSE '02' END  as monedacodigo " & _
                "  From [" & Trim(VGCNx.DefaultDatabase) & "].DBO.MAEART Where acodigo ='" & Trim(Text2(0)) & "'"
   
    
    MsgBox "Traslado de almacen satisfactorio...!!", vbInformation, "AVISO"
    Frame3.Visible = False
    Text1 = Trim(Text3)
    Text1.SetFocus
    

nerror:
    If Err Then
        MsgBox "Error: " & Err.Number & "-" & Err.Description
        Err = 0
        Exit Sub
    End If
End Sub

Private Sub cTraslado_Click()
  Limpiartexto Text2, 0, 4
  Label4(0) = "": Label4(1) = ""
  aflag = 1
  Frame3.Visible = True
  Text4.SetFocus
End Sub

Private Sub DGrid1_HeadClick(ByVal ColIndex As Integer)
  Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, ColIndex + 1, nlongi, xCondi)
  ConfigGrid xtipo
End Sub


Private Sub DGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    If DGrid1.Row >= 0 Then
        Select Case xtipo
          Case 1, 3
            nAyuda = DGrid1.Columns(0).Text
            nDetalle = Trim(DGrid1.Columns(2).Text)
          Case 2
            nAyuda = DGrid1.Columns(2).Text
            nDetalle = ""
         End Select
    End If
    Unload Me
  Else
    DGrid1.SetFocus
  End If
  xdata = ""
End Sub

Private Sub Form_Load()
  MostrarForm Me, "C"
    
  DoEvents
  nlongi(1) = 0
  DoEvents
  
  cTraslado.Visible = False
  xdato = UCase(xdato)
  If Trim(Escadena(xdata)) = "1" Then
     Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, "adescri like '" & xdato & "%'")
     Text1 = Escadena(xdato): Text1.SelStart = Len(Trim(Text1))
  ElseIf Trim(Escadena(xdata)) = "2" Then
     If VGParamSistem.stockcomp = 1 Then
        If VGParamSistem.kitvirtual = 1 Then
           Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, "adescri like '%" & xdato & "%' ")
         Else
         Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, "adescri like '%" & xdato & "%' and stalma='" & FrmPedido.Ctr_Ayuda3.xclave & "' AND STSKDIS-stskcom>0")
        End If
      Else
        If FrmPedidoVentanilla.Chkentrega.Value = 1 Then
             Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, "adescri like '" & xdato & "%'")
          Else
             Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, "adescri like '" & xdato & "%' and stalma='" & FrmPedidoVentanilla.Ctr_Ayuda3.xclave & "' AND STSKDIS>0")
        End If
 '      Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, "adescri like '" & xdato & "%' ")
     End If
     Text1 = Escadena(xdato): Text1.SelStart = Len(Trim(Text1))
     cTraslado.Visible = True
 ElseIf Trim(Escadena(xdata)) = "3" Then
     Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, "adescri like '" & xdato & "%' and stalma='" & FrmCotizacionLibre.Ctr_Ayuda3.xclave & "'")
     Text1 = Escadena(xdato): Text1.SelStart = Len(Trim(Text1))
 ElseIf Trim(Escadena(xdata)) = "4" Then
     Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, "adescri like '" & xdato & "%' and stalma='" & FrmTraslado.Ctr_Ayuda1.xclave & "' AND STSKDIS>0")
     Text1 = Escadena(xdato): Text1.SelStart = Len(Trim(Text1))
  Else
     Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, xCondi)
  End If
  ConfigGrid xtipo
End Sub


Public Property Let BFiltro(ByRef campos)
   Dim f As Integer
   Dim tam As Integer
   tam = UBound(campos)
   ReDim nfiltra(tam, 2)
   Combo1.Clear
   For f = 1 To UBound(campos)
      nfiltra(f, 1) = campos(f, 1)
      nfiltra(f, 2) = campos(f, 2)
      Combo1.AddItem campos(f, 1)
   Next f
   If xdata Like "[1234]" Then
     Combo1.ListIndex = 1
'     Text1.SetFocus
   End If
   
   ConfigGrid xtipo
End Property


Public Function ConfigGrid(xtipo As Integer)
   Dim J As Integer
      
   Select Case xtipo
     Case 1
        DGrid1.Columns(0).Width = 1200
        DGrid1.Columns(1).Width = 5100
        If DGrid1.Columns.Count = 3 Then
           DGrid1.Columns(2).Width = 1200
        End If
        If DGrid1.Columns.Count = 4 Then
           DGrid1.Columns(1).Width = 1200
           DGrid1.Columns(2).Width = 4500
           DGrid1.Columns(3).Width = 1000
           
        End If
        
     Case 2
        DGrid1.Columns(0).Width = 600
        DGrid1.Columns(1).Width = 1000
        DGrid1.Columns(2).Width = 1200
        DGrid1.Columns(3).Width = 1000
        DGrid1.Columns(4).Width = 3000
        DGrid1.Columns(5).Width = 1200
        DGrid1.Columns(5).NumberFormat = "##,###,##0.00"
   End Select
   DGrid1.Refresh
End Function

Public Property Let TipoForma(pdata As Integer)
   xtipo = pdata
End Property

Public Property Let BTabla(pdata As String)
   xtabla = pdata
End Property

Public Property Let BConexion(ByRef pdata)
   Set vcon = pdata
End Property


Public Property Let BCampos(pdata As String)
   xCampos = pdata
End Property



Public Property Let BOrden(pdata As String)
   xOrden = IIf(IsNull(pdata), "", Trim(pdata))
   AOrden = IIf(IsNull(pdata), "", Trim(pdata))
End Property

Public Property Let BCondi(pdata As String)
   xCondi = IIf(IsNull(pdata), "", Trim(pdata))
   ACondi = IIf(IsNull(pdata), "", Trim(pdata))
  
End Property

Public Property Let Bdata(ByRef pdata)
   xdata = pdata
End Property

Public Property Let Bdato(ByRef pdata)
   xdato = pdata
End Property

Private Sub TDBGrid1_Click()
   On Error Resume Next
    If TDBGrid1.ApproxCount > 0 Then
        Text2(0).Text = Trim(TDBGrid1.Columns(0).Text)
        Text2(1).Text = Text4    'almacen origen
        Text2(2).Text = numero(FrmPedidoVentanilla.MBox2(0).Text)
        Text2(3).Text = FrmPedidoVentanilla.Ctr_Ayuda3.xclave  'almacen destino
        Text2(4).Text = numero(0)
        Text2(4).SetFocus
    End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  If KeyCode = 13 Then
      Call TDBGrid1_Click
  End If

End Sub

Private Sub Text1_Change()
    Dim posi As Integer
    posi = Combo1.ListIndex + 1
    If Len(Trim(xCondi)) = 0 Then
       Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, nfiltra(posi, 2) & " Like '%" & Text1 & "%'")
    Else
       Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, "##xx_ventas", DGrid1, xCampos, xOrden, nlongi, xCondi & " and " & nfiltra(posi, 2) & " Like '%" & Trim(Text1) & "%'")
    End If
    ConfigGrid xtipo
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  Dim posi As Integer
  Dim rs As ADODB.Recordset
  
  If KeyAscii = 13 Then
     If Len(Trim(Text1)) > 0 Then
        posi = Combo1.ListIndex + 1
        If Len(Trim(xCondi)) = 0 Then
           Call DLLGENERALAYUDA.ListarEnTDBGRID(VGCNx, xtabla, DGrid1, xCampos, xOrden, nlongi, nfiltra(posi, 2) & " Like '" & Text1 & "%'")
        Else
           Call DLLGENERALAYUDA.ListarEnTDBGRID(VGCNx, xtabla, DGrid1, xCampos, xOrden, nlongi, xCondi & " and " & nfiltra(posi, 2) & " Like '%" & Text1 & "%'")
        End If
     Else
        Call DLLGENERALAYUDA.ListarEnTDBGRID(VGCNx, xtabla, DGrid1, xCampos, xOrden, nlongi, xCondi)
     End If
     ConfigGrid xtipo
     DGrid1.SetFocus
  End If

End Sub


Private Sub Text2_GotFocus(Index As Integer)
   Call DLLGENERALAYUDA.Enfoquetexto(Text2(Index))
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub Text3_Change()
    xdato = UCase(xdato)
    'Set TDBGrid1.DataSource = Nothing
    
    Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, TDBGrid1, xCampos, xOrden, nlongi, "adescri like '" & Trim(UCase(Text3.Text)) & "%' and stalma='" & Trim(Text4) & "' and stskdis>0")
    With TDBGrid1
        .Columns(0).Width = 1000
        .Columns(1).Width = 5000
        If .Columns.Count = 3 Then
           .Columns(2).Width = 1200
        End If
        .Refresh
    End With
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  If KeyAscii = 13 Then
     TDBGrid1.SetFocus
  End If
End Sub

Private Sub Text4_Change()
   If Len(Trim(Text4)) = 0 Then
       Text2(1) = ""
   Else
      Text2(1) = Trim(Text4)
   End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Trim(Text4) = Trim(FrmPedidoVentanilla.Ctr_Ayuda3.xclave) Then
        MsgBox " No Puede ser el mismo almacen...Verifique!!!", vbInformation, "AVISO"
        Text4 = "": Text4.SetFocus
        Exit Sub
     End If
     Text3.SetFocus
   End If
End Sub

Private Sub Text4_LostFocus()
     If Trim(Text4) = Trim(FrmPedidoVentanilla.Ctr_Ayuda3.xclave) Then
'        Call DLLGENERALAYUDA.Enfoquetexto(Text4)
'        MsgBox " No Puede ser el mismo almacen...Verifique!!!", vbInformation, "AVISO"
'        Text4 = "": Text4.SetFocus
'        Exit Sub
     End If
End Sub

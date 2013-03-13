VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmGuiaSal1 
   Caption         =   "Guia de salida"
   ClientHeight    =   6990
   ClientLeft      =   1215
   ClientTop       =   1920
   ClientWidth     =   12300
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   12300
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrmValida 
      BackColor       =   &H00C9955A&
      BorderStyle     =   0  'None
      Caption         =   "Pendientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   1530
      TabIndex        =   71
      Top             =   2025
      Visible         =   0   'False
      Width           =   8940
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000009&
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7470
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   225
         Width           =   1275
      End
      Begin TrueOleDBGrid70.TDBGrid GridP 
         Height          =   1650
         Left            =   180
         TabIndex        =   72
         Top             =   585
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   2910
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
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   8280
         Top             =   1530
      End
      Begin VB.Label LblFal 
         Caption         =   "Label15"
         Height          =   330
         Index           =   5
         Left            =   2025
         TabIndex        =   86
         Top             =   4590
         Width           =   2310
      End
      Begin VB.Label LblFal 
         Caption         =   "Label15"
         Height          =   330
         Index           =   4
         Left            =   2025
         TabIndex        =   85
         Top             =   4230
         Width           =   2310
      End
      Begin VB.Label LblFal 
         Caption         =   "Label15"
         Height          =   330
         Index           =   3
         Left            =   2025
         TabIndex        =   84
         Top             =   3870
         Width           =   2310
      End
      Begin VB.Label LblFal 
         Caption         =   "Label15"
         Height          =   330
         Index           =   2
         Left            =   2070
         TabIndex        =   83
         Top             =   3555
         Width           =   2310
      End
      Begin VB.Label LblFal 
         Caption         =   "Label15"
         Height          =   330
         Index           =   1
         Left            =   2115
         TabIndex        =   82
         Top             =   3195
         Width           =   2310
      End
      Begin VB.Label LblPro 
         Caption         =   "Label15"
         Height          =   330
         Index           =   5
         Left            =   405
         TabIndex        =   81
         Top             =   4680
         Width           =   1365
      End
      Begin VB.Label LblPro 
         Caption         =   "Label15"
         Height          =   330
         Index           =   4
         Left            =   405
         TabIndex        =   80
         Top             =   4230
         Width           =   1365
      End
      Begin VB.Label LblPro 
         Caption         =   "Label15"
         Height          =   330
         Index           =   3
         Left            =   405
         TabIndex        =   79
         Top             =   3825
         Width           =   1365
      End
      Begin VB.Label LblPro 
         Caption         =   "Label15"
         Height          =   330
         Index           =   2
         Left            =   405
         TabIndex        =   78
         Top             =   3465
         Width           =   1365
      End
      Begin VB.Label LblPro 
         Caption         =   "Label15"
         Height          =   330
         Index           =   1
         Left            =   450
         TabIndex        =   77
         Top             =   3105
         Width           =   1365
      End
      Begin VB.Label LblFal 
         Caption         =   "Label15"
         Height          =   330
         Index           =   0
         Left            =   2160
         TabIndex        =   76
         Top             =   2790
         Width           =   2310
      End
      Begin VB.Label LblPro 
         Caption         =   "Label15"
         Height          =   330
         Index           =   0
         Left            =   450
         TabIndex        =   75
         Top             =   2745
         Width           =   1365
      End
      Begin VB.Label Lblmensaje 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "El stock de los siguientes productos no cubren la cantidad pedida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   195
         TabIndex        =   74
         Top             =   255
         Width           =   6330
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         Height          =   2490
         Left            =   45
         Top             =   45
         Width           =   8835
      End
   End
   Begin VB.Frame FrmPen 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Pendientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   1530
      TabIndex        =   63
      Top             =   1170
      Visible         =   0   'False
      Width           =   8400
      Begin TrueOleDBGrid70.TDBGrid TDBGrid 
         Height          =   1380
         Left            =   135
         TabIndex        =   67
         Top             =   405
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   2434
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
      Begin VB.CommandButton CmdCan 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   45
         Width           =   1275
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guias Pendientes"
         Height          =   195
         Left            =   150
         TabIndex        =   64
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1740
      Left            =   60
      TabIndex        =   32
      Top             =   3450
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   3069
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FormatString    =   $"FrmGuiaSal.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3510
      Left            =   60
      TabIndex        =   22
      Top             =   -60
      Width           =   11805
      Begin VB.TextBox TxtCon 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1620
         Width           =   4590
      End
      Begin VB.Frame Frame2 
         Caption         =   "Direcciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   6750
         TabIndex        =   56
         Top             =   3330
         Visible         =   0   'False
         Width           =   4995
         Begin MSDataGridLib.DataGrid dbGrid1 
            Height          =   1380
            Left            =   135
            TabIndex        =   59
            Top             =   225
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   2434
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
               MarqueeStyle    =   3
               BeginProperty Column00 
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   3135.118
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cAcepta 
            Caption         =   "&Acepta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4260
            TabIndex        =   58
            Top             =   2160
            Width           =   1155
         End
         Begin VB.CommandButton cCerrar 
            Caption         =   "Cerrar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5550
            TabIndex        =   57
            Top             =   2160
            Width           =   1365
         End
      End
      Begin VB.TextBox Texttipdoc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3060
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   11
         Top             =   1950
         Width           =   495
      End
      Begin VB.TextBox Txtnrodoc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7785
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   60
         Top             =   1290
         Width           =   1416
      End
      Begin VB.CommandButton Bdire 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6210
         TabIndex        =   55
         Top             =   1275
         Width           =   375
      End
      Begin VB.TextBox tx_ordfab 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8250
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1950
         Width           =   1416
      End
      Begin VB.TextBox tx_codmaq 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   16
         Top             =   3045
         Width           =   1416
      End
      Begin VB.CheckBox ChkTalla 
         Alignment       =   1  'Right Justify
         Caption         =   "Salidas por Talla"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6780
         TabIndex        =   20
         Top             =   2280
         Width           =   1680
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5055
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1950
         Width           =   1416
      End
      Begin VB.ComboBox CmbSerie 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5610
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   570
         Width           =   915
      End
      Begin VB.CommandButton CmdGrabarCab 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11250
         TabIndex        =   21
         Top             =   1320
         Width           =   435
      End
      Begin VB.TextBox TxTransa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   0
         Top             =   210
         Width           =   645
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   2
         Top             =   570
         Width           =   645
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7635
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Top             =   570
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   6
         Top             =   930
         Width           =   1305
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7785
         TabIndex        =   7
         Top             =   930
         Width           =   3915
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   8
         Top             =   1290
         Width           =   4590
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1950
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10665
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   17
         Top             =   1290
         Width           =   495
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   7605
         TabIndex        =   1
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   69074945
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin VB.TextBox Text13 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9045
         TabIndex        =   19
         Top             =   180
         Width           =   2025
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuVendedor 
         Height          =   315
         Left            =   1470
         TabIndex        =   14
         Top             =   2295
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   200
         NomTabla        =   "vt_vendedor"
         TituloAyuda     =   "Ayuda de Vendedores"
         ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
         XcodCampo       =   "vendedorcodigo"
         XListCampo      =   "vendedornombres"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "vendedorcodigo,vendedornombres"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuTransporte 
         Height          =   315
         Left            =   1470
         TabIndex        =   15
         Top             =   2670
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   556
         XcodMaxLongitud =   11
         xcodwith        =   800
         NomTabla        =   "al_transporte"
         TituloAyuda     =   "Ayuda de Transporte"
         ListaCampos     =   "TRACODIGO(1),TRANOMBRE(1)"
         XcodCampo       =   "TRACODIGO"
         XListCampo      =   "TRANOMBRE"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "TRACODIGO,TRANOMBRE"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaEmpresa 
         Height          =   450
         Left            =   1485
         TabIndex        =   5
         Top             =   900
         Visible         =   0   'False
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   794
         XcodMaxLongitud =   11
         xcodwith        =   1000
         NomTabla        =   "cp_proveedor"
         TituloAyuda     =   "Ayuda de Proveedores"
         ListaCampos     =   "clientecodigo(1),clienteruc(1),clienterazonsocial(1),clientedireccion(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "clientecodigo,clienteruc,clienterazonsocial,clientedireccion"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Contacto :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   70
         Top             =   1620
         Width           =   765
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Almacen Destino :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9315
         TabIndex        =   23
         Top             =   1320
         Width           =   1830
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Nro Pedido :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6840
         TabIndex        =   65
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9955A&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   2190
         TabIndex        =   62
         Top             =   570
         Width           =   2535
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9955A&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   2190
         TabIndex        =   61
         Top             =   210
         Width           =   4245
      End
      Begin VB.Label LblCC 
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
         Left            =   8610
         TabIndex        =   40
         Top             =   2220
         Width           =   2910
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Equip./Maqui :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   3060
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Orden Fabricación :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   6720
         TabIndex        =   38
         Top             =   2010
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Orden Compra :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   2010
         Width           =   1155
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Centro Costo :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   3750
         TabIndex        =   39
         Top             =   2010
         Width           =   1065
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Top             =   2310
         Width           =   795
      End
      Begin VB.Label lblSerie 
         Caption         =   "Serie :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   35
         Top             =   630
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Transportista :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   2700
         Width           =   1065
      End
      Begin VB.Label Label19 
         Caption         =   "N°"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7770
         TabIndex        =   33
         Top             =   1020
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Cliente :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   990
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Referencia :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   630
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Doc. :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         TabIndex        =   29
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Transacción :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   270
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Num. Doc :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6795
         TabIndex        =   27
         Top             =   630
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Razon Social :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6705
         TabIndex        =   26
         Top             =   945
         Width           =   1005
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   1320
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   60
      TabIndex        =   50
      Top             =   5190
      Width           =   11805
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10140
         TabIndex        =   52
         Top             =   90
         Width           =   1425
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   51
         Top             =   120
         Width           =   1395
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total  Cantidad :"
         Height          =   195
         Index           =   0
         Left            =   8715
         TabIndex        =   54
         Top             =   150
         Width           =   1365
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total  Items :"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   53
         Top             =   150
         Width           =   1125
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   2370
      TabIndex        =   41
      Top             =   5790
      Width           =   7515
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   1005
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   90
         Width           =   1155
      End
      Begin VB.CommandButton CmdGrabarDet 
         Caption         =   "&Grabar"
         Height          =   1005
         Left            =   4755
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   90
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Eliminar"
         Height          =   1005
         Left            =   3180
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   90
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   1005
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   90
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Adicionar"
         Height          =   1005
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   90
         Visible         =   0   'False
         Width           =   1155
      End
   End
   Begin VB.Frame FrameComentario 
      Caption         =   "Comentarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   2220
      TabIndex        =   46
      Top             =   3450
      Visible         =   0   'False
      Width           =   8316
      Begin VB.CommandButton CmdComCan 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   49
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton CmdComGrabar 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   48
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxComentario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   600
         MultiLine       =   -1  'True
         TabIndex        =   47
         Top             =   240
         Width           =   5655
      End
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   210
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2070
      TabIndex        =   69
      Text            =   "Text10"
      Top             =   4890
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox TxtTransp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1395
      TabIndex        =   68
      Top             =   5115
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "FrmGuiaSal1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Este formulario es utilizado como Guia de Remision o Devolucion cualquier
'cambio se tiene que validar en ambos formulario
'VGGuiaSal es la variable global que indica si es guia o devolucion

Option Explicit
''Dim db As Database
Dim TCamb As Double
Dim hubo_error As Boolean
Dim direccion  As String
Dim xserie As String * 1
Dim VGDllGeneral As New dllgeneral.dll_general
Dim rg As New ADODB.Recordset

'Dim DbAux As Database
Dim numsal As Double       'Numero consecutivo de guia de remision
Dim Unid As String
Dim nument As Long       'Numero consecutivo de nota de ingreso
Dim precioprom As Double
Dim CANTIDAD As Double
Dim canttemp As Double
Dim Campo As String * 2  'Indica el tipo de transaccion
Dim contador As Long     'Indica el item del flex
Dim auxdisp As Long
Dim cantidadDEV As Double
Dim numserie As String    'Numero de guia remision
Dim tipo, Codigo2 As String
Dim TT_CONTADOR As Integer
Dim salir As Boolean      'Numero de la serie de guia de remision
Dim Serie As String
Dim AlmacenRF As String
Dim cTransa As String
Dim ruc As String
Dim EstadoDevolucion As String
Dim WithEvents Conex As ADODB.Connection
Attribute Conex.VB_VarHelpID = -1
Dim Completo As Boolean
Dim flagserie, flaglote As String * 1
Dim dato_invalido As Boolean
Dim serie_lote As String
Dim Rs As ADODB.Recordset  'agregado
Dim Rs2 As ADODB.Recordset  'agregado
Dim Cliente As Boolean, Requerimiento As Boolean
Dim t As Integer

Private Sub pro_xserie()
  If flagserie = "S" Then
        xserie = "S"
        Exit Sub
  End If
  If flaglote = "S" Then
        xserie = "N"
        Exit Sub
  End If
  xserie = "X"
End Sub
Function coduso(dato As String) As String
   Dim rsql As String
   Dim Rs As New ADODB.Recordset
   rsql = "select UM_ABREV from TabUniMed where UM_NOMBRE ='" & dato & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set Rs = VGCNx.Execute(rsql)
   If Rs.RecordCount = 0 Then
    coduso = ""
   Else
    coduso = Rs(0)
   End If
   Rs.Close
   Set Rs = Nothing
End Function

Function Nombre_Unidad(dato As String) As String
   Dim rsql As String
   Dim Rs As New ADODB.Recordset
   rsql = "select UM_NOMBRE from TabUniMed where UM_ABREV ='" & dato & "'" '
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set Rs = VGCNx.Execute(rsql)
   If Rs.RecordCount = 0 Then
     Nombre_Unidad = ""
   Else
     Nombre_Unidad = Rs(0)
   End If
   Rs.Close
   Set Rs = Nothing
End Function
Private Sub limpia()
   Label11 = ""
'   TxDescri = ""
'   lblUniEst = ""
'   lblPreciofin = ""
'   TxtArticulo.text = ""
'   Text4.text = ""
'   lbcantstk = ""
'   Text3.text = ""
'   TxtCantidad.Enabled = True
'   TxtCantidad.text = ""
'   Text6.BackColor = &H80000009
'   TxtArticulo.Enabled = True
'   Text6.Enabled = True
'   Text6 = ""
'   Text6.Enabled = False
'   Label11.Visible = False
'   lbEtiNum.Visible = False
'   Command1.Enabled = False
   'txEquip = ""
   'txccosto = ""
   'TxordFab = ""
'   Combo1.Clear
   
End Sub

Private Sub Bdire_Click()
On Error Resume Next

Set rg = Nothing
Set dbGrid1.DataSource = Nothing

Set rg = VGCNx.Execute("select cliedirnumero as Nro,cliedirdireccion as Direccion from vt_clientedireccion where clientecodigo='" & Text5 & "'")
If rg.RecordCount > 0 Then
    Frame2.Visible = True
    Set dbGrid1.DataSource = rg
    dbGrid1.Refresh
Else
   Frame2.Visible = False
End If
   
End Sub

Private Sub cAcepta_Click()
   Text7 = IIf(IsNull(dbGrid1.Columns(1).text), "", dbGrid1.Columns(1).text)
   Frame2.Visible = False
   Text7.SetFocus
End Sub

Private Sub cCerrar_Click()
   Frame2.Visible = False
   Text7.SetFocus
End Sub

Private Sub CmbSerie_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   SendKeys "{tab}"
   KeyAscii = 0
 End If
End Sub

Private Sub CmdCan_Click()
FrmPen.Visible = False
End Sub

Private Sub CmdComCan_Click()
'CANCELA EL COMENTARIO
Dim rpta As Integer
FrameComentario.Visible = False
crtlvisible (True)
rpta = MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso")
If rpta = vbYes Then
  '  imprimir
  imprimirguias
End If
End Sub

Private Sub CmdComGrabar_Click()
'GRABA EL COMENTARIO DE LA GUIA
Dim rsql As String
Dim rpta As String
 
rsql = "Update MovAlmCab set CAGLOSA = '" & TxComentario & "' "
rsql = rsql & "Where  CAALMA = '" & VGAlma & "'AND  CATD= '" & tipo & "' AND CANUMDOC = '" & numserie & "'" '
VGCNx.Execute rsql
FrameComentario.Visible = False
crtlvisible (True)
rpta = MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso")
If rpta = vbYes Then
    ' imprimir
    imprimirguias
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

'Agregar
Private Sub Command1_Click()
VGSeleccion = 1
FormCreacionSal.Caption = "Ingreso de Articulos"
buscar
FormCreacionSal.Show 1
End Sub
'Modificar
Private Sub Command2_Click()
If MSFlexGrid1.Rows = 1 Then
   MsgBox "No existe registros para Modificar", vbInformation, "Información"
   Exit Sub
End If
If VGGuiaSal Then
   VGSeleccion = 2
   FormCreacionSal.Caption = "Modificación de Articulos"
   buscar
   FormCreacionSal.Show 1
Else
   MSFlexGrid1_Click
End If
End Sub
'Eliminar
Private Sub Command3_Click()
Dim I As Integer
If MSFlexGrid1.Rows = 1 Then
    MsgBox "No existe registros para Eliminar", vbInformation, "Información"
    Exit Sub
End If
If MsgBox("Desea Eliminar el registro", vbQuestion + vbYesNo, "Información") = vbYes Then
    I = MSFlexGrid1.RowSel
    If MSFlexGrid1.Rows > 2 Then
        MSFlexGrid1.RemoveItem I
    Else
        MSFlexGrid1.Clear
        MSFlexGrid1.Rows = 1
        MSFlexGrid1.Row = 0
        visualizarFG2
        CmdGrabarDet.SetFocus
    End If
End If
End Sub

Private Sub CmdSalirDevol_Click()
'Frame2.Visible = False
Frame1.Visible = True
CmdSalir.SetFocus
End Sub

Private Sub CmdGrabarCab_Click()
 If CmbSerie.text = "" Then
   MsgBox "Seleccione la serie de la guia", vbInformation, "Aviso"
   CmbSerie.SetFocus
   Exit Sub
 End If
 If TxTransa = "" Then
   MsgBox "Ingrese el tipo de transaccion", vbExclamation, "Error"
   TxTransa.SetFocus:    Exit Sub
 ElseIf TxTransa.Enabled And TxTransa.Visible Then
   If Existe(1, TxTransa, "TABTRANSA", "TT_CODMOV", False, "S", "TT_TIPMOV") = False Then
       MsgBox "La Transacción no existe", vbInformation, "Información"
       TxTransa.SetFocus:   Exit Sub
   End If
 End If
 If TxTransa = "TD" Then
   If Trim(Text11) = "" Then
      MsgBox "Debe ingresar el almacen de destino", vbExclamation, "Error"
      Text11.SetFocus
      Exit Sub
   ElseIf Existe(1, Text11, "TABALM", "TAALMA", False) = False Then
      MsgBox "El Almacén no existe", vbInformation, "Información"
      Text11.SetFocus: Exit Sub
   End If
 End If
 
 
If Text1.Visible And Text1.Enabled = True Then
    If IsNumeric(Text1.text) Then
       If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
           If Existe(3, Text1, "CENTRO_COSTOS", "cencost_codigo", False) = False Then
                  MsgBox "La Cuenta no existe", vbInformation, "Mensaje"
                  Text1.SetFocus: Exit Sub
           End If
       End If
    Else
       MsgBox "Ingrese el numero de Centro de Costo", vbInformation, mensaje1
       If Text1.Enabled Then Text1.SetFocus
       Exit Sub
    End If
End If
 muestra
' Frame4.Enabled = True
' Txtarticulo.Enabled = True
' Txtarticulo.SetFocus
 If salir Then Unload Me
End Sub

Private Sub CmdGrabarDet_Click()
On Error GoTo GrabErr1
Dim contador As Integer               ' Grabar
Dim criterio As String
Dim cadena As String
Dim cadena1 As String
Dim cadena2 As String
Dim Aux As String
Dim rpta As Integer
Dim FACTOR As Double
Dim uSql As String
Dim tipo As String * 2
Dim nroguia As String
Dim veces As Integer, I As Integer
Dim cad, ncad As String
Dim rst As New ADODB.Recordset
Dim Rs As New ADODB.Recordset
Dim Productos As String


Set Conex = VGCNx

If Text3.text = "GR" And Text4.text = "" Then
   MsgBox "ATENCION !!! " & Chr(13) & "El Numero de Guias es 0 ", vbCritical, "Sistemas"
   Exit Sub
End If

CANTIDAD = 0:  veces = 0: Productos = ""
'------------------------------------------------------------------------------------------------------
'VALIDACION DE EMISION DE GUIAS
With MSFlexGrid1
    If .Rows > 1 Then
    For I = 1 To .Rows - 1
        
        Set Rs = VGCNx.Execute("select b.productocodigo as Codigo,c.adescri as Producto," _
        & " a.stskdis as Disponible," & .TextMatrix(I, 3) & " as Can_Pedida,Faltantes=(a.stskdis-" & .TextMatrix(I, 3) & ") " _
        & " from " & VGParamSistem.BDEmpresa & ".dbo.stkart a " _
        & " inner join " & VGParamSistem.BDEmpresa & ".dbo.vt_detallepedido b on a.stcodigo=b.productocodigo " _
        & " inner join " & VGParamSistem.BDEmpresa & ".dbo.maeart c on b.productocodigo=c.acodigo " _
        & " where b.pedidonumero='" & Txtnrodoc.text & "' " _
        & " and a.stskdis-" & .TextMatrix(I, 3) & "<0  and stalma='" & Text11 & "' and b.productocodigo='" & Trim(.TextMatrix(I, 0)) & "'")
        If Not Rs.EOF Then
            GridP.DataSource = Rs
            With GridP
                  .Columns(0).Width = 1000
                  .Columns(1).Width = 4000
                  .Columns(2).Width = 900
                  .Columns(3).Width = 1000
                  .Columns(4).Width = 900
            End With
            GridP.Refresh
            MsgBox "ATENCION !!! " & Chr(13) & "NO SE PUEDE EMITIR LA GUIA ", vbCritical, "Sistemas"
            FrmValida.Visible = True
            Timer1.Enabled = True
            Exit Sub
        End If
    
    Next I
    End If
End With
'-----------------------------------------------------------------------------------------------------
If Len(Trim(Text11.text)) = 0 Then
    MsgBox "Falta seleccionar almacen.", vbInformation, "Sistema"
    Text11.SetFocus
    Exit Sub
End If

'Desea grabar el registro
Set rst = Nothing
 
If MSFlexGrid1.Rows = 1 Then Exit Sub
    If Not VGGuiaSal Then
           Aux = Text4
           CmbSerie.AddItem Mid(Text4, 1, 3)
           CmbSerie.ListIndex = 0
           actualiza_guia_dev
           muestra                          'obtiene el numero de guia de salida
    Else
           Aux = Text13
           'En este opción permite cambiar por número de Guia
           numserie = Right(CmbSerie.text & Right(Text4, 8), 11)
           rpta = MsgBox("Es el número correcto de la Guia " & Chr(13) & numserie, vbInformation + vbOKCancel, "Confirmación")
           If rpta = vbCancel Then
                nroguia = LTrim(InputBox$("Ingrese el  número de guia", "Nro Guia"))
                If nroguia = "" Then
                        MsgBox "No se ha grabado la salida", vbInformation, "Aviso"
                        Exit Sub
                End If
                Do While True
                        Serie = Format(CmbSerie.text, "000")
                        numsal = CDbl(nroguia)
                        nroguia = Format(nroguia, String(7, "0"))
                        nroguia = Serie & nroguia
                        If verifica_nro_guia(nroguia) Then          'Si es verdadero es que no esiste el nro guia
                            MsgBox "Ingrese el Número Correcto", vbInformation, "Mensaje"
                            nroguia = LTrim(InputBox$("Ingrese el número de guia", "Nro Guia"))
                            If nroguia = "" Then
                                  MsgBox "No se ha grabado la salida", vbInformation, "Aviso"
                                  Exit Sub
                            End If
                        Else
                            Exit Do
                        End If
                 Loop
                 numserie = nroguia
            End If
            If verifica_nro_guia(numserie) Then
               MsgBox "Ingrese el siguiente consecutivo de " & numserie, vbInformation, "Aviso"
               Exit Sub
            End If
    End If
    
    AlmacenRF = IIf(Text11 <> "", Text11, "")
    cTransa = TxTransa
    grabacabecera
    FACTOR = 1
    contador = 1
    'If hubo_error Then Exit Sub
    'Falta controlar si se pudo grabar puede estar bloqueada  on error goto
    While MSFlexGrid1.Rows > contador
            CANTIDAD = 0
                   
            If UCase(MSFlexGrid1.TextMatrix(contador, 0)) <> "TEXTO" Then
                  If Not VGGuiaSal Then
                         'caso de devoluciones
                         If MSFlexGrid1.TextMatrix(contador, 5) <> "" Then
                                CANTIDAD = Val(MSFlexGrid1.TextMatrix(contador, 5))
                                cantidadDEV = Val(MSFlexGrid1.TextMatrix(contador, 4))
                         Else
                                CANTIDAD = Val(MSFlexGrid1.TextMatrix(contador, 3))
                                cantidadDEV = 0  'indica no hay devolucion en ese item para que no descargue
                         End If
                  Else
                         'caso de guia remision
                         CANTIDAD = Val(MSFlexGrid1.TextMatrix(contador, 3))
                         cantidadDEV = 1
                  End If
                  If MSFlexGrid1.TextMatrix(contador, 7) = "S" Then
                  '     Data2.Recordset("DESERIE") = MSFlexGrid1.TextMatrix(contador, 2)
                           ncad = "INSERT INTO movalmdet " & _
                                  "(DEALMA,DETD,DENUMDOC,DEITEM,DECODIGO,DEDESCRI,DEUNIDAD,DECANTID,DECANTENT,DECENCOS,DEORDFAB,DEQUIPO,DESERIE,DECANREF1)" & _
                                  " VALUES (" & _
                                  "'" & VGAlma & "'," & _
                                  "'GS','" & numserie & "'," & contador & "," & _
                                  "'" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & _
                                  "'" & IIf(MSFlexGrid1.TextMatrix(contador, 1) = "", " ", MSFlexGrid1.TextMatrix(contador, 1)) & "'," & _
                                  "'" & IIf(MSFlexGrid1.TextMatrix(contador, 4) = "", " ", MSFlexGrid1.TextMatrix(contador, 4)) & "'," & CANTIDAD & "," & _
                                  CANTIDAD & ",'" & Trim(MSFlexGrid1.TextMatrix(contador, 8)) & "'," & _
                                  "'" & Trim(MSFlexGrid1.TextMatrix(contador, 9)) & "'," & _
                                  "'" & Trim(MSFlexGrid1.TextMatrix(contador, 10)) & "'," & _
                                  "'" & MSFlexGrid1.TextMatrix(contador, 2) & "'," & _
                                  MSFlexGrid1.TextMatrix(contador, 11) & ")"
                  
                  ElseIf MSFlexGrid1.TextMatrix(contador, 7) = "N" Then
                  '    Data2.Recordset("DELOTE") = MSFlexGrid1.TextMatrix(contador, 2)
                       ncad = "INSERT INTO movalmdet " & _
                                  "(DEALMA,DETD,DENUMDOC,DEITEM,DECODIGO,DEDESCRI,DEUNIDAD,DECANTID,DECANTENT,DECENCOS,DEORDFAB,DEQUIPO,DELOTE,DECANREF1)" & _
                                  " VALUES (" & _
                                  "'" & VGAlma & "'," & _
                                  "'GS','" & numserie & "'," & contador & "," & _
                                  "'" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & _
                                  "'" & IIf(MSFlexGrid1.TextMatrix(contador, 1) = "", " ", MSFlexGrid1.TextMatrix(contador, 1)) & "'," & _
                                  "'" & IIf(MSFlexGrid1.TextMatrix(contador, 4) = "", " ", MSFlexGrid1.TextMatrix(contador, 4)) & "'," & CANTIDAD & "," & _
                                  CANTIDAD & ",'" & Trim(MSFlexGrid1.TextMatrix(contador, 8)) & "'," & _
                                  "'" & Trim(MSFlexGrid1.TextMatrix(contador, 9)) & "'," & _
                                  "'" & Trim(MSFlexGrid1.TextMatrix(contador, 10)) & "'," & _
                                  "'" & MSFlexGrid1.TextMatrix(contador, 2) & "'," & _
                                  MSFlexGrid1.TextMatrix(contador, 11) & ")"
                  
                  Else
                        ncad = "INSERT INTO movalmdet " & _
                                      "(DEALMA,DETD,DENUMDOC,DEITEM,DECODIGO,DEDESCRI,DEUNIDAD,DECANTID,DECANTENT,DECENCOS,DEORDFAB,DEQUIPO,DECANREF1)" & _
                                      " VALUES (" & _
                                      "'" & VGAlma & "'," & _
                                      "'GS','" & numserie & "'," & contador & "," & _
                                      "'" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & _
                                      "'" & IIf(MSFlexGrid1.TextMatrix(contador, 1) = "", " ", MSFlexGrid1.TextMatrix(contador, 1)) & "'," & _
                                      "'" & IIf(MSFlexGrid1.TextMatrix(contador, 4) = "", " ", MSFlexGrid1.TextMatrix(contador, 4)) & "'," & CANTIDAD & "," & _
                                      CANTIDAD & ",'" & Trim(MSFlexGrid1.TextMatrix(contador, 8)) & "'," & _
                                      "'" & Trim(MSFlexGrid1.TextMatrix(contador, 9)) & "'," & _
                                      "'" & Trim(MSFlexGrid1.TextMatrix(contador, 10)) & "'," & _
                                      MSFlexGrid1.TextMatrix(contador, 11) & ")"
                  End If
                
                

                  If TxTransa <> "GF" Then   'Para descargar
                          If Not IsNull(MSFlexGrid1.TextMatrix(contador, 3)) And cantidadDEV <> 0 Then
                                grabastk (contador)
                                If (IsNumeric(MSFlexGrid1.TextMatrix(contador, 6))) > 0 And VGGuiaSal Then
                                       'SE ESTA TOMANDO DEPRECIO COMO COSTO DE PRODUCTO
                                       'Data2.Recordset("DEPRECIO") = Val(MSFlexGrid1.TextMatrix(contador, 6)) * VGTipCamb '******el precio
                                       'Data2.Recordset("DEPRECI1") = Val(MSFlexGrid1.TextMatrix(contador, 5)) * VGTipCamb '******el precio
                                     VGCNx.Execute "INSERT INTO movalmdet " & _
                                                        "(DEALMA,DETD,DENUMDOC,DEITEM,DECODIGO,DEDESCRI,DEUNIDAD,DECANTID,DECANTENT,DECENCOS,DEORDFAB,DEQUIPO," & _
                                                        "DEPRECIO,DEPRECI1,DECANREF1)" & _
                                                        " VALUES (" & _
                                                        "'" & VGAlma & "'," & _
                                                        "'GS','" & numserie & "'," & contador & "," & _
                                                        "'" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & _
                                                        "'" & IIf(MSFlexGrid1.TextMatrix(contador, 1) = "", " ", MSFlexGrid1.TextMatrix(contador, 1)) & "'," & _
                                                        "'" & IIf(MSFlexGrid1.TextMatrix(contador, 4) = "", " ", MSFlexGrid1.TextMatrix(contador, 4)) & "'," & CANTIDAD & "," & _
                                                        CANTIDAD & ",'" & Trim(MSFlexGrid1.TextMatrix(contador, 8)) & "'," & _
                                                        "'" & Trim(MSFlexGrid1.TextMatrix(contador, 9)) & "'," & _
                                                        "'" & Trim(MSFlexGrid1.TextMatrix(contador, 10)) & "'," & _
                                                        Val(MSFlexGrid1.TextMatrix(contador, 6)) * VGTipCamb & "," & _
                                                        Val(MSFlexGrid1.TextMatrix(contador, 5)) * VGTipCamb & "," & _
                                                        Val(MSFlexGrid1.TextMatrix(contador, 11)) & ")"
                                Else
'                                       Data2.Recordset("DEPRECIO") = Val(Format(precioprom, "#####.000000"))
'                                       Data2.Recordset("DEPRECI1") = IIf(IsNumeric(MSFlexGrid1.TextMatrix(contador, 5)), Val(MSFlexGrid1.TextMatrix(contador, 5)) * VGTipCamb, 0)
                                     VGCNx.Execute "INSERT INTO movalmdet " & _
                                                        "(DEALMA,DETD,DENUMDOC,DEITEM,DECODIGO,DEDESCRI,DEUNIDAD,DECANTID,DECANTENT,DECENCOS,DEORDFAB,DEQUIPO," & _
                                                        "DEPRECIO,DEPRECI1,DECANREF1)" & _
                                                        " VALUES (" & _
                                                        "'" & VGAlma & "'," & _
                                                        "'GS','" & numserie & "'," & contador & "," & _
                                                        "'" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & _
                                                        "'" & IIf(MSFlexGrid1.TextMatrix(contador, 1) = "", " ", MSFlexGrid1.TextMatrix(contador, 1)) & "'," & _
                                                        "'" & IIf(MSFlexGrid1.TextMatrix(contador, 4) = "", " ", MSFlexGrid1.TextMatrix(contador, 4)) & "'," & CANTIDAD & "," & _
                                                        CANTIDAD & ",'" & Trim(MSFlexGrid1.TextMatrix(contador, 8)) & "'," & _
                                                        "'" & Trim(MSFlexGrid1.TextMatrix(contador, 9)) & "'," & _
                                                        "'" & Trim(MSFlexGrid1.TextMatrix(contador, 10)) & "'," & _
                                                        Val(Format(precioprom, "#####.000000")) & "," & _
                                                        IIf(IsNumeric(MSFlexGrid1.TextMatrix(contador, 5)), Val(MSFlexGrid1.TextMatrix(contador, 5)) * VGTipCamb, 0) & "," & _
                                                        Val(MSFlexGrid1.TextMatrix(contador, 11)) & ")"
                                
                                End If
                                
                                
                                'En devolucion no entra las guia de transferecia
                                'Data2.Recordset.Update
                                If Text11.text <> "" And TxTransa = "TD" And VGGuiaSal Then
                                      ''If Not IsNull(Data3.Recordset("STKPREPRO")) Then
                                      ''  precioprom = Val(Format(Data3.Recordset("STKPREPRO"), "#####.000000"))
                                      ''Else
                                        precioprom = 0
                                      ''End If
                                      Unid = ""
                                      ''If Not Data2.Recordset.EOF Then
                                      ''   If Not IsNull(Data2.Recordset("deunidad")) Then Unid = Trim(cNull(Data2.Recordset("deunidad")))
                                      ''End If
                                                                            
                                      'grabadetalmacen
                                      cad = insertar1
                                      'Completo = False
                                      Conex.Execute cad
                                      'Do
                                      '   DoEvents
                                      'Loop Until Completo
                                      grabastk1 (contador)
                               End If
                          End If
                  Else
                      VGCNx.Execute ncad
                      
'                      Data2.Recordset.Update
                      'En caso de Guia contra factura no realiza descarga a stock
                      Text11 = VGAlma
                      Campo = "GS"   'El valor del campo cambia a NI u otro valor cuando es transferencia
                      nument = numserie
                      Unid = MSFlexGrid1.TextMatrix(contador, 4)
                      'cad = insertar1
                      'Completo = False
                      'Conex.Execute cad
                      'Do
                      '     DoEvents
                      'Loop Until Completo
                      'grabadetalmacen
                  End If
            Else
                Conex.Execute "INSERT INTO movalmdet " & _
                              "(DEALMA,DETD,DENUMDOC,DEITEM,DECODIGO,DEDESCRI,DEUNIDAD,DECANREF1)" & _
                              " VALUES (" & _
                              "'" & VGAlma & "'," & _
                              "'GS','" & numserie & "'," & contador & "," & _
                              "'" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & _
                              "'" & IIf(MSFlexGrid1.TextMatrix(contador, 1) = "", " ", MSFlexGrid1.TextMatrix(contador, 1)) & "'," & _
                              "'" & IIf(MSFlexGrid1.TextMatrix(contador, 4) = "", " ", MSFlexGrid1.TextMatrix(contador, 4)) & "'," & _
                              Val(MSFlexGrid1.TextMatrix(contador, 11)) & ")"
                  
'                  Data2.Recordset.Update
            End If
            contador = contador + 1
            'antes habia update
   Wend
   
If Text3.text = "GR" Then
   Set rst = VGCNx.Execute("select * from vt_puntovtadocumento where puntovtacodigo='" & VGparametros.puntovta & "' and empresacodigo='" & VGparametros.empresacodigo & "' and documentocodigo='" & Trim(Text3) & "' and puntovtadocserie='" & CmbSerie.text & "'")
   If rst.RecordCount > 0 Then
      Text4.text = Trim(rst!puntovtadoccorr)
      VGCNx.Execute "UPDATE vt_puntovtadocumento " & _
              " Set puntovtadoccorr='" & Right("0000000000" & Trim(CStr(Val(Text4)) + 1), 8) & "'" & _
              " where puntovtacodigo='" & VGparametros.puntovta & "' and empresacodigo='" & VGparametros.empresacodigo & "' and documentocodigo='" & Trim(Text3) & "' and puntovtadocserie='" & CmbSerie.text & "'"
   End If
   rst.Close
End If
   
'Ctr_AyudaEmpresa.xclave
If Requerimiento Then
    VGCNx.Execute "update co_cabordcompra set estadooccodigo='5' where oc_cnumord='" & Text8.text & "' " _
    & " and tipoordencodigo='" & Texttipdoc.text & "' and empresacodigo=(select empresacodigo from co_multiempresas where empresaruc='" & Text5.text & "') " '
End If
 
Requerimiento = False

If VGGuiaSal Then
     reinicia
Else
     CmdGrabarDet.Visible = False
     limpia
     visualizarFG
     If Text8.Enabled And Text8.Visible Then Text8.SetFocus
End If
VGval = False
CmdGrabarCab.Enabled = False
rpta = MsgBox("Desea Agregar Comentarios", vbYesNo + vbQuestion, "Aviso")
If rpta = vbYes Then
    'Text12.SetFocus
    crtlvisible (False)
    FrameComentario.Visible = True
    TxComentario.SetFocus
Else
     'inicializar
     TxTransa.Enabled = True
     rpta = MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso")
     If rpta = vbYes Then
        ' imprimir
        imprimirguias
     End If
End If
If Not VGGuiaSal Then Unload Me

Exit Sub
GrabErr1:
MsgBox Err.Number & "-" & Err.Description
Exit Sub
Resume
End Sub

Private Sub imprimirguias()

Dim nguia As String
Dim ntabla As String
Dim busca As New dll_apisgen.dll_apis
Dim rb As New ADODB.Recordset
Dim rb1 As New ADODB.Recordset
Dim contador As Double
Dim contador1 As Double
Dim numguias As Integer, TCANT As Integer, nflag As Integer
Dim SQL As String
Dim inicio As Integer
Dim fin As Integer
Dim j As Integer
Dim numero As String
Dim distrito As String


ntabla = "movalmdet"
contador = 0
'
'contador = 0
'Set rb = VGCNx.Execute("select * from gtempfile ")
'If rb.RecordCount > 0 Then
'    If rb.RecordCount Mod 50 > 0 Then
'        numguias = Int(rb.RecordCount / 50) + 1
'     Else
'         numguias = Int(rb.RecordCount / 50)
'    End If
'     rb.MoveFirst
'     Do While contador < numguias
'              contador = contador + 1
'              inicio = (contador - 1) * 50 + 1
'              If contador * 50 > rb.RecordCount Then
'                 fin = rb.RecordCount
'               Else
'                 fin = contador * 50
'              End If
'
'              nguia = CmbSerie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='GR' and puntovtadocserie='" & CmbSerie & "' and puntovtacodigo='" & GPunto & "'", VGCNx), 8)
'
'  '          VGCNx.Execute "Update vt_puntovtadocumento " & _
'  '                " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(Val(nguia) + 1)), 8) & "'" & _
'  '                " Where documentocodigo='GR' and puntovtacodigo='" & GPunto & "' and puntovtadocserie='" & CmbSerie & "'"
'
'             contador1 = 0
'              If fin > rb.RecordCount Then
'                 fin = rb.RecordCount - inicio
'              End If
'              VGCNx.Execute "delete from gtempfile2filas"
'          For j = inicio To fin
'                 contador1 = contador1 + 1
'                 If contador1 <= 25 Then
'                     SQL = "INSERT INTO gtempfile2filas(item,producto1,descripcion1,cantidad1,importe1,"
'                     SQL = SQL & "cantidad2,importe2) "
'                     SQL = SQL & " VALUES ( '" & contador1 & "','" & RTrim(rb!productocodigo) & "','" & RTrim(rb!productodescripcion) & "','" & rb!detpedcantpedida & "','" & rb!detpedimpbruto & "',0,0)"
'                  Else
'                     TCANT = contador1 - 25
'                      SQL = "UPDATE gtempfile2filas set producto2 ='" & RTrim(rb!productocodigo) & "',"
'                      SQL = SQL & " descripcion2='" & RTrim(rb!productodescripcion) & "',"
'                      SQL = SQL & "cantidad2='" & rb!detpedcantpedida & "',"
'                        SQL = SQL & "importe2= '" & rb!detpedimpbruto & "'"
'                        SQL = SQL & " where item = " & TCANT & ""
'                 End If
'                 VGCNx.Execute SQL
'                 rb.MoveNext
'          Next j
'    Loop
'End If
'rb.Close

'---------------------- OPCION DE IMRPIMRI GUIAS ------------------------------------
                                   
Screen.MousePointer = 11
                                   
With oCrystalReport
        .Reset
        .ReportFileName = VGParamSistem.RutaReport & "vt_guiaimpresa_" & VGParamSistem.BDEmpresa & Text11.text & ".rpt"

       If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGcadenareport2
        End If

        .DiscardSavedData = True
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .WindowShowZoomCtl = True
        .WindowShowNavigationCtls = True
        .WindowShowPrintBtn = True
        .WindowTitle = "Impresion Guia de Remision"
        .StoredProcParam(0) = VGParamSistem.BDEmpresa
        .StoredProcParam(1) = VGAlma
        .StoredProcParam(2) = "GS"
        .StoredProcParam(3) = numserie
        .Action = 1
        
  End With
  
Screen.MousePointer = 1

Exit Sub
errores:
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
  Err = 0
  Exit Sub
  
End Sub
Private Function TraeDataSerie(nsql As String, vcon As ADODB.Connection) As String
    Dim rsbuscn As New ADODB.Recordset
    
    Set rsbuscn = vcon.Execute(nsql)
    If rsbuscn.RecordCount > 0 Then
        TraeDataSerie = rsbuscn!puntovtadoccorr
    Else
        TraeDataSerie = "1"
    End If
    Set rsbuscn = Nothing

End Function

Private Sub CmdSalir_Click()
Dim I As Integer
Dim Productos As String

If Frame1.Visible Then
     If MSFlexGrid1.Rows > 1 Then
        If vbYes = MsgBox("Desea Grabar?", vbYesNo + vbQuestion, "Aviso") Then
        With MSFlexGrid1
            If .Rows > 1 Then
                Set Rs = VGCNx.Execute("select b.productocodigo as Codigo,c.adescri as Producto," _
                & " a.stskdis as Disponible," & .TextMatrix(.Rows - 1, 3) & " as Can_Pedida,Faltantes=(a.stskdis-" & .TextMatrix(.Rows - 1, 3) & ") " _
                & " from " & VGParamSistem.BDEmpresa & ".dbo.stkart a " _
                & " inner join " & VGParamSistem.BDEmpresa & ".dbo.vt_detallepedido b on a.stcodigo=b.productocodigo " _
                & " inner join " & VGParamSistem.BDEmpresa & ".dbo.maeart c on b.productocodigo=c.acodigo " _
                & " where b.pedidonumero='" & Txtnrodoc.text & "' " _
                & " and a.stskdis-" & .TextMatrix(.Rows - 1, 3) & "<0 ")
                If Not Rs.EOF Then
                    GridP.DataSource = Rs
                    With GridP
                          .Columns(0).Width = 1000
                          .Columns(1).Width = 4000
                          .Columns(2).Width = 900
                          .Columns(3).Width = 1000
                          .Columns(4).Width = 900
                    End With
                    GridP.Refresh
                    MsgBox "ATENCION !!! " & Chr(13) & "NO SE PUEDE EMITIR LA GUIA ", vbCritical, "Sistemas"
                    FrmValida.Visible = True
                    Timer1.Enabled = True
                    Exit Sub
                End If
            End If
        End With
        
       CmdGrabarDet_Click
       End If
    End If
    
    VGval = False
    TxTransa.Enabled = True
    Text6.Enabled = True
    Text5.Enabled = True
    Text8.Enabled = True
    Unload Me
 Else
    Frame1.Visible = True
    CmdSalir.SetFocus
 End If

End Sub





Private Sub Command4_Click()
FrmValida.Visible = False
Timer1.Enabled = False
t = 0
Text3.SetFocus
End Sub

Private Sub Conex_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
  Completo = True
End Sub

Private Sub Ctr_AyudaEmpresa_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Set Rs2 = Nothing
Set Rs2 = VGCNx.Execute("select empresacodigo as Empresa,tipo as Tipo,pedidotipofac as Doc,PEDIDOENTREGA AS Destino ,[Nro Orden]=PedidoNumero," _
& " (detpedcantpedida) as Cantidad from v_almacenyventas WHERE TIPO=3  and empresacodigo='" & ColecCampos(0) & "'" _
& " ")
If Not Rs2.EOF Then
    Set TDBGrid.DataSource = Rs2
    With TDBGrid
        .Columns(0).Width = 700
        .Columns(1).Width = 500
        .Columns(2).Width = 500
        .Columns(3).Width = 3500
        .Columns(4).Width = 1200
        .Columns(5).Width = 1000
    End With
    
    TDBGrid.Refresh
    FrmPen.Visible = True
    Text5.text = ColecCampos(1)
    Text6.text = ColecCampos(2)
    Text7.text = ColecCampos(3)
Else
    MsgBox "No hay Requerimientos pendientes" & Chr(13) & "para esta empresa.", vbInformation, "Sistemas"
    Ctr_AyudaEmpresa.SetFocus
    Text5.text = ColecCampos(1)
    Text6.text = ColecCampos(2)
    Text7.text = ColecCampos(3)
    'Exit Sub
End If

'Text5.text = ColecCampos(1)
'Text6.text = ColecCampos(2)
'Text7.text = ColecCampos(3)

CmdGrabarDet.Visible = True
Command3.Visible = True
Command2.Visible = True
Command1.Visible = True



End Sub

Private Sub Ctr_AyuTransporte_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
TxtTransp.text = Ctr_AyuTransporte.xclave
End Sub

Private Sub Ctr_AyuVendedor_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
'TxVendedor.text = Ctr_AyuVendedor.xclave
End Sub

Private Sub DTPicker1_Change()
        DTPicker1.Value = UltimoCierreFech(DTPicker1.Value)
        VGTipCamb = DevolverTCambio(DTPicker1.Value)
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  If TxTransa = "TD" Then
         Text11.SetFocus
  Else
         If TxTransa.Enabled Then
            TxTransa.SetFocus
         End If
  End If
End If
End Sub

Private Sub Form_Activate()
   Dim j, kTotal As Double
   If MSFlexGrid1.Rows > 1 Then
      Text2 = Format(MSFlexGrid1.Rows - 1, "##,###,##0.00")
      kTotal = 0
      For j = 1 To MSFlexGrid1.Rows - 1
        kTotal = kTotal + CDbl(MSFlexGrid1.TextMatrix(j, 3))
      Next
      Text9 = Format(kTotal, "##,###,##0.00")
   Else
      Text2 = Format(0, "##,###,##0.00")
      Text9 = Format(0, "##,###,##0.00")
   End If
End Sub

Private Sub Form_Load()
Dim rsqli As String
Call Ctr_AyuTransporte.Conexion(VGCNx)
Call Ctr_AyuVendedor.Conexion(VGCNx)
Call Ctr_AyudaEmpresa.Conexion(VGCNx)
Cliente = False
Requerimiento = False

VGSeleccion = 1   'Indica el modo de apertura = 1 y modificacion=2
VGForm = 6
limpia

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = 800
 

salir = False
hubo_error = False
'RMM*******************************************************************
 DTPicker1.Value = UltimoCierreFech(CDate(Format(Now, "dd/MM/yyyy")))
'*******************************************************************
VGTipCamb = DevolverTCambio(DTPicker1.Value)
'SAS
Deshabilitar (False)
Codigo2 = "GUIA DE REMISION"
If VGGuiaSal Then
    FrmGuiaSal.Caption = "Registro de Guia de Salidas"
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    'CmdGrabarDet.Visible = False
    lblSerie.Visible = True
    AgregarSerie
    CmbSerie.Visible = CBool(-1)
    visualizarFG1
Else
    FrmGuiaSal.Caption = "Devolucion de Guia de Salidas"
    visualizarFG
    TxTransa.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = True
    TxTransa = "DG"
    Text3 = "GS"
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    'CmdGrabarDet.Visible = False
    CmdGrabarCab.Visible = False
    CmbSerie.Visible = False
End If


If CmbSerie.Visible Then CmbSerie.ListIndex = 0

CmdGrabarDet.Picture = MDIPrincipal.ImageList2.ListImages("Facturado").Picture
CmdSalir.Picture = MDIPrincipal.ImageList2.ListImages("Retornar").Picture
Command3.Picture = MDIPrincipal.ImageList2.ListImages("Eliminar").Picture
Command2.Picture = MDIPrincipal.ImageList2.ListImages("Modificar").Picture
Command1.Picture = MDIPrincipal.ImageList2.ListImages("Insertar").Picture

End Sub

Private Sub Grid1_DblClick()
MSFlexGrid1.Rows = 1


Set Rs = VGCNx.Execute("select A.PEDIDONUMERO,A.PEDIDONROFACT, b.PRODUCTOCODIGO,c.ADESCRI," _
& " B.DETPEDCANTPEDIDA,A.PEDIDOTOTNETO " _
& " from VT_PEDIDO a inner join VT_DETALLEPEDIDO b  on a.PEDIDONUMERO=b.PEDIDONUMERO " _
& " Inner join maeart c on b.PRODUCTOCODIGO=c.acodigo " _
& " where A.PEDIDONROFACT='" & Rs2!pedidonumero & "'")

Do While Not Rs.EOF
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = Rs!productocodigo
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = Rs!ADESCRI
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = Rs!detpedcantpedida
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = Rs!pedidototneto
    'Txtnrodoc.text = Grid1.TextMatrix(Grid1.Rows - 1, 1)
    'Text11.text = Grid1.TextMatrix(Grid1.Rows - 1, 0)
    'TxVendedor.text = Grid1.TextMatrix(Grid1.Rows - 1, 5)
    'Label5.Caption = Grid1.TextMatrix(Grid1.Rows - 1, 6)
    Rs.MoveNext
Loop

FrmPen.Visible = False

End Sub


Private Sub MSFlexGrid1_Click()
If Not VGGuiaSal Then
    Frame1.Visible = False
'    Frame2.Visible = True
    llenadatos
 End If
End Sub

Private Sub MSFlexGrid2_Click()

End Sub

Private Sub MSFlexGrid2_DblClick()
'Set Rs = VGCNx.Execute("select * from  movalmdet where denumdoc='" & Grid1.TextMatrix(Grid1.Row, 1) & "'")
'MSFlexGrid1.Rows = 1
'Do While Not Rs.EOF
'    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
'    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = Rs!decodigo
'    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = IIf(IsNull(Rs!DEDESCRI), "", Rs!DEDESCRI)
'    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = Rs!DECANTID
'    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 4) = IIf(IsNull(Rs!DEUNIDAD), "", Rs!DEUNIDAD)
''    TxVendedor.text = Rs!vendedorcodigo
''    TxtTransp.text = Rs!transportecodigo
'    Text11.text = Rs!dealma
'    Rs.MoveNext
'Loop

End Sub

Private Sub TDBGrid_DblClick()
If Rs2(0) = "1" Then
    Set Rs = VGCNx.Execute("select a.pedidonumero,a.pedidonrofact,a.productocodigo,a.ADESCRI," _
    & " Saldo=(a.detpedcantpedida)-sum(isnull(a.decantid,0)),b.PEDIDOTOTNETO,b.VENDEDORCODIGO," _
    & " a.almacencodigo , b.transportecodigo " _
    & " from " & VGParamSistem.BDEmpresa & ".dbo.v_almacenyventas a " _
    & " inner join vt_pedido b on a.pedidonumero=b.pedidonumero " _
    & " WHERE a.clienteruc='" & Text5.text & "' and a.empresacodigo='" & VGparametros.empresacodigo & "' and a.puntovtacodigo='" & VGparametros.puntovta & "' and a.pedidonumero='" & Rs2(3) & "' " _
    & " group by a.pedidonumero,a.pedidonrofact,a.productocodigo,a.ADESCRI,a.DETPEDCANTPEDIDA,b.PEDIDOTOTNETO,b.VENDEDORCODIGO," _
    & " a.almacencodigo , b.transportecodigo having (a.detpedcantpedida)-sum(isnull(a.decantid,0))>0 ")
     
ElseIf Rs2(0) = "2" Then
   Set Rs = VGCNx.Execute("select A.PEDIDONUMERO,A.PEDIDONROFACT, b.PRODUCTOCODIGO,c.ADESCRI," _
   & " B.DETPEDCANTPEDIDA,A.PEDIDOTOTNETO,A.VENDEDORCODIGO,a.almacencodigo,a.transportecodigo " _
   & " from VT_tempoPEDIDO01 a inner join VT_tempoDETALLEPEDIDO01 b  on a.PEDIDONUMERO=b.PEDIDONUMERO " _
   & " Inner join maeart c on b.PRODUCTOCODIGO=c.acodigo " _
   & " where A.PEDIDONUMERO='" & Rs2(2) & "'")
ElseIf Rs2(0) = "3" Then
   Set Rs = VGCNx.Execute("SELECT a.tipoordencodigo,b.oc_ccodigo,dbo.MAEART.ADESCRI,b.oc_ncantid,a.oc_cnumord as NroPedido," _
   & " a.oc_cnumord,a.almacenorigen FROM dbo.co_DETordcompra b " _
   & " inner JOIN dbo.co_cabordcompra a on b.OC_CNUMORD=a.OC_CNUMORD " _
   & " inner JOIN dbo.MAEART ON  b.OC_Ccodigo=maeart.acodigo inner JOIN dbo.tabalm c " _
   & " ON  a.almacendestino=c.taalma where a.estadooccodigo<=4 and a.oc_cnumord='" & Rs2(3) & "'")
End If

MSFlexGrid1.Rows = 1
If Rs2(0) = 3 Then
    Do While Not Rs.EOF
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = Rs!oc_ccodigo
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = Rs!ADESCRI
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = Rs!oc_ncantid
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = 0
        Txtnrodoc.text = Rs!oc_cnumord
        Texttipdoc.text = Rs!tipoordencodigo
        Text8.text = Rs!oc_cnumord
        Text11.text = Rs!almacenorigen
        VGAlma = Rs!almacenorigen
        Rs.MoveNext
    Loop
    Requerimiento = True
Else
    Do While Not Rs.EOF
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = Rs!productocodigo
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = Rs!ADESCRI
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = Rs!saldo
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = Rs!pedidototneto
        Txtnrodoc.text = Rs!pedidonumero
        Text11.text = Rs!almacencodigo
        VGAlma = Rs!almacencodigo
        If Not IsNull(Rs!vendedorcodigo) Then
            Ctr_AyuVendedor.xclave = Rs!vendedorcodigo: Ctr_AyuVendedor.Ejecutar
        End If
        
        If Not IsNull(Rs!transportecodigo) Then
            Ctr_AyuTransporte.xclave = Rs!transportecodigo: Ctr_AyuTransporte.Ejecutar
        End If
        
        Rs.MoveNext
    Loop
    Requerimiento = False
End If

FrmPen.Visible = False

End Sub


Private Sub Text1_DblClick()
  Dim Adodc3 As ADODB.Recordset   'Centro de Costos
  Set Adodc3 = New ADODB.Recordset
  If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
        Adodc3.Open "SELECT cencost_codigo,cencost_descripcion FROM centro_costos ", VGcnxCT, adOpenStatic, adLockOptimistic
  Else
        Adodc3.Open "SELECT cencost_codigo,cencost_descripcion FROM centro_costos ", VGCNx, adOpenStatic, adLockOptimistic
  End If
  
        frmReferencia.Conectar Adodc3, "SELECT cencost_codigo,cencost_descripcion FROM centro_costos  "
        frmReferencia.Label1.Caption = "Centro de Costos"
        frmReferencia.Show vbModal
        Adodc3.Close
        If vGUtil(1) <> "" Then
                 Text1 = vGUtil(1)
                 LblCC = vGUtil(2)
''''''                 If Not ClsTock.CCostoconSalidas(VGAlma, Text10, VGConfig) And TxTransa = "DP" Then
''''''                    MsgBox "No hay Salidas con el Centro de Costo que Selecciono", vbCritical, "Aviso "
''''''                    Text10 = ""
''''''                    LblCC = ""
''''''                 End If
        End If
        If Text1 <> "" Then Text1_KeyPress (13)

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   Text1_DblClick
ElseIf KeyCode = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'**********************CENTRO COSTO
If KeyAscii = 13 And Text10.text <> "" Then
  If IsNumeric(Text1.text) Then
      If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
          If Existe(3, Text1, "CENTRO_COSTOS", "cencost_codigo", False) = False Then
                 MsgBox "La Cuenta no existe", vbInformation, "Mensaje"
                 Text1.SetFocus: Exit Sub
          End If
      Else
          If Existe(1, Text10, "CENTRO_COSTOS", "cencost_codigo", False) = False Then
          '       MsgBox "La Cuenta no existe", vbInformation, "Mensaje"
          '       Text1.SetFocus: Exit Sub
          End If
      End If
          Tabula (KeyAscii)
          'Cmddetalle_Click
   Else
      MsgBox "Ingrese el numero de Centro de Costo", vbInformation, mensaje1
      Text1.SetFocus
   End If
End If

End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Text10_DblClick
ElseIf KeyCode = 46 Then
     'Label20 = ""
End If
End Sub
'Almacen
Private Sub Text11_DblClick()
    Dim Adodc3 As ADODB.Recordset
    Set Adodc3 = New ADODB.Recordset
    Adodc3.Open "SELECT TAALMA,TADESCRI FROM TABALM", VGCNx, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc3, "SELECT TAALMA,TADESCRI FROM TABALM"
    frmReferencia.Label1.Caption = "Almacenes"
    frmReferencia.Show vbModal
    Adodc3.Close
    If vGUtil(1) <> "" Then Text11 = (vGUtil(1))
    VGAlma = Text11
    If Text11 <> "" Then Text11_KeyPress (13)
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text11_DblClick
If KeyCode = 13 Then VGAlma = "" & Trim(Text11)
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim rst As New ADODB.Recordset
  
 If KeyCode = 112 Then
    Text3_DblClick
 ElseIf KeyCode = 13 Then
   If UCase(Text3.text) = "GR" Then
      Set rst = VGCNx.Execute("select * from vt_puntovtadocumento where empresacodigo='" & VGparametros.empresacodigo & "' and puntovtacodigo='" & VGparametros.puntovta & "' and documentocodigo='" & Trim(Text3) & "'")
      If rst.RecordCount > 0 Then
         CmbSerie.Clear
         Do Until rst.EOF
            CmbSerie.AddItem rst!puntovtadocserie
            Text4.text = Trim(rst!puntovtadoccorr)
            rst.MoveNext
         Loop
         CmbSerie.ListIndex = 0
      
      End If
      rst.Close
    End If
 End If
 
 Set rst = Nothing
End Sub

Private Sub Text3_LostFocus()
   On Error Resume Next
   Call Text3_KeyDown(13, 0)
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 And Text3 <> "" Then
    Text4_DblClick
 End If
End Sub

Private Sub Text4_LostFocus()
Text4.text = Format(Text4.text, "00000000")
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
     Text5_DblClick
ElseIf KeyCode = 8 Then
     Text6 = ""
     Text7 = ""
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
Dim rbusca As New ADODB.Recordset

If KeyAscii = 13 And Trim(Text5) <> "" Then
      Text5 = Trim(Text5)
           
      Text6 = existe_clie(Text5)
       If Text6 = "" Then
              MsgBox "No existe el codigo del cliente", vbExclamation, "Clientes"
       Else
             Text7 = Mid(direccion, 1, 25)
             'TxtCambio.SetFocus
       End If
End If
End Sub

Private Sub Text10_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT COD_FP,DES_FP  FROM FORMA_PAGO", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT COD_FP,DES_FP  FROM FORMA_PAGO"
frmReferencia.Label1.Caption = "Forma de Pago"
frmReferencia.Show vbModal
Adodc3.Close
If vGUtil(1) <> "" Then Text10 = (vGUtil(1))
If Text10 <> "" Then Text10_KeyPress (13)
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text10.text <> "" Then
     Dim Adodc3 As ADODB.Recordset
     Set Adodc3 = New ADODB.Recordset
     Adodc3.Open "SELECT COD_FP,DES_FP  FROM FORMA_PAGO WHERE COD_FP ='" & Text10 & "'", VGCNx, adOpenStatic, adLockOptimistic
     If Adodc3.RecordCount > 0 Then
            ' muestra

            If Text11.Visible And Text11.Enabled Then
              Text11.SetFocus
            Else
              SendKeys "{tab}"
            End If
      Else
             Text10.SetFocus
      End If
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TxTransa.text = "TD" And Len(TxTransa.text) = 2 Then
    If Existe(1, Text11, "TabAlm", "TAALMA", False) Then
        If VGAlma = Text11 Then
            MsgBox "No se puede Transferir al mismo Almacén", vbExclamation, "Error"
            Text11.SetFocus:  Exit Sub
        End If
        SendKeys "{tab}"
    End If
End If
End Sub

Private Sub txtCambio_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 And IsNumeric(TxtCambio) Then
'    Text10.Enabled = True
'    Text10.Visible = True
'    Text10.SetFocus
'    Exit Sub
'  End If
'  If KeyAscii = 13 And (Trim(TxtCambio) = "" Or Val(TxtCambio) = 0) Then
'    CmdGrabarCab_Click
'    If CmdGrabarDet.Enabled And CmdGrabarDet.Visible Then CmdGrabarDet.SetFocus
'    Exit Sub
'  End If
'  If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And Chr$(KeyAscii) <> "." And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim criterio As String

If KeyAscii = 13 Then
    If Len(Text3) = 2 Then
        Text3 = UCase(Text3)
        If ValidarDoc(Text3) = "" Then Exit Sub
        
        Text4.SetFocus
        Exit Sub
    Else
        Text3 = ""
        Text4 = ""
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End If

End Sub

Private Sub Timer1_Timer()
t = t + 1
Lblmensaje.Visible = IIf(t Mod 2 = 0, True, False)
If t > 21 Then
    Timer1.Enabled = False
    t = 0
End If
End Sub

Private Sub TxTransa_DblClick()
Dim Adodc3 As ADODB.Recordset

Set Adodc3 = New ADODB.Recordset
VGRegEnt = 2

Adodc3.Open "SELECT TT_CODMOV,TT_DESCRI,tt_clie FROM Tabtransa where  TT_tipmov = '" & IIf(VGRegEnt = 1, "I", "S") & "'", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT TT_CODMOV,TT_DESCRI,tt_clie FROM Tabtransa where  TT_tipmov = '" & IIf(VGRegEnt = 1, "I", "S") & "'"
frmReferencia.Label1.Caption = "Transacciones"
frmReferencia.Show vbModal
Adodc3.Close

If vGUtil(1) <> "" Then
    TxTransa = vGUtil(1)
    Label9.Caption = vGUtil(2)
End If

If vGUtil(3) <> "" Then Cliente = IIf(vGUtil(3) = "S", True, False)

If TxTransa <> "" Then
    Deshabilitar (True)
                  
    If Cliente Then
        Text5.Enabled = True
        Text5.Visible = True
        Ctr_AyudaEmpresa.Enabled = False
        Ctr_AyudaEmpresa.Visible = False
        Ctr_AyudaEmpresa.xclave = ""
        Label12.Caption = "Cod.Cliente :"
    Else
        Ctr_AyudaEmpresa.Enabled = True
        Ctr_AyudaEmpresa.Visible = True
        Text5.Enabled = False
        Text5.Visible = False
        Label12.Caption = "Cod.Empresa :"
    End If

    LIMPIACABECERA
    buscar
End If

End Sub
Private Sub TxTransa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Len(TxTransa.text) = 2 Then
    TxTransa = UCase(TxTransa)
    If (transa(TxTransa) <> "") Then
            Deshabilitar (True)
            If Cliente Then
                Text5.Enabled = True
                Text5.Visible = True
                Ctr_AyudaEmpresa.Enabled = False
                Ctr_AyudaEmpresa.Visible = False
                Ctr_AyudaEmpresa.xclave = ""
                Label12.Caption = "Cod.Cliente :"
                
            Else
                Ctr_AyudaEmpresa.Enabled = True
                Ctr_AyudaEmpresa.Visible = True
                Text5.Enabled = False
                Text5.Visible = False
                Label12.Caption = "Cod.Proveedor :"
            End If
            LIMPIACABECERA
            buscar
            CmbSerie.Clear
            Call AgregarSerie
            If TxTransa = "TD" Then
                   Text11.SetFocus
            Else
                   Text3.SetFocus
            End If
     Else
            Enfoque TxTransa
    End If
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxTransa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    TxTransa_DblClick
    CmbSerie.Clear
    Call AgregarSerie
End If

End Sub
Private Sub Text4_DblClick()
If Not VGGuiaSal Then
     FormAyuguia.Show 1
     If Text4 <> "" Then
            devolver (Text4)
     End If
End If
If Text4 <> "" Then
     Text4_KeyPress (13)
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxTransa = "GF" Then
  If Text4 = "" Then
     'Text3 = "FT"
   MsgBox "Ingrese el número de factura", vbInformation, "Aviso"
   Text4.SetFocus
  End If
  If (Text3 = "FT" Or Text3 = "FE" Or Text3 = "BV") And VGGuiaSal Then '
    llenarfactura
    Exit Sub
  End If
 End If
 If KeyAscii = 13 Then
    If Not VGGuiaSal And Trim(Text4) <> "" Then
         devolver (Text4)
    Else
         SendKeys "{tab}"
         KeyAscii = 0
    End If
 End If
End Sub

Private Sub Text5_DblClick()
If Len(Trim(TxTransa.text)) = 0 Then
    MsgBox "Primero seleccione tipo de transaccion", vbInformation, "Sistema"
    TxTransa.SetFocus
    Exit Sub
End If

If Len(Text3.text) = 0 Then
    MsgBox "Primero ingrese documento de referencia", vbInformation, "Sistema"
    Text3.SetFocus
    Exit Sub
ElseIf Len(CmbSerie.text) = 0 Then
    MsgBox "Falta seleccionar serie", vbInformation, "Sistema"
    CmbSerie.SetFocus
    Exit Sub
ElseIf Len(Text4.text) = 0 Then
    MsgBox "Falta indicar numero de documento", vbInformation, "Sistema"
    Text4.SetFocus
    Exit Sub
End If

Text5 = ""
Text6 = ""
Text7 = ""
FrmAyuCliente.Show 1
Text5 = FrmAyuCliente.cCod
Text6 = FrmAyuCliente.cNom
Text7 = Trim(FrmAyuCliente.cDir)
ruc = FrmAyuCliente.cRuc

Dim RsCliente As ADODB.Recordset

Set RsCliente = VGCNx.Execute("select empresacodigo from co_multiempresas where empresaruc='" & ruc & "'")
If RsCliente.RecordCount = 0 Then
    CmdGrabarDet.Visible = True
    Command3.Visible = True
    Command2.Visible = True
    Command1.Visible = True
    Set Rs2 = Nothing
    Set Rs2 = VGCNx.Execute("select Tipo, empresacodigo as Empresa,Almacen=almacencodigo ,Numero_Pedido=PedidoNumero," _
    & " Saldo=(detpedcantpedida)-sum(isnull(decantid,0))" _
    & " from " & VGParamSistem.BDEmpresa & ".dbo.v_almacenyventas WHERE CLIENTERUC='" & FrmAyuCliente.cCod & "' and empresacodigo='" & VGparametros.empresacodigo & "' and puntovtacodigo='" & VGparametros.puntovta & "' and Estado=0" _
    & " group by tipo,almacencodigo,pedidonumero,detpedcantpedida,empresacodigo having (detpedcantpedida)-sum(isnull(decantid,0))>0")

    If Not Rs2.EOF Then
        Set TDBGrid.DataSource = Rs2
        TDBGrid.Refresh
        FrmPen.Visible = True
    Else
        MsgBox "Este cliente no tiene Guias pendientes", vbInformation, "Sistemas"
    End If
Else
    CmdGrabarDet.Visible = True
    Command3.Visible = True
    Command2.Visible = True
    Command1.Visible = True
    Set Rs2 = Nothing
    Set Rs2 = VGCNx.Execute("select tipo as Tipo,pedidotipofac as Doc,PEDIDOENTREGA AS Destino ,[Nro Orden]=PedidoNumero," _
    & " (detpedcantpedida) as Cantidad from v_almacenyventas WHERE TIPO=3 ")
    If Not Rs2.EOF Then
        Set TDBGrid.DataSource = Rs2
        TDBGrid.Refresh
        FrmPen.Visible = True
    Else
        MsgBox "Este cliente no tiene Guias pendientes"
    End If
End If


End Sub

Private Sub Text3_DblClick()
Dim Adodc3 As ADODB.Recordset

Set Adodc3 = New ADODB.Recordset

Adodc3.Open "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU"
frmReferencia.Label1.Caption = "Tipo de Documentos"
frmReferencia.Show vbModal

Adodc3.Close

If vGUtil(1) <> "" Then Text3 = (vGUtil(1))
If vGUtil(1) <> "" Then Label10.Caption = (vGUtil(2))

If Text3 <> "" Then Text4.SetFocus


End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Trim(Text7.text) <> "" Then
       Text7.SetFocus
  End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Text7.text <> "" Then
      ' muestra
      'TxtCambio.SetFocus
  End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
  Dim criterio As String
  If KeyAscii = 13 Then             'de orden de compra
    If IsNumeric(Text8.text) Then
       'If Len(Text4.text) = 7 Then
'         criterio = "CANUMDOC = " & Chr$(34) + Text4.text + Chr$(34) & "AND  CACODCLIE = " & Chr$(34) + Text4.text + Chr$(34)
'         Data1.Recordset.FindFirst criterio             ya no va
'         If Not Data1.Recordset.NoMatch Then
'            MsgBox "El Numero documento ya ha sido registrado !"
'            Exit Sub
'         Else
'            MsgBox "ingreso"
'         End If
      'End If
      muestra
    Else
      If Trim(TxTransa) = "" Then
        CmdSalir.SetFocus
      End If
      MsgBox "Ingrese el numero de la Orden Compra", vbOKOnly + vbExclamation, "Error"
   End If
  End If
End Sub
Private Sub ocultarlabel()
Label7.Visible = True   'False
Text7.Visible = True    'False
Label9.Visible = True   'False

Label10.Visible = True   'False
Text10.Visible = True    'False
Label11.Visible = True    'False
Text11.Visible = True    'False
End Sub

Private Sub muestra()
Dim numfil As Long
Dim rsql As String
Dim ultimoserie As Long
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset
If Trim(VGAlma) <> "" Then
   rsql = "Select CTNNUMERO,CTNNUMFIN FROM NUM_DOCUMENTOS WHERE   CTNCODIGO = 'GS'  AND CTNNUMSER = '" & CmbSerie.text & "' "
   Rs.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
   If IsNull(Rs(0)) Then
      MsgBox "No se ha ingresado el numero de inicio de la serie en la Tabla ", vbInformation, "Error"
      salir = True
      Exit Sub
   End If
   If Not Rs.EOF Then
      numsal = Rs(0) + 1
      ultimoserie = Rs(1)
      Serie = Format(CmbSerie.text, "000")            ' ********************* Serie contiene   la seie de la guia de remision
      If Rs(0) > Rs(1) Then
         MsgBox "No se puede emitir guia," & Chr(13) & "La Nro. guia es mayor que número máximo", vbCritical, "Aviso"
         salir = True
         Exit Sub
      End If
      numserie = Serie & Format(numsal, "0000000")
      If VGGuiaSal And TxTransa <> "GF" Then
         sigue
      End If
   End If
   Rs.Close
Else
   MsgBox "No hay ningún Almacén Activo", vbInformation, "Información"
End If

If Not CmdGrabarDet.Visible Then sigue: CmdGrabarDet.SetFocus

End Sub


Private Sub sigue()
  Command1.Visible = True
  Command2.Visible = True
  Command3.Visible = True
  CmdGrabarDet.Visible = True
  FormCreacionSal.Caption = "Ingreso de Articulos"
  buscar
  If ChkTalla.Value = 0 Then
    FormCreacionSal.Show 1
   Else
    FrmIngTallas.Show 1
  End If
End Sub
Public Function insertar1()            ' grabadetalmacen()
'Esta funcion graba el detalle en el almacen de transferecia
 Dim cad As String
 
 If MSFlexGrid1.TextMatrix(contador, 7) = "S" Then
      cad = "insert into MovAlmDet (DEALMA,DETD,DENUMDOC,DECODIGO,DECANTID,DEPRECIO,DEITEM,DEUNIDAD,DESERIE,DETIPCAM,DECENCOS,DEORDFAB,DEQUIPO) values ('" & Text11 & "','" & Campo & "','" & Format(nument, "0000000000") & "','" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & CANTIDAD & "," & precioprom & "," & contador & ",'" & Unid & "','" & MSFlexGrid1.TextMatrix(contador, 2) & "'," & TCamb & ",'" & MSFlexGrid1.TextMatrix(contador, 8) & "','" & MSFlexGrid1.TextMatrix(contador, 9) & "','" & MSFlexGrid1.TextMatrix(contador, 10) & "' ) "
 ElseIf MSFlexGrid1.TextMatrix(contador, 7) = "N" Then
      cad = "insert into MovAlmDet (DEALMA,DETD,DENUMDOC,DECODIGO,DECANTID,DEPRECIO,DEITEM,DEUNIDAD,DELOTE,DETIPCAM,DECENCOS,DEORDFAB,DEQUIPO) values ('" & Text11 & "','" & Campo & "','" & Format(nument, "0000000000") & "','" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & CANTIDAD & "," & precioprom & "," & contador & ",'" & Unid & "','" & MSFlexGrid1.TextMatrix(contador, 2) & "' ," & TCamb & ",'" & MSFlexGrid1.TextMatrix(contador, 8) & "','" & MSFlexGrid1.TextMatrix(contador, 9) & "','" & MSFlexGrid1.TextMatrix(contador, 10) & "') "
 Else
      cad = "insert into MovAlmDet (DEALMA,DETD,DENUMDOC,DECODIGO,DECANTID,DEPRECIO,DEITEM,DEUNIDAD,DETIPCAM,DECENCOS,DEORDFAB,DEQUIPO) values ('" & Text11 & "','" & Campo & "','" & Format(nument, "0000000000") & "','" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & CANTIDAD & "," & precioprom & "," & contador & ",'" & Unid & "'," & TCamb & ",'" & MSFlexGrid1.TextMatrix(contador, 8) & "','" & MSFlexGrid1.TextMatrix(contador, 9) & "','" & MSFlexGrid1.TextMatrix(contador, 10) & "') "
 End If
 insertar1 = cad
End Function

Public Sub grabaalmacen()
'GRABA EN EL ALMCEN DESTINO
Dim uSql As String
Dim insertar1 As String
Dim Rs As New ADODB.Recordset
Dim rsql As String

rsql = "select  TANUMENT from tabAlm where TAALMA =  '" & Text11 & " ' "
Set Rs = VGCNx.Execute(rsql)
If Rs.EOF Then Exit Sub
nument = Rs(0) + 1: Campo = "NI"

If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
   TCamb = Val(Devolver_Dato(3, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
   TCamb = Val(Devolver_Dato(1, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
End If
  
insertar1 = "insert into MovAlmCab (CAALMA,CATD,CANUMDOC,CACODMOV,CAFECDOC,CATIPMOV,CASITGUI,CARFALMA,CARFTDOC,CARFNDOC,CAHORA,CAUSUARI,catipcam,contacto) values ('" & Text11 & "','" & Campo & "','" & Format(nument, "0000000000") & "','51','" & DTPicker1.Value & "','I','V','" & VGAlma & "','GS','" & numserie & "','" & Format(Time, "hh:mm:ss") & "','" & VGUsua & "'," & TCamb & ",'" & TxtCon.text & "') "
 
VGCNx.Execute insertar1
uSql = "Update TabAlm set TANUMENT = " & nument & " where TAALMA='" & Text11 & "' "
VGCNx.Execute uSql
'insertar1 = "insert into MovAlmCab (CAALMA,CATD,CANUMDOC,CACODMOV,CAFECDOC,CATIPMOV,CASITUA,CARFTDOC,CARFNDOC,CARFALMA) values ('" & Text11 & "','" & Campo & "','" & nument & "','TD','" & DTPicker1 & "','I','V','NS','" & Text4 & "','01' ) "

End Sub

Public Sub grabastk(contador)
Dim cadena As String
Dim criterio As String
Dim ncantidad As Double
Dim acmd As New ADODB.Command
Dim RSBUSCA2 As New ADODB.Recordset

   On Error GoTo GrabErr
   cadena = MSFlexGrid1.TextMatrix(contador, 0)
   criterio = " STCODIGO = '" & cadena & "'"
   criterio = criterio + " and  STALMA = '" & VGAlma & "'"
   'Data3.Recordset.FindFirst criterio
   Set RSBUSCA2 = VGCNx.Execute("SELECT * FROM STKART WHERE " & criterio)
   If RSBUSCA2.RecordCount > 0 Then
            'Data3.Recordset.Edit
            canttemp = RSBUSCA2.Fields("STSKDIS")      ' revisar si validar en creacion
                           'para la salida
            If VGGuiaSal Then
              'Data3.Recordset("STSKDIS") = Data3.Recordset("STSKDIS") - CANTIDAD
              ncantidad = RSBUSCA2.Fields("STSKDIS") - CANTIDAD
            Else
              'Data3.Recordset("STSKDIS") = Data3.Recordset("STSKDIS") + cantidadDEV
              ncantidad = RSBUSCA2.Fields("STSKDIS") + cantidadDEV
            End If
            If Not IsNull(RSBUSCA2.Fields("STKPREPRO")) Then
              If VGGuiaSal Then
                 precioprom = RSBUSCA2.Fields("STKPREPRO") * VGTipCamb
              Else
                 precioprom = RSBUSCA2.Fields("STKPREPRO")
              End If
            Else
              precioprom = 0
            End If
            
            VGCNx.Execute "UPDATE stkart " & _
                              " set STSKDIS=" & ncantidad & _
                               " WHERE " & criterio
    Else
'            Data3.Recordset.AddNew                  'existe
 '           Data3.Recordset("STALMA") = VGAlma   '"01"
 '           Data3.Recordset("STCODIGO") = MSFlexGrid1.TextMatrix(contador, 0)
 '           Data3.Recordset("STSKDIS") = CANTIDAD
            VGCNx.Execute "INSERT INTO stkart " & _
                            "(STALMA,STCODIGO,STSKDIS)" & _
                            " VALUES(" & _
                            "'" & VGAlma & "'," & _
                            MSFlexGrid1.TextMatrix(contador, 0) & "," & _
                            CANTIDAD & ")"
                            
        'Grabamos en Facturacion
         Set acmd.ActiveConnection = VGCNx
         acmd.CommandText = "al_actualizaproducto_pro"
         acmd.CommandType = adCmdStoredProc
         acmd.Prepared = True
         With acmd
             .Parameters("@baseini") = VGCNx.DefaultDatabase
             .Parameters("@basefin") = VGCNx.DefaultDatabase  'VGBase2
             .Parameters("@almacen") = VGAlma
             .Parameters("@articulo") = MSFlexGrid1.TextMatrix(contador, 0)
             .Parameters("@tipo") = "1"
         End With
         acmd.Execute
         Set acmd = Nothing
    End If
    'Data3.Recordset.Update
    RSBUSCA2.Close
    Set RSBUSCA2 = Nothing
     
     
    If MSFlexGrid1.TextMatrix(contador, 7) = "S" Then grabaserie VGAlma, MSFlexGrid1.TextMatrix(contador, 0)
    If MSFlexGrid1.TextMatrix(contador, 7) = "N" Then grabalote VGAlma, MSFlexGrid1.TextMatrix(contador, 0)
    Call ValMes(VGAlma, False)
    Exit Sub
GrabErr:
'Resume Next
    MsgBox Err.Description
    hubo_error = True
End Sub

Private Sub grabalote(alma As String, codigo As String)
Dim uSql As String
Dim Lote As String
Dim nuevo_stk As Double
Dim rsql As String
Dim Rs As New ADODB.Recordset
Dim fecfab As Date
Dim fecven As Date
   
    Lote = MSFlexGrid1.TextMatrix(contador, 2)
    rsql = "select STSLKDIS FROM STKLOTE where   STSALMA= '" & alma & "' and  STSCODIGO= '" & codigo & "' and STSLOTE= '" & Lote & "'" '
    'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
    Set Rs = VGCNx.Execute(rsql)
    If Rs.RecordCount > 0 Then
       If (Campo = "NI" And alma <> VGAlma) Then
         nuevo_stk = Rs(0) + CANTIDAD
       Else
         nuevo_stk = Rs(0) - CANTIDAD
       End If
       uSql = "Update STKLOTE set STSLKDIS = " & nuevo_stk & " WHERE  STSALMA='" & alma & "' and STSCODIGO='" & codigo & "'AND STSLOTE='" & Lote & "'"
    Else
     If (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) = "__/__/____" Then
        fecfab = Format(MSFlexGrid1.TextMatrix(contador, 9), "DD/MM/YYYY")
        uSql = "insert into STKLOTE (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB) VALUES ('" & alma & "','" & codigo & "','" & Lote & "'," & CANTIDAD & ",'" & fecfab & "') "
     ElseIf (MSFlexGrid1.TextMatrix(contador, 9)) = "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
        fecven = Format(MSFlexGrid1.TextMatrix(contador, 8), "DD/MM/YYYY")
        uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECVEN)  VALUES ('" & alma & "','" & codigo & "','" & Lote & "' ," & CANTIDAD & " ,'" & fecven & "') " 'SIN FECFAB
     ElseIf (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
        'If Not IsDate(fecfab) Then
        fecfab = Date
        'If Not IsDate(fecven) Then
        fecven = Date
        
        uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB,STSFECVEN)  VALUES ('" & alma & "','" & codigo & "','" & Lote & "' ," & CANTIDAD & " ,'" & fecfab & "','" & fecven & "') "
     Else
'        If Not IsDate(fecfab) Then
        fecfab = Date
'        If Not IsDate(fecven) Then
        fecven = Date
        
        uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS)  VALUES ('" & alma & "','" & codigo & "','" & Lote & "' ," & CANTIDAD & ") "
     End If
    End If
    VGCNx.Execute uSql
       
End Sub

Private Sub grabaserie(alma As String, codigo As String)
Dim uSql As String
Dim Serie As String
Dim valor As Integer
Dim Rs As New ADODB.Recordset
Dim rsql As String
Dim fecfab As Date
Dim fecven As Date
    Serie = MSFlexGrid1.TextMatrix(contador, 2)
    rsql = "select STSSKDIS FROM STKSERI where STSALMA= '" & alma & "' and  STSCODIGO= '" & codigo & "' and STSSERIE= '" & Serie & "'" '
'    Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
    Set Rs = VGCNx.Execute(rsql)
    If Rs.RecordCount > 0 Then
       valor = IIf((Campo = "NI" And alma <> VGAlma), 1, 0)
       uSql = "update STKSERI set STSSKDIS = " & valor & " WHERE  STSALMA='" & alma & "' and STSCODIGO='" & codigo & "'AND STSSERIE='" & Serie & "'"
    Else
       If (Campo = "NI" And alma <> VGAlma) Then
         uSql = "insert into STKSERI (STSALMA,STSCODIGO,STSSERIE,STSFECMOV,STSFECVEN,STSSKDIS,STSSKCOM) VALUES ('" & alma & "','" & codigo & "','" & Serie & "' ,'" & Date & "','" & Date & "',1,1) "
       Else
         uSql = "insert into STKSERI (STSALMA,STSCODIGO,STSSERIE,STSFECMOV,STSFECVEN,STSSKDIS,STSSKCOM) VALUES ('" & alma & "','" & codigo & "','" & Serie & "' ,'" & Date & "','" & Date & "',0,0) "
       End If
    End If
    VGCNx.Execute uSql
End Sub

Public Sub grabastk1(contador)
   Dim acmd As New ADODB.Command
   Dim criterio As String
   Dim cadena As String
   Dim ndato As Double
   Dim rsbusca As New ADODB.Recordset
   
   cadena = MSFlexGrid1.TextMatrix(contador, 0)
   criterio = " STCODIGO = '" & cadena & "'"
   criterio = criterio + "and  STALMA = '" & Text11 & "'"
 '  Data3.Recordset.FindFirst criterio
   Set rsbusca = VGCNx.Execute("SELECT * FROM STKART WHERE " & criterio)
   If rsbusca.RecordCount = 0 Then
'     Data3.Recordset.AddNew
'     Data3.Recordset("STSKDIS") = CANTIDAD
'     Data3.Recordset("STKPREPRO") = precioprom
'     Data3.Recordset("STALMA") = Text11  '"01"
'     Data3.Recordset("STCODIGO") = MSFlexGrid1.TextMatrix(contador, 0)
      VGCNx.Execute "Insert Into Stkart " & _
                        "(STSKDIS,STKPREPRO,STALMA,STCODIGO)" & _
                        " values(" & _
                        CANTIDAD & "," & precioprom & ",'" & Text11 & "'," & MSFlexGrid1.TextMatrix(contador, 0) & "')"
                        
       'Grabamos en Facturacion
        Set acmd.ActiveConnection = VGCNx
        acmd.CommandText = "al_actualizaproducto_pro"
        acmd.CommandType = adCmdStoredProc
        acmd.Prepared = True
        With acmd
            .Parameters("@baseini") = VGCNx.DefaultDatabase
            .Parameters("@basefin") = VGBase2
            .Parameters("@almacen") = Text11
            .Parameters("@articulo") = MSFlexGrid1.TextMatrix(contador, 0)
            .Parameters("@tipo") = "1"
        End With
        acmd.Execute
        Set acmd = Nothing
                        
   Else
     'Data3.Recordset.Edit
     auxdisp = rsbusca.Fields("STSKDIS")
     If rsbusca.Fields("STKPREPRO") <> 0 And (canttemp + auxdisp) <> 0 Then   'no se registrado algun precio
       'rsbusca.Fields("STKPREPRO") = VGTipCamb * (precioprom * canttemp + auxdisp * rsbusca.Fields("STKPREPRO")) / (canttemp + auxdisp)
       VGCNx.Execute "UPDATE stkart " & _
                          " SET STKPREPRO =" & VGTipCamb * (precioprom * canttemp + auxdisp * rsbusca.Fields("STKPREPRO")) / (canttemp + auxdisp) & _
                          " WHERE " & criterio
       
     End If
     'Data3.Recordset("STSKDIS") = Data3.Recordset("STSKDIS") + CANTIDAD
      VGCNx.Execute "UPDATE stkart " & _
                        " SET STSKDIS =" & rsbusca.Fields("STSKDIS") + CANTIDAD & _
                        " WHERE " & criterio
   End If
   rsbusca.Close
   Set rsbusca = Nothing
   
  ' Data3.Recordset.Update
   
   If MSFlexGrid1.TextMatrix(contador, 7) = "S" Then grabaserie Text11, MSFlexGrid1.TextMatrix(contador, 0)
   If MSFlexGrid1.TextMatrix(contador, 7) = "N" Then grabalote Text11, MSFlexGrid1.TextMatrix(contador, 0)
   'Data3.Refresh
   Call ValMes(Text11, True)
End Sub

Private Sub devolver(NumDoc As String)
   Dim adors As New ADODB.Recordset
   Dim Rs As New ADODB.Recordset
   Dim rsql As String
   
   rsql = "select  CACODCLI,CAFECDOC,CACODMON,CASITGUI  from MovAlmCab where CAALMA = '" & VGAlma & "' and CATD= 'GS'  AND  CANUMDOC= '" & NumDoc & "' AND CASITGUI IN ( 'V','P')  AND NOT CACIERRE "
   Set adors = New ADODB.Recordset
   adors.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
   If adors.RecordCount = 0 Then
       MsgBox "No existe el número, ha sido Anulado o ha sido Facturado o  se ha producido el Cierre Mensual", vbCritical, "Verificar"
       Exit Sub
   End If
   If VGGuiaSal = False Then
       Deshabilitar (False)
   Else
       Deshabilitar (True)
   End If
   DTPicker1 = adors(1)
   Text5 = adors("cacodcli")
   If Not IsNull(adors("cacodmon")) Then
      If adors("cacodmon") = "02" Then

      End If
   End If
   EstadoDevolucion = adors("casitgui")
   adors.Close
   rsql = "select  CNOMCLI,CDIRCLI from MAECLI   where CCODCLI= '" & Trim(Text5.text) & "' "
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set Rs = VGCNx.Execute(rsql)
   If Rs.RecordCount = 0 Then
       MsgBox "No existe Cliente, Documento Incompleto ", vbCritical, mensaje1
       Exit Sub
   End If
   Text6 = Rs(0)
   Text7 = Rs(1)
   Call llenarFG(Text3, Text4)
End Sub

Private Sub actualiza_guia_dev()
  Dim uSql As String
  uSql = "Update MovAlmCab set CASITGUI = 'A' where CAALMA = '" & VGAlma & "' and CATD= 'GS'  AND  CANUMDOC= '" & Text4 & "'"
  VGCNx.Execute uSql
End Sub

Function buscarclie(doc As String) As Recordset
  Dim Rs As New ADODB.Recordset
  Dim rsql As String
  rsql = "select  CNOMCLI,CDIRCLI from MAECLI   where CCODCLI= '" & doc & "' "
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  Set Rs = VGCNx.Execute(rsql)
  If Rs.EOF Then
       MsgBox "No existe Cliente ", vbCritical, mensaje1
       Exit Function
  End If
End Function

Public Sub buscarstk(Cod As String, CANTIDAD As Double, suma As Boolean)
  Dim Rs As New ADODB.Recordset
  Dim rsql As String
  rsql = "select n.STSKDIS from  StkArt  n.STALMA = '" & VGAlma & "'   and n.STCODIGO= " & Cod & " "
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  Set Rs = VGCNx.Execute(rsql)
  If Rs.EOF Then
     MsgBox "No hay dicho articulo en almacen", vbCritical, mensaje1
     Exit Sub
  End If
  If suma Then
     Rs(0) = Rs(0) + CANTIDAD
  Else
     Rs(0) = Rs(0) - CANTIDAD
  End If
End Sub

Private Sub llenarFG(tipo As String, NumDoc As String)
     Dim Adoreg1 As ADODB.Recordset
     Dim rsql As String
     Dim ser_lot As String
     Dim dato As String
     ' FG.FormatString = "Codigo| Descripcion| Serie \ Lote|  Cantidad | A Devolver| A Entregar||"
      MSFlexGrid1.Row = 0
      MSFlexGrid1.Cols = 11
      MSFlexGrid1.ColWidth(0) = 1500
      MSFlexGrid1.ColWidth(1) = 2700
      MSFlexGrid1.ColWidth(2) = 1000
      MSFlexGrid1.ColWidth(3) = 1200
      MSFlexGrid1.ColWidth(4) = 1200
      MSFlexGrid1.ColWidth(5) = 1200
      MSFlexGrid1.ColWidth(6) = 2
      MSFlexGrid1.ColWidth(7) = 2
      MSFlexGrid1.ColWidth(8) = 1100
      MSFlexGrid1.ColWidth(9) = 1100
      MSFlexGrid1.ColWidth(10) = 1100
      
      MSFlexGrid1.ColAlignment(0) = 1
      rsql = "select n.DECODIGO, n.DEDESCRI, m.AUNIDAD, n.DECANTID, n.DESERIE,n.DELOTE  from MovAlmDet n ,maeArt m where  n.DEALMA ='" & VGAlma & "' AND n.DETD = '" & tipo & "' AND n.DENUMDOC ='" & NumDoc & "' and m.acodigo=n.decodigo  ORDER BY n.DEITEM "  '

     Set Adoreg1 = New ADODB.Recordset
     Adoreg1.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
     If Adoreg1.RecordCount = 0 Then
       Exit Sub
     End If
     Adoreg1.MoveFirst
     MSFlexGrid1.Rows = 1
     While Not Adoreg1.EOF
       If IsNull(Adoreg1(4)) And IsNull(Adoreg1(5)) Then
             ser_lot = ""
             dato = "X"
       ElseIf Not IsNull(Adoreg1(4)) Then
             ser_lot = Adoreg1(4)
             dato = "S"
       Else
             ser_lot = Adoreg1(5)
             dato = "N"
       End If                '0               1               2                 3               4              5            6            7
       MSFlexGrid1.AddItem (Adoreg1(0) & vbTab & Adoreg1(1) & vbTab & ser_lot & vbTab & Adoreg1(3) & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & dato)
       Adoreg1.MoveNext
     Wend
End Sub
Private Sub visualizarFG1()
     
     MSFlexGrid1.Cols = 11
     MSFlexGrid1.Row = 0
     visualizarFG2
     MSFlexGrid1.ColWidth(0) = 1500  'cod
     MSFlexGrid1.ColWidth(1) = 2700   'des
     MSFlexGrid1.ColWidth(2) = 1200   'ser
     MSFlexGrid1.ColWidth(3) = 1500   'cant
     MSFlexGrid1.ColWidth(4) = 1300   'und
     MSFlexGrid1.ColWidth(5) = 1500   'vv.
     MSFlexGrid1.ColWidth(6) = 1500   'pv
     MSFlexGrid1.ColWidth(7) = 2   'xserie
     MSFlexGrid1.ColWidth(8) = 1100
     MSFlexGrid1.ColWidth(9) = 1100
     MSFlexGrid1.ColWidth(10) = 1100
     
     MSFlexGrid1.ColAlignment(0) = 1
End Sub
Private Sub visualizarFG2()
     'MsFlexGrid1.FormatString = "   Codigo |   Descripcion|  Serie\Lote| Cantidad|  Unidad | V.Valor |P.Valor|| "
     MSFlexGrid1.Cols = 12
     MSFlexGrid1.Row = 0
     MSFlexGrid1.TextMatrix(0, 0) = " CODIGO "
     MSFlexGrid1.TextMatrix(0, 1) = " DESCRIPCION"
     MSFlexGrid1.TextMatrix(0, 2) = " SERIE \ LOT"
     MSFlexGrid1.TextMatrix(0, 3) = " CANTIDAD "
     MSFlexGrid1.TextMatrix(0, 4) = " UNIDAD "
     MSFlexGrid1.TextMatrix(0, 5) = " V. VALOR"
     MSFlexGrid1.TextMatrix(0, 6) = " P. VALOR"
     MSFlexGrid1.TextMatrix(0, 7) = " F"   'revisar
     MSFlexGrid1.TextMatrix(0, 8) = "Cent.Costo "
     MSFlexGrid1.TextMatrix(0, 9) = "Ord.Fabri  "
     MSFlexGrid1.TextMatrix(0, 10) = "Maqu./Equi."
     MSFlexGrid1.TextMatrix(0, 11) = "Can.Ref"
     
     MSFlexGrid1.ColWidth(8) = 1100
     MSFlexGrid1.ColWidth(9) = 1100
     MSFlexGrid1.ColWidth(10) = 1100
     MSFlexGrid1.ColAlignment(0) = 1
End Sub

Private Sub visualizarFG()
  MSFlexGrid1.Clear
  MSFlexGrid1.Row = 0
  MSFlexGrid1.ColWidth(0) = 1500  'cod
  MSFlexGrid1.ColWidth(1) = 2700   'des
  MSFlexGrid1.ColWidth(2) = 1000   'ser
  MSFlexGrid1.ColWidth(3) = 1200   'cant
  MSFlexGrid1.ColWidth(4) = 1200   'und
  MSFlexGrid1.ColWidth(5) = 1200   'vv.
  MSFlexGrid1.ColWidth(6) = 2   'vv.
  MSFlexGrid1.ColWidth(8) = 1100
  MSFlexGrid1.ColWidth(9) = 1100
  MSFlexGrid1.ColWidth(10) = 1100
  
  MSFlexGrid1.Rows = 1
  MSFlexGrid1.TextMatrix(0, 0) = " CODIGO"
  MSFlexGrid1.TextMatrix(0, 1) = " DESCRIPCION"
  MSFlexGrid1.TextMatrix(0, 2) = " SERIE"
  MSFlexGrid1.TextMatrix(0, 3) = " CANTIDAD"
  MSFlexGrid1.TextMatrix(0, 4) = " A DEVOLVER"
  MSFlexGrid1.TextMatrix(0, 5) = " A ENTREGAR"
  MSFlexGrid1.TextMatrix(0, 8) = "Cent.Costo "
  MSFlexGrid1.TextMatrix(0, 9) = "Ord.Fabri  "
  MSFlexGrid1.TextMatrix(0, 10) = "Maqu./Equi."
  
  MSFlexGrid1.ColAlignment(0) = 1
End Sub

Private Sub llenadatos()
  'Text12 = ""
'  Label13 = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
'  Label15 = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
'  Label17 = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
'  Text12.SetFocus
End Sub

Private Sub reinicia()
  limpia
  MSFlexGrid1.Clear
  MSFlexGrid1.Rows = 1
  visualizarFG2
'  MSFlexGrid1.TextMatrix(0, 0) = " CODIGO"
'  MSFlexGrid1.TextMatrix(0, 1) = " DESCRIPCION"
'  MSFlexGrid1.TextMatrix(0, 2) = " CANTIDAD ING"
'  MSFlexGrid1.TextMatrix(0, 3) = " UNIDAD ING"
'  MSFlexGrid1.TextMatrix(0, 4) = " PRECIO UNIT"
'  MSFlexGrid1.TextMatrix(0, 5) = " CANT INF"
'  MSFlexGrid1.TextMatrix(0, 6) = " PRECIO INF"
  Command1.Visible = False
  Command2.Visible = False
  Command3.Visible = False
  'CmdGrabarDet.Visible = False
  CmdSalir.SetFocus
End Sub


Private Sub Deshabilitar(flag As Boolean)
  'TxtCambio.Enabled = flag
  Text3.Enabled = flag
  Text4.Enabled = flag
  Text5.Enabled = flag
  Text6.Enabled = flag
  Text7.Enabled = flag
  Text8.Enabled = flag
  Text11.Enabled = flag
End Sub

Function transa(text As TextBox) As String
 Dim Rs As Recordset
 Dim rsql As String
 Dim dato As String
  dato = "S"
  rsql = "select  TT_DESCRI,tt_clie FROM TabTransa where TT_CODMOV= '" & text & "' and TT_TIPMOV ='S'"    '& dato & "'" '
  
   Set Rs = VGCNx.Execute(rsql)
  If Rs.RecordCount > 0 Then
    transa = Rs(0)
    Label9 = Rs(0)
    Cliente = IIf(Rs(1) = "S", True, False)
  Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly + vbExclamation, "Error"
    transa = ""
  End If
   Rs.Close
End Function

Function ValidarDoc(txt As TextBox) As String
  
  Dim Rs As Recordset
  Dim rsql As String
rsql = "select TDO_DESCRI  from TIPO_DOCU  where TDO_TIPDOC='" & txt.text & "'"

Set Rs = VGCNx.Execute(rsql)
If Rs.RecordCount = 0 Then
   MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
   ValidarDoc = ""
   txt.SetFocus
   Exit Function
End If
ValidarDoc = Rs(0)
Label10.Caption = Rs(0)
Rs.Close
End Function

Private Sub grabacabecera()
Dim uSql As String
Dim Data1 As New ADODB.Recordset

'If Text4.text <> "" Then
 On Error GoTo GrabErr
 
      Data1.Open "movalmcab", VGCNx, adOpenDynamic, adLockOptimistic
      Data1.AddNew
      Data1("CAALMA") = VGAlma     '"0
      Data1("CATIPMOV") = "S"
      Data1("CATD") = "GS"
      Data1("CAUSUARI") = VGUsua
      tipo = Data1("CATD")
      Data1("CACOTIZA") = IIf(Len(Trim(tx_ordfab)) = 0, " ", tx_ordfab)
      If Trim(Text3.text) <> "" Then
         Data1("CARFTDOC") = Trim(Text3.text)
      Else
         Data1("CARFTDOC") = " "
      End If
      If Trim(Text4.text) <> "" Then
         Data1("CARFNDOC") = numserie
      End If
      Data1("CAFECDOC") = DTPicker1
      'guardar el nro del doc referencial
      Data1("CATIPGUI") = TxTransa
      Data1("CAHORA") = Format(Time, "hh:mm:ss")
      Data1("CAFECACT") = Date
      If Trim(TxTransa.text) <> "" Then
         Data1("CACODMOV") = UCase$(Trim(TxTransa.text))
      Else
         Data1("CACODMOV") = " "
      End If
      If Trim(Text6.text) <> "" Then
         Data1("CANOMCLI") = LTrim(Text6)
      Else
         Data1("CANOMCLI") = " "
      End If
      If Trim(Text7.text) <> "" Then
         Data1("CADIRENV") = LTrim(RTrim(Text7))
      Else
         Data1("CADIRENV") = " "
      End If
      If Text5.Visible And Trim(Text5.text) <> "" Then
         Data1("CACODCLI") = Trim(Text5.text)
         Data1("CARUC") = IIf(ruc <> "", ruc, " ")
      Else
         Data1("CACODCLI") = " "
         Data1("CARUC") = " "
      End If
      If Trim(Text8.text) <> "" Then
         Data1("CANUMORD") = Mid$(UCase$(Text8.text), 1, 10)
      Else
         Data1("CANUMORD") = " "
      End If
         Data1("CATIPCAM") = DevolverTCambio(DTPicker1.Value)
         Data1("CAFORVEN") = " "
      If Trim(TxtTransp) = "" Then
         Data1("CACODTRAN") = " "
      Else
          Data1("CACODTRAN") = Trim(TxtTransp)
      End If

      Data1("CAVENDE") = Ctr_AyuVendedor.xclave

      If Text11.Visible And Trim(Text11.text) <> "" Then
         Data1("CARFALMA") = Mid$(UCase$(Text11.text), 1, 2)
         grabaalmacen     'graba al almacen de referencia
      Else
         Data1("CARFALMA") = " "
      End If
      If Not VGGuiaSal Then                 'para devolucion
         Data1("CASITGUI") = Trim(EstadoDevolucion)
      Else
        TxTransa = UCase(TxTransa)
        If TxTransa = "GF" Then
           Data1("CASITGUI") = "E"    'para guia facturada
        ElseIf TxTransa = "GV" Then
           Data1("CASITGUI") = "P"    'para guia por facturar
        Else
           Data1("CASITGUI") = "V"    'para guia de remision cualquiera
        End If
      End If
      Data1("CAESTIMP") = "V"
      Data1("CAFECACT") = Date
      Data1("CANUMDOC") = numserie
      Data1("canroped") = Txtnrodoc
      Data1("Contacto") = TxtCon.text
      Data1.Update
   'End If
   'Data1.Refresh
   Data1.Close
   Set Data1 = Nothing
   
   
   uSql = "Update NUM_DOCUMENTOS set   CTNNUMERO= " & numsal & " where  CTNCODIGO = 'GS' AND CTNNUMSER= '" & CmbSerie.text & "'"
   VGCNx.Execute uSql
   hubo_error = False
   Exit Sub
GrabErr:
Resume
    MsgBox Err.Description
    hubo_error = True
End Sub

Private Sub buscar()
  Dim criterio As String
  Dim Rs As Recordset
  
  Dim rsql As String
   TxTransa = UCase(Trim(TxTransa))
   'Busco la transaccion
   rsql = "select  *  from TabTransa  where TT_CODMOV ='" & TxTransa.text & "' and TT_TIPMOV ='S'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set Rs = VGCNx.Execute(rsql)
   If Rs.RecordCount = 0 Then
      MsgBox "El tipo de transaccion no existe !", vbOKOnly, "Error"
      LIMPIACABECERA
      TxTransa.SetFocus
      Exit Sub
   End If

   If Not IsNull(Rs("TT_CONT")) Then
            TT_CONTADOR = Rs("TT_CONT")
   Else
       MsgBox "El tipo de transaccion no esta inicialida !", vbOKOnly, "Error"
       Exit Sub
   End If
   If Rs("TT_ALMA") = "N" Then
      Text11.Enabled = False
      Label11.Visible = True  'False
      Text11.Visible = True   'False
   Else
      Label11.Visible = True
      Text11.Visible = True
   End If
   If Rs("TT_OC") = "N" Then
'      Text8.Enabled = False
   End If
   If Rs("TT_CLIE") = "S" Then
         'Text5.Enabled = True
         Text6.Enabled = True
         Text7.Enabled = True
   Else
         'Text5.Enabled = False
         Text6.Enabled = False
         Text7.Enabled = False
   End If
   'MsgBox "Transaccion correcta", vbOKOnly, "Aviso"
   '*RMM*************************************
   If Rs("TT_CC") = "N" Then
      Text1.Enabled = False
      Label27.Visible = True  'False
      'Label10.Visible = False
      Text1.Visible = True   'False
      FormCreacionSal.txccosto.Visible = False
      FormCreacionSal.lblccosto.Visible = False
   Else
      Label27.Visible = True
      Text1.Visible = True
      Text1.Enabled = True
      FormCreacionSal.txccosto.Visible = True
      FormCreacionSal.lblccosto.Visible = True
      FormCreacionSal.txccosto = Text1
   End If
   '*RMM*************************************
           
   If Rs("TT_ORDFAB") = "S" Then
      tx_ordfab.Visible = True
      Label25.Visible = True
      FormCreacionSal.lblordfab.Visible = True
      FormCreacionSal.TxordFab.Visible = True
   Else
      tx_ordfab.Visible = True  'False
      Label25.Visible = True   'False
      FormCreacionSal.lblordfab.Visible = False
      FormCreacionSal.TxordFab.Visible = False
   End If
   
   If Rs("TT_EQUIP") = "S" Then
      tx_codmaq.Visible = True
      Label26.Visible = True
      FormCreacionSal.lblMaq.Visible = True
      FormCreacionSal.txEquip.Visible = True
   Else
      tx_codmaq.Visible = True  ' False
      Label26.Visible = True    'False
      FormCreacionSal.lblMaq.Visible = False
      FormCreacionSal.txEquip.Visible = False
   End If
           
   'lbltrans = Mid(lbltrans, 1, 21)
   If Text3.Enabled Then
      Text3.SetFocus
   ElseIf Text5.Enabled Then
      Text5.SetFocus
   ElseIf Text6.Enabled Then
      Text6.SetFocus
   ElseIf Text7.Enabled Then
      Text7.SetFocus
   ElseIf Text8.Enabled Then
      Text8.SetFocus
   ElseIf Text11.Enabled Then
      Text11.SetFocus
   Else
      TxTransa.SetFocus
   End If
   CmdGrabarCab.Enabled = True
   
End Sub
Private Sub ValMes(almacen As String, entrada As Boolean)
  Dim cadena As String
  Dim criterio As String
 
  Dim adors As New ADODB.Recordset
  Dim rsql As String
  Dim Cantent As Double
  Dim Cantsal As Double
  Dim mespro As String
  Dim uSql As String
   mespro = Year(DTPicker1) & Format(Month(DTPicker1), "00")
   cadena = MSFlexGrid1.TextMatrix(contador, 0) 'codigo del art
   rsql = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & almacen & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & cadena & "'"  '
   
  'Set adors = New ADODB.Recordset
  Set adors = VGCNx.Execute(rsql)
  ' adors.Open RSQL, Vgcnx, adOpenDynamic, adLockOptimistic
  If adors.RecordCount <> 0 Then
      If Not VGGuiaSal Then
             Cantent = adors(1) - cantidadDEV   '1
             uSql = "Update MoResMes set SMCANSAL = " & Cantent & "  where SMALMA='" & almacen & "'  and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
       Else
             If entrada Then
                    Cantent = adors(0) + CANTIDAD
                    uSql = "Update MoResMes set SMCANENT = " & Cantent & " where SMALMA='" & almacen & "' and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
             Else
                    Cantsal = adors(1) + CANTIDAD
                    uSql = "Update MoResMes set SMCANSAL = " & Cantsal & " where SMALMA='" & almacen & "' and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
             End If
       End If
   Else
       Cantsal = IIf(entrada, 0, CANTIDAD)
       Cantent = IIf(entrada, CANTIDAD, 0)
       uSql = "insert into MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMSALDOINI) VALUES ('" & almacen & "','" & cadena & "','" & mespro & "' ," & Cantent & "," & Cantsal & ",0) "
   End If
   VGCNx.Execute uSql
   adors.Close
End Sub


'******* ********************************  Factura  *************
Private Sub llenarfactura()
 Dim Rs As ADODB.Recordset
 Dim rsql As String
 Dim dato As String
 Dim NumDoc As String
 Dim numserie1 As String
 numserie1 = Mid(Text4, 1, 3)
 NumDoc = Mid(Text4, 4, 10)
 rsql = "Select * from FACCAB where  cfnumser = '" & numserie1 & "' "  'and cfnrocaj ='" & vGPtoVenta & "'
 rsql = rsql & "and cfnumdoc = '" & NumDoc & "' AND cftd= '" & Text3 & "'"
 
 Set Rs = New ADODB.Recordset
 Rs.Open rsql, VGCNx, adOpenStatic
 If Rs.RecordCount > 0 Then
   If Rs("CFFACGUI") = "S" Then  'cuando la graba
      MsgBox "Documento de referencia tiene guia, no procede", vbExclamation, "Aviso"
      Rs.Close
      Exit Sub
   End If
   If Rs("CFALMA") <> VGAlma Then
   MsgBox "Almacen de facturacion no es igual al almacen actual", vbExclamation, "Aviso"
      Rs.Close
      Exit Sub
   End If
    'TxSerie = Rs("CFNUMSER") & Rs("CFNUMDOC")
   If Not IsNull(Rs("CFFECDOC")) Then DTPicker1 = Rs("CFFECDOC")
   If Not IsNull(Rs("CFCODCLI")) Then Text5 = Rs("CFCODCLI")
    'If Not IsNull(Rs("CFRUC")) Then TxRuc = Rs("CFRUC")
   If Not IsNull(Rs("CFNOMBRE")) Then Text6 = Rs("CFNOMBRE")
   If Not IsNull(Rs("CFDIRECC")) Then Text7 = Rs("CFDIRECC")
    ' OJO If Not IsNull(Rs("CFALMA")) Then TxReferencia = Rs("CFALMA")
    'If Not IsNull(Rs("CFRFNUMSER")) Then TxNumSer = Rs("CFRFNUMSER")
    'If Not IsNull(Rs("CFRFNUMDOC")) Then TxNumDoc = Rs("CFRFNUMDOC")
    'If Not IsNull(Rs("CFVENDE")) Then TxVendedor = Rs("CFVENDE")
   If Not IsNull(Rs("CFCODMON")) Then
     If Rs("CFCODMON") = "01" Then
          'Combo2.ListIndex = 0
     Else
          'Combo2.ListIndex = 1
     End If
    End If
    If Not IsNull(Rs("CFFORVEN")) Then Text10 = Rs("CFFORVEN")
    
    If Not IsNull(Rs("CFORDCOM")) Then Text8 = Rs("CFORDCOM")
    detallefact
    CmdGrabarDet.Visible = True
Else
    MsgBox "No existe Factura", vbInformation, "Mensaje"
End If
Rs.Close
End Sub
Private Sub detallefact()
 Dim Rs As ADODB.Recordset
 Dim rsql As String
 Dim dato As String
 Dim NumDoc As String
 Dim numserie1 As String
 Dim Serie As String
 numserie1 = Mid(Text4, 1, 3)  'obtengo la serie
 NumDoc = Mid(Text4, 4, 10)   'obtengo el numero de doc

 rsql = "Select * from FACDET where  dfnumser = '" & numserie1 & "' "  'and cfnrocaj ='" & vGPtoVenta & "'                    '    A Inner Join FACCAB B on a.DFTD = B.CFTD and "
 rsql = rsql & "and dfnumdoc = '" & NumDoc & "'  AND    dftd ='" & Text3 & "' "
 Set Rs = New ADODB.Recordset
 Rs.Open rsql, VGCNx, adOpenStatic
 If Rs.RecordCount > 0 Then
   If MSFlexGrid1.Rows > 1 Then
        MSFlexGrid1.Rows = 1
   End If
   MSFlexGrid1.Refresh
   While Not Rs.EOF
     If Not IsNull(Rs("DFSERIE")) Then
             Serie = Rs("DFSERIE")
     ElseIf Not IsNull(Rs("DFLOTE")) Then
             Serie = Rs("DFLOTE")
     Else
             Serie = ""
     End If
     MSFlexGrid1.AddItem (Rs("dfcodigo") & vbTab & Rs("dfdescri") & vbTab & Serie & vbTab & Rs("dfcantid") & vbTab & Rs("dfunidad") & vbTab & Rs("dfprec_ven") & vbTab & Rs("dfprec_ori"))
     Rs.MoveNext
   Wend
 Else
   MsgBox "No existe el registro en Detalle de Factura", vbInformation, "Mensaje"
 End If
 Rs.Close
End Sub

Function verifica_nro_guia(nroguia As String) As Boolean
Dim csql As String
Dim adors As ADODB.Recordset
   verifica_nro_guia = True
   If Len(nroguia) <> 11 Then
     Exit Function
   End If
   csql = "SELECT CASITGUI FROM MovAlmCab where CATD='GS'  and  CANUMDOC ='" & nroguia & "' "
   Set adors = New ADODB.Recordset
   adors.Open csql, VGCNx, adOpenDynamic, adLockOptimistic
   If adors.RecordCount = 0 Then
        verifica_nro_guia = False
   Else
        MsgBox "El número de Guia de remisión ya fue grabado", vbInformation, "Aviso"
   End If
   
End Function

Private Sub LIMPIACABECERA()
Label10.Caption = ""
Text3 = ""
Ctr_AyudaEmpresa.xclave = "": Ctr_AyudaEmpresa.xnombre = ""
Text5.text = ""
Text6.text = ""
Text7.text = ""
Text8.text = ""
Txtnrodoc.text = ""
Text11.text = ""
TxtCon.text = ""
Texttipdoc.text = ""
Text1.text = ""
tx_ordfab.text = ""
ChkTalla.Value = False
Ctr_AyuVendedor.xclave = "": Ctr_AyuVendedor.xnombre = ""
Ctr_AyuTransporte.xclave = "": Ctr_AyuTransporte.xnombre = ""
tx_codmaq.text = ""
MSFlexGrid1.Rows = 1
End Sub


Private Sub GRABO_DET()
'(((((((((((((((((((((((((((
'esto es para facturacion
'Dim csql As String
'
'csql = "Insert Into FACWORDET (dfNUMPED,dfVENDE,dfNROCAJ,dfsecuen,dfCODIGO,"
'If Trim(TxArticulo) = "TEXTO" Then
'    csql = csql & "dfTEXTO,"
'Else
'    csql = csql & "dfDESCRI,"
'End If
'csql = csql & "DFPRECI1,DFPRECIO,DFPORDES,DFDESCTO,DFDESCLI,DFDESESP,DFCANTID,DFCANREF,DFPRESUP,DFORDEN,DFESTADO,"
'csql = csql & "DFFECDOC,DFIGV,DFIGVPOR,DFIMPUS,DFIMPMN,DFSERIE,DFLOTE,DFALMA) VALUES ("
'If cTip = "GS" Then
'    csql = csql & "'A',"
'Else
'    csql = csql & "'F',"
'End If
'csql = csql '" & FrmPFacD04.cNumPed & "','" & FrmPFacD04.TxVendedor & "','" & vGPtoVenta & "',"
'csql = csql & "'" & Format(Str(FrmPFacD04.nItem), "00") & "','" & TxArticulo & "','" & TxDescripcion & "',"
'csql = csql & "" & nPreci1 & "," & TxPrecio & "," & TxDescuento & "," & FrmPFacD04.nImpDsAr & "," & FrmPFacD04.nImpDsCl & ","
'csql = csql & "" & FrmPFacD04.nImpDsEs & "," & TxCantidad & "," & TxCantRefe & ",'" & TxPresupuesto & "','" & TxNum & "','V', '" & Format(FrmPFacD04.MaskEdBox1, "mm/dd/yyyy") & "',"
'csql = csql & "" & FrmPFacD04.nwIgv & "," & nIgvPor & "," & FrmPFacD04.nImpUS & "," & FrmPFacD04.nImpMN & ",'" & FrmPFacD04.TxSerie & "',"
'csql = csql & "'" & TxLote & "','" & FrmPFacD04.TxPto & "')"
'
'VGBaseDatos.Execute csql
'
End Sub

Private Sub imprimir()
'    Dim CADENA As String
'    Dim cFormato As String
'    Dim cDireccion As String
'    Dim cNomRepor  As String
'    Dim aBusca As New ADODB.Recordset
'    Dim cRuc As String
'
'    On Error GoTo ErrImp
'            'CrystalReport1.WindowTitle = "Sistema de Inventarios - inv017"
'            'CrystalReport1.ReportFileName =  VGParamSistem.RutaReport & "inv017.rpt"
'
'            'BUSCA EN TABLA DE DOCUMENTOS Y EXTRAE EL FORMATO, LUEGO BUSCA EN TABLA DE FORMATOS Y EXTRAE EL NOMBRE DEL REPORTE
'            cFormato = Devolver_Dato(1, Text3, "Num_Documentos", "CtnCodigo", False, "CTNFORMATO")
'            If Trim(cFormato) <> "" Then
'                    If cTransa = "TD" Then
'                            cDireccion = Devolver_Dato(1, AlmacenRF, "TabAlm", "TAALMA", False, "TADIRECC")
'                    Else
'                            cDireccion = Devolver_Dato(2, VGCODEMPRESA, "Empresa", "EMP_CODIGO", False, "EMP_DIRECCION")
'
'                    End If
'                    cRuc = Devolver_Dato(2, VGCODEMPRESA, "Empresa", "EMP_CODIGO", False, "EMP_RUC_DOCUMENTO")
'                   ' cNomRepor = Devolver_Dato(2, VGCODEMPRESA, "Formato", "COD_EMP", False, "NOM_REP", cFormato, "COD_FOR", Text3, "TIPO_DOC")   'TxTransa
'                    cNomRepor = "REPGUIAREM.RPT"
'                    If Trim(cNomRepor) <> "" Then
'                            CrystalReport1.Reset
'                            CrystalReport1.ReportFileName = VGParamSistem.RutaReport & cNomRepor
'
'                            CrystalReport1.Connect = VGcadenareport2
'                            CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
'                            CrystalReport1.StoredProcParam(1) = VGAlma
'                            CrystalReport1.StoredProcParam(2) = "GS"
'                            CrystalReport1.StoredProcParam(3) = numserie
'
'                            'Ubi_Tab CrystalReport1
'                            ''cadena = "{MOVALMCAB.CAALMA} = '" & VGAlma & "'  and {MOVALMCAB.CATD} = 'GS' and {MOVALMCAB.CANUMDOC} = '" & numserie & "'"
'
'                            CrystalReport1.DiscardSavedData = True
'                            CrystalReport1.Destination = crptToWindow
'                            ''CrystalReport1.SelectionFormula = cadena
'                            ''CrystalReport1.Formulas(0) = "Empresa = '" & VGparametros.RucEmpresa & "'"
'                            ''CrystalReport1.Formulas(1) = "Direccion = '" & cDireccion & "' "
'                            ''CrystalReport1.Formulas(2) = "Ruc = '" & cRuc & "' "
'                            ''CrystalReport1.Formulas(3) = "Tipo = '" & Devolver_Dato(1, cTransa, "TABTRANSA", "TT_CODMOV", False, "TT_DESCRI", "S", "TT_TIPMOV") & "' "
'                            'CrystalReport1.Formulas(0) = "nota ='" & Codigo2 & "'"
'                            'CrystalReport1.Formulas(1) = "empresa ='" & VGparametros.RucEmpresa & "'"
'                            'CrystalReport1.Formulas(2) = "hora ='" & Time & "'"
'
'                            CrystalReport1.formulas(0) = "fecha='" & CStr(Day(CDate(DTPicker1.Value))) & "     " & VGDllGeneral.DesMes(Month(CDate(DTPicker1.Value))) & "                       " & Right(CStr(Year(CDate(DTPicker1.Value))), 1) & "'"                 'DTPicker1.Value & "'"
'                            CrystalReport1.formulas(1) = "condicion='" & Text10 & "-" & Label20 & "'"
'                            CrystalReport1.formulas(2) = "destinado='" & Text5 & "-" & Text6 & "'"
'                            CrystalReport1.formulas(3) = "direccion='" & Text7 & "'"
'                            CrystalReport1.formulas(4) = "ruc='" & "" & "'"
'                            CrystalReport1.formulas(5) = "partida='" & VGDIRE & "'"
'                            CrystalReport1.formulas(6) = "llegada='" & Text7 & "'"
'                            CrystalReport1.formulas(7) = "compra='" & "x" & "'"
'                            CrystalReport1.formulas(8) = "venta='" & "x" & "'"
'                            CrystalReport1.formulas(9) = "transforma='" & "x" & "'"
'                            CrystalReport1.formulas(10) = "devolucion='" & "x" & "'"
'                            CrystalReport1.formulas(11) = "traslado='" & "x" & "'"
'                            CrystalReport1.formulas(12) = "trasladoemi='" & "x" & "'"
'                            CrystalReport1.formulas(13) = "importa='" & "x" & "'"
'                            CrystalReport1.formulas(14) = "exporta='" & "x" & "'"
'                            CrystalReport1.formulas(15) = "otro='" & "x" & "'"
'                            CrystalReport1.formulas(16) = "transporte='" & TxtTransp.text & "-" & Label23 & "'"
'                            Set aBusca = VGCNx.Execute("select * from al_transporte where TRACODIGO='" & TxtTransp.text & "'")
'                            If aBusca.RecordCount > 0 Then
'                                CrystalReport1.formulas(17) = "transdire='" & aBusca("TRADIR") & "'"
'                                CrystalReport1.formulas(18) = "transruc='" & aBusca(" TRARUC") & "'"
'                                CrystalReport1.formulas(19) = "transplaca='" & aBusca("TRAPLACA") & "'"
'                            Else
'                                CrystalReport1.formulas(17) = "transdire='" & " " & "'"
'                                CrystalReport1.formulas(18) = "transruc='" & " " & "'"
'                                CrystalReport1.formulas(19) = "transplaca='" & " " & "'"
'                            End If
'                            aBusca.Close
'                            Set aBusca = Nothing
'
'                            CrystalReport1.formulas(20) = "fechatraslado='" & DTPicker1.Value & "'"
'
'
'                            CrystalReport1.WindowShowPrintBtn = True
'                            CrystalReport1.WindowShowRefreshBtn = True
'                            CrystalReport1.WindowShowSearchBtn = True
'                            CrystalReport1.WindowShowPrintSetupBtn = True
'                            CrystalReport1.WindowState = crptMaximized
'                            If CrystalReport1.Status <> 2 Then
'                                CrystalReport1.Action = 1
'                                VGCNx.Execute "Update MovAlmCab Set CaEstImp = 'I' Where CATD = 'GS' and CANUMDOC = '" & numserie & "'"
'                            End If
'                    Else
'                            MsgBox "No existe el nombre del Reporte, verifique en Formatos", vbInformation, "Información"
'                            Exit Sub
'                    End If
'            Else
'                    MsgBox "No existe el Formato del Documento, verifique en la Tabla de Documentos", vbInformation, "Información"
'                    Exit Sub
'            End If
'        Exit Sub
'ErrImp:
'     MsgBox Err.Description
'     Resume Next
End Sub

Private Sub crtlvisible(dato As Boolean)
   MSFlexGrid1.Visible = dato
   Command1.Visible = dato
   Command2.Visible = dato
   Command3.Visible = dato
   CmdGrabarDet.Visible = dato
   CmdSalir.Visible = dato
   
End Sub

Function existe_clie(text As TextBox) As String
  Dim rsql As String
  Dim Rs As New ADODB.Recordset
  direccion = ""
  'RSQL = "SELECT CNOMCLI ,CDIRCLI FROM maecli where CCODCLI= '" & text & "'"
  rsql = "Select clienterazonsocial as cnomcli,clientedireccion as cdircli " & _
       " FROM vt_cliente where clientecodigo='" & text & "'"
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  Set Rs = VGCNx.Execute(rsql)
   If Not Rs.EOF Then 'existe
     existe_clie = Rs(0)
     direccion = IIf(IsNull(Rs(1)), " ", Rs(1))
   Else
     existe_clie = ""
  End If
  Rs.Close
End Function


Private Sub TxtTransp_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT TRACODIGO,TRANOMBRE FROM al_transporte", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT TRACODIGO,TRANOMBRE FROM al_transporte"
frmReferencia.Label1.Caption = "Transportista"
frmReferencia.Show vbModal
Adodc3.Close
If vGUtil(1) <> "" Then TxtTransp = (vGUtil(1))

If TxtTransp <> "" Then TxtTransp_KeyPress (13)
End Sub

Private Sub TxtTransp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   TxtTransp_DblClick
ElseIf KeyCode = 48 Then
    
ElseIf KeyCode = 8 Then
    
End If
End Sub

Private Sub TxtTransp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
         If TxtTransp <> "" Then
            Dim Adodc3 As ADODB.Recordset
            Set Adodc3 = New ADODB.Recordset
            TxtTransp = Trim(TxtTransp)
            Adodc3.Open "SELECT TRACODIGO,TRANOMBRE FROM al_transporte WHERE  TRACODIGO= '" & TxtTransp & "'", VGCNx, adOpenStatic, adLockOptimistic
            If Adodc3.EOF Then
              If vbYes = MsgBox("El código de Transportista no existe," & Chr(13) & "desea agregarlo ", vbInformation + vbYesNo, "Aviso") Then
                  VGtransp = False
                  FrmTranspor.Show 1
                  VGtransp = True
               Else
                 TxtTransp = ""
              End If
            Else
              
            End If
            If CmdGrabarCab.Visible = True And CmdGrabarCab.Enabled = True Then
              CmdGrabarCab.SetFocus
            End If
         Else
            SendKeys "{tab}"
         End If
 End If
End Sub

Private Sub AgregarSerie()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
'TxTransa
Adodc3.Open "SELECT CTNNUMSER FROM NUM_DOCUMENTOS   WHERE CTNCODIGO = '" & "GR" & "'", VGCNx, adOpenStatic, adLockOptimistic
 While Not Adodc3.EOF
    CmbSerie.AddItem Adodc3(0)
    Adodc3.MoveNext
 Wend
 Adodc3.Close
 CmbSerie.Visible = True
 CmbSerie.Enabled = True
 'CmbSerie.ListIndex = 0
End Sub

Private Sub TxVendedor_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT vendedorcodigo,vendedornombres  FROM vt_VENDEDOR", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT vendedorcodigo,vendedornombres  FROM vt_VENDEDOR"
frmReferencia.Label1.Caption = "Vendedor"
frmReferencia.Show vbModal
Adodc3.Close

End Sub

Private Sub TxVendedor_GotFocus()
'Enfoque TxVendedor
End Sub

Private Sub TxVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxVendedor_DblClick
End Sub

Private Sub TxVendedor_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Existe(1, TxVendedor, "vt_vendedor", "vendedorcodigo", False) Then
'        SendKeys "{tab}"
'    Else
'        If Trim(TxVendedor) = "" Then
'            MsgBox "Ingrese Vendedor", vbInformation, "Mensaje"
'        Else
'            MsgBox "El Vendedor no existe", vbInformation, "Mensaje"
'        End If
'        TxVendedor.SetFocus
'    End If
'End If
End Sub

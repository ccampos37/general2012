VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmMantLibro 
   Caption         =   "Libros"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   6960
   Begin TabDlg.SSTab SSTab1 
      Height          =   5715
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   10081
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "frmMantLibro.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblNumReg"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TDBGridAsientos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmbotones"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "frmMantLibro.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cAcepta"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cCancela"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Correlativos"
      TabPicture(2)   =   "frmMantLibro.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl(4)"
      Tab(2).Control(1)=   "lbl(15)"
      Tab(2).Control(2)=   "lbl(14)"
      Tab(2).Control(3)=   "lbl(13)"
      Tab(2).Control(4)=   "lbl(12)"
      Tab(2).Control(5)=   "lbl(11)"
      Tab(2).Control(6)=   "lbl(10)"
      Tab(2).Control(7)=   "lbl(9)"
      Tab(2).Control(8)=   "lbl(8)"
      Tab(2).Control(9)=   "lbl(7)"
      Tab(2).Control(10)=   "lbl(6)"
      Tab(2).Control(11)=   "lbl(5)"
      Tab(2).Control(12)=   "txt(14)"
      Tab(2).Control(13)=   "txt(3)"
      Tab(2).Control(14)=   "txt(13)"
      Tab(2).Control(15)=   "txt(12)"
      Tab(2).Control(16)=   "txt(11)"
      Tab(2).Control(17)=   "txt(7)"
      Tab(2).Control(18)=   "txt(6)"
      Tab(2).Control(19)=   "txt(5)"
      Tab(2).Control(20)=   "txt(4)"
      Tab(2).Control(21)=   "txt(10)"
      Tab(2).Control(22)=   "txt(9)"
      Tab(2).Control(23)=   "txt(8)"
      Tab(2).ControlCount=   24
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   3465
         TabIndex        =   15
         Top             =   3735
         Width           =   1140
      End
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   2025
         TabIndex        =   14
         Top             =   3735
         Width           =   1140
      End
      Begin VB.Frame Frame1 
         Height          =   3135
         Left            =   90
         TabIndex        =   10
         Top             =   420
         Width           =   6570
         Begin VB.CheckBox chk 
            Height          =   300
            Index           =   0
            Left            =   1830
            TabIndex        =   2
            Top             =   855
            Width           =   450
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   0
            Left            =   1815
            TabIndex        =   0
            Top             =   195
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   529
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   2
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            NoCaracteres    =   "0123456789"
            MarcarTextoAlEnfoque=   -1  'True
            NoRangoCadena   =   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   1
            Left            =   1815
            TabIndex        =   1
            Top             =   510
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   529
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   30
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "Flag Control"
            Height          =   270
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   1470
         End
         Begin VB.Label lbl 
            Caption         =   "Descripción"
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   510
            Width           =   1575
         End
         Begin VB.Label lbl 
            Caption         =   "Código"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   210
            Width           =   2310
         End
      End
      Begin VB.Frame frmbotones 
         Height          =   555
         Left            =   -74520
         TabIndex        =   4
         Top             =   4830
         Width           =   5730
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Imprimir"
            Height          =   330
            Index           =   3
            Left            =   3435
            TabIndex        =   9
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Salir"
            Height          =   330
            Index           =   4
            Left            =   4560
            TabIndex        =   8
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Eliminar"
            Height          =   330
            Index           =   2
            Left            =   2310
            TabIndex        =   7
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "E&ditar"
            Height          =   330
            Index           =   1
            Left            =   1185
            TabIndex        =   6
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Nuevo"
            Height          =   330
            Index           =   0
            Left            =   60
            TabIndex        =   5
            Top             =   165
            Width           =   1080
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGridAsientos 
         Height          =   3270
         Left            =   -74925
         TabIndex        =   16
         Top             =   915
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   5768
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Código"
         Columns(0).DataField=   "librocodigo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descripción"
         Columns(1).DataField=   "librodescripcion"
         Columns(1).DataWidth=   1700
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   4
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Flag Control"
         Columns(2).DataField=   "flagcontrol"
         Columns(2).DataWidth=   1700
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Corr. Enero"
         Columns(3).DataField=   "libronumcorr01"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Corr. Febrero"
         Columns(4).DataField=   "libronumcorr02"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Corr. Marzo"
         Columns(5).DataField=   "libronumcorr03"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Corr. Abril"
         Columns(6).DataField=   "libronumcorr04"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Corr. Mayo"
         Columns(7).DataField=   "libronumcorr05"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Corr. Junio"
         Columns(8).DataField=   "libronumcorr06"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Corr. Julio"
         Columns(9).DataField=   "libronumcorr07"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Corr. Agosto"
         Columns(10).DataField=   "libronumcorr08"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Corr. Setiembre"
         Columns(11).DataField=   "libronumcorr09"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Corr. Octubre"
         Columns(12).DataField=   "libronumcorr10"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Corr. Noviembre"
         Columns(13).DataField=   "libronumcorr11"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Corr. Diciembre"
         Columns(14).DataField=   "libronumcorr12"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   15
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=15"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(45)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(49)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(53)=   "Column(13).Width=2725"
         Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(57)=   "Column(14).Width=2725"
         Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=2646"
         Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
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
         MultiSelect     =   2
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=84,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.alignment=3,.bold=0,.fontsize=825"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=106,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=102,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=98,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=95,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=96,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=97,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=86,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=83,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=84,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=85,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=82,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=79,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=80,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=81,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=78,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=75,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=76,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=77,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=74,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=70,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=62,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=59,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=60,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=61,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=58,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=54,.parent=13"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=51,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=52,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=53,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=50,.parent=13"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=47,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=48,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=49,.parent=17"
         _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=46,.parent=13"
         _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=43,.parent=14"
         _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=44,.parent=15"
         _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=45,.parent=17"
         _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=32,.parent=13"
         _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=29,.parent=14"
         _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=30,.parent=15"
         _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=31,.parent=17"
         _StyleDefs(96)  =   "Named:id=33:Normal"
         _StyleDefs(97)  =   ":id=33,.parent=0"
         _StyleDefs(98)  =   "Named:id=34:Heading"
         _StyleDefs(99)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(100) =   ":id=34,.wraptext=-1"
         _StyleDefs(101) =   "Named:id=35:Footing"
         _StyleDefs(102) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(103) =   "Named:id=36:Selected"
         _StyleDefs(104) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(105) =   "Named:id=37:Caption"
         _StyleDefs(106) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(107) =   "Named:id=38:HighlightRow"
         _StyleDefs(108) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(109) =   "Named:id=39:EvenRow"
         _StyleDefs(110) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(111) =   "Named:id=40:OddRow"
         _StyleDefs(112) =   ":id=40,.parent=33"
         _StyleDefs(113) =   "Named:id=41:RecordSelector"
         _StyleDefs(114) =   ":id=41,.parent=34"
         _StyleDefs(115) =   "Named:id=42:FilterBar"
         _StyleDefs(116) =   ":id=42,.parent=33"
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   8
         Left            =   -73095
         TabIndex        =   22
         Top             =   2460
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   9
         Left            =   -69960
         TabIndex        =   23
         Top             =   735
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   10
         Left            =   -69960
         TabIndex        =   24
         Top             =   1050
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   4
         Left            =   -73095
         TabIndex        =   18
         Top             =   1080
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   5
         Left            =   -73095
         TabIndex        =   19
         Top             =   1425
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   6
         Left            =   -73095
         TabIndex        =   20
         Top             =   1770
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   7
         Left            =   -73095
         TabIndex        =   21
         Top             =   2115
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   11
         Left            =   -69960
         TabIndex        =   25
         Top             =   1395
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   12
         Left            =   -69960
         TabIndex        =   26
         Top             =   1740
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   13
         Left            =   -69960
         TabIndex        =   27
         Top             =   2085
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   3
         Left            =   -73095
         TabIndex        =   17
         Top             =   720
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   14
         Left            =   -69960
         TabIndex        =   28
         Top             =   2415
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         Valor           =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Registros"
         Height          =   270
         Left            =   -70260
         TabIndex        =   42
         Top             =   4320
         Width           =   900
      End
      Begin VB.Label lblNumReg 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNumReg"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   -69315
         TabIndex        =   41
         Top             =   4305
         Width           =   915
      End
      Begin VB.Label lbl 
         Caption         =   "Febrero"
         Height          =   285
         Index           =   5
         Left            =   -74805
         TabIndex        =   40
         Top             =   1170
         Width           =   780
      End
      Begin VB.Label lbl 
         Caption         =   "Marzo"
         Height          =   285
         Index           =   6
         Left            =   -74805
         TabIndex        =   39
         Top             =   1515
         Width           =   780
      End
      Begin VB.Label lbl 
         Caption         =   "Abril"
         Height          =   285
         Index           =   7
         Left            =   -74805
         TabIndex        =   38
         Top             =   1860
         Width           =   780
      End
      Begin VB.Label lbl 
         Caption         =   "Mayo"
         Height          =   285
         Index           =   8
         Left            =   -74805
         TabIndex        =   37
         Top             =   2175
         Width           =   780
      End
      Begin VB.Label lbl 
         Caption         =   "Junio"
         Height          =   285
         Index           =   9
         Left            =   -74805
         TabIndex        =   36
         Top             =   2535
         Width           =   780
      End
      Begin VB.Label lbl 
         Caption         =   "Julio"
         Height          =   285
         Index           =   10
         Left            =   -71430
         TabIndex        =   35
         Top             =   765
         Width           =   660
      End
      Begin VB.Label lbl 
         Caption         =   "Agosto"
         Height          =   285
         Index           =   11
         Left            =   -71445
         TabIndex        =   34
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label lbl 
         Caption         =   "Setiembre"
         Height          =   285
         Index           =   12
         Left            =   -71445
         TabIndex        =   33
         Top             =   1455
         Width           =   1125
      End
      Begin VB.Label lbl 
         Caption         =   "Octubre"
         Height          =   285
         Index           =   13
         Left            =   -71445
         TabIndex        =   32
         Top             =   1785
         Width           =   1125
      End
      Begin VB.Label lbl 
         Caption         =   "Noviembre"
         Height          =   285
         Index           =   14
         Left            =   -71445
         TabIndex        =   31
         Top             =   2145
         Width           =   1125
      End
      Begin VB.Label lbl 
         Caption         =   "Diciembre"
         Height          =   285
         Index           =   15
         Left            =   -71445
         TabIndex        =   30
         Top             =   2505
         Width           =   1125
      End
      Begin VB.Label lbl 
         Caption         =   "Enero"
         Height          =   285
         Index           =   4
         Left            =   -74805
         TabIndex        =   29
         Top             =   810
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmMantLibro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim rs As New ADODB.Recordset
Dim rsLibro As ADODB.Recordset

Private Sub Form_Load()
  Call ConfiguraForm
  Call MuestraDatosLibro
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
  Set rsLibro = Nothing
  Set VGvardllgen = Nothing
End Sub

Sub ConfiguraForm()
  SSTab1.Tab = 0
  SSTab1.TabEnabled(1) = False
  SSTab1.TabEnabled(2) = False
  cAcepta.Enabled = False
  lblNumReg.Caption = Empty
  Me.Width = 7050
  Me.Height = 6255
End Sub

'FIXIT: Declare 'MuestraDatosLibro' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function MuestraDatosLibro()
 Dim SQL As String
  
  SQL = "SELECT A.librocodigo, A.librodescripcion,A.flagcontrol,B.libronumcorr01,"
  SQL = SQL & "B.libronumcorr02, B.libronumcorr03, B.libronumcorr04,B.libronumcorr05,"
  SQL = SQL & "B.libronumcorr06, B.libronumcorr07, B.libronumcorr08,B.libronumcorr09,"
  SQL = SQL & "B.libronumcorr10 , B.libronumcorr11, B.libronumcorr12 "
  SQL = SQL & "FROM  ct_libro A left join ct_librocorre B on b.empresacodigo='" & VGParametros.empresacodigo & "' and a.librocodigo<>'00' "
  SQL = SQL & "and A.librocodigo=B.librocodigo and B.libroanno='" & VGParamSistem.Anoproceso & "'"
  
  Set rs = VGCNx.Execute(SQL)
  Set TDBGridAsientos.DataSource = rs
  Call ConfiguraGridLibros
  lblNumReg.Caption = rs.RecordCount
  
End Function

Private Sub cmdBotones_Click(Index As Integer)
  Dim J As Integer
  Dim spos As Integer
  Dim SQL As String
  Dim SQL1 As String
  
  On Error GoTo x
  
  Select Case Index
     Case 0   'nuevo
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
        modoinsert = True
        Call LimpiarForm(frmMantLibro, "")
        txt(0).SetFocus
        Call ModoEditable(True, frmMantLibro, "")
        frmbotones.Visible = False
        
     Case 1   'modificar
        If TDBGridAsientos.Row < 0 Then
          Exit Sub
        End If
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
        modoedit = True
        frmbotones.Visible = False
        Call Editarlibro
        Call ModoEditable(True, frmMantLibro, "")
      
     Case 2   'eliminar
       If MsgBox("Desea eliminar el registro de Libro Nº " & TDBGridAsientos.Columns(0).Value & "?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          SQL = "DELETE FROM CT_LIBRO WHERE librocodigo='" & txt(0).Text & "'"
          SQL1 = "DELETE FROM CT_LIBROCORRE WHERE "
          SQL1 = SQL1 & " empresacodigo='" & VGParametros.empresacodigo & "' and librocodigo='" & txt(0).Text & "' AND libroanno='" & VGParamSistem.Anoproceso & "'"
          VGCNx.Execute (SQL)
          VGCNx.Execute (SQL1)
          Call MuestraDatosLibro
       End If
        
     Case 3   'imprimir
       Call Impresion("rptlibro.rpt")
     
     Case 4  ' salir
       Unload Me
  End Select
  
  Exit Sub
   
x:
  If Index = 2 And err.Number = -2147217873 Then
    MsgBox "Registro no podrá Eliminarse mientras exista Información en la Tablas Relacionadas", vbInformation, Caption
  Else
    MsgBox "Error inesperado: " & err.Description & "  " & err.Number, vbInformation, Caption
  End If
   
End Sub

Private Sub cAcepta_Click()
  Dim SQL As String
  Dim SQL1 As String
  Dim xx As New ADODB.Recordset
  On Error GoTo x
  Dim Cont As Integer
  Set VGvardllgen = New dllgeneral.dll_general
  VGCNx.BeginTrans
  Cont = 0
  If modoinsert = True Then
    SQL = "INSERT INTO CT_libro (librocodigo,librodescripcion,flagcontrol,"
    SQL = SQL & "libronumcorr01,libronumcorr02,libronumcorr03,libronumcorr04,libronumcorr05,libronumcorr06,libronumcorr07,libronumcorr08,libronumcorr09,libronumcorr10,libronumcorr11,libronumcorr12,usuariocodigo,fechaact) "
    SQL = SQL & "VALUES ('" & txt(0).Text & "','" & txt(1).Text & "'," & chk(0).Value & ","
    SQL = SQL & VGvardllgen.ESNULO(txt(3).Text, 0) & "," & VGvardllgen.ESNULO(txt(4).Text, 0) & "," & VGvardllgen.ESNULO(txt(5).Text, 0) & ","
    SQL = SQL & VGvardllgen.ESNULO(txt(6).Text, 0) & "," & VGvardllgen.ESNULO(txt(7).Text, 0) & "," & VGvardllgen.ESNULO(txt(8).Text, 0) & ","
    SQL = SQL & VGvardllgen.ESNULO(txt(9).Text, 0) & "," & VGvardllgen.ESNULO(txt(10).Text, 0) & "," & VGvardllgen.ESNULO(txt(11).Text, 0) & "," & VGvardllgen.ESNULO(txt(12).Text, 0) & "," & VGvardllgen.ESNULO(txt(13).Text, 0) & "," & VGvardllgen.ESNULO(txt(14).Text, 0) & ",'"
    SQL = SQL & VGusuario & "','" & Format(Date, "dd/mm/yyyy") & "')"
    
    SQL1 = " select  * from CT_libroCorre where empresacodigo='" & VGParametros.empresacodigo & "'"
    SQL1 = SQL1 & " and librocodigo='" & txt(0).Text & "' and libroanno='" & VGParamSistem.Anoproceso & "'"
    Set xx = Nothing
    Set xx = VGCNx.Execute(SQL1)
    If xx.RecordCount = 0 Then
       Cont = 1
    End If
      SQL1 = "INSERT INTO CT_libroCorre (empresacodigo,librocodigo,libroanno,"
      SQL1 = SQL1 & "libronumcorr01,libronumcorr02,libronumcorr03,libronumcorr04,libronumcorr05,libronumcorr06,libronumcorr07,libronumcorr08,libronumcorr09,libronumcorr10,libronumcorr11,libronumcorr12,usuariocodigo,fechaact) "
      SQL1 = SQL1 & "VALUES ('" & VGParametros.empresacodigo & "',"
      SQL1 = SQL1 & "'" & txt(0).Text & " ','" & VGParamSistem.Anoproceso & "',"
      SQL1 = SQL1 & VGvardllgen.ESNULO(txt(3).Text, 0) & "," & VGvardllgen.ESNULO(txt(4).Text, 0) & "," & VGvardllgen.ESNULO(txt(5).Text, 0) & ","
      SQL1 = SQL1 & VGvardllgen.ESNULO(txt(6).Text, 0) & "," & VGvardllgen.ESNULO(txt(7).Text, 0) & "," & VGvardllgen.ESNULO(txt(8).Text, 0) & ","
      SQL1 = SQL1 & VGvardllgen.ESNULO(txt(9).Text, 0) & "," & VGvardllgen.ESNULO(txt(10).Text, 0) & "," & VGvardllgen.ESNULO(txt(11).Text, 0) & "," & VGvardllgen.ESNULO(txt(12).Text, 0) & "," & VGvardllgen.ESNULO(txt(13).Text, 0) & "," & VGvardllgen.ESNULO(txt(14).Text, 0) & ",'"
      SQL1 = SQL1 & VGusuario & "','" & Format(Date, "dd/mm/yyyy") & "')"
                  
  ElseIf modoedit = True Then
    SQL = "UPDATE CT_LIBRO SET librodescripcion='" & Trim$(UCase$(txt(1).Text)) & "',"
    SQL = SQL & "flagcontrol=" & chk(0).Value & ","
    SQL = SQL & "libronumcorr01=" & VGvardllgen.ESNULO(txt(3).Text, 0) & ",libronumcorr02=" & VGvardllgen.ESNULO(txt(4).Text, 0) & ",libronumcorr03=" & VGvardllgen.ESNULO(txt(5).Text, 0) & ","
    SQL = SQL & "libronumcorr04=" & VGvardllgen.ESNULO(txt(6).Text, 0) & ",libronumcorr05=" & VGvardllgen.ESNULO(txt(7).Text, 0) & ",libronumcorr06=" & VGvardllgen.ESNULO(txt(8).Text, 0) & ","
    SQL = SQL & "libronumcorr07=" & VGvardllgen.ESNULO(txt(9).Text, 0) & ",libronumcorr08=" & VGvardllgen.ESNULO(txt(10).Text, 0) & ",libronumcorr09=" & VGvardllgen.ESNULO(txt(11).Text, 0) & ","
    SQL = SQL & "libronumcorr10=" & VGvardllgen.ESNULO(txt(12).Text, 0) & ",libronumcorr11=" & VGvardllgen.ESNULO(txt(13).Text, 0) & ",libronumcorr12=" & VGvardllgen.ESNULO(txt(14).Text, 0) & ","
    SQL = SQL & "usuariocodigo='" & VGusuario & "',fechaact='" & Format(Date, "dd/mm/yyyy") & "' "
    SQL = SQL & "WHERE librocodigo='" & txt(0).Text & "'"
    
    SQL1 = " select  * from CT_libroCorre where empresacodigo='" & VGParametros.empresacodigo & "'"
    SQL1 = SQL1 & " and librocodigo='" & txt(0).Text & "' and libroanno='" & VGParamSistem.Anoproceso & "'"
    Set xx = Nothing
    Set xx = VGCNx.Execute(SQL1)
    If xx.RecordCount = 0 Then
      SQL1 = "INSERT INTO CT_libroCorre (empresacodigo,librocodigo,libroanno,"
      SQL1 = SQL1 & "libronumcorr01,libronumcorr02,libronumcorr03,libronumcorr04,libronumcorr05,libronumcorr06,libronumcorr07,libronumcorr08,libronumcorr09,libronumcorr10,libronumcorr11,libronumcorr12,usuariocodigo,fechaact) "
      SQL1 = SQL1 & "VALUES ('" & VGParametros.empresacodigo & "',"
      SQL1 = SQL1 & "'" & txt(0).Text & " ','" & VGParamSistem.Anoproceso & "',"
      SQL1 = SQL1 & VGvardllgen.ESNULO(txt(3).Text, 0) & "," & VGvardllgen.ESNULO(txt(4).Text, 0) & "," & VGvardllgen.ESNULO(txt(5).Text, 0) & ","
      SQL1 = SQL1 & VGvardllgen.ESNULO(txt(6).Text, 0) & "," & VGvardllgen.ESNULO(txt(7).Text, 0) & "," & VGvardllgen.ESNULO(txt(8).Text, 0) & ","
      SQL1 = SQL1 & VGvardllgen.ESNULO(txt(9).Text, 0) & "," & VGvardllgen.ESNULO(txt(10).Text, 0) & "," & VGvardllgen.ESNULO(txt(11).Text, 0) & "," & VGvardllgen.ESNULO(txt(12).Text, 0) & "," & VGvardllgen.ESNULO(txt(13).Text, 0) & "," & VGvardllgen.ESNULO(txt(14).Text, 0) & ",'"
      SQL1 = SQL1 & VGusuario & "','" & Format(Date, "dd/mm/yyyy") & "')"
   Else
      SQL1 = "UPDATE CT_LIBROCorre SET "
      SQL1 = SQL1 & "libronumcorr01=" & VGvardllgen.ESNULO(txt(3).Text, 0) & ",libronumcorr02=" & VGvardllgen.ESNULO(txt(4).Text, 0) & ",libronumcorr03=" & VGvardllgen.ESNULO(txt(5).Text, 0) & ","
      SQL1 = SQL1 & "libronumcorr04=" & VGvardllgen.ESNULO(txt(6).Text, 0) & ",libronumcorr05=" & VGvardllgen.ESNULO(txt(7).Text, 0) & ",libronumcorr06=" & VGvardllgen.ESNULO(txt(8).Text, 0) & ","
      SQL1 = SQL1 & "libronumcorr07=" & VGvardllgen.ESNULO(txt(9).Text, 0) & ",libronumcorr08=" & VGvardllgen.ESNULO(txt(10).Text, 0) & ",libronumcorr09=" & VGvardllgen.ESNULO(txt(11).Text, 0) & ","
      SQL1 = SQL1 & "libronumcorr10=" & VGvardllgen.ESNULO(txt(12).Text, 0) & ",libronumcorr11=" & VGvardllgen.ESNULO(txt(13).Text, 0) & ",libronumcorr12=" & VGvardllgen.ESNULO(txt(14).Text, 0) & ","
      SQL1 = SQL1 & "usuariocodigo='" & VGusuario & "',fechaact='" & Format(Date, "dd/mm/yyyy") & "' "
      SQL1 = SQL1 & "WHERE empresacodigo='" & VGParametros.empresacodigo & "' and librocodigo='" & txt(0).Text & "' AND "
      SQL1 = SQL1 & "libroanno='" & VGParamSistem.Anoproceso & "'"
    End If
  End If
  VGCNx.Execute (SQL)
  If Cont = 0 Then
     VGCNx.Execute (SQL1)
  End If
  VGCNx.CommitTrans
  
  Set VGvardllgen = Nothing
  frmbotones.Visible = True
  modoinsert = False: modoedit = False
  Call MuestraDatosLibro
  cAcepta.Enabled = False
  Set VGvardllgen = Nothing
  Call ModoEditable(False, frmMantLibro, "")
  SSTab1.Tab = 0
  SSTab1.TabEnabled(1) = False
  SSTab1.TabEnabled(2) = False
  Exit Sub

x:
  If err.Number = -2147217873 Then
    MsgBox "Esta intentando registrar Código de Sublibro Existente ", vbInformation, Caption
    txt(0).SetFocus
  Else
    MsgBox "Error inesperado: " & err.Number & " " & err.Description, vbInformation, Caption
  End If
  VGCNx.RollbackTrans
     
End Sub

Private Sub cCancela_Click()
  frmbotones.Visible = True
  modoinsert = False: modoedit = False
  cAcepta.Enabled = False
  SSTab1.Tab = 0
  SSTab1.TabEnabled(0) = True
  SSTab1.TabEnabled(1) = False
  SSTab1.TabEnabled(2) = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  If PreviousTab = 0 Then SSTab1.TabEnabled(PreviousTab) = False
End Sub

Private Sub TDBGridAsientos_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
    If rs.Sort = Empty Then
        rs.Sort = TDBGridAsientos.Columns.Item(ColIndex).DataField & " asc"
     ElseIf Right(rs.Sort, 3) = "asc" Then
        rs.Sort = TDBGridAsientos.Columns.Item(ColIndex).DataField & " desc"
     ElseIf Right(rs.Sort, 4) = "desc" Then
        rs.Sort = TDBGridAsientos.Columns.Item(ColIndex).DataField & " asc"
    End If
    TDBGridAsientos.Refresh
End Sub

Private Sub TDBGridAsientos_DblClick()
    If rs.RecordCount > 0 Then Call cmdBotones_Click(1)
End Sub

'FIXIT: Declare 'LastRow' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Private Sub TDBGridAsientos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Call Editarlibro
End Sub

Private Sub txt_Change(Index As Integer)
 If modoinsert = True Or modoedit = True Then
   cAcepta.Enabled = ValidaDataIngreso()
 End If
End Sub

Private Sub chk_Click(Index As Integer)
    If modoinsert = True Or modoedit = True Then cAcepta.Enabled = ValidaDataIngreso()
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If Index = 2 And KeyCode = 13 Then
    cAcepta.SetFocus
 End If

End Sub

Private Sub txt_LostFocus(Index As Integer)
 If Index = 0 Then
   If Not IsNull(txt(0).Text) Then txt(0).Text = Format(txt(0).Text, "00")
 Else
   txt(Index).Text = UCase$(txt(Index).Text)
 End If
End Sub

Sub Editarlibro()
 Dim i As Integer
 
 If rs.RecordCount > 0 Then
    With TDBGridAsientos
        txt(0).Text = .Columns(0).Value
        txt(1).Text = .Columns(1).Value
        chk(0).Value = IIf(.Columns(2).Value = 0, 0, 1)
        For i = 1 To 12
             txt(i + 2).Text = .Columns(i + 2).Value
        Next
    End With
 End If
End Sub

Sub ConfiguraGridLibros()
 Dim i As Integer
 With TDBGridAsientos
   .Columns(0).Width = 1000
   .Columns(1).Width = 3800
   .Columns(2).Width = 1350
   For i = 1 To 12
      .Columns(i + 2).Visible = False
   Next
 End With

End Sub

Function ValidaDataIngreso() As Boolean
 Dim i As Integer
  For i = 0 To 1
   If txt(i).Text = Empty Then
     ValidaDataIngreso = False
     Exit Function
   End If
  Next
  ValidaDataIngreso = True
End Function

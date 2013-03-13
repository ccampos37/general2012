VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmMantSubAsiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sub Asiento"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   6705
   Begin TabDlg.SSTab SSTab1 
      Height          =   7395
      Left            =   30
      TabIndex        =   6
      Top             =   15
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   13044
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Consulta x Asiento"
      TabPicture(0)   =   "frmMantSubAsiento.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblNumReg"
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "TDBGrid1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "frmMantSubAsiento.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblMensaje"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cAcepta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cCancela"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "frmbotones"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Correlativos"
      TabPicture(2)   =   "frmMantSubAsiento.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl(4)"
      Tab(2).Control(1)=   "lbl(5)"
      Tab(2).Control(2)=   "lbl(6)"
      Tab(2).Control(3)=   "lbl(7)"
      Tab(2).Control(4)=   "lbl(8)"
      Tab(2).Control(5)=   "lbl(9)"
      Tab(2).Control(6)=   "lbl(10)"
      Tab(2).Control(7)=   "lbl(11)"
      Tab(2).Control(8)=   "lbl(12)"
      Tab(2).Control(9)=   "lbl(13)"
      Tab(2).Control(10)=   "lbl(14)"
      Tab(2).Control(11)=   "lbl(15)"
      Tab(2).Control(12)=   "txt(3)"
      Tab(2).Control(13)=   "txt(2)"
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
      Begin VB.Frame frmbotones 
         Height          =   555
         Left            =   510
         TabIndex        =   31
         Top             =   6765
         Width           =   5730
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Nuevo"
            Height          =   330
            Index           =   0
            Left            =   60
            TabIndex        =   36
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "E&ditar"
            Height          =   330
            Index           =   1
            Left            =   1185
            TabIndex        =   35
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Eliminar"
            Height          =   330
            Index           =   2
            Left            =   2310
            TabIndex        =   34
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Salir"
            Height          =   330
            Index           =   4
            Left            =   4560
            TabIndex        =   33
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Imprimir"
            Height          =   330
            Index           =   3
            Left            =   3435
            TabIndex        =   32
            Top             =   165
            Width           =   1080
         End
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   3465
         TabIndex        =   28
         Top             =   6360
         Width           =   1140
      End
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   2025
         TabIndex        =   27
         Top             =   6360
         Width           =   1140
      End
      Begin VB.Frame Frame1 
         Height          =   5910
         Left            =   45
         TabIndex        =   7
         Top             =   330
         Width           =   6540
         Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
            Height          =   3270
            Left            =   30
            TabIndex        =   37
            Top             =   2520
            Width           =   6480
            _ExtentX        =   11430
            _ExtentY        =   5768
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "C�digo"
            Columns(0).DataField=   "subasientocodigo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripci�n"
            Columns(1).DataField=   "subasientodescripcion"
            Columns(1).DataWidth=   1700
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Glosa"
            Columns(2).DataField=   "subasientoglosa"
            Columns(2).DataWidth=   1700
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   4
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Repite Doc."
            Columns(3).DataField=   "subasientorepitedoc"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Moneda"
            Columns(4).DataField=   "monedadescripcion"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Corr. Enero"
            Columns(5).DataField=   "subasientonumcorr01"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Corr. Febrero"
            Columns(6).DataField=   "subasientonumcorr02"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Corr. Marzo"
            Columns(7).DataField=   "subasientonumcorr03"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Corr. Abril"
            Columns(8).DataField=   "subasientonumcorr04"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Corr. Mayo"
            Columns(9).DataField=   "subasientonumcorr05"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Corr. Junio"
            Columns(10).DataField=   "subasientonumcorr06"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Corr. Julio"
            Columns(11).DataField=   "subasientonumcorr07"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "Corr. Agosto"
            Columns(12).DataField=   "subasientonumcorr08"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "Corr. Setiembre"
            Columns(13).DataField=   "subasientonumcorr09"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "Corr. Octubre"
            Columns(14).DataField=   "subasientonumcorr10"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "Corr. Noviembre"
            Columns(15).DataField=   "subasientonumcorr11"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(16)._VlistStyle=   0
            Columns(16)._MaxComboItems=   5
            Columns(16).Caption=   "Corr. Diciembre"
            Columns(16).DataField=   "subasientonumcorr12"
            Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(17)._VlistStyle=   0
            Columns(17)._MaxComboItems=   5
            Columns(17).Caption=   "Cod. Moneda"
            Columns(17).DataField=   "monedacodigo"
            Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(18)._VlistStyle=   0
            Columns(18)._MaxComboItems=   5
            Columns(18).DataField=   ""
            Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   19
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=19"
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
            Splits(0)._ColumnProps(16)=   "Column(3)._ColStyle=1"
            Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(18)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(22)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(26)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(30)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(34)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(38)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(41)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(42)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(43)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(44)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(45)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(46)=   "Column(11).Width=2725"
            Splits(0)._ColumnProps(47)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(48)=   "Column(11)._WidthInPix=2646"
            Splits(0)._ColumnProps(49)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(50)=   "Column(12).Width=2725"
            Splits(0)._ColumnProps(51)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(12)._WidthInPix=2646"
            Splits(0)._ColumnProps(53)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(54)=   "Column(13).Width=2725"
            Splits(0)._ColumnProps(55)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(56)=   "Column(13)._WidthInPix=2646"
            Splits(0)._ColumnProps(57)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(58)=   "Column(14).Width=2725"
            Splits(0)._ColumnProps(59)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(60)=   "Column(14)._WidthInPix=2646"
            Splits(0)._ColumnProps(61)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(62)=   "Column(15).Width=2725"
            Splits(0)._ColumnProps(63)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(64)=   "Column(15)._WidthInPix=2646"
            Splits(0)._ColumnProps(65)=   "Column(15).Order=16"
            Splits(0)._ColumnProps(66)=   "Column(16).Width=2725"
            Splits(0)._ColumnProps(67)=   "Column(16).DividerColor=0"
            Splits(0)._ColumnProps(68)=   "Column(16)._WidthInPix=2646"
            Splits(0)._ColumnProps(69)=   "Column(16).Order=17"
            Splits(0)._ColumnProps(70)=   "Column(17).Width=2725"
            Splits(0)._ColumnProps(71)=   "Column(17).DividerColor=0"
            Splits(0)._ColumnProps(72)=   "Column(17)._WidthInPix=2646"
            Splits(0)._ColumnProps(73)=   "Column(17).Order=18"
            Splits(0)._ColumnProps(74)=   "Column(18).Width=2725"
            Splits(0)._ColumnProps(75)=   "Column(18).DividerColor=0"
            Splits(0)._ColumnProps(76)=   "Column(18)._WidthInPix=2646"
            Splits(0)._ColumnProps(77)=   "Column(18).Order=19"
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
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=94,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=91,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=92,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=93,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=90,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=87,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=88,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=89,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=86,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=83,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=84,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=85,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=82,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=79,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=80,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=81,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=78,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=66,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=63,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=64,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=65,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=62,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=59,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=60,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=61,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=58,.parent=13"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=55,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=56,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=57,.parent=17"
            _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=54,.parent=13"
            _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=51,.parent=14"
            _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=52,.parent=15"
            _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=53,.parent=17"
            _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=50,.parent=13"
            _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=47,.parent=14"
            _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=48,.parent=15"
            _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=49,.parent=17"
            _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=46,.parent=13"
            _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=43,.parent=14"
            _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=44,.parent=15"
            _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=45,.parent=17"
            _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=32,.parent=13"
            _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=29,.parent=14"
            _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=30,.parent=15"
            _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=31,.parent=17"
            _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=110,.parent=13"
            _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=107,.parent=14"
            _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=108,.parent=15"
            _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=109,.parent=17"
            _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=28,.parent=13"
            _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=25,.parent=14"
            _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=26,.parent=15"
            _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=27,.parent=17"
            _StyleDefs(112) =   "Named:id=33:Normal"
            _StyleDefs(113) =   ":id=33,.parent=0"
            _StyleDefs(114) =   "Named:id=34:Heading"
            _StyleDefs(115) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(116) =   ":id=34,.wraptext=-1"
            _StyleDefs(117) =   "Named:id=35:Footing"
            _StyleDefs(118) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(119) =   "Named:id=36:Selected"
            _StyleDefs(120) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(121) =   "Named:id=37:Caption"
            _StyleDefs(122) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(123) =   "Named:id=38:HighlightRow"
            _StyleDefs(124) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(125) =   "Named:id=39:EvenRow"
            _StyleDefs(126) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(127) =   "Named:id=40:OddRow"
            _StyleDefs(128) =   ":id=40,.parent=33"
            _StyleDefs(129) =   "Named:id=41:RecordSelector"
            _StyleDefs(130) =   ":id=41,.parent=34"
            _StyleDefs(131) =   "Named:id=42:FilterBar"
            _StyleDefs(132) =   ":id=42,.parent=33"
         End
         Begin VB.CheckBox chk 
            Height          =   300
            Left            =   1890
            TabIndex        =   5
            Top             =   1725
            Width           =   450
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   0
            Left            =   1875
            TabIndex        =   1
            Top             =   465
            Width           =   1455
            _ExtentX        =   2566
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
            MaxLength       =   4
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
            Height          =   345
            Left            =   1890
            TabIndex        =   2
            Top             =   780
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   609
            XcodMaxLongitud =   2
            xcodwith        =   500
            NomTabla        =   "gr_moneda"
            ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
            XcodCampo       =   "monedacodigo"
            XListCampo      =   "monedadescripcion"
            ListaCamposDescrip=   "C�digo,Descripci�n"
            ListaCamposText =   "monedacodigo,monedadescripcion"
            Requerido       =   0   'False
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
            Height          =   315
            Left            =   1890
            TabIndex        =   0
            Top             =   150
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            XcodMaxLongitud =   3
            xcodwith        =   500
            NomTabla        =   "ct_asiento"
            ListaCampos     =   "asientocodigo(1),asientodescripcion(1)"
            XcodCampo       =   "asientocodigo"
            XListCampo      =   "asientodescripcion"
            ListaCamposDescrip=   "C�digo,Descripci�n"
            ListaCamposText =   "asientocodigo,asientodescripcion"
            Requerido       =   0   'False
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   1
            Left            =   1875
            TabIndex        =   3
            Top             =   1110
            Width           =   4590
            _ExtentX        =   8096
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
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   14
            Left            =   1875
            TabIndex        =   4
            Top             =   1425
            Width           =   4590
            _ExtentX        =   8096
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
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "Documento Repetido"
            Height          =   270
            Left            =   135
            TabIndex        =   30
            Top             =   1785
            Width           =   1680
         End
         Begin VB.Label lbl 
            Caption         =   "Glosa"
            Height          =   285
            Index           =   16
            Left            =   120
            TabIndex        =   29
            Top             =   1470
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Descripci�n"
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   1110
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Moneda"
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   810
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Sub Asiento"
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   510
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Asiento"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   210
            Width           =   2310
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3525
         Left            =   -74970
         TabIndex        =   8
         Top             =   1035
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   6218
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
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
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
         MultipleLines   =   1
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=64,.bold=0,.fontsize=825,.italic=0"
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
      Begin TextFer.TxFer txt 
         Height          =   345
         Index           =   8
         Left            =   -72675
         TabIndex        =   16
         Top             =   2925
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
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
         Height          =   345
         Index           =   9
         Left            =   -72675
         TabIndex        =   17
         Top             =   3270
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
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
         Height          =   345
         Index           =   10
         Left            =   -72675
         TabIndex        =   18
         Top             =   3615
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
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
         Height          =   345
         Index           =   4
         Left            =   -72675
         TabIndex        =   12
         Top             =   1545
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
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
         Height          =   345
         Index           =   5
         Left            =   -72675
         TabIndex        =   13
         Top             =   1890
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
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
         Height          =   345
         Index           =   6
         Left            =   -72675
         TabIndex        =   14
         Top             =   2235
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
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
         Height          =   345
         Index           =   7
         Left            =   -72675
         TabIndex        =   15
         Top             =   2580
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
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
         Height          =   345
         Index           =   11
         Left            =   -72675
         TabIndex        =   19
         Top             =   3960
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
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
         Height          =   345
         Index           =   12
         Left            =   -72675
         TabIndex        =   20
         Top             =   4305
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
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
         Height          =   345
         Index           =   13
         Left            =   -72675
         TabIndex        =   21
         Top             =   4650
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
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
         Height          =   345
         Index           =   2
         Left            =   -72675
         TabIndex        =   10
         Top             =   855
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
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
         Enabled         =   0   'False
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   345
         Index           =   3
         Left            =   -72675
         TabIndex        =   11
         Top             =   1200
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
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
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2475
         TabIndex        =   51
         Top             =   6855
         Width           =   1605
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seleccionar un Asiento"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   -74970
         TabIndex        =   50
         Top             =   720
         Width           =   6570
      End
      Begin VB.Label lbl 
         Caption         =   "Diciembre"
         Height          =   285
         Index           =   15
         Left            =   -74310
         TabIndex        =   49
         Top             =   4815
         Width           =   1245
      End
      Begin VB.Label lbl 
         Caption         =   "Noviembre"
         Height          =   285
         Index           =   14
         Left            =   -74310
         TabIndex        =   48
         Top             =   4455
         Width           =   1245
      End
      Begin VB.Label lbl 
         Caption         =   "Octubre"
         Height          =   285
         Index           =   13
         Left            =   -74310
         TabIndex        =   47
         Top             =   4095
         Width           =   1245
      End
      Begin VB.Label lbl 
         Caption         =   "Setiembre"
         Height          =   285
         Index           =   12
         Left            =   -74310
         TabIndex        =   46
         Top             =   3765
         Width           =   1245
      End
      Begin VB.Label lbl 
         Caption         =   "Agosto"
         Height          =   285
         Index           =   11
         Left            =   -74310
         TabIndex        =   45
         Top             =   3405
         Width           =   1245
      End
      Begin VB.Label lbl 
         Caption         =   "Julio"
         Height          =   285
         Index           =   10
         Left            =   -74310
         TabIndex        =   44
         Top             =   3045
         Width           =   1245
      End
      Begin VB.Label lbl 
         Caption         =   "Junio"
         Height          =   285
         Index           =   9
         Left            =   -74310
         TabIndex        =   43
         Top             =   2670
         Width           =   1245
      End
      Begin VB.Label lbl 
         Caption         =   "Mayo"
         Height          =   285
         Index           =   8
         Left            =   -74310
         TabIndex        =   42
         Top             =   2295
         Width           =   1245
      End
      Begin VB.Label lbl 
         Caption         =   "Abril"
         Height          =   285
         Index           =   7
         Left            =   -74310
         TabIndex        =   41
         Top             =   1935
         Width           =   1245
      End
      Begin VB.Label lbl 
         Caption         =   "Marzo"
         Height          =   285
         Index           =   6
         Left            =   -74310
         TabIndex        =   40
         Top             =   1605
         Width           =   1245
      End
      Begin VB.Label lbl 
         Caption         =   "Febrero"
         Height          =   285
         Index           =   5
         Left            =   -74310
         TabIndex        =   39
         Top             =   1290
         Width           =   1245
      End
      Begin VB.Label lbl 
         Caption         =   "Enero"
         Height          =   285
         Index           =   4
         Left            =   -74310
         TabIndex        =   38
         Top             =   915
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "N� Registros"
         Height          =   270
         Left            =   -70260
         TabIndex        =   22
         Top             =   4650
         Width           =   900
      End
      Begin VB.Label lblNumReg 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNumReg"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   -69315
         TabIndex        =   9
         Top             =   4635
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmMantSubAsiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim rs As New ADODB.Recordset
Dim rsAsiento As ADODB.Recordset

Private Sub Form_Load()
  Call ConfiguraForm
  Call MuestraDatosAsiento
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
  Set rsAsiento = Nothing
  Set VGvardllgen = Nothing
End Sub

Sub ConfiguraForm()
  SSTab1.Tab = 0
  SSTab1.TabEnabled(1) = False
  SSTab1.TabEnabled(2) = False
  Ctr_Ayuda1.conexion VGCNx
  Ctr_Ayuda2.conexion VGCNx
  Ctr_Ayuda1.Filtro = "asientocodigo<>'00'"
  'Ctr_Ayuda2.Filtro = "monedacodigo<>'00'"
  cAcepta.Enabled = False
  lblNumReg.Caption = Empty
  Me.Width = 6825
  Me.Height = 7815
End Sub

Sub MuestraDatosAsiento()
 Dim SQL  As String
    Set rsAsiento = New ADODB.Recordset
    SQL = "SELECT asientocodigo as Codigo,asientodescripcion as Descripci�n FROM ct_asiento WHERE asientocodigo<>'00' ORDER BY 1"
    Set rsAsiento = VGCNx.Execute(SQL)
    Set TDBGrid1.DataSource = rsAsiento
    TDBGrid1.Columns(0).Width = 800
    TDBGrid1.Columns(1).Width = 900
    lblNumReg.Caption = rsAsiento.RecordCount
End Sub

Private Sub Ctr_Ayuda1_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Call MuestraDatosSubAsiento
End Sub

'FIXIT: Declare 'MuestraDatosSubAsiento' con un tipo de datos de enlace en tiempo de compilaci�n     FixIT90210ae-R1672-R1B8ZE
Public Function MuestraDatosSubAsiento()
 Dim SQL As String
  SQL = "SELECT ct_subasiento.subasientocodigo, ct_subasiento.subasientodescripcion,"
  SQL = SQL & "ct_subasiento.subasientoglosa,ct_subasiento.subasientorepitedoc,gr_moneda.monedadescripcion,"
  SQL = SQL & "ct_subasiento.subasientonumcorr01, ct_subasiento.subasientonumcorr02, ct_subasiento.subasientonumcorr03, ct_subasiento.subasientonumcorr04,"
  SQL = SQL & "ct_subasiento.subasientonumcorr05, ct_subasiento.subasientonumcorr06, ct_subasiento.subasientonumcorr07, ct_subasiento.subasientonumcorr08,"
  SQL = SQL & "ct_subasiento.subasientonumcorr09, ct_subasiento.subasientonumcorr10, ct_subasiento.subasientonumcorr11, ct_subasiento.subasientonumcorr12,ct_subasiento.monedacodigo "
  SQL = SQL & "FROM ct_subasiento INNER JOIN gr_moneda ON ct_subasiento.monedacodigo = gr_moneda.monedacodigo "
  SQL = SQL & "INNER JOIN ct_asiento ON ct_subasiento.asientocodigo = ct_asiento.asientocodigo "
  SQL = SQL & "WHERE ct_subasiento.subasientocodigo<>'00' AND "
  SQL = SQL & "ct_subasiento.asientocodigo='" & Trim$(Ctr_Ayuda1.xclave) & "' "
  SQL = SQL & "ORDER BY 1,2"
  Set rs = VGCNx.Execute(SQL)
  Set TDBGrid2.DataSource = rs
  Call ConfiguraGridSubAsientos
  If rs.RecordCount <= 0 Then Call LimpiarForm(frmMantSubAsiento, "ctr_ayuda1")
  
End Function

Private Sub cmdBotones_Click(Index As Integer)
  Dim J As Integer
  Dim spos As Integer
  Dim SQL As String
  
  On Error GoTo X
  
  Select Case Index
     Case 0   'nuevo
        SSTab1.TabEnabled(2) = True
        SSTab1.Tab = 1
        'Call LimpiarValores
        
        Call LimpiarForm(frmMantSubAsiento, "Ctr_Ayuda1")
        
        txt(0).Text = GeneraCodigo(VGCNx, "Select max(subasientocodigo) from ct_subasiento where asientocodigo='" & Ctr_Ayuda1.xclave & "'", "0000")
        txt(0).SetFocus
        Call ModoEditable(True, frmMantSubAsiento, "Ctr_Ayuda1")
        frmbotones.Visible = False
        modoinsert = True
        lblMensaje.Caption = "Nuevo"
        
     Case 1   'modificar
        If TDBGrid1.Row < 0 Then
          Exit Sub
        End If
        SSTab1.TabEnabled(2) = True
        SSTab1.Tab = 1
        modoedit = True
        frmbotones.Visible = False
        Call ModoEditable(True, frmMantSubAsiento, "Ctr_Ayuda1")
        lblMensaje.Caption = "Editar"
      
     Case 2   'eliminar
       If MsgBox("Desea eliminar el registro de SubAsiento N� " & TDBGrid2.Columns(0).Value & "?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          SQL = "DELETE FROM CT_SUBASIENTO WHERE subasientocodigo='" & Trim$(TDBGrid2.Columns(0).Value) & "' AND "
          SQL = SQL & "asientocodigo='" & Trim$(Ctr_Ayuda1.xclave) & "'"
          VGCNx.Execute (SQL)
          Call MuestraDatosSubAsiento
       End If
        
     Case 3   'imprimir
       Call Impresion("rptSubAsiento.rpt")
     
     Case 4  ' salir
       Unload Me
  End Select
  
  Exit Sub
   
X:
  If Index = 2 And err.Number = -2147217873 Then
    MsgBox "Registro no podr� Eliminarse mientras exista Informaci�n en la Tablas Relacionadas", vbInformation, Caption
  Else
    MsgBox "Error inesperado: " & err.Description & "  " & err.Number, vbInformation, Caption
  End If
   
End Sub

Private Sub cAcepta_Click()
  Dim SQL As String
  On Error GoTo X
  
  Set VGvardllgen = New dllgeneral.dll_general
  VGCNx.BeginTrans
  
  If modoinsert = True Then
    SQL = "INSERT INTO CT_SUBASIENTO (subasientocodigo,asientocodigo,monedacodigo,subasientodescripcion,"
    SQL = SQL & "subasientonumcorr01,subasientonumcorr02,subasientonumcorr03,subasientonumcorr04,subasientonumcorr05,subasientonumcorr06,subasientonumcorr07,subasientonumcorr08,subasientonumcorr09,subasientonumcorr10,subasientonumcorr11,subasientonumcorr12,subasientoglosa,subasientorepitedoc,usuariocodigo,fechaact) "
    SQL = SQL & "VALUES ('" & txt(0).Text & "','" & Ctr_Ayuda1.xclave & "','" & Ctr_Ayuda2.xclave & "','" & Trim$(UCase$(txt(1).Text)) & "',"
    SQL = SQL & VGvardllgen.ESNULO(txt(2).Text, 0) & "," & VGvardllgen.ESNULO(txt(3).Text, 0) & "," & VGvardllgen.ESNULO(txt(4).Text, 0) & ","
    SQL = SQL & VGvardllgen.ESNULO(txt(5).Text, 0) & "," & VGvardllgen.ESNULO(txt(6).Text, 0) & "," & VGvardllgen.ESNULO(txt(7).Text, 0) & ","
    SQL = SQL & VGvardllgen.ESNULO(txt(8).Text, 0) & "," & VGvardllgen.ESNULO(txt(9).Text, 0) & "," & VGvardllgen.ESNULO(txt(10).Text, 0) & "," & VGvardllgen.ESNULO(txt(11).Text, 0) & "," & VGvardllgen.ESNULO(txt(12).Text, 0) & "," & VGvardllgen.ESNULO(txt(13).Text, 0) & ",'" & txt(14).Text & "'," & chk.Value & ",'"
    SQL = SQL & VGusuario & "','" & Format(Date, "dd/mm/yyyy") & "')"
                  
  ElseIf modoedit = True Then
    SQL = "UPDATE CT_SUBASIENTO SET monedacodigo='" & Ctr_Ayuda2.xclave & "',"
    SQL = SQL & "subasientodescripcion='" & Trim$(UCase$(txt(1).Text)) & "',"
    SQL = SQL & "subasientonumcorr01=" & VGvardllgen.ESNULO(txt(2).Text, 0) & ",subasientonumcorr02=" & VGvardllgen.ESNULO(txt(3).Text, 0) & ",subasientonumcorr03=" & VGvardllgen.ESNULO(txt(4).Text, 0) & ","
    SQL = SQL & "subasientonumcorr04=" & VGvardllgen.ESNULO(txt(5).Text, 0) & ",subasientonumcorr05=" & VGvardllgen.ESNULO(txt(6).Text, 0) & ",subasientonumcorr06=" & VGvardllgen.ESNULO(txt(7).Text, 0) & ","
    SQL = SQL & "subasientonumcorr07=" & VGvardllgen.ESNULO(txt(8).Text, 0) & ",subasientonumcorr08=" & VGvardllgen.ESNULO(txt(9).Text, 0) & ",subasientonumcorr09=" & VGvardllgen.ESNULO(txt(10).Text, 0) & ","
    SQL = SQL & "subasientonumcorr10=" & VGvardllgen.ESNULO(txt(11).Text, 0) & ",subasientonumcorr11=" & VGvardllgen.ESNULO(txt(12).Text, 0) & ",subasientonumcorr12=" & VGvardllgen.ESNULO(txt(13).Text, 0) & ","
    SQL = SQL & "subasientoglosa='" & txt(14).Text & "',subasientorepitedoc=" & chk.Value & ","
    SQL = SQL & "usuariocodigo='" & VGusuario & "',fechaact='" & Format(Date, "dd/mm/yyyy") & "' "
    SQL = SQL & "WHERE subasientocodigo='" & txt(0).Text & "' AND asientocodigo='" & Ctr_Ayuda1.xclave & "'"
  End If
  
  VGCNx.Execute (SQL)
  VGCNx.CommitTrans
  
  Set VGvardllgen = Nothing
  frmbotones.Visible = True
  modoinsert = False: modoedit = False: lblMensaje.Caption = Empty
  Call MuestraDatosSubAsiento
  cAcepta.Enabled = False
  Set VGvardllgen = Nothing
  Call ModoEditable(False, frmMantSubAsiento, "")
  Exit Sub

X:
  If err.Number = -2147217873 Then
    MsgBox "Esta intentando registrar C�digo de SubAsiento Existente ", vbInformation, Caption
    txt(0).SetFocus
  Else
    MsgBox "Error inesperado: " & err.Number & " " & err.Description, vbInformation, Caption
  End If
  VGCNx.RollbackTrans
     
End Sub

Private Sub cCancela_Click()
  frmbotones.Visible = True
  modoinsert = False: modoedit = False: lblMensaje.Caption = Empty
  cAcepta.Enabled = False
  SSTab1.Tab = 0
  SSTab1.TabEnabled(0) = True
  SSTab1.TabEnabled(1) = False
  SSTab1.TabEnabled(2) = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  If PreviousTab = 0 Then SSTab1.TabEnabled(PreviousTab) = False
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
    If rs.Sort = Empty Then
        rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " asc"
     ElseIf Right(rs.Sort, 3) = "asc" Then
        rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " desc"
     ElseIf Right(rs.Sort, 4) = "desc" Then
        rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " asc"
    End If
    TDBGrid1.Refresh
End Sub

Private Sub TDBGrid1_DblClick()
 If rsAsiento.RecordCount > 0 Then
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = False
   SSTab1.Tab = 1
   Ctr_Ayuda1.xclave = TDBGrid1.Columns(0).Text: Ctr_Ayuda1.Ejecutar
   Ctr_Ayuda1.Enabled = False
   Call ModoEditable(False, frmMantSubAsiento, "Ctr_Ayuda1")
   cAcepta.Enabled = False
 End If
End Sub

'FIXIT: Declare 'LastRow' con un tipo de datos de enlace en tiempo de compilaci�n          FixIT90210ae-R1672-R1B8ZE
Private Sub TDBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Call EditarSubAsiento
End Sub

Private Sub txt_Change(Index As Integer)
 If modoinsert = True Or modoedit = True Then
   cAcepta.Enabled = ValidaDataIngreso()
 End If
End Sub

Private Sub chk_Click()
    If modoinsert = True Or modoedit = True Then cAcepta.Enabled = ValidaDataIngreso()
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 And Index = 15 Then
    cAcepta.SetFocus
    Call cAcepta_Click
  End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
 If Index = 0 Then
   If Not IsNull(txt(0).Text) Then txt(0).Text = Format(txt(0).Text, "0000")
 Else
   txt(Index).Text = UCase$(txt(Index).Text)
 End If
End Sub

Sub EditarSubAsiento()
 Dim i As Integer
 
 If rs.RecordCount > 0 Then
    With TDBGrid2
        For i = 2 To 13
             txt(i).Text = TDBGrid2.Columns(i + 3).Value
        Next
        txt(0).Text = .Columns(0).Value
        txt(1).Text = .Columns(1).Value
        txt(14).Text = .Columns(2).Value
        chk.Value = IIf(.Columns(3).Value = 0, 0, 1)
        Ctr_Ayuda2.xclave = .Columns(17).Value: Ctr_Ayuda2.Ejecutar
    End With
 
 End If
End Sub

Sub ConfiguraGridSubAsientos()
 Dim i As Integer
 With TDBGrid2
   .Columns(0).Width = 700
   .Columns(1).Width = 2500
   .Columns(2).Width = 1900
   .Columns(3).Width = 1000
   .Columns(4).Width = 900
   For i = 5 To 16
      .Columns(i).Width = 1250
   Next
 End With

End Sub

Function ValidaDataIngreso() As Boolean
 Dim i As Integer
  If Ctr_Ayuda1.xclave = Empty Then
    ValidaDataIngreso = False
    Exit Function
  End If
   
  If Ctr_Ayuda2.xclave = Empty Then
    ValidaDataIngreso = False
    Exit Function
  End If
  
  For i = 0 To 1
   If txt(i).Text = Empty Then
     ValidaDataIngreso = False
     Exit Function
   End If
   
  Next
  
  If txt(14).Text = Empty Then
     ValidaDataIngreso = False
     Exit Function
  End If

  ValidaDataIngreso = True
End Function


'Private Function LimpiarValores()
'Dim I As Integer
' Ctr_Ayuda2.xclave = Empty: Ctr_Ayuda2.Ejecutar
' For I = 0 To 14
'  txt(I).Text = Empty
' Next
' chk.Value = 0
'
'End Function


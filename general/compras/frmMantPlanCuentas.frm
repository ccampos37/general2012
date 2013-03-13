VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "textfer.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmMantPlanCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan de Cuentas"
   ClientHeight    =   7212
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   11184
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7212
   ScaleWidth      =   11184
   Begin TabDlg.SSTab SSTab1 
      Height          =   6585
      Left            =   4590
      TabIndex        =   19
      Top             =   630
      Width           =   6570
      _ExtentX        =   11599
      _ExtentY        =   11621
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "frmMantPlanCuentas.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "lblNumReg"
      Tab(0).Control(2)=   "TDBGrid1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "frmMantPlanCuentas.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cAcepta"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cCancela"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Cuentas Distribución"
      TabPicture(2)   =   "frmMantPlanCuentas.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdSigue"
      Tab(2).Control(1)=   "txtPorcen"
      Tab(2).Control(2)=   "cmdAceptar"
      Tab(2).Control(3)=   "Ctr_Ayuda3"
      Tab(2).Control(4)=   "Ctr_Ayuda2"
      Tab(2).Control(5)=   "TDBGrid2"
      Tab(2).Control(6)=   "lblPorcen"
      Tab(2).Control(7)=   "Label5"
      Tab(2).Control(8)=   "lbl(13)"
      Tab(2).Control(9)=   "lbl(12)"
      Tab(2).Control(10)=   "Label4"
      Tab(2).ControlCount=   11
      Begin VB.CommandButton cmdSigue 
         Caption         =   "&Aceptar"
         Height          =   360
         Left            =   -72480
         TabIndex        =   52
         Top             =   5175
         Width           =   1125
      End
      Begin TextFer.TxFer txtPorcen 
         Height          =   315
         Left            =   -72075
         TabIndex        =   24
         Top             =   2595
         Width           =   1305
         _ExtentX        =   2307
         _ExtentY        =   550
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         NumeroDecimales =   2
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "..."
         Height          =   300
         Left            =   -70725
         TabIndex        =   25
         Top             =   2610
         Width           =   270
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   5490
         Left            =   -74970
         TabIndex        =   38
         Top             =   675
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   9694
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Cuenta"
         Columns(0).DataField=   "cuentacodigo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descripción"
         Columns(1).DataField=   "cuentadescripcion"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Linea Act."
         Columns(2).DataField=   "cuentalineaactivo"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Linea Pas."
         Columns(3).DataField=   "cuentalineapasivo"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Cta. Nat."
         Columns(4).DataField=   "cuentanatu"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Cta Linea EGP"
         Columns(5).DataField=   "cuentalineaegp"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Cta Nat EGP"
         Columns(6).DataField=   "cuentanategp"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   4
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Cta. Analitico"
         Columns(7).DataField=   "cuentaestadoanalitico"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Tipo Analitico"
         Columns(8).DataField=   "tipoanaliticocodigo"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   4
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Cta Costos"
         Columns(9).DataField=   "cuentaestadocostos"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   4
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Cta. Doc"
         Columns(10).DataField=   "cuentadocumento"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Cta Nivel"
         Columns(11).DataField=   "cuentanivel"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Tipo Analit Desc."
         Columns(12).DataField=   "tipoanaliticodescripcion"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   4
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Cta. Est Dist."
         Columns(13).DataField=   "cuentaestadodistribucion"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   4
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Cta. Estado"
         Columns(14).DataField=   "cuentaestado"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "Tipo Cuenta"
         Columns(15).DataField=   "tipocuentacodigo"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "Tipo Ajuste"
         Columns(16).DataField=   "tipoajuste"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   17
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   508
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=17"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2731"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2731"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2731"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2731"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2731"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=2731"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2731"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=2731"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=2731"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=2731"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=2731"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(45)=   "Column(11).Width=2731"
         Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(49)=   "Column(12).Width=2731"
         Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(53)=   "Column(13).Width=2731"
         Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(57)=   "Column(14).Width=2731"
         Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=2646"
         Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(61)=   "Column(15).Width=2731"
         Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=2646"
         Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(65)=   "Column(16).Width=2731"
         Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=2646"
         Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000014&,.bold=0,.fontsize=825"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
         _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
         _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
         _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
         _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
         _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
         _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
         _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
         _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
         _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=98,.parent=13"
         _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=95,.parent=14"
         _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=96,.parent=15"
         _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=97,.parent=17"
         _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=102,.parent=13"
         _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=14"
         _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=15"
         _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=17"
         _StyleDefs(104) =   "Named:id=33:Normal"
         _StyleDefs(105) =   ":id=33,.parent=0"
         _StyleDefs(106) =   "Named:id=34:Heading"
         _StyleDefs(107) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(108) =   ":id=34,.wraptext=-1"
         _StyleDefs(109) =   "Named:id=35:Footing"
         _StyleDefs(110) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(111) =   "Named:id=36:Selected"
         _StyleDefs(112) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(113) =   "Named:id=37:Caption"
         _StyleDefs(114) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(115) =   "Named:id=38:HighlightRow"
         _StyleDefs(116) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(117) =   "Named:id=39:EvenRow"
         _StyleDefs(118) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(119) =   "Named:id=40:OddRow"
         _StyleDefs(120) =   ":id=40,.parent=33"
         _StyleDefs(121) =   "Named:id=41:RecordSelector"
         _StyleDefs(122) =   ":id=41,.parent=34"
         _StyleDefs(123) =   "Named:id=42:FilterBar"
         _StyleDefs(124) =   ":id=42,.parent=33"
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         Height          =   360
         Left            =   3795
         TabIndex        =   36
         Top             =   5730
         Width           =   1125
      End
      Begin VB.Frame Frame2 
         Height          =   4800
         Left            =   45
         TabIndex        =   21
         Top             =   660
         Width           =   6480
         Begin VB.ComboBox cboTipoAjuste 
            Height          =   315
            ItemData        =   "frmMantPlanCuentas.frx":0054
            Left            =   2730
            List            =   "frmMantPlanCuentas.frx":0061
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   2625
            Width           =   3690
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda4 
            Height          =   330
            Left            =   2745
            TabIndex        =   7
            Top             =   2310
            Width           =   3660
            _ExtentX        =   6456
            _ExtentY        =   572
            Enabled         =   0   'False
            XcodMaxLongitud =   2
            xcodwith        =   500
            NomTabla        =   "ct_tipocuenta"
            ListaCampos     =   "tipocuentacodigo(1),tipocuentadescripcion(1)"
            XcodCampo       =   "tipocuentacodigo"
            XListCampo      =   "tipocuentadescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "tipocuentacodigo,tipocuentadescripcion"
         End
         Begin VB.CommandButton cmdDistribucion 
            Caption         =   "..."
            Height          =   210
            Left            =   2985
            TabIndex        =   15
            Top             =   4110
            Width           =   255
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   4
            Left            =   2730
            TabIndex        =   14
            Top             =   4395
            Width           =   210
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   3
            Left            =   2730
            TabIndex        =   13
            Top             =   4110
            Width           =   195
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   1
            Left            =   2730
            TabIndex        =   11
            Top             =   3570
            Width           =   285
         End
         Begin VB.CheckBox chk 
            Height          =   210
            Index           =   0
            Left            =   2730
            TabIndex        =   9
            Top             =   2985
            Width           =   225
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
            Height          =   315
            Left            =   2730
            TabIndex        =   10
            Top             =   3180
            Width           =   3660
            _ExtentX        =   6456
            _ExtentY        =   550
            XcodMaxLongitud =   3
            xcodwith        =   500
            NomTabla        =   "ct_tipoanalitico"
            ListaCampos     =   "tipoanaliticocodigo(1),tipoanaliticodescripcion(1)"
            XcodCampo       =   "tipoanaliticocodigo"
            XListCampo      =   "tipoanaliticodescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "tipoanaliticocodigo,tipoanaliticodescripcion"
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   0
            Left            =   2745
            TabIndex        =   0
            Top             =   150
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   529
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   20
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
            Index           =   1
            Left            =   2745
            TabIndex        =   1
            Top             =   480
            Width           =   3630
            _ExtentX        =   6414
            _ExtentY        =   529
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   35
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   2
            Left            =   2745
            TabIndex        =   2
            Top             =   810
            Width           =   3630
            _ExtentX        =   6414
            _ExtentY        =   529
            BackColor       =   16777215
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   10
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
            Left            =   2745
            TabIndex        =   3
            Top             =   1110
            Width           =   3630
            _ExtentX        =   6414
            _ExtentY        =   529
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   10
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
            Left            =   2730
            TabIndex        =   4
            Top             =   1410
            Width           =   825
            _ExtentX        =   1461
            _ExtentY        =   529
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   1
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   5
            Left            =   2730
            TabIndex        =   5
            Top             =   1710
            Width           =   3630
            _ExtentX        =   6414
            _ExtentY        =   529
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   10
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
            Left            =   2730
            TabIndex        =   6
            Top             =   2010
            Width           =   825
            _ExtentX        =   1461
            _ExtentY        =   529
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   1
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   2
            Left            =   2730
            TabIndex        =   12
            Top             =   3840
            Width           =   240
         End
         Begin VB.Label lbl 
            Caption         =   "Tipo de Ajuste"
            Height          =   240
            Index           =   15
            Left            =   135
            TabIndex        =   55
            Top             =   2700
            Width           =   2325
         End
         Begin VB.Label lbl 
            Caption         =   "Tipo Cuenta"
            Height          =   240
            Index           =   14
            Left            =   135
            TabIndex        =   53
            Top             =   2400
            Width           =   2325
         End
         Begin VB.Label lbl 
            Caption         =   "Cuenta Activa"
            Height          =   240
            Index           =   11
            Left            =   150
            TabIndex        =   45
            Top             =   4410
            Width           =   2325
         End
         Begin VB.Label lbl 
            Caption         =   "Asiento Automático"
            Height          =   240
            Index           =   10
            Left            =   150
            TabIndex        =   44
            Top             =   4140
            Width           =   2325
         End
         Begin VB.Label lblNivel 
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   5760
            TabIndex        =   43
            Top             =   180
            Width           =   600
         End
         Begin VB.Label Label3 
            Caption         =   "Nivel Cuenta"
            Height          =   270
            Left            =   4710
            TabIndex        =   42
            Top             =   210
            Width           =   945
         End
         Begin VB.Label lbl 
            Caption         =   "Controla Documento"
            Height          =   240
            Index           =   9
            Left            =   150
            TabIndex        =   41
            Top             =   3855
            Width           =   2325
         End
         Begin VB.Label lbl 
            Caption         =   "Controla Centro Costo"
            Height          =   240
            Index           =   8
            Left            =   150
            TabIndex        =   35
            Top             =   3570
            Width           =   2565
         End
         Begin VB.Label lbl 
            Caption         =   "Tipo Analítico"
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   34
            Top             =   3285
            Width           =   2550
         End
         Begin VB.Label lbl 
            Caption         =   "Controla Código Análisis"
            Height          =   330
            Index           =   6
            Left            =   135
            TabIndex        =   33
            Top             =   3000
            Width           =   2550
         End
         Begin VB.Label lbl 
            Caption         =   "Naturaleza de Estado Gan. y Pérd."
            Height          =   330
            Index           =   5
            Left            =   135
            TabIndex        =   32
            Top             =   2115
            Width           =   2550
         End
         Begin VB.Label lbl 
            Caption         =   "Línea de Estado de Gan. y Pérd."
            Height          =   330
            Index           =   4
            Left            =   135
            TabIndex        =   31
            Top             =   1815
            Width           =   2550
         End
         Begin VB.Label lbl 
            Caption         =   "Naturaleza de Balance"
            Height          =   330
            Index           =   3
            Left            =   135
            TabIndex        =   30
            Top             =   1500
            Width           =   2550
         End
         Begin VB.Label lbl 
            Caption         =   "Línea de Balance de Pasivo"
            Height          =   330
            Index           =   2
            Left            =   135
            TabIndex        =   29
            Top             =   1200
            Width           =   2550
         End
         Begin VB.Label lbl 
            Caption         =   "Línea de Balance de Activo"
            Height          =   330
            Index           =   1
            Left            =   135
            TabIndex        =   28
            Top             =   900
            Width           =   2550
         End
         Begin VB.Label lbl 
            Caption         =   "Descripción Cuenta"
            Height          =   225
            Index           =   0
            Left            =   135
            TabIndex        =   27
            Top             =   570
            Width           =   2550
         End
         Begin VB.Label Label2 
            Caption         =   "Código Cuenta"
            Height          =   315
            Left            =   135
            TabIndex        =   26
            Top             =   240
            Width           =   2550
         End
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda3 
         Height          =   450
         Left            =   -74850
         TabIndex        =   23
         Top             =   1875
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   804
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Cuenta,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   345
         Left            =   -74835
         TabIndex        =   22
         Top             =   1155
         Width           =   4350
         _ExtentX        =   7684
         _ExtentY        =   614
         XcodMaxLongitud =   20
         xcodwith        =   800
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Cuenta,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Height          =   1635
         Left            =   -74970
         TabIndex        =   46
         Top             =   3015
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   2879
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
         Splits(0).RecordSelectorWidth=   508
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2731"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2731"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
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
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         Height          =   360
         Left            =   2010
         TabIndex        =   37
         Top             =   5730
         Width           =   1125
      End
      Begin VB.Label lblPorcen 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   -69705
         TabIndex        =   51
         Top             =   4695
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Total (%)"
         Height          =   255
         Left            =   -70440
         TabIndex        =   50
         Top             =   4740
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Seleccionar Cuenta al Cargo"
         Height          =   255
         Index           =   13
         Left            =   -74790
         TabIndex        =   49
         Top             =   900
         Width           =   4035
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "Seleccionar Cuenta al Abono"
         Height          =   255
         Index           =   12
         Left            =   -74835
         TabIndex        =   48
         Top             =   1635
         Width           =   4095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Porcentaje de Distribución (%)"
         ForeColor       =   &H0080FFFF&
         Height          =   300
         Left            =   -74955
         TabIndex        =   47
         Top             =   2610
         Width           =   2895
      End
      Begin VB.Label lblNumReg 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNumReg"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   -69405
         TabIndex        =   40
         Top             =   6225
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Registros"
         Height          =   270
         Left            =   -70350
         TabIndex        =   39
         Top             =   6240
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   504
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11184
      _ExtentX        =   19727
      _ExtentY        =   889
      ButtonWidth     =   1397
      ButtonHeight    =   847
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
            Object.ToolTipText     =   "Añadir registro al Plan Cuentas"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&ditar"
            Object.ToolTipText     =   "Modificar Plan de Cuentas"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Eliminar"
            Object.ToolTipText     =   "Eliminar Plan de Cuentas"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Imprimir"
            Object.ToolTipText     =   "Listar Plan de Cuentas"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Object.ToolTipText     =   "Cerrar la Ventana "
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   7305
      Left            =   -15
      TabIndex        =   16
      Top             =   -105
      Width           =   4575
      Begin TextFer.TxFer txtBuscar 
         Height          =   300
         Left            =   45
         TabIndex        =   54
         Top             =   765
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   529
         BackColor       =   16777215
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         Enabled         =   0   'False
         ColorIlumina    =   -2147483624
         Valor           =   ""
         NoCaracteres    =   "0123456789"
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
         NoRangoCadena   =   -1  'True
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   135
         Top             =   6615
         _ExtentX        =   995
         _ExtentY        =   995
         BackColor       =   -2147483643
         ImageWidth      =   13
         ImageHeight     =   13
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantPlanCuentas.frx":00A3
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantPlanCuentas.frx":01AB
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   300
         Left            =   4260
         TabIndex        =   20
         Top             =   780
         Width           =   285
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   6210
         Left            =   45
         TabIndex        =   17
         Top             =   1080
         Width           =   4485
         _ExtentX        =   7916
         _ExtentY        =   10964
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "frmMantPlanCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim i_filaorigen As Integer
Dim rs As New ADODB.Recordset
Dim rsDist As New ADODB.Recordset
Dim cCta As String
Dim dCta As String
Dim xCuenta As String
Dim xdllgen As New dll_general
Dim FLAGMOVIMIENTODISTRI As Boolean
Dim FLAGDISTRIBUCION As Boolean
Dim l_error As String

Private Sub Form_Load()
  Screen.MousePointer = 11
  l_error = Empty
  Call ConfiguraForm
  Call MuestraDatos(Empty)
  Call Arbol(txtBuscar.Text)
  Set xdllgen = New dll_general
  If Len(l_error) > 0 Then
    frmError.RichError.Text = l_error
    Screen.MousePointer = 1
    frmError.Show 1
  End If
  Screen.MousePointer = 1
  TDBGrid1.FetchRowStyle = True
  xCuenta = "%"
End Sub

Sub ConfiguraForm()
  SSTab1.Tab = 0
  SSTab1.TabEnabled(1) = False
  SSTab1.TabEnabled(2) = False
  Ctr_Ayuda1.conexion VGcnx
  Ctr_Ayuda2.conexion VGcnx
  Ctr_Ayuda3.conexion VGcnx
  Ctr_Ayuda2.Filtro = "cuentanivel=" & VGnumniveles
  Ctr_Ayuda3.Filtro = "cuentanivel=" & VGnumniveles
  Ctr_Ayuda4.conexion VGcnx
  cAcepta.Enabled = False
  lblNumReg.Caption = Empty
  frmMantPlanCuentas.Width = 11310
  frmMantPlanCuentas.Height = 7650
  FLAGMOVIMIENTODISTRI = False
  FLAGDISTRIBUCION = False
  Call IniciaGridDist
End Sub

Public Function MuestraDatos(xCta As String)
  Dim SQL As String
   SQL = "SELECT ct_cuenta.cuentacodigo, ct_cuenta.cuentadescripcion, ct_cuenta.cuentalineaactivo, ct_cuenta.cuentalineapasivo,"
   SQL = SQL & "ct_cuenta.cuentanatu, ct_cuenta.cuentalineaegp,ct_cuenta.cuentanategp, ct_cuenta.cuentaestadoanalitico,"
   SQL = SQL & "ct_cuenta.tipoanaliticocodigo, ct_cuenta.cuentaestadoccostos,ct_cuenta.cuentadocumento,ct_cuenta.cuentanivel,ct_tipoanalitico.tipoanaliticodescripcion,ct_cuenta.cuentaestadodistribucion,ct_cuenta.cuentaestado,tipocuentacodigo,tipoajuste "
   SQL = SQL & "FROM ct_cuenta INNER JOIN  ct_tipoanalitico ON ct_cuenta.tipoanaliticocodigo = ct_tipoanalitico.tipoanaliticocodigo "
   SQL = SQL & "WHERE ct_cuenta.cuentacodigo<>'00'"
   If xCta <> Empty Then
     SQL = SQL & "AND ct_cuenta.cuentacodigo like '" & Trim(xCta) & "%' "
   End If
   SQL = SQL & "ORDER BY 1"
   Set rs = VGcnx.Execute(SQL)
   Set TDBGrid1.DataSource = rs
   Call ConfiguraTdbgrid
   lblNumReg.Caption = rs.RecordCount
   SSTab1.Tab = 0
End Function

Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
      Call cmdBuscar_Click
    End If
End Sub

Private Sub cmdBuscar_Click()
  Set VGvardllgen = New dllgeneral.dll_general
  txtBuscar.Text = VGvardllgen.ESNULO(txtBuscar.Text, "%")
  Call MuestraDatos(txtBuscar.Text)
  Call Arbol(txtBuscar.Text)
End Sub

Sub EditarValores()
 Dim i As Integer
  With TDBGrid1
    For i = 0 To 6
      txt(i).Text = Trim(xdllgen.ESNULO(.Columns(i).Text, Empty))
    Next
    chk(0).Value = IIf(Trim(.Columns(7).Text) = -1, 1, 0)
    Ctr_Ayuda1.xclave = Trim(xdllgen.ESNULO(.Columns(8).Text, Empty)): Ctr_Ayuda1.Ejecutar
    chk(1).Value = IIf(Trim(.Columns(9).Text) = -1, 1, 0)
    chk(2).Value = IIf(Trim(.Columns(10).Text) = -1, 1, 0)
    chk(3).Value = IIf(Trim(.Columns(13).Text) = -1, 1, 0)
    chk(4).Value = IIf(Trim(.Columns(14).Text) = -1, 1, 0)
    lblNivel.Caption = Trim(xdllgen.ESNULO(.Columns(11).Text, Empty))
    Ctr_Ayuda4.xclave = xdllgen.ESNULO(.Columns(15).Text, Empty): Ctr_Ayuda4.Ejecutar
    
    cboTipoAjuste.ListIndex = CInt(.Columns(16).Text)
    
  End With
  Call ConfiguraModoEdicion
End Sub

Sub ConfiguraModoEdicion()
    If lblNivel.Caption = Empty Then
        MsgBox "Debe registrar el Código de Cuenta Contable", vbInformation, Caption
        Call ModoNormal  'Deshabilitar todos los objetos de ingreso
        txt(0).SetFocus
    Else
        If lblNivel.Caption = 1 Then  'Preguntar por 1º Nivel
           Ctr_Ayuda4.Enabled = True
        Else
           Ctr_Ayuda4.Enabled = False
        End If
        
        If lblNivel.Caption = VGnumniveles Then
           Call ModoEdicion(True)
        Else
           Call ModoEdicion(False)
        End If
    End If
End Sub

Public Function LimpiarValores()
 Dim i As Integer
  Ctr_Ayuda1.xclave = Empty: Ctr_Ayuda1.Ejecutar
  Ctr_Ayuda4.xclave = Empty: Ctr_Ayuda4.Ejecutar
  For i = 0 To 6
    txt(i).Text = Empty
  Next
  For i = 0 To 3
    chk(i).Value = 0
  Next
  chk(4).Value = 1
  lblNivel.Caption = Empty
  'cboTipoAjuste.SelText = Empty
  
End Function

Private Sub cAcepta_Click()
 If ValidarData = True Then
   Call GrabarData
 End If
End Sub

Private Sub cCancela_Click()
  SSTab1.TabEnabled(0) = True
  SSTab1.Tab = 0
  SSTab1.SetFocus
  Toolbar1.Visible = True
  TreeView1.Enabled = True
  modoinsert = False
  modoedit = False
  i_filaorigen = -1
  FLAGDISTRIBUCION = False
End Sub

Sub GrabarData()
  Dim SQL As String
  'On Error GoTo X
  
  SSTab1.TabEnabled(0) = True
  If cboTipoAjuste.ListIndex < 0 Then cboTipoAjuste.ListIndex = 0
  
  If modoinsert = True Then
    VGcnx.BeginTrans
    SQL = GrabarPlanCuenta(0, txt(0).Text, txt(1).Text, Val(txt(2).Text), Val(txt(3).Text), Val(txt(5).Text), txt(4).Text, txt(6).Text, chk(1).Value, chk(0).Value, chk(2).Value, CInt(lblNivel.Caption), Ctr_Ayuda1.xclave, chk(3).Value, Ctr_Ayuda4.xclave, Left(cboTipoAjuste.List(cboTipoAjuste.ListIndex), 2))
    VGcnx.Execute (SQL)
    
    If CInt(lblNivel.Caption) = VGnumniveles And FLAGDISTRIBUCION = True Then
       Call GrabarCuentaDistribucion
       Call GrabarTablaSaldos
    End If
    
    VGcnx.CommitTrans
    Call Arbol(txtBuscar.Text)
                  
  ElseIf modoedit = True Then
    VGcnx.BeginTrans
    SQL = GrabarPlanCuenta(1, txt(0).Text, txt(1).Text, Val(txt(2).Text), Val(txt(3).Text), Val(txt(5).Text), txt(4).Text, txt(6).Text, chk(1).Value, chk(0).Value, chk(2).Value, CInt(lblNivel.Caption), Ctr_Ayuda1.xclave, chk(3).Value, Ctr_Ayuda4.xclave, Left(cboTipoAjuste.List(cboTipoAjuste.ListIndex), 2))
    VGcnx.Execute (SQL)
    
    If CInt(lblNivel.Caption) = VGnumniveles And FLAGDISTRIBUCION = True Then
        Call GrabarCuentaDistribucion
    End If
    If lblNivel.Caption = 1 Then
        Call GrabaTipoCuenta(txt(0).Text, Ctr_Ayuda4.xclave, lblNivel.Caption)
    End If
    
    VGcnx.CommitTrans
  End If
  
  Call MuestraDatos(Right(Trim(cCta), CLng(Len(Trim(cCta))) - 1))
  Toolbar1.Visible = True: TreeView1.Enabled = True: txt(0).Enabled = True
  modoinsert = False: modoedit = False
  i_filaorigen = -1
  Ctr_Ayuda1.Enabled = False
  FLAGDISTRIBUCION = False
  FLAGMOVIMIENTODISTRI = False
  Set rsDist = Nothing
  Exit Sub

X:
  If Err.Number = -2147217873 Then
    MsgBox "Esta intentando registrar uno Existente " & Err.Description, vbInformation, Caption
  Else
    MsgBox "Error inesperado: " & Err.Number & " " & Err.Description, Caption
  End If
  VGcnx.RollbackTrans

End Sub

Function ValidarData() As Boolean
 Dim i As Integer
 Dim SQL As String
  If lblNivel.Caption = Empty Then
    MsgBox "No se ha podido registrar el Número de Nivel de la Cuenta Contable", vbInformation, Caption
    ValidarData = False
    txt(0).SetFocus
    Exit Function
  End If
  
  If VG_aNIVELES(lblNivel.Caption - 1) <> CLng(Len(txt(0).Text)) Then
    MsgBox "La Cuenta a registrar no corresponde con el Nivel de Cuenta", vbInformation, Caption
    ValidarData = False
    txt(0).SetFocus
    Exit Function
  End If
  
  If CLng(lblNivel.Caption) > 1 Then
   SQL = "SELECT cuentacodigo FROM CT_CUENTA WHERE cuentacodigo='" & Left(txt(0).Text, VG_aNIVELES(lblNivel.Caption - 2)) & "'"
   If xdllgen.VerificaDatoExistente(VGcnx, SQL) <= 0 Then
     MsgBox "La Cuenta a registrar no tiene la Cuenta Superior Correspondiente ", vbInformation, Caption
     ValidarData = False
     txt(0).SetFocus
     Exit Function
   End If
  End If
  
  SQL = "SELECT cuentacodigo FROM CT_CUENTA WHERE cuentacodigo='" & txt(0).Text & "'"
  If modoinsert = True And xdllgen.VerificaDatoExistente(VGcnx, SQL) > 0 Then
    MsgBox "La Cuenta se encuentra registrada en la Base Datos, Debe registrar otra", vbInformation, Caption
    ValidarData = False
    txt(0).SetFocus
    Exit Function
  End If
  
  If Ctr_Ayuda4.xclave = Empty Then
     MsgBox "No existe Código de Tipo de Cuenta en el registro editado", vbInformation, Caption
     If lblNivel.Caption = 1 Then
        Ctr_Ayuda4.SetFocus
     Else
        txt(0).SetFocus
     End If
     ValidarData = False
     Exit Function
  End If
  
  If CLng(lblNivel.Caption) > 1 Then
   If chk(0).Value = 1 And Ctr_Ayuda1.xclave = Empty Then
     MsgBox "Falta indicar el Tipo de Analítico", vbInformation, Caption
     ValidarData = False
     Ctr_Ayuda1.SetFocus
     Exit Function
   End If
  End If
  
  If chk(3).Value = 1 And FLAGDISTRIBUCION = False Then
      MsgBox "No Existe Porcentaje de Distribución para esta cuenta, Deshabilitar el check", vbInformation, Caption
      ValidarData = False
      chk(3).SetFocus
      Exit Function
  End If
   
  ValidarData = True
End Function

Private Sub chk_Click(INDEX As Integer)
  Select Case INDEX
    Case 0
       Ctr_Ayuda1.Enabled = IIf(chk(0).Value = 0, False, True)
       If chk(0).Value = 0 Then Ctr_Ayuda1.xclave = Empty: Ctr_Ayuda1.xnombre = Empty
       
    Case 3
       If chk(3).Value = 1 Then
         If FLAGMOVIMIENTODISTRI = False Then
            cmdDistribucion.Visible = True
            SSTab1.TabEnabled(2) = True
            SSTab1.Tab = 2
            FLAGMOVIMIENTODISTRI = False
            Call LlenarPorcentajes
         End If
       Else
         cmdDistribucion.Visible = False
         SSTab1.TabEnabled(2) = False
       End If
  End Select
  
  If modoedit = True Then
     cAcepta.Enabled = True
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
  Set rsDist = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  SSTab1.TabEnabled(PreviousTab) = False
End Sub

Private Sub TDBGrid1_DblClick()
  If rs.RecordCount > 0 And (modoedit = False And modoinsert = False) Then Call Mantenimiento(1)
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
 Dim rsX As ADODB.Recordset
 Set rsX = rs.Clone(adLockReadOnly)
 rsX.Bookmark = Bookmark
 If rsX!cuentanivel = 1 Then
   RowStyle.BackColor = &H80000018
 End If

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
  Call ConfiguraTdbgrid
  TDBGrid1.Refresh
End Sub

Private Sub ConfiguraTdbgrid()
    With TDBGrid1
        .Columns(0).Width = 900
        .Columns(1).Width = 4100
        .Columns(2).Width = 800
        .Columns(3).Width = 800
        .Columns(4).Width = 800
        .Columns(5).Width = 800
        .Columns(6).Width = 800
    End With

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Call Mantenimiento(Button.INDEX - 1)
End Sub

Sub Mantenimiento(indice As Integer)
  Dim J As Integer
  Dim spos As Integer
  Dim SQL As String

  On Error GoTo X
  SSTab1.TabEnabled(1) = True

  Select Case indice
     Case 0   'nuevo
        SSTab1.Tab = 1
        SSTab1.TabEnabled(2) = False
        Call LimpiarValores
        Call ModoNormal
        Toolbar1.Visible = False
        TreeView1.Enabled = False
        modoinsert = True
        FLAGDISTRIBUCION = False

     Case 1   'modificar
        If TDBGrid1.Row < 0 Then
          Exit Sub
        End If
        Call EditarValores
        modoedit = True
        SSTab1.Tab = 1
        Toolbar1.Visible = False
        TreeView1.Enabled = False
        i_filaorigen = TDBGrid1.Row
        txt(0).Enabled = False
        cAcepta.Enabled = False
        FLAGDISTRIBUCION = False

     Case 2   'eliminar
       If MsgBox("Desea eliminar el registro?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          Dim rs As ADODB.Recordset
          Set rs = New ADODB.Recordset
          SQL = "Select isnull(count(*),0) from ct_cuenta where cuentacodigo like '" & Trim(TDBGrid1.Columns(0).Value) & "%'"
          Set rs = VGcnx.Execute(SQL)
          If rs(0) > 1 Then
             MsgBox "La Cuenta a Eliminar tiene Cuentas Dependientes al Nivel Inferior" & Chr(10) & Chr(13) & "Deben Eliminarse primero las Cuentas de Nivel Inferior", vbInformation, Caption
             Exit Sub
          End If
          SQL = "DELETE FROM CT_CUENTA WHERE cuentacodigo = '" & Trim(TDBGrid1.Columns(0).Value) & "'"
          VGcnx.Execute (SQL)
          Call MuestraDatos(Trim(TDBGrid1.Columns(0).Value))
       End If

     Case 3   'Imprimir
       With MDIPrincipal
          .cryRpt.Destination = crptToWindow
          .cryRpt.WindowState = crptMaximized
          .cryRpt.StoredProcParam(0) = VGParamSistem.BDEmpresa
          .cryRpt.StoredProcParam(1) = Trim(xCuenta) & "%"
          .cryRpt.Formulas(0) = "@Emp='" & VGParamConta.NomEmpresa & "'"
          .cryRpt.ReportFileName = App.Path & "\ReportesLaser\rptPlanCuentas.rpt"
          .cryRpt.Connect = vgCADENAREPORT
          .cryRpt.DiscardSavedData = True
          .cryRpt.Action = 1
       End With

     Case 4  ' salir
       Unload Me
  End Select
  Exit Sub

X:
  If indice = 2 And Err.Number = -2147217873 Then
    MsgBox "Registro no podrá Eliminarse mientras exista Información en las Tablas Relacionadas", vbInformation, Caption
  Else
    MsgBox "Error inesperado: " & Err.Description & "  " & Err.Number, vbInformation, Caption
  End If
End Sub

Private Sub txt_Change(INDEX As Integer)
  cAcepta.Enabled = IIf(xdllgen.ESNULO(lblNivel.Caption, 0) = VGnumniveles, ValidarCuentaUltimoNivel(), ValidarCuentaNivel())
End Sub

 Private Sub txt_KeyPress(INDEX As Integer, KeyAscii As Integer)
  If KeyAscii = 13 And INDEX = 6 And cAcepta.Value = True Then
    cAcepta.SetFocus
    Call cAcepta_Click
  End If
End Sub

Private Sub txt_LostFocus(INDEX As Integer)
 Dim i As Integer
    txt(INDEX).Text = UCase(txt(INDEX).Text)
    If modoinsert = True And INDEX = 0 Then
      For i = 1 To VGnumniveles
       If VG_aNIVELES(i - 1) = Len(Trim(txt(0).Text)) Then
         lblNivel.Caption = i
         Call ConfiguraModoEdicion
         cboTipoAjuste.ListIndex = 0
        
         If i = 1 And UBound(VG_aNIVELES) > 1 Then
           Ctr_Ayuda4.Enabled = True
         Else
           Ctr_Ayuda4.Enabled = False
           Ctr_Ayuda4.xclave = DevuelveTipoCuenta(): Ctr_Ayuda4.Ejecutar
         End If
         
         Exit For
       Else
         lblNivel.Caption = Empty
         Ctr_Ayuda4.xclave = Empty: Ctr_Ayuda4.Ejecutar
       End If
      Next
    End If
  
  If INDEX = 1 Then Call ConfiguraModoEdicion
  
End Sub

Private Sub Ctr_Ayuda4_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
   cAcepta.Enabled = IIf(xdllgen.ESNULO(lblNivel.Caption, 0) = VGnumniveles, ValidarCuentaUltimoNivel(), ValidarCuentaNivel())
End Sub

Private Sub cboTipoAjuste_Click()
  cAcepta.Enabled = IIf(xdllgen.ESNULO(lblNivel.Caption, 0) = VGnumniveles, ValidarCuentaUltimoNivel(), ValidarCuentaNivel())
End Sub

Function ValidarCuentaUltimoNivel() As Boolean
 Dim i As Integer
  For i = 0 To 1
    If txt(i).Text = Empty Then
     ValidarCuentaUltimoNivel = False
     Exit Function
    End If
  Next

  ValidarCuentaUltimoNivel = True
End Function

Function ValidarCuentaNivel() As Boolean
 Dim i As Integer
  For i = 0 To 6
    If txt(i).Text = Empty Then
     ValidarCuentaNivel = False
     Exit Function
    End If
  Next

  ValidarCuentaNivel = True
End Function

Sub ModoEdicion(flagULTIMONIVEL As Boolean) 'True: Ultimo Nivel  False:Otros Niveles
  Dim i As Integer
   For i = 0 To 1
    txt(i).Enabled = True
  Next
  For i = 2 To 6
    txt(i).BackColor = IIf(flagULTIMONIVEL = True, ColorHabilitado, ColorDesHabilitado)
    txt(i).Enabled = Not flagULTIMONIVEL
    txt(i).Text = 0
  Next
  For i = 0 To 3
    chk(i).Enabled = flagULTIMONIVEL
  Next
  cmdDistribucion.Visible = flagULTIMONIVEL
  Ctr_Ayuda1.Enabled = flagULTIMONIVEL
  cboTipoAjuste.Enabled = flagULTIMONIVEL
  
End Sub

Sub ModoNormal()
 Dim i As Integer
   For i = 0 To 1
     txt(i).Enabled = True
  Next
  For i = 2 To 6
     txt(i).Enabled = True
  Next
  For i = 0 To 4
     chk(i).Enabled = True
  Next
  Ctr_Ayuda1.Enabled = True
  Ctr_Ayuda4.Enabled = True
  cboTipoAjuste.Enabled = True
  
End Sub

Function GrabarPlanCuenta(tipooperacion As Integer, xCod As String, xDes As String, xLinAct As Double, xLinPas As Double, xLinEgp As Double, xNatu As String, xNatuegp As String, xCtaCostos As Integer, xCtaAnalitico As Integer, xCtaDoc As Integer, xCtaNivel As Integer, xTipoAnalitico As String, xCtaDist As Integer, xTipoCuenta As String, xTipoAJuste As String) As String
 Dim strSQL As String
 
 xTipoAnalitico = IIf(IsNull(xTipoAnalitico) Or xTipoAnalitico = Empty, "00", xTipoAnalitico)
 Select Case tipooperacion
   Case 0
     strSQL = "INSERT INTO ct_cuenta (cuentacodigo, cuentadescripcion, cuentalineaactivo,cuentalineapasivo, cuentalineaegp, cuentanatu, cuentanategp,"
     strSQL = strSQL & "cuentaestadoccostos, cuentaestadoanalitico,cuentadocumento, cuentanivel, tipoanaliticocodigo,cuentaestadodistribucion,tipocuentacodigo,usuariocodigo, fechaact, CUENTAGRUPO,tipoajuste,cuentaestado) "
     strSQL = strSQL & "VALUES ('" & xCod & "','" & xDes & "'," & xLinAct & "," & xLinPas & "," & xLinEgp & ",'" & xNatu & "','" & xNatuegp & "'," & xCtaCostos & "," & xCtaAnalitico & "," & xCtaDoc & "," & xCtaNivel & ",'" & xTipoAnalitico & "'," & xCtaDist & ",'" & xTipoCuenta & "','" & VGusuario & "','" & Format(Now, "dd/mm/yyyy") & "','01','" & xTipoAJuste & "','1')"
   
   Case 1
     strSQL = "UPDATE CT_CUENTA SET "
     strSQL = strSQL & "cuentadescripcion='" & xDes & "',"
     strSQL = strSQL & "cuentalineaactivo=" & xdllgen.ESNULO(xLinAct, 0) & ","
     strSQL = strSQL & "cuentalineapasivo=" & xdllgen.ESNULO(xLinPas, 0) & ","
     strSQL = strSQL & "cuentalineaegp=" & xdllgen.ESNULO(xLinEgp, 0) & ","
     strSQL = strSQL & "cuentanatu='" & xdllgen.ESNULO(xNatu, "") & "',"
     strSQL = strSQL & "cuentanategp='" & xdllgen.ESNULO(xNatuegp, "") & "',"
     strSQL = strSQL & "cuentaestadoccostos=" & xdllgen.ESNULO(xCtaCostos, 0) & ","
     strSQL = strSQL & "cuentaestadoanalitico=" & xdllgen.ESNULO(xCtaAnalitico, 0) & ","
     strSQL = strSQL & "cuentadocumento=" & xdllgen.ESNULO(xCtaDoc, 0) & ","
     strSQL = strSQL & "cuentanivel=" & xdllgen.ESNULO(xCtaNivel, 0) & ","
     strSQL = strSQL & "tipoanaliticocodigo='" & xTipoAnalitico & "',"
     strSQL = strSQL & "cuentaestadodistribucion=" & xdllgen.ESNULO(xCtaDist, 0) & ","
     strSQL = strSQL & "tipocuentacodigo='" & xTipoCuenta & "',"
     strSQL = strSQL & "usuariocodigo='" & VGusuario & "',"
     strSQL = strSQL & "fechaact='" & Format(Now, "dd/mm/yyyy") & "',"
     strSQL = strSQL & "cuentagrupo='01',"
     strSQL = strSQL & "tipoajuste='" & xTipoAJuste & "' "
     strSQL = strSQL & "WHERE cuentacodigo='" & xCod & "'"

 End Select
 GrabarPlanCuenta = strSQL

End Function

Function GrabarCuentaDistribucion()
 Dim i As Long
 Dim SQL As String
 
 If rsDist.RecordCount > 0 Then
   SQL = "DELETE FROM ct_distribucion WHERE cuentacodigo='" & txt(0).Text & "'"
   VGcnx.Execute (SQL)
   rsDist.MoveFirst
   For i = 0 To rsDist.RecordCount - 1
     SQL = "INSERT ct_distribucion (cuentacodigo,distribucioncargo,distribucionabono,distribucionporcen,usuariocodigo,fechaact) VALUES "
     SQL = SQL & "('" & rsDist(0) & "','" & rsDist(1) & "','" & rsDist(2) & "'," & rsDist(3) & ",'" & VGusuario & "','" & Date & "')"
     VGcnx.Execute (SQL)
     rsDist.MoveNext
   Next
 End If

End Function

Function GrabarTablaSaldos()
 Dim SQL As String
 Dim NombreTabla As String
    NombreTabla = "CT_SALDOS" & VGParamSistem.Anoproceso
    SQL = "INSERT " & NombreTabla & "(cuentacodigo,usuariocodigo,fechaact)"
    SQL = SQL & "VALUES ('" & txt(0).Text & "','" & VGusuario & "','" & Date & "')"
    VGcnx.Execute (SQL)

End Function

Private Sub cmdSigue_Click()
 If rsDist.RecordCount > 0 And CDbl(lblPorcen.Caption) <> 100 Then
   MsgBox "El Total de % Distribución no se ha completado al 100%", vbInformation, Caption
 Else
   cAcepta.Enabled = IIf(rsDist.RecordCount > 0 And FLAGMOVIMIENTODISTRI = True, True, False)
   chk(3).Value = IIf(rsDist.RecordCount > 0, 1, 0)
   SSTab1.Tab = 1
   SSTab1.TabEnabled(1) = True
 End If

End Sub

Private Sub cmdDistribucion_Click()
  SSTab1.TabEnabled(2) = True
  SSTab1.Tab = 2
  FLAGMOVIMIENTODISTRI = False
  Call LlenarPorcentajes
End Sub

Sub LlenarPorcentajes()
  Set rsDist = Nothing
  Call IniciaGridDist
  Call CargaGridDist
  Call ConfigGridDist
  FLAGDISTRIBUCION = True
End Sub

Private Function IniciaGridDist()
  Call rsDist.Fields.Append("Cuenta", adVarChar, 20)
  Call rsDist.Fields.Append("Cuenta Cargo", adVarChar, 20)
  Call rsDist.Fields.Append("Cuenta Abono", adVarChar, 20)
  Call rsDist.Fields.Append("Porcentaje", adDouble)
  Call rsDist.Fields.Append("Item", adInteger)
  rsDist.Open
End Function

Private Sub CargaGridDist()
  Dim xRs As ADODB.Recordset
  Dim nConta As Long
  Dim SQL As String
  Set xRs = New ADODB.Recordset
  SQL = "SELECT cuentacodigo,distribucioncargo,distribucionabono,distribucionporcen "
  SQL = SQL & "FROM ct_distribucion WHERE cuentacodigo='" & txt(0).Text & "'"
  Set xRs = VGcnx.Execute(SQL)
  nConta = 1
  If xRs.RecordCount > 0 Then
     xRs.MoveFirst
     Do While Not xRs.EOF
       rsDist.AddNew
       rsDist.Fields(0) = xRs(0).Value
       rsDist.Fields(1) = xRs(1).Value
       rsDist.Fields(2) = xRs(2).Value
       rsDist.Fields(3) = xRs(3).Value
       rsDist.Fields(4) = nConta
       nConta = nConta + 1
       rsDist.Update
       xRs.MoveNext
     Loop
  End If
  Set xRs = Nothing
End Sub

Private Function ConfigGridDist()
 Dim i As Integer
  Set TDBGrid2.DataSource = rsDist
  With TDBGrid2
    For i = 0 To 4
      .Columns(i).AllowSizing = False
    Next
    .Columns(0).Visible = False
    .Columns(0).Caption = "Cuenta"
    .Columns(1).Width = 1700
    .Columns(1).Caption = "Cuenta Cargo"
    .Columns(2).Width = 1700
    .Columns(2).Caption = "Cuenta Abono"
    .Columns(3).Width = 1900
    .Columns(3).Caption = "Distribución(%)"
    .Columns(4).Width = 800
    .Columns(4).Caption = "Item"
  End With
  lblPorcen.Caption = DevuelveTotPor()
  TDBGrid2.Refresh
End Function

Private Function ActualizaGridDist()
 Dim nReg As Long
 nReg = DevuelveNumReg() + 1
 With rsDist
   .AddNew
   .Fields(0) = txt(0).Text
   .Fields(1) = Ctr_Ayuda2.xclave
   .Fields(2) = Ctr_Ayuda3.xclave
   .Fields(3) = CDbl(txtPorcen.Text)
   .Fields(4) = CLng(nReg)
   .Update
 End With
End Function

Private Sub txtbuscarcuenta_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   Call cmdBuscar_Click
 End If
End Sub


Private Sub txtPorcen_Change()
 If Ctr_Ayuda2.xclave <> Empty And Ctr_Ayuda3.xclave <> Empty Then
   If txtPorcen.Text <> Empty Then
     CmdAceptar.Enabled = True
     Exit Sub
   End If
 End If
 CmdAceptar.Enabled = False
End Sub

Function ValidarGridDist() As Boolean
 
 If DevuelveTotPor() + Val(txtPorcen.Text) > 100 Then
   MsgBox "La Sumatoria de Porcentaje excede al 100%", vbInformation, Caption
   txtPorcen.SetFocus
   SendKeys "{HOME}+{END}"
   ValidarGridDist = False
   Exit Function
 End If
  
 If Val(txtPorcen.Text) = 0 Then
   MsgBox "Valor de Porcentaje  no permitido", vbInformation, Caption
   txtPorcen.SetFocus
   SendKeys "{HOME}+{END}"
   ValidarGridDist = False
   Exit Function
 End If
  
 ValidarGridDist = True
End Function

Function DevuelveTotPor() As Double
  Dim i As Integer
  Dim nSum As Double
  nSum = 0
  If rsDist.RecordCount > 0 Then
    rsDist.MoveFirst
    Do While Not rsDist.EOF
      nSum = nSum + rsDist.Fields(3).Value
      rsDist.MoveNext
    Loop
  End If
  DevuelveTotPor = Format(nSum, "###0.#0")
End Function

Private Sub TDBGrid2_Click()
  TDBGrid2.SetFocus
End Sub

Private Sub TDBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim nvalor As String
  If rsDist.RecordCount > 0 Then
    If KeyCode = 46 Then
      nvalor = TDBGrid2.Columns(4).Text
      If rsDist.RecordCount > 0 Then
         rsDist.MoveFirst
         Do Until rsDist.EOF
           If rsDist.Fields(4) = nvalor Then
             rsDist.Delete adAffectCurrent
             rsDist.Update
             Exit Do
           End If
           rsDist.MoveNext
         Loop
         ConfigGridDist
         lblPorcen.Caption = DevuelveTotPor
      End If
    End If
  End If
End Sub

Function DevuelveNumReg() As Long
 Dim nUlt As Long
 If rsDist.RecordCount > 0 Then
   rsDist.MoveFirst
   Do While Not rsDist.EOF
     nUlt = rsDist.Fields(4)
     rsDist.MoveNext
   Loop
   DevuelveNumReg = nUlt
 Else
   DevuelveNumReg = 0
 End If
End Function

Private Sub Arbol(xCta As String)
On Error GoTo xx
  Dim rs As ADODB.Recordset
  Dim SQL As String
  Dim CodCan As String
  Dim i As Long
  Dim K As Integer
  Dim nodX As NODE
  
  Set VGvardllgen = New dllgeneral.dll_general
  xCta = VGvardllgen.ESNULO(xCta, "%")
  
  SQL = "SELECT A.cuentacodigo,A.cuentadescripcion, A.cuentanivel FROM CT_CUENTA A "
  SQL = SQL & "WHERE A.cuentacodigo<>'00' and A.cuentacodigo like '" & xCta & "%' "
  SQL = SQL & "ORDER BY 1"
  Set rs = New ADODB.Recordset
  Set rs = VGcnx.Execute(SQL)
  
  If (rs.EOF Or rs.BOF) Then
    Exit Sub
  End If
  TreeView1.Nodes.Clear
  Set nodX = TreeView1.Nodes.Add(, , "P", "Plan de Cuentas", 1)
  nodX.EnsureVisible
  
  For i = 1 To rs.RecordCount
    cCta = "P" & Trim(rs(0))
    If rs(2) > 1 Then
      Set nodX = TreeView1.Nodes.Add("P" & Trim(Mid(cCta, 2, VG_aNIVELES(rs(2) - 2))), tvwChild, cCta, rs(0), 1)
    Else
      Set nodX = TreeView1.Nodes.Add("P", tvwChild, cCta, rs(0), 1)
      nodX.EnsureVisible
    End If
    rs.MoveNext
   Next i
   rs.MoveFirst
   Exit Sub
xx:
  'MsgBox "Error de Inconsistencia en Base de Datos: " & "No existe Plan de Cuentas superior para la Cuenta " & rs(0) & " - " & rs(1), vbInformation, Caption
  l_error = l_error & "No existe Cuenta superior para la Cuenta " & rs(0) & " - " & rs(1) & Chr(13)
  Resume Next
End Sub

Private Sub TREEVIEW1_EXPAND(ByVal NODE As MSComctlLib.NODE)
  NODE.Image = 2
End Sub

Private Sub TREEVIEW1_COLLAPSE(ByVal NODE As MSComctlLib.NODE)
  NODE.Image = 1
End Sub

Private Sub TreeView1_NodeClick(ByVal NODE As MSComctlLib.NODE)
  If NODE.Key <> Empty Then
    xCuenta = Right(Trim(NODE.Key), Len(Trim(NODE.Key)) - 1)
    MuestraDatos (Right(Trim(NODE.Key), Len(Trim(NODE.Key)) - 1))
  End If
End Sub

Private Sub cmdAceptar_Click()
 If ValidarGridDist = True Then
   Call ActualizaGridDist
   Call ConfigGridDist
   Ctr_Ayuda2.xclave = Empty
   Ctr_Ayuda3.xclave = Empty
   Ctr_Ayuda2.xnombre = Empty
   Ctr_Ayuda3.xnombre = Empty
   txtPorcen.Text = Empty
   FLAGMOVIMIENTODISTRI = True
 End If
End Sub

Function DevuelveTipoCuenta()
 Dim rsX As New ADODB.Recordset
 Dim SQL As String
 
 Set rsX = New ADODB.Recordset
 SQL = "Select tipocuentacodigo from ct_cuenta where cuentacodigo=left('" & Trim(txt(0).Text) & "',2)"
 Set rsX = VGcnx.Execute(SQL)
 If rsX.RecordCount > 0 Then
   DevuelveTipoCuenta = rsX(0)
 Else
   DevuelveTipoCuenta = 0
 End If
 Set rsX = Nothing

End Function

Sub GrabaTipoCuenta(xCta As String, xValor As String, xNivel As Integer)
 Dim SQL As String
    SQL = "UPDATE ct_cuenta SET tipocuentacodigo='" & xValor & "' "
    SQL = SQL & "WHERE left(cuentacodigo,2)='" & xCta & "' AND cuentanivel>" & xNivel
    VGcnx.Execute (SQL)
End Sub

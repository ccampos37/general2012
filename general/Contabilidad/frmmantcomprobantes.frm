VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmantcomprobantes 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Comprobantes Contables"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmmantcomprobantes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   11475
   Begin TabDlg.SSTab SSTabMant 
      Height          =   8985
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   15849
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "frmmantcomprobantes.frx":1272
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameConsulta"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "frmmantcomprobantes.frx":128E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shilu2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "SSTab2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "StBar"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "frameGrid"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "FrameCabecera"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.Frame FrameCabecera 
         BackColor       =   &H00C0FFFF&
         Height          =   2100
         Left            =   105
         TabIndex        =   24
         Top             =   315
         Width           =   11175
         Begin VB.CheckBox ChkGrabado 
            Alignment       =   1  'Right Justify
            Caption         =   "Oper. Grabada"
            ForeColor       =   &H000000C0&
            Height          =   225
            Left            =   150
            TabIndex        =   22
            Top             =   1740
            Width           =   1710
         End
         Begin MSComCtl2.DTPicker DTPFechaComprobCab 
            Height          =   300
            Left            =   1665
            TabIndex        =   18
            Top             =   540
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   529
            _Version        =   393216
            Format          =   16318465
            CurrentDate     =   37469
         End
         Begin TextFer.TxFer TxGlosaComprobCab 
            Height          =   300
            Left            =   1650
            TabIndex        =   20
            Top             =   855
            Width           =   4605
            _ExtentX        =   8123
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
         End
         Begin TextFer.TxFer TxObsComprobCab 
            Height          =   540
            Left            =   1650
            TabIndex        =   21
            Top             =   1170
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   953
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
            MaxLength       =   150
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
         End
         Begin TextFer.TxFer TxCtrNref 
            Height          =   300
            Left            =   4245
            TabIndex        =   23
            Top             =   1725
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   529
            BackColor       =   13056
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   65280
            MaxLength       =   16
            Text            =   "00000"
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   "00000"
            ColorTextoAlEnfocar=   16711680
         End
         Begin VB.Label lbNumComprobCablibro 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2FDFE&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000001"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   4770
            TabIndex        =   93
            Top             =   240
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label leNComprob 
            AutoSize        =   -1  'True
            Caption         =   "Nº. Comp. Libro :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   3345
            TabIndex        =   92
            Top             =   285
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label lbnemoref 
            AutoSize        =   -1  'True
            Caption         =   "Nemotecnico "
            Height          =   195
            Left            =   1995
            TabIndex        =   90
            Top             =   1755
            Visible         =   0   'False
            Width           =   2070
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00808080&
            Height          =   1785
            Left            =   6480
            Shape           =   4  'Rounded Rectangle
            Top             =   210
            Width           =   4560
         End
         Begin VB.Label leDebe 
            AutoSize        =   -1  'True
            Caption         =   "DEBE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   6660
            TabIndex        =   69
            Top             =   750
            Width           =   510
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FFFFFF&
            Height          =   1800
            Left            =   6495
            Shape           =   4  'Rounded Rectangle
            Top             =   180
            Width           =   4560
         End
         Begin VB.Label leHaber 
            AutoSize        =   -1  'True
            Caption         =   "HABER"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   6660
            TabIndex        =   68
            Top             =   1050
            Width           =   645
         End
         Begin VB.Label leSoles 
            AutoSize        =   -1  'True
            Caption         =   "SOLES S/."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   7770
            TabIndex        =   67
            Top             =   360
            Width           =   945
         End
         Begin VB.Label LeDolares 
            AutoSize        =   -1  'True
            Caption         =   "DOLARES US$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   9450
            TabIndex        =   66
            Top             =   360
            Width           =   1305
         End
         Begin VB.Label LbTotales 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FBE3D9&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   7530
            TabIndex        =   65
            Top             =   705
            Width           =   1635
         End
         Begin VB.Label LbTotales 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FBE3D9&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   7530
            TabIndex        =   64
            Top             =   1005
            Width           =   1635
         End
         Begin VB.Label LbTotales 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBE0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   9315
            TabIndex        =   63
            Top             =   690
            Width           =   1635
         End
         Begin VB.Label LbTotales 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBE0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   9300
            TabIndex        =   62
            Top             =   1005
            Width           =   1635
         End
         Begin VB.Label leDifer 
            AutoSize        =   -1  'True
            Caption         =   "DIFER."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   6675
            TabIndex        =   61
            Top             =   1560
            Width           =   630
         End
         Begin VB.Label LbTotales 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FEFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   7530
            TabIndex        =   60
            Top             =   1515
            Width           =   1635
         End
         Begin VB.Label LbTotales 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00F2FEFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   5
            Left            =   9315
            TabIndex        =   59
            Top             =   1515
            Width           =   1635
         End
         Begin VB.Label leNComprob 
            AutoSize        =   -1  'True
            Caption         =   "Nº. Comprobante :"
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   52
            Top             =   225
            Width           =   1305
         End
         Begin VB.Label lbNumComprobCab 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2FDFE&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0000010000"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1665
            TabIndex        =   51
            Top             =   225
            Width           =   1515
         End
         Begin VB.Label leFechaComprob 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fecha Comprob. :"
            Height          =   195
            Index           =   7
            Left            =   165
            TabIndex        =   50
            Top             =   555
            Width           =   1260
         End
         Begin VB.Label leGlosa 
            Caption         =   "Glosa Comprob. :"
            Height          =   345
            Left            =   165
            TabIndex        =   49
            Top             =   900
            Width           =   1425
         End
         Begin VB.Label leObservaciones 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones :"
            Height          =   195
            Left            =   165
            TabIndex        =   48
            Top             =   1305
            Width           =   1155
         End
      End
      Begin VB.Frame frameGrid 
         BackColor       =   &H00FFFFC0&
         Height          =   2925
         Left            =   120
         TabIndex        =   81
         Top             =   2400
         Width           =   11175
         Begin TrueOleDBGrid70.TDBGrid TDBG_Det 
            Height          =   2280
            Left            =   75
            TabIndex        =   45
            Top             =   180
            Width           =   11040
            _ExtentX        =   19473
            _ExtentY        =   4022
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
            MultipleLines   =   0
            CellTips        =   2
            CellTipsWidth   =   0
            MultiSelect     =   2
            DataView        =   1
            AnimateWindow   =   2
            DeadAreaBackColor=   16777215
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
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
            _StyleDefs(24)  =   "Splits(0).Style:id=47,.parent=1"
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
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=154,.parent=47,.alignment=1,.bgcolor=&H80000018&"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=151,.parent=48,.alignment=2"
            _StyleDefs(70)  =   ":id=151,.bgcolor=&H8000000F&"
            _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=152,.parent=49"
            _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=153,.parent=51,.bgcolor=&H80000018&"
            _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=158,.parent=47,.alignment=1,.bgcolor=&H80000018&"
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
         Begin VB.Shape Shape10 
            BackColor       =   &H8000000B&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   90
            Left            =   0
            Top             =   0
            Width           =   11265
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Registros :"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   8940
            TabIndex        =   83
            Top             =   2625
            Width           =   975
         End
         Begin VB.Label lbnregdetalle 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "0 "
            Height          =   255
            Left            =   10035
            TabIndex        =   82
            Top             =   2595
            Width           =   1050
         End
         Begin VB.Shape Shape9 
            BorderColor     =   &H00404040&
            Height          =   285
            Left            =   10020
            Top             =   2580
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000B&
         Height          =   1785
         Left            =   -74910
         TabIndex        =   54
         Top             =   390
         Width           =   11250
         Begin MSComCtl2.DTPicker DTPFechaContab 
            Height          =   300
            Left            =   9660
            TabIndex        =   87
            Top             =   510
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            Format          =   16318465
            CurrentDate     =   37489
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_SubAsiento 
            Height          =   315
            Left            =   5025
            TabIndex        =   56
            Top             =   1350
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   556
            XcodMaxLongitud =   4
            xcodwith        =   450
            NomTabla        =   "ct_subasiento"
            TituloAyuda     =   "Busqueda de  SubAsiento"
            ListaCampos     =   "subasientocodigo(1),subasientodescripcion(1),monedacodigo(1),subasientoglosa(1),subasientorepitedoc(2)"
            XcodCampo       =   "subasientocodigo"
            XListCampo      =   "subasientodescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion,Moneda"
            ListaCamposText =   "subasientocodigo,subasientodescripcion,monedacodigo,subasientoglosa,subasientorepitedoc"
            Requerido       =   0   'False
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Asiento 
            Height          =   300
            Left            =   195
            TabIndex        =   55
            Top             =   1350
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   529
            XcodMaxLongitud =   3
            xcodwith        =   150
            NomTabla        =   "ct_asiento"
            TituloAyuda     =   "Busqueda de Asiento"
            ListaCampos     =   "asientocodigo(1), asientodescripcion(1),flaggrabado(2),controlnref(2),nemotecref(1),librocodigo(1)"
            XcodCampo       =   "asientocodigo"
            XListCampo      =   "asientodescripcion"
            ListaCamposDescrip=   "Codigo,Descripción,OperGraba"
            ListaCamposText =   "asientocodigo,asientodescripcion,flaggrabado,controlnref,nemotecref,librocodigo"
            Requerido       =   0   'False
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Contabilizacion :"
            Height          =   225
            Left            =   7530
            TabIndex        =   86
            Top             =   585
            Width           =   1950
         End
         Begin VB.Image Image1 
            Height          =   465
            Left            =   135
            Picture         =   "frmmantcomprobantes.frx":12AA
            Stretch         =   -1  'True
            Top             =   210
            Width           =   450
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   " Consulta e Ingreso de Comprobantes Contables"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   690
            TabIndex        =   80
            Top             =   495
            Width           =   4935
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H00808080&
            Height          =   15
            Left            =   60
            Top             =   900
            Width           =   11130
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   45
            TabIndex        =   79
            Top             =   135
            Width           =   11145
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Asiento :"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   195
            TabIndex        =   58
            Top             =   1140
            Width           =   1935
         End
         Begin VB.Label lbSubAsiento 
            BackStyle       =   0  'Transparent
            Caption         =   "Subasiento :"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5025
            TabIndex        =   57
            Top             =   1140
            Width           =   1590
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H00FFFFFF&
            Height          =   15
            Left            =   60
            Top             =   915
            Width           =   11130
         End
      End
      Begin MSComctlLib.StatusBar StBar 
         Height          =   285
         Left            =   90
         TabIndex        =   53
         Top             =   8580
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   4
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Object.Width           =   5371
               MinWidth        =   5363
               Text            =   "Asiento : "
               TextSave        =   "Asiento : "
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   7832
               MinWidth        =   7832
               Text            =   "Sub Asiento :"
               TextSave        =   "Sub Asiento :"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               TextSave        =   "27/10/2012"
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               AutoSize        =   1
               Object.Width           =   3916
               Picture         =   "frmmantcomprobantes.frx":251C
               Text            =   "Estado :"
               TextSave        =   "Estado :"
            EndProperty
         EndProperty
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3195
         Left            =   120
         TabIndex        =   1
         Top             =   5355
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   5636
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         BackColor       =   12648447
         MouseIcon       =   "frmmantcomprobantes.frx":379E
         TabCaption(0)   =   "&Ingreso del detalle"
         TabPicture(0)   =   "frmmantcomprobantes.frx":37BA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Shilu1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "FramDetalle"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         Begin VB.Frame FramDetalle 
            BackColor       =   &H00C0FFFF&
            Height          =   2730
            Left            =   75
            TabIndex        =   2
            Top             =   345
            Width           =   11085
            Begin VB.CommandButton CmdDocPend 
               Caption         =   "..."
               Height          =   300
               Left            =   6255
               TabIndex        =   91
               Top             =   1410
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.CheckBox ChkInafecto 
               Alignment       =   1  'Right Justify
               Caption         =   "Inafecto"
               Height          =   255
               Left            =   6660
               TabIndex        =   39
               Top             =   2070
               Width           =   900
            End
            Begin TextFer.TxFer TxValor 
               Height          =   330
               Left            =   8775
               TabIndex        =   47
               Top             =   2280
               Visible         =   0   'False
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   582
               Alignment       =   1
               BackColor       =   -2147483624
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
               Text            =   "0.00"
               ColorIlumina    =   -2147483624
               SaltarAlEnter   =   -1  'True
               Valor           =   "0.00"
               TipoDato        =   1
               SignodeMiles    =   -1  'True
               NumeroDecimales =   2
               SignoNegativo   =   0   'False
               Formato         =   "###,###,###.00"
               MarcarTextoAlEnfoque=   -1  'True
            End
            Begin VB.CheckBox ChkAjusta 
               Alignment       =   1  'Right Justify
               Caption         =   "Ajustar por el Usuario"
               Height          =   210
               Left            =   8760
               TabIndex        =   46
               Top             =   2055
               Width           =   2205
            End
            Begin TextFer.TxFer TxMonto 
               Height          =   315
               Left            =   8760
               TabIndex        =   44
               Top             =   1560
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   556
               Alignment       =   1
               BackColor       =   16384
               Object.CausesValidation=   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   12648447
               MaxLength       =   15
               Text            =   "0.00"
               ColorIlumina    =   12648447
               SaltarAlEnter   =   -1  'True
               Valor           =   "0.00"
               TipoDato        =   1
               NumeroDecimales =   3
               SignoNegativo   =   0   'False
               Formato         =   "###,###,###,###.##"
               MarcarTextoAlEnfoque=   -1  'True
               ColorTextoAlEnfocar=   16711680
            End
            Begin TextFer.TxFer TxNdoc 
               Height          =   300
               Left            =   4515
               TabIndex        =   33
               Top             =   1410
               Width           =   1710
               _ExtentX        =   3016
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
               MaxLength       =   10
               Text            =   ""
               ColorIlumina    =   -2147483624
               SaltarAlEnter   =   -1  'True
               Valor           =   ""
            End
            Begin TextFer.TxFer TxSerie 
               Height          =   300
               Left            =   3900
               TabIndex        =   32
               Top             =   1410
               Width           =   555
               _ExtentX        =   979
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
            End
            Begin TextFer.TxFer txRuc 
               Height          =   315
               Left            =   5820
               TabIndex        =   30
               Top             =   1095
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   556
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
               MaxLength       =   11
               Text            =   ""
               ColorIlumina    =   -2147483624
               SaltarAlEnter   =   -1  'True
               Valor           =   ""
               NoCaracteres    =   "0123456789"
               MarcarTextoAlEnfoque=   -1  'True
               NoRangoCadena   =   -1  'True
            End
            Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Moneda 
               Height          =   315
               Left            =   8775
               TabIndex        =   42
               Top             =   510
               Width           =   2190
               _ExtentX        =   3863
               _ExtentY        =   556
               Enabled         =   0   'False
               XcodMaxLongitud =   2
               NomTabla        =   "gr_moneda"
               TituloAyuda     =   "Busqueda de Moneda"
               ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
               XcodCampo       =   "monedacodigo"
               XListCampo      =   "monedadescripcion"
               ListaCamposDescrip=   "Código,Descripción"
               ListaCamposText =   "monedacodigo,monedadescripcion"
            End
            Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_TipDoc 
               Height          =   315
               Left            =   1125
               TabIndex        =   31
               Top             =   1395
               Width           =   2070
               _ExtentX        =   3651
               _ExtentY        =   556
               XcodMaxLongitud =   2
               NomTabla        =   "gr_documento"
               TituloAyuda     =   "Busqueda de Tipo de  Documento"
               ListaCampos     =   "documentocodigo(1),documentodescripcion(1),documentonotacredito(2)"
               XcodCampo       =   "documentocodigo"
               XListCampo      =   "documentodescripcion"
               ListaCamposDescrip=   "Código,Descripción"
               ListaCamposText =   "documentocodigo,documentodescripcion,documentonotacredito"
            End
            Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Analitico 
               Height          =   300
               Left            =   1125
               TabIndex        =   29
               Top             =   1095
               Width           =   4065
               _ExtentX        =   7170
               _ExtentY        =   529
               XcodMaxLongitud =   15
               xcodwith        =   1100
               NomTabla        =   "v_analiticoentidad"
               TituloAyuda     =   "Busqueda de Analitico"
               ListaCampos     =   "analiticocodigo(1),entidadrazonsocial(1),entidadruc(1)"
               XcodCampo       =   "analiticocodigo"
               XListCampo      =   "entidadrazonsocial"
               ListaCamposDescrip=   "Codigo,Descripcion,Ruc"
               ListaCamposText =   "analiticocodigo,entidadrazonsocial,entidadruc"
            End
            Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_TipAnal 
               Height          =   315
               Left            =   5175
               TabIndex        =   28
               Top             =   795
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   556
               XcodMaxLongitud =   3
               xcodwith        =   100
               NomTabla        =   "ct_tipoanalitico"
               TituloAyuda     =   "Busqueda de Tipo de Analitico"
               ListaCampos     =   "tipoanaliticocodigo(1),tipoanaliticodescripcion(1)"
               XcodCampo       =   "tipoanaliticocodigo"
               XListCampo      =   "tipoanaliticodescripcion"
               ListaCamposDescrip=   "Código,Descripción"
               ListaCamposText =   "tipoanaliticocodigo,tipoanaliticodescripcion"
            End
            Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_CCosto 
               Height          =   330
               Left            =   1125
               TabIndex        =   27
               Top             =   795
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   582
               XcodMaxLongitud =   6
               xcodwith        =   300
               NomTabla        =   "ct_centrocosto"
               ListaCampos     =   "centrocostocodigo(1),centrocostodescripcion(1)"
               XcodCampo       =   "centrocostocodigo"
               XListCampo      =   "centrocostodescripcion"
               ListaCamposDescrip=   "Codigo,Descripcion"
               ListaCamposText =   "centrocostocodigo,centrocostodescripcion"
            End
            Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Cuenta 
               Height          =   315
               Left            =   1125
               TabIndex        =   26
               Top             =   495
               Width           =   6450
               _ExtentX        =   11377
               _ExtentY        =   556
               XcodMaxLongitud =   20
               xcodwith        =   1000
               NomTabla        =   "ct_cuenta"
               TituloAyuda     =   "Busqueda de Cuenta"
               ListaCampos     =   $"frmmantcomprobantes.frx":37D6
               XcodCampo       =   "cuentacodigo"
               XListCampo      =   "cuentadescripcion"
               ListaCamposDescrip=   "Código,Descripción"
               ListaCamposText =   "cuentacodigo,cuentadescripcion,cuentaestadoccostos,cuentaestadoanalitico,cuentadocumento,tipoanaliticocodigo,tipoajuste"
            End
            Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Opera 
               Height          =   315
               Left            =   1125
               TabIndex        =   25
               Top             =   180
               Width           =   6465
               _ExtentX        =   11404
               _ExtentY        =   556
               XcodMaxLongitud =   2
               NomTabla        =   "ct_operacion"
               TituloAyuda     =   "Busqueda de Operacion"
               ListaCampos     =   "operacioncodigo(1),operaciondescripcion(1),operaciondocumentoanulado(1),facturacionanticipada(1)"
               XcodCampo       =   "operacioncodigo"
               XListCampo      =   "operaciondescripcion"
               ListaCamposDescrip=   "Código,Descripción,Doc.Anulado,Factura anticipada"
               ListaCamposText =   "operacioncodigo,operaciondescripcion,operaciondocumentoanulado,facturacionanticipada"
            End
            Begin VB.ComboBox CmbID 
               Height          =   315
               ItemData        =   "frmmantcomprobantes.frx":3866
               Left            =   8760
               List            =   "frmmantcomprobantes.frx":3870
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   195
               Width           =   2220
            End
            Begin VB.ComboBox CmbTcambio 
               Height          =   315
               ItemData        =   "frmmantcomprobantes.frx":3889
               Left            =   8775
               List            =   "frmmantcomprobantes.frx":3896
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   855
               Width           =   2220
            End
            Begin MSComCtl2.DTPicker Dtp_FechaDoc 
               Height          =   315
               Left            =   1140
               TabIndex        =   37
               Top             =   2025
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               Format          =   16318465
               CurrentDate     =   37469
            End
            Begin MSComCtl2.DTPicker DtpFech_Ven 
               Height          =   315
               Left            =   3795
               TabIndex        =   38
               Top             =   2040
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   556
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   16318465
               CurrentDate     =   37469
            End
            Begin TextFer.TxFer TxGlosa 
               Height          =   300
               Left            =   1125
               TabIndex        =   40
               Top             =   2355
               Width           =   6435
               _ExtentX        =   11351
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
               MaxLength       =   50
               Text            =   ""
               ColorIlumina    =   -2147483624
               SaltarAlEnter   =   -1  'True
               Valor           =   ""
            End
            Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_TipRef 
               Height          =   315
               Left            =   1125
               TabIndex        =   34
               Top             =   1695
               Width           =   2070
               _ExtentX        =   3651
               _ExtentY        =   556
               XcodMaxLongitud =   2
               NomTabla        =   "gr_documento"
               TituloAyuda     =   "Busqueda de Tipo de  Documento"
               ListaCampos     =   "documentocodigo(1),documentodescripcion(1)"
               XcodCampo       =   "documentocodigo"
               XListCampo      =   "documentodescripcion"
               ListaCamposDescrip=   "Código,Descripción"
               ListaCamposText =   "documentocodigo,documentodescripcion"
               Requerido       =   0   'False
            End
            Begin TextFer.TxFer TxNref 
               Height          =   300
               Left            =   3885
               TabIndex        =   35
               Top             =   1695
               Width           =   2175
               _ExtentX        =   3836
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
               MaxLength       =   20
               Text            =   ""
               ColorIlumina    =   -2147483624
               SaltarAlEnter   =   -1  'True
               Valor           =   ""
            End
            Begin MSComCtl2.DTPicker Dtp_FechaDocRef 
               Height          =   285
               Left            =   6075
               TabIndex        =   36
               Top             =   1725
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   503
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   16318465
               CurrentDate     =   37469
            End
            Begin VB.Label lbnref 
               AutoSize        =   -1  'True
               Caption         =   "Nº . :"
               Height          =   195
               Left            =   3240
               TabIndex        =   89
               Top             =   1755
               Width           =   360
            End
            Begin VB.Label lbtipref 
               Caption         =   "T.D. Ref. :"
               Height          =   255
               Left            =   120
               TabIndex        =   88
               Top             =   1740
               Width           =   1020
            End
            Begin VB.Label lbconv 
               Caption         =   "Conversión :"
               ForeColor       =   &H00800000&
               Height          =   240
               Left            =   7785
               TabIndex        =   84
               Top             =   2355
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.Label lb_vcambio 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00F2FEFF&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   8775
               TabIndex        =   70
               Top             =   1230
               Width           =   2220
            End
            Begin VB.Label lbTipAnal 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Analitico :"
               Height          =   195
               Index           =   0
               Left            =   4020
               TabIndex        =   19
               Top             =   870
               Width           =   1050
            End
            Begin VB.Label lbAnalitico 
               AutoSize        =   -1  'True
               Caption         =   "Analitico :"
               Height          =   195
               Left            =   120
               TabIndex        =   17
               Top             =   1170
               Width           =   690
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Cod Oper. :"
               Height          =   195
               Left            =   120
               TabIndex        =   16
               Top             =   270
               Width           =   810
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Cuenta :"
               Height          =   195
               Left            =   120
               TabIndex        =   15
               Top             =   555
               Width           =   600
            End
            Begin VB.Label lbccosto 
               AutoSize        =   -1  'True
               Caption         =   "C.Costo :"
               Height          =   195
               Left            =   120
               TabIndex        =   14
               Top             =   855
               Width           =   645
            End
            Begin VB.Label lbruc 
               AutoSize        =   -1  'True
               Caption         =   "R.U.C. :"
               Height          =   195
               Left            =   5235
               TabIndex        =   13
               Top             =   1170
               Width           =   570
            End
            Begin VB.Shape Shape3 
               BorderColor     =   &H00808080&
               Height          =   2385
               Left            =   7680
               Top             =   225
               Width           =   15
            End
            Begin VB.Shape Shape4 
               BorderColor     =   &H00FFFFFF&
               Height          =   2385
               Left            =   7695
               Top             =   225
               Width           =   15
            End
            Begin VB.Label lbtipdoc 
               Caption         =   "Tipo doc. :"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   1440
               Width           =   1020
            End
            Begin VB.Label lbndocum 
               AutoSize        =   -1  'True
               Caption         =   "Nº doc. :"
               Height          =   195
               Left            =   3240
               TabIndex        =   11
               Top             =   1455
               Width           =   630
            End
            Begin VB.Label lbFechaDoc 
               AutoSize        =   -1  'True
               Caption         =   "Fecha doc. :"
               Height          =   195
               Left            =   120
               TabIndex        =   10
               Top             =   2115
               Width           =   900
            End
            Begin VB.Label lbFechVen 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Venc. :"
               Height          =   195
               Left            =   2505
               TabIndex        =   9
               Top             =   2100
               Width           =   1230
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Indicador :"
               Height          =   195
               Left            =   7770
               TabIndex        =   8
               Top             =   225
               Width           =   750
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "T/Cambio :"
               Height          =   195
               Left            =   7770
               TabIndex        =   7
               Top             =   930
               Width           =   795
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "V/Cambio :"
               Height          =   195
               Left            =   7800
               TabIndex        =   6
               Top             =   1275
               Width           =   795
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Monto :"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   7785
               TabIndex        =   5
               Top             =   1605
               Width           =   540
            End
            Begin VB.Label Label23 
               Caption         =   "Glosa :"
               Height          =   195
               Left            =   120
               TabIndex        =   4
               Top             =   2400
               Width           =   495
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Moneda :"
               Height          =   195
               Left            =   7800
               TabIndex        =   3
               Top             =   615
               Width           =   675
            End
         End
         Begin VB.Shape Shilu1 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            Height          =   30
            Left            =   1575
            Top             =   15
            Visible         =   0   'False
            Width           =   9630
         End
      End
      Begin VB.Frame FrameConsulta 
         BackColor       =   &H00808080&
         Height          =   6735
         Left            =   -74910
         TabIndex        =   71
         Top             =   2160
         Width           =   11250
         Begin TextFer.TxFer TxEjecutar 
            Height          =   300
            Left            =   120
            TabIndex        =   85
            Top             =   465
            Width           =   7485
            _ExtentX        =   13203
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
            Valor           =   ""
         End
         Begin VB.CheckBox ChkTodos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808080&
            Caption         =   "Todos"
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   7650
            TabIndex        =   78
            Top             =   480
            Width           =   855
         End
         Begin MSDataListLib.DataCombo Dtc_Campo 
            Height          =   315
            Left            =   9375
            TabIndex        =   77
            Top             =   435
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "nombre"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin TrueOleDBGrid70.TDBGrid TDBG_Consulta 
            Height          =   5460
            Left            =   120
            TabIndex        =   72
            Top             =   810
            Width           =   11040
            _ExtentX        =   19473
            _ExtentY        =   9631
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Comprobante"
            Columns(0).DataField=   "cabcomprobnumero"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Fecha Cont."
            Columns(1).DataField=   "cabcomprobfeccontable"
            Columns(1).NumberFormat=   "Short Date"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Total Debe S/."
            Columns(2).DataField=   "cabcomprobtotdebe"
            Columns(2).NumberFormat=   "###,###,###,###.00"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Total Haber S/."
            Columns(3).DataField=   "cabcomprobtothaber"
            Columns(3).NumberFormat=   "###,###,###,###.00"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Total Deb US$"
            Columns(4).DataField=   "cabcomprobtotussdebe"
            Columns(4).NumberFormat=   "###,###,###,###.00"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Total Haber US $"
            Columns(5).DataField=   "cabcomprobtotusshaber"
            Columns(5).NumberFormat=   "###,###,###,###.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Estado"
            Columns(6).DataField=   "estcomprobcodigo"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2805"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2725"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2566"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2487"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=3149"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3069"
            Splits(0)._ColumnProps(12)=   "Column(2)._ColStyle=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(14)=   "Column(3).Width=2990"
            Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2910"
            Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(19)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=2"
            Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(24)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=2"
            Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(29)=   "Column(6).Width=979"
            Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=900"
            Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
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
            Caption         =   "Resultados de La Consulta"
            MultipleLines   =   0
            CellTips        =   2
            CellTipsWidth   =   0
            MultiSelect     =   2
            DeadAreaBackColor=   13160660
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=1,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&H8000000F&"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H344A87&"
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
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1,.bgcolor=&H80000018&"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1,.bgcolor=&H80000018&"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1,.bgcolor=&H80000018&"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1,.bgcolor=&H80000018&"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(64)  =   "Named:id=33:Normal"
            _StyleDefs(65)  =   ":id=33,.parent=0"
            _StyleDefs(66)  =   "Named:id=34:Heading"
            _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   ":id=34,.wraptext=-1"
            _StyleDefs(69)  =   "Named:id=35:Footing"
            _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   "Named:id=36:Selected"
            _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=37:Caption"
            _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(75)  =   "Named:id=38:HighlightRow"
            _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=39:EvenRow"
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
         End
         Begin VB.Shape Shape8 
            BackColor       =   &H8000000B&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   90
            Left            =   0
            Top             =   0
            Width           =   11265
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H00404040&
            Height          =   285
            Left            =   10065
            Top             =   6345
            Width           =   1095
         End
         Begin VB.Label lbl_nregconsulta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "0 "
            Height          =   285
            Left            =   10080
            TabIndex        =   76
            Top             =   6345
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Registros :"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   9000
            TabIndex        =   75
            Top             =   6390
            Width           =   975
         End
         Begin VB.Label Label2 
            BackColor       =   &H00808080&
            Caption         =   "Valor :"
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   120
            TabIndex        =   74
            Top             =   210
            Width           =   2085
         End
         Begin VB.Label Label4 
            BackColor       =   &H00808080&
            Caption         =   "Criterio :"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   8715
            TabIndex        =   73
            Top             =   510
            Width           =   570
         End
      End
      Begin VB.Shape Shilu2 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         Height          =   2865
         Left            =   11295
         Top             =   2490
         Visible         =   0   'False
         Width           =   30
      End
   End
End
Attribute VB_Name = "frmantcomprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ClsMM1 As ClsMantMov1
Dim rscampo As ADODB.Recordset
Dim rscabecera As ADODB.Recordset
Dim WithEvents rsmantenimiento As ADODB.Recordset
Attribute rsmantenimiento.VB_VarHelpID = -1
Public IMant As Integer
Dim adReasonAux As ADODB.EventReasonEnum
Public VPAsiento As String, VPSubAsiento As String
Dim VlUltAccion As Integer
Public VlGrabada As Boolean
Public VlNref As Boolean
Public Vllabelsref As String
'FIXIT: Declare 'VlCtaAjuste' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
Dim VlCtaAjuste As Variant
Dim VlLibro As String
Dim VlNotaCredito As Boolean
Dim VlExisteRef As Boolean
Dim m_CodComprob As String

Private Sub ChkAjusta_Click()
    If ChkAjusta.Value = 1 Then
        lbconv.Visible = True
        TxValor.Visible = True
        If CtrAyu_Moneda.xclave = VGParametros.monedabase Then
           TxValor.Text = Format(rsmantenimiento!montouss, "###,###,###.00")
           TxValor.valor = rsmantenimiento!montouss
          Else
           TxValor.Text = Format(rsmantenimiento!montosol, "###,###,###.00")
           TxValor.valor = rsmantenimiento!montosol
        End If
       Else
        lbconv.Visible = False
        TxValor.Visible = False
    End If
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    If VGMoverRegistro Then Exit Sub
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, Montos)
    Call CalcularTotales(rsmantenimiento)
End Sub

Private Sub ChkGrabado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call SendKeys("{TAB}")
End Sub

Private Sub ChkInafecto_Click()
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, plantillaasientoinafecto)
End Sub

Private Sub ChkInafecto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call SendKeys("{TAB}")
End Sub

Private Sub ChkTodos_Click()
    If ChkTodos.Value = 1 Then
        Call EjecutarConsulta("", True)
      Else
        Call EjecutarConsulta("", False)
    End If
End Sub
Private Sub CmbID_Click()
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, indicador)
    Call ClsMM1.CalculoIGV(rsmantenimiento)
    Call ClsMM1.CalculodeAjuste(rsmantenimiento)
    Call CalcularTotales(rsmantenimiento)
End Sub

Private Sub CmbID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub
Private Sub CmbTcambio_Click()
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Set ClsMM1 = New ClsMantMov1
    VGValorCambio = ClsMM1.RecuperaTipoCambio(Format(Dtp_FechaDoc, "dd/mm/yyyy"), CmbTcambio.ListIndex + 1)
    lb_vcambio.Caption = Format(VGValorCambio, "#.000 ")
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, Montos)
    Call ClsMM1.CalculoIGV(rsmantenimiento)
    Call ClsMM1.CalculodeAjuste(rsmantenimiento)
    Call CalcularTotales(rsmantenimiento)
End Sub
Private Sub CmbTcambio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub

Private Sub CmdDocPend_Click()
    Call MuestraDocPend
End Sub

Private Sub CtrAyu_Analitico_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, analiticocodigo)
    txRuc.Text = ESNULO(Trim$(ColecCampos("entidadruc").Value), ""): txRuc.Locked = True
    
End Sub

Private Sub CtrAyu_Analitico_AlNoDevolverNada()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, analiticocodigo)
    txRuc.Text = "": txRuc.Locked = False
    
End Sub

Private Sub CtrAyu_Asiento_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    CtrAyu_SubAsiento.Filtro = "asientocodigo='" & Trim$(CtrAyu_Asiento.xclave) & "'"
    CtrAyu_SubAsiento.xclave = "": CtrAyu_SubAsiento.xnombre = ""
    
'    CtrAyu_SubAsiento.Enabled = True
'    lbSubAsiento.Enabled = True
    VlGrabada = ESNULO(ColecCampos("flaggrabado").Value, 0)
    VlNref = ESNULO(ColecCampos("controlnref").Value, 0)
    Vllabelsref = ESNULO(ColecCampos("nemotecref").Value, "")
    VlLibro = ColecCampos("librocodigo").Value
End Sub

Private Sub CtrAyu_Asiento_AlNoDevolverNada()
    CtrAyu_SubAsiento.xclave = "": CtrAyu_SubAsiento.xnombre = ""
    VlGrabada = False
    VlNref = False
    Vllabelsref = ""
    VlLibro = ""
'    CtrAyu_SubAsiento.Enabled = False
'    lbSubAsiento.Enabled = False
End Sub

Private Sub CtrAyu_CCosto_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, centrocostocodigo)
End Sub

Private Sub CtrAyu_CCosto_AlNoDevolverNada()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, centrocostocodigo)
    VlCtaAjuste = ""
End Sub

Private Sub CtrAyu_Cuenta_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
     If Not VGflaglimpia Then Exit Sub
     Call HabilitarSegunCuenta(ColecCampos("cuentaestadoccostos").Value, ColecCampos("cuentaestadoanalitico").Value, ColecCampos("cuentadocumento").Value)
     Set ClsMM1 = New ClsMantMov1
     Call ClsMM1.ActualizarDetalle(rsmantenimiento, cuentacodigo)
     Set VGvardllgen = New dllgeneral.dll_general
     CtrAyu_TipAnal.xclave = Trim$(VGvardllgen.ESNULO(ColecCampos("tipoanaliticocodigo").Value, ""))
     CtrAyu_TipAnal.Ejecutar
     VlCtaAjuste = Trim$(VGvardllgen.ESNULO(ColecCampos("tipoajuste").Value, ""))
   
     
End Sub
Private Sub HabilitarSegunCuenta(SiCCostos As Boolean, SiAnalitico As Boolean, SiDocumento As Boolean)
    CtrAyu_CCosto.Visible = SiCCostos
    lbccosto.Visible = SiCCostos
    If Not SiCCostos Then
        CtrAyu_CCosto.xclave = "00": CtrAyu_CCosto.xnombre = "(Ninguno)"
    End If
    CtrAyu_TipAnal.Visible = SiAnalitico
    CtrAyu_TipAnal.Enabled = Not SiAnalitico
    lbTipAnal(0).Visible = SiAnalitico
    CtrAyu_Analitico.Visible = SiAnalitico
    lbAnalitico.Visible = SiAnalitico
    If Not SiAnalitico Then
        CtrAyu_TipAnal.xclave = "00": CtrAyu_TipAnal.xnombre = "(Ninguno)"
        CtrAyu_Analitico.xclave = "00": CtrAyu_Analitico.xnombre = "(Ninguno)"
       Else
        CtrAyu_TipAnal.xclave = "": CtrAyu_TipAnal.xnombre = ""
        CtrAyu_Analitico.xclave = "": CtrAyu_Analitico.xnombre = ""
        txRuc.Text = ""
    End If
    txRuc.Visible = SiAnalitico
    lbruc.Visible = SiAnalitico
    CtrAyu_TipDoc.Visible = SiDocumento
    lbtipdoc.Visible = SiDocumento
    lbtipref.Visible = SiDocumento
    CtrAyu_TipRef.Visible = SiDocumento
    If Not SiDocumento Then
        CtrAyu_TipDoc.xclave = "00": CtrAyu_TipDoc.xnombre = "(Ninguno)"
        CtrAyu_TipRef.xclave = "00": CtrAyu_TipRef.xnombre = "(Ninguno)"
       Else
        CtrAyu_TipDoc.xclave = "": CtrAyu_TipDoc.xnombre = ""
        CtrAyu_TipRef.xclave = "": CtrAyu_TipRef.xnombre = ""
    End If
    TxSerie.Visible = SiDocumento
    TxNdoc.Visible = SiDocumento
    lbndocum.Visible = SiDocumento
    TxNref.Visible = SiDocumento
    lbnref.Visible = SiDocumento
    CmdDocPend.Visible = SiDocumento
    Dtp_FechaDoc.Visible = SiDocumento
'    lbFechaDoc.Visible = SiDocumento
    DtpFech_Ven.Visible = SiDocumento
    lbFechVen.Visible = SiDocumento
    Dtp_FechaDocRef.Visible = SiDocumento
End Sub


Private Sub CtrAyu_Cuenta_AlNoDevolverNada()
    If Not VGflaglimpia Then Exit Sub
    Call HabilitarSegunCuenta(False, False, False)
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, cuentacodigo)
End Sub

Private Sub CtrAyu_Moneda_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, monedacodigo)
    'Poner un flag para que no entre cuando es recursivo
    If vgcont = 2 Then Exit Sub
    Call ClsMM1.CalculoIGV(rsmantenimiento)
    Call ClsMM1.CalculodeAjuste(rsmantenimiento)
    Call CalcularTotales(rsmantenimiento)
End Sub

Private Sub CtrAyu_Moneda_AlNoDevolverNada()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, monedacodigo)
End Sub

Private Sub CtrAyu_Opera_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, operacioncodigo)
    Vgdocumentoanulado = IIf(ESNULO(ColecCampos!operaciondocumentoanulado, False), ColecCampos!operacioncodigo, "")
End Sub

Private Sub CtrAyu_Opera_AlNoDevolverNada()
If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, operacioncodigo)
End Sub

Private Sub CtrAyu_SubAsiento_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
        VGGlosa = ColecCampos("subasientoglosa").Value
        VGRepiteDoc = ColecCampos("subasientorepitedoc").Value
        VGMonSubAsiento = ColecCampos("monedacodigo").Value
End Sub

Private Sub CtrAyu_TipAnal_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    CtrAyu_Analitico.Filtro = "tipoanaliticocodigo='" & Trim$(CtrAyu_TipAnal.xclave) & "'"
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, tipoanaliticocodigo)
End Sub

Private Sub CtrAyu_TipAnal_AlNoDevolverNada()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, tipoanaliticocodigo)
End Sub

Private Sub CtrAyu_TipDoc_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, documentocodigo)
    VlNotaCredito = ColecCampos("documentonotacredito").Value
End Sub

Private Sub CtrAyu_TipDoc_AlNoDevolverNada()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, documentocodigo)
    VlNotaCredito = False
End Sub

Private Sub CtrAyu_TipRef_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, documentocodigo)
End Sub

Private Sub CtrAyu_TipRef_AlNoDevolverNada()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, documentocodigo)
End Sub

Private Sub Dtc_Campo_Change()
    Label2.Caption = Dtc_Campo.Text
End Sub

Public Sub Dtp_FechaDoc_Change()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Dim POSANTERIOR As Integer
    If Not VlNotaCredito Then
        VGValorCambio = ClsMM1.RecuperaTipoCambio(Format(Dtp_FechaDoc, "dd/mm/yyyy"), CmbTcambio.ListIndex + 1)
        lb_vcambio.Caption = Format(VGValorCambio, "#.000 ")
        Dim flagcambio As Boolean
        flagcambio = False
        If VGValorCambio = 0 Then
            MsgBox "No Existe tipo cambio para esta fecha", vbInformation
            Dtp_FechaDoc.SetFocus
            flagcambio = True
        End If
    End If
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, detcomprobfechaemision)
    If VGMoverRegistro Then Exit Sub
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, documentocodigo)
    If VGRepiteDoc Then
      With rsmantenimiento
         POSANTERIOR = .AbsolutePosition
        .MoveFirst
        While Not .EOF
            !detcomprobfechaemision = Dtp_FechaDoc
'            !documentocodigo = frmantcomprobantes.CtrAyu_TipDoc.xclave
            .Update
            .MoveNext
        Wend
        .AbsolutePosition = POSANTERIOR
      End With
      Screen.MousePointer = 1
    End If
    
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, Montos)
    If vgcont = 2 Then Exit Sub
    Call ClsMM1.CalculoIGV(rsmantenimiento)
    Call ClsMM1.CalculodeAjuste(rsmantenimiento, flagcambio)
    Call CalcularTotales(rsmantenimiento)
End Sub

Private Sub Dtp_FechaDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub

Private Sub Dtp_FechaDocRef_Change()
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, detcomprobfecharef)
    If Not VlNotaCredito Then Exit Sub
    Call DTPFechaComprobCab_Change
        
        
End Sub

Private Sub Dtp_FechaDocRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub

Private Sub DtpFech_Ven_Change()
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, detcomprobfechavencimiento)
End Sub

Private Sub DtpFech_Ven_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub

Private Sub DTPFechaComprobCab_Change()
    Set ClsMM1 = New ClsMantMov1
Dim Fecha As Date
    
    If Not VlNotaCredito Then
        VGValorCambio = ClsMM1.RecuperaTipoCambio(Format(DTPFechaComprobCab, "dd/mm/yyyy"), Venta)
        lb_vcambio.Caption = Format(VGValorCambio, "#.000 ")
        Fecha = DTPFechaComprobCab
     Else
        VGValorCambio = ClsMM1.RecuperaTipoCambio(Format(Dtp_FechaDocRef, "dd/mm/yyyy"), CmbTcambio.ListIndex + 1)
        lb_vcambio.Caption = Format(VGValorCambio, "#.000 ")
        Fecha = Dtp_FechaDocRef
    End If
    If rsmantenimiento Is Nothing Then Exit Sub
    If rsmantenimiento.RecordCount = 0 Then Exit Sub
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
'    If Not VlNotaCredito Then Dtp_FechaDoc.Value = DTPFechaComprobCab
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, monedacodigo, True, Fecha)
    
    Call ClsMM1.CalculoIGV(rsmantenimiento)
    Call ClsMM1.CalculodeAjuste(rsmantenimiento)
    Call CalcularTotales(rsmantenimiento)
End Sub

Private Sub DTPFechaComprobCab_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call SendKeys("{TAB}")
End Sub

Private Sub DTPFechaComprobCab_LostFocus()
    Set VGvardllgen = New dllgeneral.dll_general
    If VGValorCambio = 0 And IMant = 1 Then
        MsgBox "No se encuentra el tipo de cambio para esta fecha " & _
               "Por lo tanto no se podra realizar niguna transaccion", vbInformation
        Call Cancelar
        Call PBoton(VlUltAccion)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub DTPFechaContab_Change()
    DTPFechaComprobCab.Value = DTPFechaContab.Value
End Sub

Private Sub Form_Activate()
    MDIPrincipal.ToolComprob.Visible = True
    MDIPrincipal.mnu00.Visible = True
    Call PBoton(VlUltAccion)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Top = 0
    Left = 0
    DTPFechaContab.Value = VGParamSistem.FechaTrabajo
    Dtp_FechaDoc.Value = VGParamSistem.FechaTrabajo
    IMant = 0
    VlUltAccion = 0
    Set VGvardllgen = New dllgeneral.dll_general
    Width = 11565
    Height = 9465
    Set rscabecera = New ADODB.Recordset
    Set ClsMM1 = New ClsMantMov1
    ClsMM1.CargarAyudas
    Set TDBG_Consulta.DataSource = Nothing
    TDBG_Det.FetchRowStyle = True
    Call PrepararTemporalDetalle
    If rsmantenimiento.RecordCount = 0 Then
        Call HabilitarDetalle(False, FramDetalle)
     Else
        Call HabilitarDetalle(True, FramDetalle)
    End If
    Call VGvardllgen.ActivaTab(0, 1, SSTabMant)
    Call GetCamposdeConsulta
    
    If m_CodComprob <> Empty Then
       TxEjecutar.Text = m_CodComprob
       SendKeys "{ENTER}"
    End If
   ' Label2.Caption = Dtc_Campo.Text
    'Cargar los tools
End Sub
Public Sub AlMoverRegistro()
Dim vardllgen As New dllgeneral.dll_general
Dim pos As Integer
    If VGactulizodoc Then Exit Sub 'Estoy Actualizando documentos
    VGMoverRegistro = True
    On Error Resume Next
    With rsmantenimiento
        FramDetalle.Enabled = Not (!detcomprobauto Or VGParametros.cierremes = True)
        CtrAyu_Opera.xclave = !operacioncodigo: CtrAyu_Opera.Ejecutar
        CtrAyu_TipAnal.xclave = !tipoanaliticocodigo: CtrAyu_TipAnal.Ejecutar
        CtrAyu_Cuenta.xclave = !cuentacodigo: CtrAyu_Cuenta.Ejecutar
        CtrAyu_CCosto.xclave = !centrocostocodigo: CtrAyu_CCosto.Ejecutar
        CtrAyu_Analitico.xclave = Trim$(!analiticocodigo): CtrAyu_Analitico.Ejecutar
        If !cuentacodigo = "00" Or vardllgen.ESNULO(!cuentacodigo, "") = "" Then
            If vardllgen.ESNULO(!comodin, "") <> "" Then
                CtrAyu_Cuenta.Filtro = "(empresacodigo='" & VGParametros.empresacodigo & "' and (cuentanivel=" & VGnumnivelescuenta & ") and (cuentacodigo<>'00') and (" & !comodin & "))"
              Else
                CtrAyu_Cuenta.Filtro = "(empresacodigo='" & VGParametros.empresacodigo & "' and (cuentacodigo<>'00')" & " and (cuentanivel=" & VGnumnivelescuenta & "))"
            End If
           Else
                CtrAyu_Cuenta.Filtro = "(empresacodigo='" & VGParametros.empresacodigo & "' and (cuentacodigo<>'00')" & " and (cuentanivel=" & VGnumnivelescuenta & "))"
        End If
        txRuc.Text = !detcomprobruc
        CtrAyu_TipDoc.xclave = !documentocodigo: CtrAyu_TipDoc.Ejecutar
'        pos = InStr(1, !detcomprobnumdocumento, "-", vbTextCompare)
        If Trim$(Len(!detcomprobnumdocumento)) > 4 Then
            TxSerie.Text = Left(Trim$(!detcomprobnumdocumento), TxSerie.MaxLength)
            TxNdoc.Text = Mid$(Trim$(!detcomprobnumdocumento), TxSerie.MaxLength + 1, Len(Trim$(!detcomprobnumdocumento)) - TxSerie.MaxLength)
        Else
            TxSerie.Text = Trim$(!detcomprobnumdocumento)
            TxNdoc.Text = ""
        End If
        Dtp_FechaDoc = Format(!detcomprobfechaemision, "dd/mm/yyyy")
        DtpFech_Ven = Format(ESNULO(!detcomprobfechavencimiento, !detcomprobfechaemision), "dd/mm/yyyy")
        Dtp_FechaDocRef = !detcomprobfecharef
        TxGlosa.Text = !detcomprobglosa
        CmbID.ListIndex = IIf(!indicador = "D", 0, 1)
        CtrAyu_Moneda.xclave = !monedacodigo: CtrAyu_Moneda.Ejecutar
    Select Case !tcambio
            Case "01" 'Compra
                CmbTcambio.ListIndex = 0
            Case "02" 'Venta
                CmbTcambio.ListIndex = 1
            Case "03" 'Promedio
                CmbTcambio.ListIndex = 2
        End Select
        lb_vcambio.Caption = Format(vardllgen.ESNULO(!valcambio, 0), "#.000 ")

        If !monedacodigo = VGParametros.monedabase Then
            TxMonto.Text = Format(vardllgen.ESNULO(!montosol, 0), "#.00")
          Else
            TxMonto.Text = Format(vardllgen.ESNULO(!montouss, 0), "#.00")
        End If
        ChkAjusta.Value = IIf(!detcomprobajusteuser, 1, 0)
        'ChkInafecto.Visible = !plantillaasientoinafecto
        ChkInafecto.Value = IIf(!plantillaasientoinafecto, 1, 0)
        CtrAyu_TipRef.xclave = !tipdocref: CtrAyu_TipRef.Ejecutar
        TxNref.Text = !detcomprobnumref
    End With
    VGMoverRegistro = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.ToolComprob.Visible = False
    MDIPrincipal.mnu00.Visible = False
End Sub

Private Sub rsmantenimiento_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If (adReason = adRsnMove Or adReason = adRsnMoveNext) And pRecordset.RecordCount > 0 And adReasonAux <> adRsnAddNew Then
        Call AlMoverRegistro
    End If
    If adReasonAux = adRsnAddNew Then adReasonAux = adRsnMove
End Sub

Private Sub rsmantenimiento_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    adReasonAux = adReason
End Sub

Private Sub SSTabMant_Click(PreviousTab As Integer)
Dim Existelibro As Boolean
    If PreviousTab = 0 Then
        StBar.Panels(1).Text = "Asiento : " + VPAsiento & " - " & CtrAyu_Asiento.xnombre
        StBar.Panels(2).Text = "Sub Asiento : " + VPSubAsiento & " - " & CtrAyu_SubAsiento.xnombre
        CtrAyu_Opera.Requerido = True
        CtrAyu_Cuenta.Requerido = True
        CtrAyu_CCosto.Requerido = True
        CtrAyu_TipAnal.Requerido = True
        CtrAyu_Analitico.Requerido = True
        CtrAyu_TipDoc.Requerido = True
        CtrAyu_TipRef.Requerido = False
        CtrAyu_Moneda.Requerido = True
  '      CtrAyu_Moneda.Enabled = True
        MDIPrincipal.mnu00_01(9).Visible = True
        If VlGrabada Then
            ChkGrabado.Visible = True
          Else
            ChkGrabado.Visible = False
        End If
        If VlNref Then
            lbnemoref.Visible = True
            lbnemoref.Caption = Vllabelsref
            TxCtrNref.Visible = True
          Else
            lbnemoref.Visible = False
            TxCtrNref.Visible = False
        End If
        Existelibro = ExisteSQL(VGCNx, "Select flagcontrol From ct_libro where librocodigo='" & _
                          Trim$(VlLibro) & "' and flagcontrol <> 0 ")
        If Existelibro Then
            leNComprob(0).Visible = True
            lbNumComprobCablibro.Visible = True
          Else
            leNComprob(0).Visible = False
            lbNumComprobCablibro.Visible = False
        End If
        VlNotaCredito = False
       Else
        CtrAyu_Opera.Requerido = False
        CtrAyu_Cuenta.Requerido = False
        CtrAyu_CCosto.Requerido = False
        CtrAyu_TipAnal.Requerido = False
        CtrAyu_Analitico.Requerido = False
        CtrAyu_TipDoc.Requerido = False
        CtrAyu_TipRef.Requerido = False
        CtrAyu_Moneda.Requerido = False
        MDIPrincipal.mnu00_01(9).Visible = False
    End If
End Sub

Private Sub TDBG_Consulta_DblClick()
    Call Modificar
End Sub

Private Sub TDBG_Consulta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Modificar
    End If
End Sub

'FIXIT: Declare 'LastRow' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Private Sub TDBG_Consulta_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rscabecera.State = 0 Then Exit Sub
    If rscabecera.RecordCount = 0 Then Exit Sub
    CtrAyu_Asiento.xclave = rscabecera!asientocodigo: CtrAyu_Asiento.Ejecutar
    CtrAyu_SubAsiento.xclave = rscabecera!subasientocodigo: CtrAyu_SubAsiento.Ejecutar
End Sub


Private Sub TDBG_Det_DblClick()
    MsgBox rsmantenimiento!Index
End Sub

'FIXIT: Declare 'Bookmark' con un tipo de datos de enlace en tiempo de compilación         FixIT90210ae-R1672-R1B8ZE
Private Sub TDBG_Det_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
Dim rsclone As New ADODB.Recordset
On Error Resume Next
Set rsclone = rsmantenimiento.Clone(adLockReadOnly)
If rsclone.RecordCount = 0 Then Exit Sub
rsclone.Bookmark = Bookmark
If rsclone!detcomprobauto Then
   RowStyle.BackColor = RGB(185, 251, 236)
End If
End Sub

Private Sub TDBG_Det_GotFocus()
    'frameGrid.BackColor = &H628837
'FIXIT: En Visual Basic .NET no hay control Shape. Los controles Shape cuadrado y rectangular se actualizan a controles Label. Los controles Shape ovalados y circulares no se actualizan a Visual Basic .NET.     FixIT90210ae-R2149-R57265
'FIXIT: En Visual Basic .NET no hay control Shape. Los controles Shape cuadrado y rectangular se actualizan a controles Label. Los controles Shape ovalados y circulares no se actualizan a Visual Basic .NET.     FixIT90210ae-R2149-R57265
'FIXIT: En Visual Basic .NET no hay control Shape. Los controles Shape cuadrado y rectangular se actualizan a controles Label. Los controles Shape ovalados y circulares no se actualizan a Visual Basic .NET.     FixIT90210ae-R2149-R57265
'FIXIT: En Visual Basic .NET no hay control Shape. Los controles Shape cuadrado y rectangular se actualizan a controles Label. Los controles Shape ovalados y circulares no se actualizan a Visual Basic .NET.     FixIT90210ae-R2149-R57265
    Shilu1.Visible = True: Shilu2.Visible = True
End Sub

Private Sub TDBG_Det_LostFocus()
'FIXIT: En Visual Basic .NET no hay control Shape. Los controles Shape cuadrado y rectangular se actualizan a controles Label. Los controles Shape ovalados y circulares no se actualizan a Visual Basic .NET.     FixIT90210ae-R2149-R57265
'FIXIT: En Visual Basic .NET no hay control Shape. Los controles Shape cuadrado y rectangular se actualizan a controles Label. Los controles Shape ovalados y circulares no se actualizan a Visual Basic .NET.     FixIT90210ae-R2149-R57265
'FIXIT: En Visual Basic .NET no hay control Shape. Los controles Shape cuadrado y rectangular se actualizan a controles Label. Los controles Shape ovalados y circulares no se actualizan a Visual Basic .NET.     FixIT90210ae-R2149-R57265
'FIXIT: En Visual Basic .NET no hay control Shape. Los controles Shape cuadrado y rectangular se actualizan a controles Label. Los controles Shape ovalados y circulares no se actualizan a Visual Basic .NET.     FixIT90210ae-R2149-R57265
    Shilu1.Visible = False: Shilu2.Visible = False
    'frameGrid.BackColor = &H808080
End Sub

Private Sub TxEjecutar_KeyDown(KeyCode As Integer, Shift As Integer)
Dim cad As String
    If KeyCode = 13 Then
        cad = Dtc_Campo.BoundText & " like '" & Trim$(TxEjecutar.Text) & "%'"
        Call EjecutarConsulta(cad, False)
    End If
End Sub

Private Sub TxGlosa_Change()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, detcomprobglosa)
End Sub

Private Sub TxGlosaComprobCab_Change()
    If rsmantenimiento Is Nothing Then Exit Sub
    If rsmantenimiento.RecordCount = 0 Then Exit Sub
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, 0)
End Sub

Private Sub TxMonto_Change()
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, Montos)
    Call ClsMM1.CalculoIGV(rsmantenimiento)
    Call ClsMM1.CalculodeAjuste(rsmantenimiento)
    Call CalcularTotales(rsmantenimiento)
End Sub

Private Sub TxMonto_GotFocus()
    If Trim$(CtrAyu_Moneda.xclave) = "" Or CtrAyu_Moneda.xclave = "00" Then
        MsgBox "Antes de colocar el monto debe elegir la moneda ", vbInformation
        CtrAyu_Moneda.SetFocus
        Exit Sub
    End If
End Sub

Private Sub TxNdoc_Change()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, detcomprobnumdocumento)
End Sub

Private Sub TxNref_Change()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, detcomprobnumdocumento)
End Sub

Private Sub TxNref_LostFocus()
Dim rsfecha As ADODB.Recordset
    If Not VlNotaCredito Then Exit Sub
    Dtp_FechaDocRef.Value = Dtp_FechaDoc.Value
    Set rsfecha = New ADODB.Recordset
    rsfecha.Open "select detcomprobfechaemision from ct_detcomprob" & VGParamSistem.Anoproceso & _
 _
                 " where documentocodigo='" & CtrAyu_TipRef.xclave & "' and " & _
                 "       detcomprobnumdocumento='" & Trim$(TxNref.Text) & "' and " & _
                 "       analiticocodigo='" & Trim$(CtrAyu_Analitico.xclave) & "'", VGCNx, adOpenKeyset, adLockReadOnly
    If rsfecha.RecordCount = 0 Then
        MsgBox "No se encuentro el documento de referencia", vbExclamation
        Exit Sub
    End If
    Dtp_FechaDocRef.Value = rsfecha!detcomprobfechaemision
    Call Dtp_FechaDocRef_Change
End Sub

Private Sub txRuc_Change()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, detcomprobruc)
End Sub
Private Sub TxSerie_Change()
    If Not VGflaglimpia Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, detcomprobnumdocumento)
End Sub
Public Sub CalcularTotales(ByVal rs As Recordset)
On Error GoTo ERRX
Dim RSAUX As ADODB.Recordset
Dim vardllgen As New dllgeneral.dll_general

Set RSAUX = rs.Clone(adLockReadOnly)
Dim montosoles As Double, montodolares As Double
Dim difsoles As Double, difdolares As Double
Dim montosolesDebe As Double, montodolaresDebe As Double
Dim montosolesHaber As Double, montodolaresHaber As Double

montosolesDebe = 0: montodolaresDebe = 0:
montosolesHaber = 0: montodolaresHaber = 0:
difsoles = 0: difdolares = 0
RSAUX.MoveFirst
    While Not RSAUX.EOF
        If RSAUX!indicador = "D" Then
            montosolesDebe = montosolesDebe + Round(vardllgen.ESNULO(RSAUX!montosol, 0), 2)
            montodolaresDebe = montodolaresDebe + Round(vardllgen.ESNULO(RSAUX!montouss, 0), 2)
          Else
            montosolesHaber = montosolesHaber + Round(vardllgen.ESNULO(RSAUX!montosol, 0), 2)
            montodolaresHaber = montodolaresHaber + Round(vardllgen.ESNULO(RSAUX!montouss, 0), 2)
        End If
        RSAUX.MoveNext
    Wend
    difsoles = montosolesDebe - montosolesHaber
    difdolares = montodolaresDebe - montodolaresHaber
    'Soles
    LbTotales(0).Caption = Format(montosolesDebe, "###,###,###,###.00 ") ' Debe
    LbTotales(1).Caption = Format(montosolesHaber, "###,###,###,###.00 ") ' Haber
    LbTotales(2).Caption = Format(difsoles, "###,###,###,###.00 ") ' Diferencia
    'Dolares
    LbTotales(3).Caption = Format(montodolaresDebe, "###,###,###,###.00 ") ' Debe
    LbTotales(4).Caption = Format(montodolaresHaber, "###,###,###,###.00 ") ' Haber
    LbTotales(5).Caption = Format(difdolares, "###,###,###,###.00 ") ' Diferencia
    
ERRX:
End Sub

Private Sub GetCamposdeConsulta()
    Set rscampo = New ADODB.Recordset
    Call rscampo.Fields.Append("codigo", adVarChar, 60)
    Call rscampo.Fields.Append("Nombre", adVarChar, 50)
    rscampo.Open
    rscampo.AddNew
    rscampo!codigo = "cabcomprobnumero"
    rscampo!nombre = "Nro. Comprobante"
    rscampo.Update
    rscampo.AddNew
    rscampo!codigo = "convert(varchar(10),cabcomprobfeccontable,103)"
    rscampo!nombre = "Fecha de Comprobante"
    rscampo.Update
    Set Dtc_Campo.RowSource = rscampo
    Dtc_Campo.BoundText = "cabcomprobnumero"
End Sub
Private Sub EjecutarConsulta(ByVal criterio As String, Optional ByVal todos As Boolean)
Dim cad As String
Dim sqlcad As String, xasiento As String, xsubasiento As String
    Set rscabecera = New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    xasiento = Trim$(CtrAyu_Asiento.xclave): xsubasiento = Trim$(CtrAyu_SubAsiento.xclave)
    If criterio = "" Then
        cad = " where empresacodigo='" & VGParametros.empresacodigo & "'"
        cad = cad & " and cabcomprobmes=" & CInt(VGParamSistem.Mesproceso) & " and  asientocodigo like '" & VGvardllgen.ESNULO(xasiento, "%%") & "' and  subasientocodigo='" & VGvardllgen.ESNULO(xsubasiento, "%%") & "' and 1=0 "
      Else
        cad = " where empresacodigo='" & VGParametros.empresacodigo & "'"
        cad = cad & " and cabcomprobmes=" & CInt(VGParamSistem.Mesproceso) & " and asientocodigo like '" & VGvardllgen.ESNULO(xasiento, "%%") & "' and  subasientocodigo like '" & VGvardllgen.ESNULO(xsubasiento, "%%") & "' and "
    End If
    If todos Then cad = " where empresacodigo='" & VGParametros.empresacodigo & "' and cabcomprobmes=" & CInt(VGParamSistem.Mesproceso) & "  "
    sqlcad = "select * from " & VGParamSistem.TablaCabcomprob & " " & cad & criterio
    rscabecera.Open sqlcad, VGCNx, adOpenKeyset, adLockReadOnly
    If rscabecera.RecordCount > 0 Then
        lbl_nregconsulta.Caption = Format(rscabecera.RecordCount, "0 ")
        TDBG_Consulta.SetFocus
      Else
        lbl_nregconsulta.Caption = Format(0, "0 ")
        TxEjecutar.SetFocus
    End If
    Set TDBG_Consulta.DataSource = rscabecera
End Sub
Private Sub Mostrar()
    If rscabecera.State = 0 Then Exit Sub
    If rscabecera.RecordCount = 0 Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Set VGvardllgen = New dllgeneral.dll_general
    Call ClearControlsInframe(FrameCabecera)
    
    Call ClsMM1.MostrarCabecera(rscabecera.Fields)
    Call ClsMM1.Limpia
    Call PrepararTemporalDetalle
    Call ClsMM1.MostrarDetalle(rsmantenimiento)
    Call HabilitarDetalle(True, FramDetalle)
    Call VGvardllgen.ActivaTab(1, 1, SSTabMant)
    VlUltAccion = 4
    Call PBoton(VlUltAccion)
End Sub
Private Sub PrepararTemporalDetalle()
    Set rsmantenimiento = New ADODB.Recordset
    Call ClsMM1.CreaRsTempDetalle(rsmantenimiento)
    rsmantenimiento.Open
    rsmantenimiento.Sort = "detcomprobitem asc,index asc"
    Set TDBG_Det.DataSource = rsmantenimiento
End Sub
Public Sub Botones(ByRef tool As Toolbar, Nuevo As Boolean, Grabar As Boolean, eliminar As Boolean, _
                   Modificar As Boolean, Cancelar As Boolean, Anadet As Boolean, EliDet As Boolean)
    With tool.Buttons
        .Item(1).Enabled = Nuevo
        .Item(2).Enabled = Grabar
        .Item(3).Enabled = eliminar
        .Item(4).Enabled = Modificar
        .Item(5).Enabled = Cancelar
        .Item(6).Visible = True
        .Item(7).Visible = True
        .Item(8).Visible = True
        .Item(7).Enabled = Anadet
        .Item(8).Enabled = EliDet
    End With
    With MDIPrincipal
        .mnu00_01(1).Enabled = Nuevo
        .mnu00_01(2).Enabled = Grabar
        .mnu00_01(3).Enabled = eliminar
        .mnu00_01(4).Enabled = Modificar
        .mnu00_01(5).Enabled = Cancelar
        .mnu00_01(6).Visible = True
        .mnu00_01(7).Visible = True
        .mnu00_01(6).Enabled = Anadet
        .mnu00_01(7).Enabled = EliDet
    End With
End Sub
Public Sub Xnuevo()
    'Validacion
    Call PrepararTemporalDetalle
    If Trim$(CtrAyu_Asiento.xclave) = "" Or Trim$(CtrAyu_SubAsiento.xclave) = "" Then
        MsgBox "Necesita seleccionar el Asiento y el SubAsiento para poder ingresar " & Chr(13) & _
               "Un Nuevo Comprobante", vbInformation
        Exit Sub
    End If
    Set VGvardllgen = New dllgeneral.dll_general
    Call ClearControlsInframe(FrameCabecera)
    Call DTPFechaComprobCab_Change
    VPAsiento = CtrAyu_Asiento.xclave
    VPSubAsiento = CtrAyu_SubAsiento.xclave
    lbnregdetalle.Caption = "0 "
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.Limpia
    Call VGvardllgen.ActivaTab(1, 1, SSTabMant)
    DTPFechaComprobCab.SetFocus
    If ClsMM1.CargarPlantillaAsiento(rsmantenimiento, VPAsiento, VPSubAsiento) Then
        Call HabilitarDetalle(True, FramDetalle)
      Else
        Call HabilitarDetalle(False, FramDetalle)
    End If
    TxGlosaComprobCab.Text = VGGlosa
    lbNumComprobCab.Caption = ClsMM1.NumeroComprob(CInt(VGParamSistem.Mesproceso))
    Dim Existelibro As Boolean
    
    Existelibro = ExisteSQL(VGCNx, "Select flagcontrol From ct_libro where librocodigo='" & _
                          Trim$(VlLibro) & "' and flagcontrol <> 0 ")
    lbNumComprobCablibro.Caption = ClsMM1.NumeroComprobLibro(CInt(VGParamSistem.Mesproceso), VlLibro)
                          
    ChkGrabado.Value = 1
    IMant = 1
    VlUltAccion = 1
End Sub
Public Sub Grabar()
Dim xnumerocompro As String, nnumerocorrcomprob As Double
Dim xnumerocomprolibro As String, nnumerocorrcomproblibro As Long
Dim Existelibro As Boolean

Dim varnerror As Integer
Set VGvardllgen = New dllgeneral.dll_general
On Error GoTo ErrorGrabar
Dim xcon As Long
xnumerocomprolibro = "": nnumerocorrcomproblibro = 0
VGvarVerifica = True
VGErrorString = ""
varnerror = 0
rsmantenimiento.Filter = 0
    Set ClsMM1 = New ClsMantMov1
    If Not ClsMM1.ValidarGrabarCabecera(rsmantenimiento.RecordCount) Then Exit Sub
    If Not ClsMM1.ValidarRsDetalle(rsmantenimiento) Then Exit Sub
    
    xcon = rsmantenimiento.RecordCount

    If Vgdocumentoanulado <> "" Then
       rsmantenimiento.Filter = "operacioncodigo='" & Vgdocumentoanulado & "'"
     Else
         rsmantenimiento.Filter = "(montosol<>0 or montouss <> 0)"
     End If
    If rsmantenimiento.RecordCount <= 1 Then
        MsgBox "Por lo Menos debe Existir dos registro con valores ", vbExclamation
        Exit Sub
    End If
    If rsmantenimiento.RecordCount <> xcon Then
        If MsgBox("Esta Seguro de Grabar ? " & Chr(13) & _
                  "Al momento de grabar se eliminaran lo registro ceros ", vbQuestion + vbOKCancel) = vbCancel _
                  Then
            rsmantenimiento.Filter = 0
            Exit Sub
        End If
    End If
    Existelibro = ExisteSQL(VGCNx, "Select flagcontrol From ct_libro where librocodigo='" & _
                          Trim$(VlLibro) & "' and flagcontrol <> 0 ")
    
    
    VGGeneral.BeginTrans 'Inicio la transaccion
    Screen.MousePointer = vbHourglass
    '1=>Paso Genera el Correlativo del Comprobante
    If IMant = 1 Then
        xnumerocompro = ClsMM1.NumeroComprob(CInt(VGParamSistem.Mesproceso), nnumerocorrcomprob)
        If Existelibro Then
            xnumerocomprolibro = ClsMM1.NumeroComprobLibro(CInt(VGParamSistem.Mesproceso), VlLibro, nnumerocorrcomproblibro)
        End If
        '6=>Paso Actualizo el Correlativo en la Tabla SubAsiento si es que ingrese un nuevo
        'Comprobante
        If IMant = 1 Then
            Call ClsMM1.ActualizaCorrelComprob(nnumerocorrcomprob)
            If Existelibro Then Call ClsMM1.ActualizaCorrelComprobLibro(nnumerocorrcomproblibro, VlLibro)
        End If
        If Not VGvarVerifica Then varnerror = 6: GoTo ErrorGrabar
      Else
        xnumerocompro = Trim$(lbNumComprobCab.Caption)
        xnumerocomprolibro = Trim$(lbNumComprobCablibro.Caption)
    End If
    If Not VGvarVerifica Then varnerror = 1: GoTo ErrorGrabar
    '2=>Paso Grabo la Cabecera del Comprobante
    Call ClsMM1.GrabarCabecera(IMant, xnumerocompro, xnumerocomprolibro)
    If Not VGvarVerifica Then varnerror = 2: GoTo ErrorGrabar
    '3=>Paso Grabo los Detalle del Comprobante
    
    
    Call ClsMM1.GrabarDetalle(rsmantenimiento, xnumerocompro, xnumerocomprolibro)
    If Not VGvarVerifica Then varnerror = 3: GoTo ErrorGrabar
    '4=>Generar Asientos Automaticos
    Call ClsMM1.GrabaAsientoAuto(xnumerocompro)
    If Not VGvarVerifica Then varnerror = 4: GoTo ErrorGrabar
    
    '5=>Calcular el total de Cabecera de Comprobante
    Call ClsMM1.CalculaComprob(xnumerocompro)
    If Not VGvarVerifica Then varnerror = 5: GoTo ErrorGrabar
        
    VGGeneral.CommitTrans 'Acepto toda la transaccion porque es correcta
    If IMant = 1 Then
        MsgBox "Se grabo Satisfactoriamente  El numero de Comprobante Generado Es :" & Chr(13) & _
           "Nro: " & xnumerocompro & _
            IIf(Existelibro, Chr(13) & "El Numero de Libro es : " & xnumerocomprolibro, "") _
           , vbInformation
        If VGParametros.ImpresionAsiento Then
            Call ImprimirComprob(xnumerocompro, VPAsiento, VPSubAsiento)
        End If
      Else
        MsgBox "Se Actualizo Satisfactoriamente  ", vbInformation
    End If
    Screen.MousePointer = vbDefault
    IMant = 0
    If rscabecera.State = 1 Then
        rscabecera.Requery
    End If
    Call Cancelar
'Esta Forma es grabar y Actualizar
'    rscabecera.Filter = "cabcomprobmes=" & Month(DTPFechaComprobCab) & " and cabcomprobnumero='" & Trim$(xnumerocompro) & "' and " & _
'                    "subasientocodigo='" & Trim$(VPSubAsiento) & "' and  asientocodigo='" & VPAsiento & "'"
'    Call Modificar
'    rscabecera.Filter = 0

    Exit Sub
    'Validando Errores
ErrorGrabar:
    Select Case varnerror
        Case 1
            MsgBox "No se Genero Correctamente el numero del Comprobante" & Chr(13) & VGErrorString, vbExclamation
        Case 2, 3, 4, 5, 6
            VGGeneral.RollbackTrans
            MsgBox "Hubo Errores al Grabar" & Chr(13) & VGErrorString, vbExclamation
            Call Cancelar
        Case Else
            MsgBox "Errores Desconocidos " & Chr(13) & err.Description
    End Select
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume Next
End Sub
Public Sub Modificar()
    IMant = 2
    Call Mostrar
End Sub
Public Sub eliminar()

    If MsgBox("Esta Seguro que desea Eliminar este Comprobante", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    VGGeneral.BeginTrans
    Screen.MousePointer = vbHourglass
    Call ClsMM1.GrabarCabecera(3, Trim$(lbNumComprobCab.Caption))
    Screen.MousePointer = vbHourglass
    VGGeneral.CommitTrans
    If rscabecera.State = 1 Then
        rscabecera.Requery
    End If
      
    MsgBox "El Registro se Elimino Correctamente"
    Call Cancelar
    Screen.MousePointer = vbDefault
    VlUltAccion = 3
End Sub
Public Sub Cancelar()
Set VGvardllgen = New dllgeneral.dll_general
    If SSTabMant.Tab = 1 Then
        Call VGvardllgen.ActivaTab(0, 1, SSTabMant)
        VlUltAccion = 5
        Set rsmantenimiento = Nothing
    End If
    
End Sub
Public Sub AñadirDetalle()
    Set ClsMM1 = New ClsMantMov1
    If rsmantenimiento.RecordCount > 0 Then
        If Not ClsMM1.ValidarGrabarDetalle Then Exit Sub
    End If
    Call HabilitarDetalle(True, FramDetalle)
    Call ClsMM1.AñadiralDetalle(rsmantenimiento)
    lbnregdetalle.Caption = Format(rsmantenimiento.RecordCount, "0 ")
    CtrAyu_Opera.SetFocus
    Call HabilitarSegunCuenta(False, False, False)
End Sub
Public Sub EliminarDetalle()
    Dim num As Integer
    Dim reg As Long
    On Error Resume Next
    If rsmantenimiento.State = 0 Then Exit Sub
    If rsmantenimiento.RecordCount = 0 Then Exit Sub
    If IMant = 1 Then
        If VerificaItemPlant(rsmantenimiento!NumPlantilla) Then
            MsgBox "Este item corresponde a una plantilla " & Chr(13) & _
                   "por lo tanto no se podra Eliminar ", vbExclamation
            Exit Sub
        End If
    End If
    If rsmantenimiento!detcomprobauto Then
        MsgBox "No se puede eliminar los asientos automaticos", vbExclamation
        Exit Sub
    End If
    Set ClsMM1 = New ClsMantMov1
    If rsmantenimiento.RecordCount = 1 Then
        ClsMM1.Limpia
    End If
    num = CInt(rsmantenimiento!detcomprobitem)
    reg = rsmantenimiento.RecordCount
    rsmantenimiento.Delete
    If num = reg Then
        rsmantenimiento.MoveNext
      Else
        Call ClsMM1.ActualizaNumItems(rsmantenimiento, num)
    End If
    Call ClsMM1.VerfiSiEsPlantilla(rsmantenimiento)
End Sub
Private Function VerificaItemPlant(dato As Integer) As Boolean
    Dim RSAUX As ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    VerificaItemPlant = False
    Dim sqlcad As String
    sqlcad = "Select * from ct_plantillaasiento where subasientocodigo='" & Trim$(VPSubAsiento) & "' and " & _
             "asientocodigo='" & Trim$(VPAsiento) & "' and plantillaasientocorrela=" & dato & " and (plantillaasientoctaajuste =0 and plantillaasientoinafecto = 0)"
    RSAUX.Open sqlcad, VGCNx, adOpenKeyset, adLockReadOnly
    If RSAUX.RecordCount > 0 Then
        VerificaItemPlant = True
    End If
End Function

Private Sub TxValor_Change()
    If Not VGflaglimpia Then Exit Sub
    If VGMoverRegistro Then Exit Sub
    Set ClsMM1 = New ClsMantMov1
    Call ClsMM1.ActualizarDetalle(rsmantenimiento, Montos)
    Call CalcularTotales(rsmantenimiento)
End Sub
Public Sub imprimir()
    If rscabecera Is Nothing Then Exit Sub
    If rscabecera.State = 0 Then Exit Sub
    If rscabecera.RecordCount = 0 Then Exit Sub
    Call ImprimirComprob(rscabecera!cabcomprobnumero, rscabecera!asientocodigo, rscabecera!subasientocodigo)
End Sub
Private Sub ImprimirComprob(Ncomprob As String, Asiento As String, SubAsiento As String)
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(0) As Variant, arrparm(7) As Variant
Screen.MousePointer = 11
    arrparm(0) = Trim$(VGParamSistem.BDEmpresa)
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = Trim$(VGParamSistem.Anoproceso)
    arrparm(3) = Trim$(VGParamSistem.Mesproceso)
    arrparm(4) = Trim$(Ncomprob)
    arrparm(5) = Trim$(Asiento)
    arrparm(6) = Trim$(SubAsiento)
    Call ImpresionRptProc("rptVoucherComprob.rpt", arrform, arrparm)
Screen.MousePointer = 1
End Sub

Public Sub PMant(Index As Integer)
    Select Case Index
        Case 1
            Call Xnuevo
        Case 2
            Call Grabar
        Case 3 'Eliminar
            Call eliminar
        Case 4 'Modificar
            Call Modificar
        Case 5
            Call Cancelar
        Case 6
            Call AñadirDetalle
        Case 7
            Call EliminarDetalle
        Case 8
            Call imprimir
    End Select
    Call PBoton(VlUltAccion)
End Sub
Public Sub Pavant(Index As Integer)
   If rsmantenimiento.RecordCount = 0 Then
        MsgBox "No puede utilizar esta funcion porque no " & Chr(13) & _
               "Existe ningun registro en el comprobante", vbInformation
   End If
   Select Case Index
    Case 1
        Me.TxMonto.SetFocus
    Case 2
        Me.CtrAyu_Opera.SetFocus
    End Select
End Sub
Private Sub PBoton(Index As Integer)
    Select Case Index
        Case 0, 5
            Call Botones(MDIPrincipal.ToolComprob, True, False, False, True, False, False, False)
        Case 1 'nuevo
            Call Botones(MDIPrincipal.ToolComprob, False, True, False, False, True, True, True)
        Case 3 'Eliminar
            Call Botones(MDIPrincipal.ToolComprob, True, False, False, True, False, False, False)
        Case 4 'Modificar
            Call Botones(MDIPrincipal.ToolComprob, False, True, True, False, True, True, True)
    End Select
End Sub
Private Sub MuestraDocPend()
Dim RsPend As ADODB.Recordset
Dim CamposPend As ADODB.Fields
On Error GoTo DocPend
    If Not (CtrAyu_Opera.xclave <> "00" And CtrAyu_Opera.xclave <> "01") Then
        MsgBox "Aqui se ven documentos pendientes a un analitico " & Chr(13) & _
               "cuando el tipo de operacion es diferente de una provisión", vbInformation
        Exit Sub
    End If
    If CtrAyu_Cuenta.xclave <> "00" And CtrAyu_Cuenta.xclave = "" Then
        MsgBox "Tiene que seleccionar un cuenta ", vbInformation
        Exit Sub
    End If
    If CtrAyu_Analitico.xclave <> "00" And CtrAyu_Analitico.xclave = "" Then
        MsgBox "Tiene que seleccionar un Analitico ", vbInformation
        Exit Sub
    End If
    Screen.MousePointer = 11
    Set RsPend = New ADODB.Recordset
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_MuestraPend_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@Ano") = VGParamSistem.Anoproceso
        .Parameters("@cuenta") = CtrAyu_Cuenta.xclave
        .Parameters("@analitico") = RTrim$(CtrAyu_Analitico.xclave)
        Set RsPend = .Execute
    End With
    Screen.MousePointer = 1
    If RsPend.RecordCount = 0 Then
        MsgBox "No existe es registro de Documentos Pendientes"
        Exit Sub
    End If
    Call FrmDocPend.RecibeRs(RsPend, CamposPend)
    If Not (CamposPend Is Nothing) Then
        TxNdoc.Text = "": TxSerie.Text = ""
        CtrAyu_TipDoc.xclave = Trim$(CamposPend("DocumentoCodigo").Value): CtrAyu_TipDoc.Ejecutar
        TxSerie.Text = Mid$(Trim$(CamposPend("ctacteanaliticonumdocumento").Value), 1, TxSerie.MaxLength)
        TxNdoc.Text = Mid$(Trim$(CamposPend("ctacteanaliticonumdocumento").Value), TxSerie.MaxLength + 1, Len(Trim$(CamposPend("ctacteanaliticonumdocumento").Value - TxSerie.MaxLength)))
        Dtp_FechaDoc.Value = CamposPend("ctacteanaliticofechadoc").Value
        Call Dtp_FechaDoc_Change
  '      CmbTcambio.ListIndex = CInt(CamposPend("detcomprobformacambio").Value) - 1
  '      VGValorCambio = CamposPend("detcomprobtipocambio").Value
  '     lb_vcambio.Caption = Format(VGValorCambio, "#.000 ")
        CtrAyu_Moneda.xclave = CamposPend("monedacodigo").Value: CtrAyu_Moneda.Ejecutar
        TxMonto.valor = CamposPend("Saldo").Value
        TxMonto.Text = CamposPend("Saldo").Value
        
    End If
    Exit Sub
DocPend:
    Screen.MousePointer = 1
    MsgBox "No se Puede mostrar los Documentos Pendientes " & Chr(13) & _
           err.Description, vbExclamation
End Sub

Property Let CodComprob(valor As String)
  m_CodComprob = valor
End Property

Private Function VerififcaAdicionCargo() As Boolean
    Dim rsX As ADODB.Recordset
    Dim sqlcad As String
'FIXIT: Declare 'AsientoCargo' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
    Dim AsientoCargo, CuentaCargo As String
    
    Set rsX = New ADODB.Recordset
    VerififcaAdicionCargo = False

    sqlcad = "Select cuentaadicionacargo from ct_cuenta where cuentacodigo='" & CtrAyu_Cuenta.xclave & "' "
    rsX.Open sqlcad, VGCNx, adOpenKeyset, adLockReadOnly
    If rsX.RecordCount > 0 Then
        CuentaCargo = rsX("cuentaadicionacargo")
    End If
End Function

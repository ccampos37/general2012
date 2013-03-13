VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmConsultaComprobantes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Comprobantes"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Consulta por Libro y/o Glosa"
      Height          =   645
      Left            =   0
      TabIndex        =   19
      Top             =   675
      Width           =   11055
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   2
         Left            =   5745
         TabIndex        =   3
         Top             =   240
         Width           =   5250
         _ExtentX        =   9260
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
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Libro 
         Height          =   315
         Left            =   795
         TabIndex        =   2
         Top             =   240
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   556
         XcodMaxLongitud =   0
         xcodwith        =   300
         NomTabla        =   "ct_libro"
         ListaCampos     =   "librocodigo(1),librodescripcion(1)"
         XcodCampo       =   "librocodigo"
         XListCampo      =   "librodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "librocodigo,librodescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label Label3 
         Caption         =   "Libro"
         Height          =   285
         Left            =   180
         TabIndex        =   21
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Glosa"
         Height          =   240
         Left            =   4755
         TabIndex        =   20
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consulta por Números de Control"
      Height          =   615
      Left            =   15
      TabIndex        =   16
      Top             =   1410
      Width           =   10095
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   0
         Left            =   1530
         TabIndex        =   4
         Top             =   225
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
         Left            =   4365
         TabIndex        =   5
         Top             =   180
         Width           =   2520
         _ExtentX        =   4445
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
         NoCaracteres    =   "0123456789"
         MarcarTextoAlEnfoque=   -1  'True
         NoRangoCadena   =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   7875
         TabIndex        =   23
         Top             =   180
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   16318465
         CurrentDate     =   37541
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha:"
         Height          =   270
         Left            =   7005
         TabIndex        =   22
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Comprobante"
         Height          =   255
         Left            =   105
         TabIndex        =   18
         Top             =   285
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Libro"
         Height          =   255
         Left            =   3525
         TabIndex        =   17
         Top             =   255
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consulta por Asiento y/o SubAsiento"
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   15
      Width           =   11055
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_SubAsiento 
         Height          =   360
         Left            =   5775
         TabIndex        =   1
         Top             =   225
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   635
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
         Height          =   330
         Left            =   810
         TabIndex        =   0
         Top             =   240
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   582
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "ct_asiento"
         TituloAyuda     =   "Busqueda de Asiento"
         ListaCampos     =   "asientocodigo(1), asientodescripcion(1),flaggrabado(2),controlnref(2),nemotecref(1),librocodigo(1)"
         XcodCampo       =   "asientocodigo"
         XListCampo      =   "asientodescripcion"
         ListaCamposDescrip=   "Codigo,Descripción,OperGraba"
         ListaCamposText =   "asientocodigo,asientodescripcion,flaggrabado,controlnref,nemotecref,librocodigo"
         Requerido       =   0   'False
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Asiento :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   75
         TabIndex        =   13
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbSubAsiento 
         BackStyle       =   0  'Transparent
         Caption         =   "Subasiento :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4710
         TabIndex        =   12
         Top             =   285
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "..."
      Height          =   360
      Left            =   10275
      TabIndex        =   8
      Top             =   1590
      Width           =   675
   End
   Begin TrueOleDBGrid70.TDBGrid TDBG_ConsultaDetalle 
      Height          =   2280
      Left            =   45
      TabIndex        =   6
      Top             =   5145
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
      Columns(6).DataWidth=   13
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Moneda"
      Columns(7).DataField=   "monedacodigo"
      Columns(7).DataWidth=   2
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "ID"
      Columns(8).DataField=   "indicador"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Monto Soles"
      Columns(9).DataField=   "montosol"
      Columns(9).NumberFormat=   "###,###,###,###.00"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Monto Dolares"
      Columns(10).DataField=   "montouss"
      Columns(10).NumberFormat=   "###,###,###,###.00"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   4
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Auto"
      Columns(11).DataField=   "detcomprobauto"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
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
      Splits(0)._ColumnProps(40)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(44)=   "Column(8).Width=609"
      Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=529"
      Splits(0)._ColumnProps(47)=   "Column(8).AllowSizing=0"
      Splits(0)._ColumnProps(48)=   "Column(8)._ColStyle=260"
      Splits(0)._ColumnProps(49)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(50)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(51)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(53)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(54)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(55)=   "Column(10).Width=2752"
      Splits(0)._ColumnProps(56)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(10)._WidthInPix=2672"
      Splits(0)._ColumnProps(58)=   "Column(10).AllowSizing=0"
      Splits(0)._ColumnProps(59)=   "Column(10)._ColStyle=514"
      Splits(0)._ColumnProps(60)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(61)=   "Column(11).Width=1402"
      Splits(0)._ColumnProps(62)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(11)._WidthInPix=1323"
      Splits(0)._ColumnProps(64)=   "Column(11).AllowSizing=0"
      Splits(0)._ColumnProps(65)=   "Column(11)._ColStyle=513"
      Splits(0)._ColumnProps(66)=   "Column(11).Order=12"
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
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=20,.parent=47"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=17,.parent=48"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=18,.parent=49"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=19,.parent=51"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=138,.parent=47"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=135,.parent=48,.alignment=0"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=136,.parent=49"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=137,.parent=51"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=154,.parent=47,.alignment=1,.bgcolor=&H80000014&"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=151,.parent=48,.alignment=2"
      _StyleDefs(74)  =   ":id=151,.bgcolor=&H8000000F&"
      _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=152,.parent=49"
      _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=153,.parent=51,.bgcolor=&H80000018&"
      _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=158,.parent=47,.alignment=1,.bgcolor=&H80000014&"
      _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=155,.parent=48,.alignment=2"
      _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=156,.parent=49"
      _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=157,.parent=51,.bgcolor=&H80000018&"
      _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=162,.parent=47,.alignment=2"
      _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=159,.parent=48,.alignment=2"
      _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=160,.parent=49"
      _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=161,.parent=51"
      _StyleDefs(85)  =   "Named:id=33:Normal"
      _StyleDefs(86)  =   ":id=33,.parent=0"
      _StyleDefs(87)  =   "Named:id=34:Heading"
      _StyleDefs(88)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(89)  =   ":id=34,.wraptext=-1"
      _StyleDefs(90)  =   "Named:id=35:Footing"
      _StyleDefs(91)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(92)  =   "Named:id=36:Selected"
      _StyleDefs(93)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(94)  =   "Named:id=37:Caption"
      _StyleDefs(95)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(96)  =   "Named:id=38:HighlightRow"
      _StyleDefs(97)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(98)  =   "Named:id=39:EvenRow"
      _StyleDefs(99)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(100) =   "Named:id=40:OddRow"
      _StyleDefs(101) =   ":id=40,.parent=33"
      _StyleDefs(102) =   "Named:id=41:RecordSelector"
      _StyleDefs(103) =   ":id=41,.parent=34"
      _StyleDefs(104) =   "Named:id=42:FilterBar"
      _StyleDefs(105) =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid70.TDBGrid TDBG_ConsultaCabecera 
      Height          =   2610
      Left            =   15
      TabIndex        =   7
      Top             =   2175
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   4604
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
      Columns(1).Caption=   "Nª Auxiliar"
      Columns(1).DataField=   "cabcomprobnlibro"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Fecha Cont."
      Columns(2).DataField=   "cabcomprobfeccontable"
      Columns(2).NumberFormat=   "Short Date"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Total Debe S/."
      Columns(3).DataField=   "cabcomprobtotdebe"
      Columns(3).NumberFormat=   "###,###,###,###.00"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Total Debe S/."
      Columns(4).DataField=   "cabcomprobtothaber"
      Columns(4).NumberFormat=   "###,###,###,###.00"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Total Deb US$"
      Columns(5).DataField=   "cabcomprobtotussdebe"
      Columns(5).NumberFormat=   "###,###,###,###.00"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Total Haber US $"
      Columns(6).DataField=   "cabcomprobtotusshaber"
      Columns(6).NumberFormat=   "###,###,###,###.00"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Glosa"
      Columns(7).DataField=   "cabcomprobglosa"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Estado"
      Columns(8).DataField=   "estcomprobcodigo"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2805"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2725"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2223"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2143"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2805"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2725"
      Splits(0)._ColumnProps(16)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(4).Width=2805"
      Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2725"
      Splits(0)._ColumnProps(21)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(5).Width=2805"
      Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=2725"
      Splits(0)._ColumnProps(26)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(28)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(31)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(33)=   "Column(7).Width=4577"
      Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=4498"
      Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(37)=   "Column(8).Width=979"
      Splits(0)._ColumnProps(38)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(8)._WidthInPix=900"
      Splits(0)._ColumnProps(40)=   "Column(8).Order=9"
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
      Caption         =   "Resultados de la Consulta"
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
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1,.bgcolor=&H80000018&"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1,.bgcolor=&H80000018&"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1,.bgcolor=&H80000018&"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1,.bgcolor=&H80000018&"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=70,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=62,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
      _StyleDefs(72)  =   "Named:id=33:Normal"
      _StyleDefs(73)  =   ":id=33,.parent=0"
      _StyleDefs(74)  =   "Named:id=34:Heading"
      _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(76)  =   ":id=34,.wraptext=-1"
      _StyleDefs(77)  =   "Named:id=35:Footing"
      _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(79)  =   "Named:id=36:Selected"
      _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=37:Caption"
      _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(83)  =   "Named:id=38:HighlightRow"
      _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=39:EvenRow"
      _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(87)  =   "Named:id=40:OddRow"
      _StyleDefs(88)  =   ":id=40,.parent=33"
      _StyleDefs(89)  =   "Named:id=41:RecordSelector"
      _StyleDefs(90)  =   ":id=41,.parent=34"
      _StyleDefs(91)  =   "Named:id=42:FilterBar"
      _StyleDefs(92)  =   ":id=42,.parent=33"
   End
   Begin VB.Label lblNroReg_Det 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9990
      TabIndex        =   15
      Top             =   7455
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Nº Registros"
      Height          =   240
      Left            =   8820
      TabIndex        =   14
      Top             =   7440
      Width           =   1110
   End
   Begin VB.Label lblNro_Reg 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9855
      TabIndex        =   10
      Top             =   4800
      Width           =   1200
   End
   Begin VB.Label Label5 
      Caption         =   "Nº Registros:"
      Height          =   225
      Left            =   8745
      TabIndex        =   9
      Top             =   4800
      Width           =   1050
   End
End
Attribute VB_Name = "frmConsultaComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rscabecera As New ADODB.Recordset
Dim rsdetalle As New ADODB.Recordset

Private Sub Form_Load()
    CtrAyu_Asiento.conexion VGCNx
    CtrAyu_SubAsiento.conexion VGCNx
    CtrAyu_Libro.conexion VGCNx
    TDBG_ConsultaDetalle.FetchRowStyle = True
    Me.Width = 11280
    Me.Height = 8115
    DTPicker1.Value = Null
End Sub

Private Sub EjecutarConsulta()
Dim cad As String
Dim sqlcad As String, xasiento As String, xsubasiento As String
    Set rscabecera = New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    xasiento = Trim$(CtrAyu_Asiento.xclave): xsubasiento = Trim$(CtrAyu_SubAsiento.xclave)
    sqlcad = "SELECT * FROM " & VGParamSistem.TablaCabcomprob & " "
    sqlcad = sqlcad & "WHERE cabcomprobmes='" & CInt(VGParamSistem.Mesproceso) & "' "
    
    If txt(0).Text <> Empty Then sqlcad = sqlcad & "AND cast(Right(rtrim$(cabcomprobnumero),5) as int)=" & CLng(Right(Trim$(txt(0).Text), 5)) & " "
    If txt(1).Text <> Empty Then sqlcad = sqlcad & "AND cast(Right(rtrim$(cabcomprobnlibro),5) as int)=" & CLng(Right(Trim$(txt(1).Text), 5)) & " "
    If txt(2).Text <> Empty Then sqlcad = sqlcad & "AND cabcomprobglosa like '%" & txt(2).Text & "%' "
    If Not IsNull(DTPicker1.Value) Then sqlcad = sqlcad & "AND cabcomprobfeccontable='" & CDate(DTPicker1.Value) & "'"
    
    If xasiento <> Empty Then sqlcad = sqlcad & "AND asientocodigo='" & xasiento & "' "
    If xsubasiento <> Empty Then sqlcad = sqlcad & "AND subasientocodigo='" & Trim$(xsubasiento) & "' "
    If Trim$(CtrAyu_Libro.xclave) <> Empty Then sqlcad = sqlcad & "AND detcomprobnlibro='" & Trim$(CtrAyu_Libro.xclave) & "' "
    
    Set rscabecera = VGCNx.Execute(sqlcad)
    
    If rscabecera.RecordCount > 0 Then
        lblNro_Reg.Caption = Format(rscabecera.RecordCount, "0 ")
        TDBG_ConsultaCabecera.SetFocus
    Else
        lblNro_Reg.Caption = Format(0, "0 ")
    End If
    Set TDBG_ConsultaCabecera.DataSource = rscabecera
End Sub

Private Sub EjecutarConsultaDetalle(xParam As String)
  Dim sqlcad As String
  sqlcad = "SELECT plantillaasientoinafecto,detcomprobitem,operacioncodigo,analiticocodigo,cuentacodigo,"
  sqlcad = sqlcad & "documentocodigo,detcomprobnumdocumento,indicador= case when detcomprobdebe>0 then 'D' else 'H' end,"
  sqlcad = sqlcad & "montosol=case when detcomprobdebe>0 then detcomprobdebe else detcomprobhaber end,"
  sqlcad = sqlcad & "montouss=case when detcomprobussdebe>0 then detcomprobussdebe else detcomprobusshaber end,detcomprobauto,monedacodigo "
  sqlcad = sqlcad & "FROM " & VGParamSistem.TablaDetcomprob & " WHERE cabcomprobnumero='" & xParam & " '"
  
  Set rsdetalle = New ADODB.Recordset
  Set rsdetalle = VGCNx.Execute(sqlcad)
  Set TDBG_ConsultaDetalle.DataSource = rsdetalle
  lblNroReg_Det.Caption = rsdetalle.RecordCount

End Sub

Private Sub cmdConsultar_Click()
  Call EjecutarConsulta
End Sub

Private Sub TDBG_ConsultaCabecera_DblClick()
   frmantcomprobantes.CodComprob = TDBG_ConsultaCabecera.Columns(0).Value
   frmantcomprobantes.Show
End Sub

Private Sub TDBG_ConsultaCabecera_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
        frmantcomprobantes.CodComprob = TDBG_ConsultaCabecera.Columns(0).Value
        frmantcomprobantes.Show
   End If
End Sub

Private Sub TDBG_ConsultaCabecera_HeadClick(ByVal ColIndex As Integer)
 With rscabecera
    If .Sort = Empty Then
        .Sort = TDBG_ConsultaCabecera.Columns.Item(ColIndex).DataField & " asc"
    ElseIf Right(.Sort, 3) = "asc" Then
        .Sort = TDBG_ConsultaCabecera.Columns.Item(ColIndex).DataField & " desc"
    ElseIf Right(.Sort, 4) = "desc" Then
        .Sort = TDBG_ConsultaCabecera.Columns.Item(ColIndex).DataField & " asc"
    End If
    TDBG_ConsultaCabecera.Refresh
 End With
End Sub

'FIXIT: Declare 'LastRow' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Private Sub TDBG_ConsultaCabecera_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If IsNull(TDBG_ConsultaCabecera.Columns(0).Value) Then Exit Sub
    Call EjecutarConsultaDetalle(TDBG_ConsultaCabecera.Columns(0).Value)
End Sub

'FIXIT: Declare 'Bookmark' con un tipo de datos de enlace en tiempo de compilación         FixIT90210ae-R1672-R1B8ZE
Private Sub TDBG_ConsultaDetalle_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    Dim rsX As ADODB.Recordset
    Set rsX = rsdetalle.Clone(adLockReadOnly)
    rsX.Bookmark = Bookmark
    If rsX!detcomprobauto Then
        RowStyle.BackColor = RGB(185, 251, 236)
    End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 1 Then
        If KeyCode = 13 Then Call EjecutarConsulta
    End If
End Sub

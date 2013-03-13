VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form CstVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Ventas"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   12150
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Height          =   1245
      Index           =   1
      Left            =   270
      TabIndex        =   8
      Top             =   180
      Width           =   11625
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   180
         TabIndex        =   15
         Top             =   120
         Width           =   2655
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   330
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   540
            Width           =   2025
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "TIPO DOCUMENTO"
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
            Index           =   1
            Left            =   390
            TabIndex        =   16
            Top             =   240
            Width           =   1875
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   6540
         TabIndex        =   13
         Top             =   120
         Width           =   2625
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   300
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   510
            Width           =   1995
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "MODO-VENTA"
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
            Index           =   1
            Left            =   330
            TabIndex        =   14
            Top             =   240
            Width           =   1905
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1005
         Left            =   2910
         TabIndex        =   9
         Top             =   120
         Width           =   3555
         Begin MSMask.MaskEdBox MBox1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "d/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   630
            TabIndex        =   1
            Top             =   540
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "d/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   2250
            TabIndex        =   2
            Top             =   540
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Del"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   210
            TabIndex        =   12
            Top             =   600
            Width           =   465
         End
         Begin VB.Label Label2 
            Caption         =   "Al"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1950
            TabIndex        =   11
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "FECHAS A CONSULTAR"
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
            Index           =   0
            Left            =   210
            TabIndex        =   10
            Top             =   210
            Width           =   3195
         End
      End
      Begin VB.CommandButton cBusca 
         Caption         =   "&Buscar"
         Height          =   405
         Left            =   9330
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   7365
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   714
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
   Begin VB.Frame Frame1 
      Height          =   5805
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   11685
      Begin VB.Frame Frame5 
         Height          =   585
         Index           =   0
         Left            =   9270
         TabIndex        =   17
         Top             =   5130
         Width           =   2265
         Begin VB.TextBox TReg 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1170
            TabIndex        =   18
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Total Reg."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   210
            TabIndex        =   19
            Top             =   240
            Width           =   1035
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   4815
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   8493
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
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=10,.bold=0,.fontsize=825,.italic=0"
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
   End
End
Attribute VB_Name = "CstVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim adll As New dllgeneral.dll_general
Dim nlongi(1) As Integer


Private Sub cBusca_Click()
    Dim nsql As String
    Dim forden As String
   
   
    If Len(Trim(MBox1(0).ClipText)) = 0 Or Len(Trim(MBox1(1).ClipText)) = 0 Then
        MsgBox "Falta Ingresar el Rango de Fecha a Consultar..!!!", vbInformation, MsgTitle
        MBox1(0).SetFocus
        Exit Sub
    End If
    If Len(Trim(MBox1(0).ClipText)) < 8 Or Len(Trim(MBox1(1).ClipText)) < 8 Then
        MsgBox "Fechas No Validas..!!!", vbInformation, MsgTitle
        MBox1(0).SetFocus
        Exit Sub
    End If
    
    If Combo2.ListCount = 0 Then
        MsgBox "Falta Ingresar Modo de Venta..!!!", vbInformation, MsgTitle
        Combo2.SetFocus
        Exit Sub
    End If
    
    Set TDBGrid1.DataSource = Nothing
    TDBGrid1.ClearFields
    TDBGrid1.Refresh
    DoEvents
    If adll.ComboDato(Combo2.Text) = g_tipobol Then
        nsql = "CASE pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofechafact as Fecha,pedidonrofact as Boleta,pedidonumero as Pedido,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
        forden = "pedidofechafact,pedidonrofact"
    ElseIf adll.ComboDato(Combo2.Text) = g_tipofac Then
        nsql = "CASE pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofechafact as Fecha,pedidonrofact as Factura,pedidonumero as Pedido,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
        forden = "pedidofechafact,pedidonrofact"
    ElseIf adll.ComboDato(Combo2.Text) = g_tipoped Then
        nsql = "CASE pedidoestado WHEN '1' THEN '*' ELSE '' END,pedidofecha as Fecha,pedidonumero as Pedido,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
        forden = "pedidofecha,pedidonumero"
    ElseIf adll.ComboDato(Combo2.Text) = g_Todos Then
        nsql = "CASE pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofechafact as Fecha,pedidotipofac,pedidonrofact,pedidonumero as Pedido,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
        forden = "pedidofecha"
    Else
       Exit Sub
    End If
    If Combo1.ListCount > 0 Then
       If adll.ComboDato(Combo2.Text) = g_tipoped Then
           Call adll.ListarEnTDBGRID(VGCNx, "vt_pedido", TDBGrid1, nsql, forden, nlongi, "pedidofecha>='" & MBox1(0) & "' and pedidofecha<='" & MBox1(1) & "' and modovtacodigo='" & adll.ComboDato(Combo1.Text) & "'")
       ElseIf adll.ComboDato(Combo2.Text) = g_tipobol Then
           Call adll.ListarEnTDBGRID(VGCNx, "vt_pedido", TDBGrid1, nsql, forden, nlongi, "pedidofechafact>='" & MBox1(0) & "' and pedidofechafact<='" & MBox1(1) & "' and modovtacodigo='" & adll.ComboDato(Combo1.Text) & "' and pedidotipofac='" & g_tipobol & "'")
       ElseIf adll.ComboDato(Combo2.Text) = g_tipofac Then
           Call adll.ListarEnTDBGRID(VGCNx, "vt_pedido", TDBGrid1, nsql, forden, nlongi, "pedidofechafact>='" & MBox1(0) & "' and pedidofechafact<='" & MBox1(1) & "' and modovtacodigo='" & adll.ComboDato(Combo1.Text) & "' and pedidotipofac='" & g_tipofac & "'")
       ElseIf adll.ComboDato(Combo2.Text) = g_Todos Then
           Call adll.ListarEnTDBGRID(VGCNx, "vt_pedido", TDBGrid1, nsql, forden, nlongi, "pedidofecha>='" & MBox1(0) & "' and pedidofecha<='" & MBox1(1) & "' and modovtacodigo='" & adll.ComboDato(Combo1.Text) & "'")
       End If
    Else
       If adll.ComboDato(Combo2.Text) = g_tipoped Then
          Call adll.ListarEnTDBGRID(VGCNx, "vt_pedido", TDBGrid1, nsql, forden, nlongi, "pedidofecha>='" & MBox1(0) & "' and pedidofecha<='" & MBox1(1) & "'")
       ElseIf adll.ComboDato(Combo2.Text) = g_tipobol Then
          Call adll.ListarEnTDBGRID(VGCNx, "vt_pedido", TDBGrid1, nsql, forden, nlongi, "pedidofechafact>='" & MBox1(0) & "' and pedidofechafact<='" & MBox1(1) & "' and pedidotipofac='" & g_tipobol & "'")
       ElseIf adll.ComboDato(Combo2.Text) = g_tipofac Then
          Call adll.ListarEnTDBGRID(VGCNx, "vt_pedido", TDBGrid1, nsql, forden, nlongi, "pedidofechafact>='" & MBox1(0) & "' and pedidofechafact<='" & MBox1(1) & "' and pedidotipofac='" & g_tipofac & "'")
       ElseIf adll.ComboDato(Combo2.Text) = g_Todos Then
          Call adll.ListarEnTDBGRID(VGCNx, "vt_pedido", TDBGrid1, nsql, forden, nlongi, "pedidofecha>='" & MBox1(0) & "' and pedidofecha<='" & MBox1(1) & "'")
       End If
    End If
    Call ConfiguraGrid
    If TDBGrid1.ApproxCount > 0 Then
       TDBGrid1.SetFocus
    Else
      Combo2.SetFocus
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   Call Seguir(Combo1, KeyAscii)
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   Call Seguir(Combo2, KeyAscii)
      
End Sub

Private Sub Form_Load()
   Dim nsql As String
   Dim forden As String
   
   MostrarFormVentas Me, "I"
   nlongi(1) = 0
   nsql = "CASE pedidoestado WHEN '1' THEN '*' ELSE '' END,pedidofecha as Fecha,pedidonrofact as Factura,pedidonumero as Documento,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
   forden = "pedidonumero"
   Call CargarTipoVentas(Combo2, 6)
   Combo2.AddItem g_Todos & "-Todos"    'Todos
   
   Call adll.llenacombo(Combo1, "select modovtacodigo,modovtadescripcion from vt_modoventa", VGCNx)
   Call adll.ListarEnTDBGRID(VGCNx, "vt_pedido", TDBGrid1, nsql, forden, nlongi, "pedidonumero='*'")
   Call ConfiguraGrid
End Sub


Public Function ConfiguraGrid()
   TReg.Text = Format(TDBGrid1.ApproxCount, "#####0")
   With TDBGrid1
       .Columns(0).Caption = "Sit"
       .Columns(0).Width = 600
       .Columns(0).HeadAlignment = dbgCenter
       .Columns(1).Caption = "Fecha"
       .Columns(1).Width = 1000
       .Columns(1).HeadAlignment = dbgCenter
       .Columns(1).NumberFormat = "dd/mm/yyyy"
       If adll.ComboDato(Combo2.Text) = g_tipobol Or adll.ComboDato(Combo2.Text) = g_tipofac Or adll.ComboDato(Combo2.Text) = g_Todos Then
             If adll.ComboDato(Combo2.Text) = g_Todos Then
                .Columns(2).Caption = "Factura"
                .Columns(2).HeadAlignment = dbgCenter
                .Columns(2).Width = 1100
                .Columns(3).Caption = "Boleta"
                .Columns(3).HeadAlignment = dbgCenter
                .Columns(3).Width = 1100
                .Columns(4).Caption = "No Pedido"
                .Columns(4).Width = 1100
                .Columns(4).HeadAlignment = dbgCenter
                .Columns(5).Caption = "Codigo"
                .Columns(5).Width = 1100
                .Columns(5).HeadAlignment = dbgCenter
                .Columns(6).Caption = "Descripcion"
                .Columns(6).Width = 3800
                .Columns(6).HeadAlignment = dbgCenter
                .Columns(7).Caption = "Total"
                .Columns(7).NumberFormat = "##,###,##0.00"
                .Columns(7).Width = 1200
                .Columns(7).HeadAlignment = dbgCenter
             Else
                .Columns(2).HeadAlignment = dbgCenter
                .Columns(2).Width = 1100
                .Columns(3).Caption = "No Pedido"
                .Columns(3).Width = 1100
                .Columns(3).HeadAlignment = dbgCenter
                .Columns(4).Caption = "Codigo"
                .Columns(4).Width = 1100
                .Columns(4).HeadAlignment = dbgCenter
                .Columns(5).Caption = "Descripcion"
                .Columns(5).Width = 4500
                .Columns(5).HeadAlignment = dbgCenter
                .Columns(6).Caption = "Total"
                .Columns(6).NumberFormat = "##,###,##0.00"
                .Columns(6).Width = 1200
                .Columns(6).HeadAlignment = dbgCenter
             End If
       Else                                             'If adll.ComboDato(Combo2.Text) = g_tipoped Then
            .Columns(2).Caption = "No Pedido"
            .Columns(2).HeadAlignment = dbgCenter
            .Columns(2).Width = 1100
            .Columns(3).Caption = "Codigo"
            .Columns(3).Width = 1100
            .Columns(3).HeadAlignment = dbgCenter
            .Columns(4).Caption = "Descripcion"
            .Columns(4).Width = 5800
            .Columns(4).HeadAlignment = dbgCenter
            .Columns(5).Caption = "Total"
            .Columns(5).NumberFormat = "##,###,##0.00"
            .Columns(5).Width = 1200
            .Columns(5).HeadAlignment = dbgCenter
       End If
       .Refresh
   End With
   
End Function

Private Sub MBox1_KeyPress(Index As Integer, KeyAscii As Integer)
  If Not IsDate(MBox1(Index)) And KeyAscii = 13 Then
     MsgBox "Fecha no Valida...!!", vbInformation, MsgTitle
     MBox1(Index).SetFocus
     Exit Sub
  End If
  Call Seguir(MBox1(Index), KeyAscii)
End Sub

Private Sub TDBGrid1_Click()
  If TDBGrid1.ApproxCount > 0 Then
      TDBGrid1.SetFocus
  End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim ltipo As String
   Dim ldocu As String
   
   ltipo = g_tipoped
   If KeyCode = 13 And TDBGrid1.ApproxCount > 0 Then
       
       If adll.ComboDato(Combo2.Text) = g_tipobol Then
         ldocu = Escadena(TDBGrid1.Columns(3).Text)
       ElseIf adll.ComboDato(Combo2.Text) = g_tipofac Then
         ldocu = Escadena(TDBGrid1.Columns(3).Text)
       ElseIf adll.ComboDato(Combo2.Text) = g_tipoped Then
         ldocu = Escadena(TDBGrid1.Columns(2).Text)
       End If
       CstDetalleDocumento.Btipo = ltipo
       CstDetalleDocumento.BNumero = ldocu
       CstDetalleDocumento.Show
   End If

End Sub


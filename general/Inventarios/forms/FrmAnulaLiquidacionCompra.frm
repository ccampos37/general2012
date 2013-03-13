VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmAnulaLiquidacionCompra 
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fr1 
      Height          =   7395
      Left            =   210
      TabIndex        =   8
      Top             =   750
      Width           =   10035
      Begin VB.Frame Frame5 
         Height          =   585
         Index           =   0
         Left            =   7590
         TabIndex        =   10
         Top             =   6660
         Width           =   2265
         Begin VB.TextBox TReg 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1170
            TabIndex        =   11
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
            Left            =   150
            TabIndex        =   12
            Top             =   240
            Width           =   1035
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   5955
         Left            =   270
         TabIndex        =   9
         Top             =   630
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   10504
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
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "PEDIDOS EMITIDOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   9525
      End
   End
   Begin VB.Frame Fr2 
      Height          =   645
      Index           =   0
      Left            =   270
      TabIndex        =   1
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cBusca 
         BackColor       =   &H80000008&
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox aBusca 
         Height          =   285
         Index           =   0
         Left            =   2370
         MaxLength       =   3
         TabIndex        =   4
         Top             =   210
         Width           =   435
      End
      Begin VB.TextBox aBusca 
         Height          =   285
         Index           =   1
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   3
         Top             =   210
         Width           =   885
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "Documento"
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   195
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8265
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmAnulaLiquidacionCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim nLongicampo(1) As Integer

Private Sub aBusca_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim ldato As String
  If KeyAscii = 13 And Index = 1 Then
     TDBGrid1.ClearFields
     Set TDBGrid1.DataSource = Nothing
     aBusca(0) = Right("0000000000" & Trim(aBusca(0)), aBusca(0).MaxLength)
     aBusca(1) = Right("0000000000" & Trim(aBusca(1)), aBusca(1).MaxLength)
     If (Val(Trim(aBusca(1).text)) = 0 And Val(Trim(aBusca(1).text)) = 0) Then
       Listado
     Else
       If adll.ComboDato(Combo1.text) = g_tipoped Then
          Call adll.ListarEnTDBGRID(VGCNx, "vt_pedido", TDBGrid1, "CASE  pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofecha,pedidonumero,pedidotipofac,pedidonrofact,clienterazonsocial,pedidototneto", "pedidofecha", nLongicampo, "pedidonumero='" & Trim(aBusca(0) & aBusca(1)) & "'")
       Else
          Call adll.ListarEnTDBGRID(VGCNx, "vt_pedido", TDBGrid1, "CASE  pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofecha,pedidonumero,pedidotipofac,pedidonrofact,clienterazonsocial,pedidototneto", "pedidofecha", nLongicampo, "pedidonrofact='" & Trim(aBusca(0) & aBusca(1)) & "' and pedidotipofac='" & adll.ComboDato(Combo1.text) & "'")
       End If
     End If
     ConfiguraGrid
  
  ElseIf KeyAscii = 13 Then
      SendKeys "{tab}"
      Exit Sub
  End If
End Sub

Private Sub aBusca_LostFocus(Index As Integer)
  If Index = 0 Then
     aBusca(0) = Right("0000000000" & Trim(aBusca(0)), aBusca(0).MaxLength)
  Else
     aBusca(1) = Right("0000000000" & Trim(aBusca(1)), aBusca(1).MaxLength)
  End If
End Sub

Private Sub cBusca_Click()
  If Len(Trim(aBusca(0).text)) = 0 Or Len(Trim(aBusca(1).text)) = 0 Then
     Listado
  Else
     aBusca_KeyPress 1, 13
  End If
End Sub


Private Sub Form_Activate()
  Listado
End Sub

Private Sub Form_Load()
   
  MostrarForm Me, "C"
  g_tipofac = "04"
  g_tipoped = "PL"
  
  Call CargarTipo(Combo1, 6)
  nLongicampo(1) = 0
  Listado
  ConfiguraGrid

End Sub


Public Function Listado()
  On Error Resume Next
  Set TDBGrid1.DataSource = Nothing
  TDBGrid1.ClearFields
  TDBGrid1.Refresh
  Call adll.ListarEnTDBGRID(VGCNx, "al_liquidacionCompra", TDBGrid1, "CASE  pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofecha,pedidonumero,pedidotipofac,pedidonrofact,clienterazonsocial,pedidototneto", "pedidofecha", nLongicampo)
  End Function

Public Function ConfiguraGrid()
   On Error Resume Next
   With TDBGrid1
       .Columns(0).Caption = "Sit"
       .Columns(0).Width = 600
       .Columns(0).HeadAlignment = dbgCenter
       .Columns(1).Caption = "Fecha"
       .Columns(1).Width = 1200
       .Columns(1).HeadAlignment = dbgCenter
       .Columns(2).Caption = "No.Pedido"
       .Columns(2).Width = 1300
       .Columns(2).HeadAlignment = dbgCenter
       .Columns(3).Caption = "T/Doc."
       .Columns(3).Width = 700
       .Columns(3).HeadAlignment = dbgCenter
       .Columns(4).Caption = "No.Documento"
       .Columns(4).Width = 2500
       .Columns(4).HeadAlignment = dbgCenter
       .Columns(5).Caption = "Descripcion"
       .Columns(5).Width = 4800
       .Columns(5).HeadAlignment = dbgCenter
       .Columns(6).Caption = "Total"
       .Columns(6).NumberFormat = "##,###,##0.00"
       .Columns(6).Width = 1200
       .Columns(6).HeadAlignment = dbgCenter
       .Refresh
   End With
   TReg(1) = Format(TDBGrid1.ApproxCount, "#####0")
   
End Function


Private Sub tdbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  If KeyCode = 13 Then
     If TDBGrid1.ApproxCount > 0 And TDBGrid1.Columns(0).text <> "*" Then
       FrmAnulaLiquidacionCompraDetalle.Btipo = g_tipoped
        FrmAnulaLiquidacionCompraDetalle.BNumero = TDBGrid1.Columns(2).text
        FrmAnulaLiquidacionCompraDetalle.Show 1
        Listado
     End If
  End If

End Sub
Private Sub CargarTipo(xcombo As ComboBox, xtipo)
  
  Select Case xtipo
    Case 1     '--condicion documento--
     xcombo.Clear
     xcombo.AddItem "0-Activo"
     xcombo.AddItem "1-Anulado"
     xcombo.ListIndex = 0
   Case 2   '--tipodocumento --
     xcombo.Clear
     xcombo.AddItem g_tipofac & "-Luidacion de Compra"
     xcombo.ListIndex = 0
   Case 3   '---estado
     xcombo.Clear
     xcombo.AddItem "S-SI"
     xcombo.AddItem "N-NO"
     xcombo.ListIndex = 1
   Case 4  '-- Tipo persona
     xcombo.Clear
     xcombo.AddItem "1-NATURAL"
     xcombo.AddItem "2-JURIDICA"
     xcombo.ListIndex = 0
   Case 5  '-tipo pais
     xcombo.Clear
     xcombo.AddItem "1-PERUANA"
     xcombo.AddItem "2-EXTRANJERA"
     xcombo.ListIndex = 0
   Case 6   '--todos los tipodocumentos --
     xcombo.Clear
     'Call adll2.llenacombo(xcombo, "select documentocodigo,documentodescripcion from vt_documento order by documentodescripcion",VGcnx)
     xcombo.AddItem g_tipofac & "-Liquidacion"
     xcombo.AddItem g_tipoped & "-Pedido"
     xcombo.ListIndex = 0
     
  End Select
End Sub



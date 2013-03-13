VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmAyudacontado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda de Documento al Contado"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin TrueOleDBGrid70.TDBGrid DGrid1 
      Height          =   5805
      Left            =   240
      TabIndex        =   3
      Top             =   210
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   10239
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=101,.bold=0,.fontsize=825,.italic=0"
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
   Begin VB.CommandButton cAcepto 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Acepta"
      Height          =   435
      Left            =   8610
      TabIndex        =   1
      Top             =   6285
      Width           =   1170
   End
   Begin VB.CommandButton cCerrar 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Cerrar"
      Height          =   435
      Left            =   9870
      TabIndex        =   0
      Top             =   6285
      Width           =   1170
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6915
      Width           =   11205
      _ExtentX        =   19764
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
End
Attribute VB_Name = "FrmAyudacontado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DLLGENERALAYUDA As New dllgeneral.dll_general
Dim rsbusca As New ADODB.Recordset
Dim rsdetab As New ADODB.Recordset
Dim xfecha As Date
Dim nlongi(1)

Public Property Let bfecha(pdata As Date)
    xfecha = pdata
End Property
Private Sub cAcepto_Click()
   Dim J As Double
   
   If rsdetab.RecordCount > 0 Then
      rsdetab.MoveFirst
      Do Until rsdetab.EOF
         FrmPlanillaCobranza.TDBGrid1.AllowAddNew = True
         FrmPlanillaCobranza.TDBGrid1.Columns(0).Text = rsdetab.Fields(1)
         FrmPlanillaCobranza.TDBGrid1.Columns(1).Text = rsdetab.Fields(2)
         FrmPlanillaCobranza.TDBGrid1.Columns(2).Text = rsdetab.Fields(3)
         FrmPlanillaCobranza.TDBGrid1.Columns(3).Text = rsdetab.Fields(4)
         FrmPlanillaCobranza.TDBGrid1.Columns(4).Text = rsdetab.Fields(5)
         FrmPlanillaCobranza.TDBGrid1.Columns(5).Text = rsdetab.Fields(6)
         FrmPlanillaCobranza.TDBGrid1.Columns(6).Text = rsdetab.Fields(7)
         FrmPlanillaCobranza.TDBGrid1.Columns(7).Text = rsdetab.Fields(8)
         FrmPlanillaCobranza.TDBGrid1.Columns(8).Text = rsdetab.Fields(9)
         FrmPlanillaCobranza.TDBGrid1.Columns(9).Text = rsdetab.Fields(10)
         FrmPlanillaCobranza.TDBGrid1.Columns(10).Text = rsdetab.Fields(11)
         FrmPlanillaCobranza.TDBGrid1.Columns(11).Text = rsdetab.Fields(12)
         FrmPlanillaCobranza.TDBGrid1.Update
         rsdetab.MoveNext
      Loop
   End If
   Set rsdetab = Nothing
   nAyuda = "": nDetalle = ""
   Unload Me
End Sub

Private Sub cCerrar_Click()
  nAyuda = "": nDetalle = ""
  Unload Me
End Sub


Private Sub DGrid1_DblClick()
  If rsdetab.RecordCount > 0 Then
    DGrid1.Columns(0).Text = ""
    DGrid1.Update
  End If
  
End Sub

Private Sub DGrid1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
     If rsdetab.RecordCount > 0 Then
      DGrid1.Columns(0).Text = "*"
      DGrid1.Update
    End If
 End If
End Sub

Private Sub Form_Load()
  MostrarForm Me, "C"
    
  nlongi(1) = 0
  DoEvents
  Call Carga_Grilla
  Set rsbusca = VGCNx.Execute("select * from cp_cargo where cargoapefecemi='" & xfecha & "' order by documentocargo")
  If rsbusca.RecordCount > 0 Then
    rsbusca.MoveFirst
    Do Until rsbusca.EOF
        rsdetab.AddNew
        rsdetab.Fields("flag") = "*"
        rsdetab.Fields("Cliente") = rsbusca.Fields("clientecodigo")
        rsdetab.Fields("TD") = rsbusca.Fields("documentocargo")
        rsdetab.Fields("Serie") = Left$(rsbusca.Fields("cargonumdoc"), 4)
        rsdetab.Fields("Numero") = Right$(rsbusca.Fields("cargonumdoc"), 10)
        rsdetab.Fields("P/T") = "T"
        rsdetab.Fields("TDp") = "10"
        rsdetab.Fields("Seriep") = "0000"
        rsdetab.Fields("Numerop") = "0000000000"
        rsdetab.Fields("Moneda") = rsbusca.Fields("monedacodigo")
        rsdetab.Fields("Banco") = ""
        rsdetab.Fields("Importe") = rsbusca.Fields("cargoapeimpape")
        rsdetab.Fields("TCambio") = rsbusca.Fields("cargoapetipcam")
        rsbusca.MoveNext
    Loop
  End If
  rsbusca.Close
  Set rsbusca = Nothing
  
End Sub


Public Function Carga_Grilla()
   Set rsdetab = Nothing
    
   Call rsdetab.Fields.Append("flag", adChar, 1)
   Call rsdetab.Fields.Append("Cliente", adChar, 11)
   Call rsdetab.Fields.Append("TD", adChar, 2)
   Call rsdetab.Fields.Append("Serie", adChar, 4)
   Call rsdetab.Fields.Append("Numero", adChar, 10)
   Call rsdetab.Fields.Append("P/T", adChar, 1)
   Call rsdetab.Fields.Append("TDp", adChar, 2)
   Call rsdetab.Fields.Append("Seriep", adChar, 4)
   Call rsdetab.Fields.Append("Numerop", adChar, 10)
   Call rsdetab.Fields.Append("Moneda", adChar, 2)
   Call rsdetab.Fields.Append("Banco", adChar, 2)
   Call rsdetab.Fields.Append("Importe", adDouble)
   Call rsdetab.Fields.Append("TCambio", adDouble)
   
   rsdetab.Open
   Set DGrid1.DataSource = rsdetab
   DGrid1.Refresh
   Call ConfigGrid

End Function


Public Function ConfigGrid()
    With DGrid1
        .Columns(0).Width = 300
        .Columns(1).Width = 1200
        .Columns(2).Width = 500
        .Columns(3).Width = 500
        .Columns(4).Width = 1300
        .Columns(5).Width = 800
        .Columns(6).Width = 800
        .Columns(7).Width = 800
        .Columns(8).Width = 1300
        .Columns(9).Width = 800
        .Columns(10).Width = 1000
        .Columns(11).Width = 1300
        .Columns(11).NumberFormat = "##,###,###,##0.00"
        .Columns(12).Width = 1000
        .Columns(12).NumberFormat = "##,###,###,##0.00"
    End With
    DGrid1.Refresh
End Function



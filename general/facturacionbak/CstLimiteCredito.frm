VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CstLimiteCredito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Limite de Credito"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5805
      Left            =   180
      TabIndex        =   9
      Top             =   1410
      Width           =   11685
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   4815
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=128,.bold=0,.fontsize=825,.italic=0"
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
      Begin VB.Frame Frame5 
         Height          =   585
         Index           =   0
         Left            =   9270
         TabIndex        =   10
         Top             =   5130
         Width           =   2265
         Begin VB.TextBox TReg 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
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
            Left            =   210
            TabIndex        =   12
            Top             =   240
            Width           =   1035
         End
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1245
      Index           =   1
      Left            =   210
      TabIndex        =   5
      Top             =   150
      Width           =   11625
      Begin VB.Frame Frame3 
         Height          =   1005
         Left            =   1890
         TabIndex        =   13
         Top             =   120
         Width           =   9675
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1620
            TabIndex        =   2
            Top             =   540
            Width           =   6405
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   510
            Width           =   1245
         End
         Begin VB.CommandButton cBusca 
            Caption         =   "&Buscar"
            Height          =   405
            Left            =   8250
            TabIndex        =   3
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "BUSQUEDA"
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
            Index           =   2
            Left            =   150
            TabIndex        =   14
            Top             =   240
            Width           =   6915
         End
      End
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
         TabIndex        =   6
         Top             =   120
         Width           =   1665
         Begin MSMask.MaskEdBox MTCambio 
            Height          =   315
            Left            =   270
            TabIndex        =   0
            Top             =   540
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "TIPO CAMBIO"
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
            Left            =   180
            TabIndex        =   7
            Top             =   240
            Width           =   1245
         End
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   7380
      Width           =   12105
      _ExtentX        =   21352
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
End
Attribute VB_Name = "CstLimiteCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim rsdeta As New ADODB.Recordset
Dim nlongi(1) As Integer
Dim vsql As String

Private Sub cBusca_Click()

    If Len(Trim(Text1)) > 0 Then
        If Combo1.ListIndex = 0 Then
            vsql = "select * from vt_cliente where clientecodigo like '" & Trim(Text1) & " %'"
        ElseIf Combo1.ListIndex = 1 Then
           vsql = "select * from vt_cliente where clienterazonsocial like '" & Trim(Text1) & "%'"
        End If
    Else
        vsql = "select * from vt_cliente"
    End If
    Call CargarDatos(vsql)
    
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
   Seguir Combo1, KeyAscii
End Sub

Private Sub Form_Load()
   
   MostrarForm Me, "C"
   nlongi(1) = 0
   Combo1.Clear
   Combo1.AddItem "Codigo"
   Combo1.AddItem "Descripcion"
   Combo1.ListIndex = 0
   MTCambio = numero(1)
   vsql = "select * from vt_cliente"
   Call DefineGrilla
   Call CargarDatos(vsql)
   
End Sub

Public Sub CargarDatos(nsql As String)
   Dim rs As New ADODB.Recordset
   Dim lsoles As Double
   Dim ldolares As Double
   Dim ssoles As Double
   Dim sdolares As Double
   
   
   Set rsdeta = Nothing
   Call DefineGrilla
   Set rs = VGcnx.Execute(nsql)
   If rs.RecordCount > 0 Then
      rs.MoveFirst
      Do Until rs.EOF
          rsdeta.AddNew
          rsdeta.Fields("Codigo") = Escadena(rs!clientecodigo)
          rsdeta.Fields("Descripcion") = Escadena(Left(rs!clienterazonsocial, 40))
          rsdeta.Fields("D") = ""
          rsdeta.Fields("P") = ""
          If IsNull(rs!clientesaldosoles) Then
             ssoles = 0
          Else
             ssoles = (rs!clientesaldosoles / CDbl(MTCambio.ClipText))
          End If
          If IsNull(rs!clientesaldodolares) Then
             sdolares = 0
          Else
             sdolares = (rs!clientesaldodolares)
          End If
          
          If IsNull(rs!clientelimitecredsoles) Then
             lsoles = 0
          Else
             lsoles = (rs!clientelimitecredsoles / CDbl(MTCambio.ClipText))
          End If
          If IsNull(rs!clientelimitecreddolar) Then
             ldolares = 0
          Else
             ldolares = (rs!clientelimitecreddolar)
          End If
          rsdeta.Fields("Limite Cred.") = numero(lsoles + ldolares)
          rsdeta.Fields("L.VTa") = ""
          rsdeta.Fields("Total Saldo") = numero(ssoles + sdolares)
          rsdeta.Update
          rs.MoveNext
      Loop
   End If
   rs.Close
   Set rs = Nothing
   
  ConfiguraGrid
End Sub


Public Sub DefineGrilla()
   Set rsdeta = Nothing
   
   Call rsdeta.Fields.Append("Codigo", adChar, 20)
   Call rsdeta.Fields.Append("Descripcion", adChar, 100)
   Call rsdeta.Fields.Append("D", adChar, 3)
   Call rsdeta.Fields.Append("P", adChar, 3)
   Call rsdeta.Fields.Append("Limite Cred.", adDouble)
   Call rsdeta.Fields.Append("L.VTa", adChar, 3)
   Call rsdeta.Fields.Append("Total Saldo", adDouble)
   
   rsdeta.Open
   ConfiguraGrid
   
End Sub
   
   
   
Public Function ConfiguraGrid()
      
   Set TDBGrid1.DataSource = rsdeta
   TReg.Text = Format(TDBGrid1.ApproxCount, "#####0")
   With TDBGrid1
       .Columns(0).Width = 1100
       .Columns(0).HeadAlignment = dbgCenter
       .Columns(1).Width = 5500
       .Columns(1).HeadAlignment = dbgCenter
       .Columns(2).Width = 600
       .Columns(2).HeadAlignment = dbgCenter
       .Columns(3).Width = 600
       .Columns(3).HeadAlignment = dbgCenter
       .Columns(4).Width = 1200
       .Columns(4).NumberFormat = "##,###,##0.00"
       .Columns(4).BackColor = &HC0FFFF
       .Columns(4).HeadAlignment = dbgCenter
       .Columns(5).Width = 600
       .Columns(5).HeadAlignment = dbgCenter
       .Columns(6).Width = 1200
       .Columns(6).NumberFormat = "##,###,##0.00"
       .Columns(6).BackColor = &H55ABFF
       .Columns(6).HeadAlignment = dbgCenter
       .Refresh
   End With
   
End Function

Private Sub MTCambio_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     MTCambio = numero(MTCambio)
     Call CargarDatos(vsql)
     Seguir MTCambio, KeyCode
  End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Seguir Combo1, KeyAscii
End Sub

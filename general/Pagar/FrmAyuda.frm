VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmAyuda 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayuda"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9060
   Begin TrueOleDBGrid70.TDBGrid DGrid1 
      Height          =   3945
      Left            =   120
      TabIndex        =   2
      Top             =   1620
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   6959
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
      Appearance      =   2
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=192,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=900"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Arial"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Arial"
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
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   6210
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   556
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
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   90
      TabIndex        =   7
      Top             =   60
      Width           =   8745
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Elija su Tipo de Busqueda e Ingrese su Dato a buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Top             =   330
         Width           =   6585
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   450
         Picture         =   "FrmAyuda.frx":0000
         Top             =   210
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   6
      Top             =   870
      Width           =   8715
      Begin VB.CommandButton cBusca 
         BackColor       =   &H0000C0C0&
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   7230
         TabIndex        =   1
         Top             =   270
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1980
         TabIndex        =   0
         Top             =   300
         Width           =   5175
      End
   End
   Begin VB.CommandButton cCerrar 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Cerrar"
      Height          =   435
      Left            =   1410
      TabIndex        =   4
      Top             =   5655
      Width           =   1230
   End
   Begin VB.CommandButton cAcepto 
      BackColor       =   &H0000C0C0&
      Caption         =   "&Acepta"
      Height          =   435
      Left            =   150
      TabIndex        =   3
      Top             =   5670
      Width           =   1170
   End
   Begin VB.Frame Frame3 
      Height          =   525
      Left            =   6600
      TabIndex        =   10
      Top             =   5520
      Width           =   2235
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   990
         TabIndex        =   12
         Top             =   180
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "Total Reg."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   11
         Top             =   210
         Width           =   915
      End
   End
End
Attribute VB_Name = "FrmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xtipo As Integer
Dim AOrden, ACondi As String
Dim DLLGENERALAYUDA As New dllgeneral.dll_general
Dim vcon As New ADODB.Connection
Dim xtabla, xCampos, xOrden, xCondi As String
Dim xdata, xdato As String

Dim nlongi(1) As Integer
Dim nfiltra() As String

Private Sub cAcepto_Click()
  Call DGrid1_KeyDown(13, 0)
End Sub

Private Sub cBusca_Click()
    Call Text1_KeyPress(13)
End Sub

Private Sub cCerrar_Click()
  nAyuda = "": nDetalle = ""
  Unload Me
End Sub

Private Sub DGrid1_HeadClick(ByVal ColIndex As Integer)
  Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, ColIndex + 1, nlongi, xCondi)
  ConfigGrid xtipo
End Sub

Private Sub DGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    If DGrid1.Row >= 0 Then
        Select Case xtipo
          Case 1, 3, 4
            nAyuda = DGrid1.Columns(0).Text
            nDetalle = Trim$(DGrid1.Columns(1).Text)
          Case 2
            nAyuda = DGrid1.Columns(2).Text
            nDetalle = ""
          Case 5
            nAyuda = DGrid1.Columns(0).Text
            nDetalle = Trim$(DGrid1.Columns(1).Text)
            nSaldo = IIf(IsNull(DGrid1.Columns(4).Text), 0, CDbl(DGrid1.Columns(4).Text))
            
         End Select
    End If
    xdata = ""
    Unload Me
    Exit Sub
  Else
    DGrid1.SetFocus
  End If
  xdata = ""
End Sub

Private Sub Form_Load()
    
  DoEvents
  nlongi(1) = 0
  DoEvents
  
  If Trim$(Escadena(xdata)) = "1" Then
     Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, "productodescripcion like '" & xdato & "%'")
     Text1 = Escadena(xdato): Text1.SelStart = Len(Trim$(Text1))
  ElseIf Trim$(Escadena(xdata)) = "2" Then
     Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, "clienterazonsocial like '" & xdato & "%'")
     Text1 = Escadena(xdato): Text1.SelStart = Len(Trim$(Text1))
  Else
     Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, xCondi)
  End If
  ConfigGrid xtipo
End Sub

Public Property Let BFiltro(ByRef campos)
   Dim f As Integer
   Dim tam As Integer
   tam = UBound(campos)
   ReDim nfiltra(tam, 2)
   Combo1.Clear
   For f = 1 To UBound(campos)
      nfiltra(f, 1) = campos(f, 1)
      nfiltra(f, 2) = campos(f, 2)
      Combo1.AddItem campos(f, 1)
   Next f
   If xdata = "1" Or xdata = "2" Then
     Combo1.ListIndex = 0
   End If
   
   ConfigGrid xtipo
End Property

Public Function ConfigGrid(xtipo As Integer)
   Dim J As Integer
      
   Select Case xtipo
     Case 1
        DGrid1.Columns(0).Width = 1000
        DGrid1.Columns(1).Width = 3000
        If DGrid1.Columns.Count = 3 Then
           DGrid1.Columns(2).Width = 1200
        End If
     Case 2
        DGrid1.Columns(0).Width = 600
        DGrid1.Columns(1).Width = 1000
        DGrid1.Columns(2).Width = 1200
        DGrid1.Columns(3).Width = 1000
        DGrid1.Columns(4).Width = 3000
        DGrid1.Columns(5).Width = 1200
        DGrid1.Columns(5).NumberFormat = "##,###,##0.00"
     Case 4
        DGrid1.Columns(0).Width = 1200
        DGrid1.Columns(1).Width = 6000
     Case 5
        DGrid1.Columns(0).Width = 800
        DGrid1.Columns(1).Width = 1600
        DGrid1.Columns(2).Width = 800
        DGrid1.Columns(3).Width = 1800
        DGrid1.Columns(4).Width = 1800
        DGrid1.Columns(4).NumberFormat = "##,###,##0.00"
        DGrid1.Columns(5).Width = 1200
        DGrid1.Columns(5).NumberFormat = "dd/mm/yyyy"
   End Select
   DGrid1.Refresh
   Text2.Text = Numero(DGrid1.ApproxCount)
End Function

Public Property Let TipoForma(pdata As Integer)
   xtipo = pdata
End Property

Public Property Let BTabla(pdata As String)
   xtabla = pdata
End Property

Public Property Let BConexion(ByRef pdata)
   Set vcon = pdata
End Property

Public Property Let BCampos(pdata As String)
   xCampos = pdata
End Property

Public Property Let BOrden(pdata As String)
   xOrden = IIf(IsNull(pdata), "", Trim$(pdata))
   AOrden = IIf(IsNull(pdata), "", Trim$(pdata))
End Property

Public Property Let BCondi(pdata As String)
   xCondi = IIf(IsNull(pdata), "", Trim$(pdata))
   ACondi = IIf(IsNull(pdata), "", Trim$(pdata))
  
End Property

Public Property Let Bdata(ByRef pdata)
   xdata = pdata
End Property

Public Property Let Bdato(ByRef pdata)
   xdato = pdata
End Property

Private Sub Text1_Change()
    Dim posi As Integer
    posi = Combo1.ListIndex + 1
    If Len(Trim$(xCondi)) = 0 Then
       Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, nfiltra(posi, 2) & " Like '" & Text1 & "%'")
    Else
       Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, xCondi & " and " & nfiltra(posi, 2) & " Like '" & Text1 & "%'")
    End If
    ConfigGrid xtipo
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  Dim posi As Integer
  If KeyAscii = 13 Then
     If Len(Trim$(Text1)) > 0 Then
        posi = Combo1.ListIndex + 1
        If Len(Trim$(xCondi)) = 0 Then
           Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, nfiltra(posi, 2) & " Like '" & Text1 & "%'")
        Else
           Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, xCondi & " and " & nfiltra(posi, 2) & " Like '" & Text1 & "%'")
        End If
     Else
        Call DLLGENERALAYUDA.ListarEnTDBGRID(vcon, xtabla, DGrid1, xCampos, xOrden, nlongi, xCondi)
     End If
     ConfigGrid xtipo
     DGrid1.SetFocus
  End If
End Sub

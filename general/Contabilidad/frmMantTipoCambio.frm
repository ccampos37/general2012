VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmMantTipoCambio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Cambio"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6705
   Begin VB.Frame frmbotones 
      Height          =   555
      Left            =   450
      TabIndex        =   5
      Top             =   4590
      Width           =   5730
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Imprimir"
         Height          =   330
         Index           =   3
         Left            =   3435
         TabIndex        =   10
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         Height          =   330
         Index           =   4
         Left            =   4560
         TabIndex        =   9
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         Height          =   330
         Index           =   2
         Left            =   2310
         TabIndex        =   8
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "E&ditar"
         Height          =   330
         Index           =   1
         Left            =   1185
         TabIndex        =   7
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         Height          =   330
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   165
         Width           =   1080
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4530
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   7990
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "frmMantTipoCambio.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TDBGrid1"
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(2)=   "lblNumReg"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "frmMantTipoCambio.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cCancela"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cAcepta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   2025
         TabIndex        =   4
         Top             =   4065
         Width           =   1140
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   3465
         TabIndex        =   17
         Top             =   4065
         Width           =   1140
      End
      Begin VB.Frame Frame1 
         Height          =   3645
         Left            =   30
         TabIndex        =   12
         Top             =   330
         Width           =   6555
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   2730
            TabIndex        =   0
            ToolTipText     =   "Seleccionar Fechas en el Mes de Proceso"
            Top             =   135
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   556
            _Version        =   393216
            Format          =   54132737
            CurrentDate     =   37501
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   0
            Left            =   2715
            TabIndex        =   1
            Top             =   450
            Width           =   3780
            _ExtentX        =   6668
            _ExtentY        =   529
            BackColor       =   16777215
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
            MaxLength       =   6
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            NumeroDecimales =   3
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   1
            Left            =   2715
            TabIndex        =   2
            Top             =   765
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   529
            BackColor       =   -2147483639
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
            MaxLength       =   40
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            NoCaracteres    =   "',"
            NumeroDecimales =   3
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   2
            Left            =   2715
            TabIndex        =   3
            Top             =   1080
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   529
            BackColor       =   16777215
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
            MaxLength       =   25
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            NoCaracteres    =   "',"
            NumeroDecimales =   3
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin VB.Label lbl 
            Caption         =   "Tipo Cambio Promedio"
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
            Index           =   3
            Left            =   120
            TabIndex        =   16
            Top             =   1110
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Tipo Cambio Venta"
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
            Index           =   2
            Left            =   135
            TabIndex        =   15
            Top             =   810
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Tipo Cambio Compra"
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
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   510
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Fecha"
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
            Left            =   120
            TabIndex        =   13
            Top             =   210
            Width           =   2310
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   4110
         Left            =   -74955
         TabIndex        =   18
         Top             =   360
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   7250
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Fecha"
         Columns(0).DataField=   "tipocambiofecha"
         Columns(0).NumberFormat=   "Short Date"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Valor Compra"
         Columns(1).DataField=   "tipocambiocompra"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Valor Venta"
         Columns(2).DataField=   "tipocambioventa"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Valor Promedio"
         Columns(3).DataField=   "tipocambiopromedio"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label1 
         Caption         =   "N� Registros"
         Height          =   270
         Left            =   -70260
         TabIndex        =   20
         Top             =   5475
         Width           =   900
      End
      Begin VB.Label lblNumReg 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNumReg"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   -69315
         TabIndex        =   19
         Top             =   5460
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmMantTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim i_filaorigen As Integer
Dim rs As New ADODB.Recordset
Dim FLAG_CHECK As Boolean

Private Sub Form_Initialize()
    If IsNumeric(VGParamSistem.Anoproceso) Then
       DTPicker1.MinDate = Format("01/" & VGParamSistem.Mesproceso & "/" & CInt(VGParamSistem.Anoproceso), "dd/mm/yyyy")
       DTPicker1.MaxDate = DateAdd("d", -1, DateAdd("m", 1, DTPicker1.MinDate))
       'DTPicker1.MaxDate = DateAdd("d", 1, DateAdd("m", 1, DTPicker1.MinDate))
       If Month(Now) = Month(DTPicker1.Value) And Year(Now) = Year(DTPicker1.Value) Then
           DTPicker1.Value = Now
       End If
    End If
End Sub

Private Sub Form_Load()
  Call ConfiguraForm
  Call MuestraDatos
End Sub

Sub ConfiguraForm()
  SSTab1.Tab = 0
  SSTab1.TabEnabled(1) = False
  cAcepta.Enabled = False
  lblNumReg.Caption = Empty
  FLAG_CHECK = False
  Me.Width = 6780
  Me.Height = 5595
End Sub

Private Function MuestraDatos()
 Dim SQL As String
  SQL = "SELECT tipocambiofecha, tipocambiocompra, tipocambioventa, tipocambiopromedio "
  SQL = SQL & "FROM CT_TIPOCAMBIO WHERE MONTH(tipocambiofecha)=" & VGParamSistem.Mesproceso & " And Year(tipocambiofecha) =" & VGParamSistem.Anoproceso & " "
  SQL = SQL & "ORDER BY 1"
  Set rs = VGCNx.Execute(SQL)
  Set TDBGrid1.DataSource = rs
  lblNumReg.Caption = rs.RecordCount
  SSTab1.Tab = 0
End Function

Private Sub cCancela_Click()
  SSTab1.TabEnabled(0) = True
  SSTab1.Tab = 0
  SSTab1.SetFocus
  frmbotones.Visible = True
  modoinsert = False
  modoedit = False
  i_filaorigen = -1
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim j As Integer
  Dim spos As Integer
  Dim SQL As String
  
  'On Error GoTo X
  SSTab1.TabEnabled(1) = True
  
  Select Case Index
     Case 0   'nuevo
        modoinsert = True
        frmbotones.Visible = False
        SSTab1.Tab = 1
        Call LimpiarValores
        DTPicker1.SetFocus
        
     Case 1   'editar
        If TDBGrid1.Row < 0 Then
          Exit Sub
        End If
        modoedit = True
        frmbotones.Visible = False
        SSTab1.Tab = 1
        Call EditarValores
      
     Case 2   'eliminar
       If MsgBox("Desea eliminar el registro de Fecha " & TDBGrid1.Columns(0).Value, vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          SQL = "DELETE FROM ct_tipocambio WHERE convert(varchar(10),tipocambiofecha,103)='" & TDBGrid1.Columns(0).Value & "'"
          VGCNx.Execute (SQL)
          Call MuestraDatos
       End If
        
     Case 3   'imprimir
       Call Impresion("rptTipoCambio.rpt")
     
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

Sub EditarValores()
 Dim I As Integer
  With TDBGrid1
    DTPicker1.Value = .Columns(0).Text
    For I = 1 To 3
      txt(I - 1).Text = Trim(.Columns(I).Text)
    Next
  End With
End Sub

Public Function LimpiarValores()
 Dim I As Integer
  If Month(Now) = Month(DTPicker1.Value) And Year(Now) = Year(DTPicker1.Value) Then
    DTPicker1.Value = Now
  End If
  For I = 0 To 2
    txt(I).Text = Empty
  Next
End Function

Private Sub cAcepta_Click()
 If ValidaData() = True Then
    Call GrabaData
 End If
     
End Sub

Sub GrabaData()
  Dim SQL As String
  On Error GoTo X
  
  SSTab1.TabEnabled(0) = True
  
  If modoinsert = True Then
        SQL = "INSERT CT_TIPOCAMBIO(tipocambiofecha, tipocambiocompra, tipocambioventa, tipocambiopromedio,usuariocodigo,fechaact) "
        SQL = SQL & "VALUES (" & DateSQL(DTPicker1.Value) & "," & txt(0).Text & "," & txt(1).Text & "," & txt(2).Text & ",'" & VGusuario & "'," & DateSQL(Date) & ")"
        VGCNx.BeginTrans
        VGCNx.Execute (SQL)
        VGCNx.CommitTrans
    
  ElseIf modoedit = True Then
    SQL = "UPDATE CT_TIPOCAMBIO SET tipocambiocompra=" & txt(0).Text & ","
    SQL = SQL & "tipocambioventa =" & txt(1).Text & ","
    SQL = SQL & "tipocambiopromedio=" & txt(2).Text & ","
    SQL = SQL & "usuariocodigo='" & VGusuario & "',fechaact=" & DateSQL(Date) & " "
    SQL = SQL & "WHERE convert(varchar(10),tipocambiofecha,103)='" & Format(DTPicker1.Value, "DD/MM/YYYY") & "'"
    VGCNx.BeginTrans
    VGCNx.Execute (SQL)
    VGCNx.CommitTrans
    
  End If
  
  Call MuestraDatos
  frmbotones.Visible = True
  modoinsert = False: modoedit = False: FLAG_CHECK = False
  i_filaorigen = -1
  Exit Sub

X:
  If err.Number = -2147217873 Then
    MsgBox "Esta intentando registrar uno Existente " & err.Description, vbInformation, Caption
  Else
    MsgBox "Error inesperado: " & err.Number & " " & err.Description
  End If
  VGCNx.RollbackTrans

End Sub

Function ValidaData() As Boolean
 Dim rsX As ADODB.Recordset
 Dim SQL As String
 Dim I As Integer
   
   SQL = "select tipocambiofecha from ct_tipocambio where convert(varchar(10),tipocambiofecha,103)='" & Format(DTPicker1.Value, "dd/mm/yyyy") & "'"
   Set VGvardllgen = New dllgeneral.dll_general
   If VGvardllgen.VerificaDatoExistente(VGCNx, SQL) > 0 And modoinsert = True Then
     MsgBox "La Fecha de Tipo Cambio a Registrar Existe en la Base", vbInformation, Caption
     ValidaData = False
     Exit Function
   End If

  ValidaData = True
End Function

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  SSTab1.TabEnabled(PreviousTab) = False
  cAcepta.Enabled = False
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

Private Sub TDBGrid1_DblClick()
 If rs.RecordCount > 0 And (modoedit = False And modoinsert = False) Then
   Call cmdBotones_Click(1)
 End If
End Sub

Private Sub ConfiguraTdbgrid()
  Dim I As Integer
  Dim i_total As Integer
  Dim i_width As Integer
'  TDBGrid1.Columns(1).Visible = False
'  TDBGrid1.Columns(2).Visible = False
  TDBGrid1.Columns(0).Width = 800
  TDBGrid1.Columns(1).Width = 1000
  TDBGrid1.Columns(2).Width = 1000
  TDBGrid1.Columns(3).Width = 1000

End Sub

Function ValidaDataIngreso() As Boolean
 Dim I As Integer
  For I = 0 To 2
   If txt(I).Text = Empty Then
     ValidaDataIngreso = False
     Exit Function
   End If
  Next

  ValidaDataIngreso = True
End Function

Private Sub txt_Change(Index As Integer)
  cAcepta.Enabled = ValidaDataIngreso()
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 And Index = 2 Then
    cAcepta.SetFocus
    Call cAcepta_Click
  End If
End Sub

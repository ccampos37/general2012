VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmLimiteCredito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento Límite Crédito"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmbotones 
      Height          =   1050
      Left            =   2280
      TabIndex        =   18
      Top             =   5520
      Width           =   4815
      Begin VB.CommandButton cmdBotones 
         Caption         =   "E&ditar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   1
         Left            =   360
         Picture         =   "FrmLimiteCredito.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   2
         Left            =   1485
         Picture         =   "FrmLimiteCredito.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   870
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   4
         Left            =   3630
         Picture         =   "FrmLimiteCredito.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   3
         Left            =   2550
         Picture         =   "FrmLimiteCredito.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   870
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "FrmLimiteCredito.frx":1108
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "FrmLimiteCredito.frx":1124
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cCancela"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cAcepta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame2 
         Caption         =   "Búsqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74640
         TabIndex        =   20
         Top             =   360
         Width           =   8595
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7080
            TabIndex        =   2
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtBusqueda 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3000
            TabIndex        =   1
            Top             =   360
            Width           =   3855
         End
         Begin VB.ComboBox cmbBusqueda 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         HelpContextID   =   7
         Left            =   2760
         TabIndex        =   8
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         HelpContextID   =   8
         Left            =   4680
         TabIndex        =   9
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   3735
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   8625
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3480
            MaxLength       =   3
            TabIndex        =   4
            Top             =   960
            Width           =   4905
         End
         Begin VB.CheckBox chk 
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   7
            Top             =   2880
            Width           =   375
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   6
            Top             =   2160
            Width           =   2145
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   3480
            MaxLength       =   3
            TabIndex        =   3
            Top             =   480
            Width           =   2145
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   5
            Top             =   1560
            Width           =   2145
         End
         Begin VB.Label lbl 
            Caption         =   "Razón Social"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   1680
            TabIndex        =   23
            Top             =   1080
            Width           =   1605
         End
         Begin VB.Label lbl 
            Caption         =   "Suspendido"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   1680
            TabIndex        =   22
            Top             =   2880
            Width           =   1365
         End
         Begin VB.Label lbl 
            Caption         =   "Crédito Dólares"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   1680
            TabIndex        =   21
            Top             =   2280
            Width           =   1725
         End
         Begin VB.Label lbl 
            Caption         =   "Cód.Cliente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   1680
            TabIndex        =   17
            Top             =   600
            Width           =   1605
         End
         Begin VB.Label lbl 
            Caption         =   "Crédito Soles"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   1680
            TabIndex        =   16
            Top             =   1680
            Width           =   1485
         End
      End
      Begin TrueOleDBGrid70.TDBGrid DGrid1 
         Height          =   3465
         Left            =   -74640
         TabIndex        =   19
         Top             =   1440
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   6112
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=236,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HC0FFFF&,.bold=0,.fontsize=900"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
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
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HFFFFFF&"
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
Attribute VB_Name = "FrmLimiteCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim i_codigocliente As String
'''' Busqueda
'FIXIT: Declare 'ArregloBusqueda' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim ArregloBusqueda()
Dim i_indexComboBusqueda As Integer

Private Sub cAcepta_Click()
    
   Dim SQL As String
   Dim s_codigodocumento As String
'FIXIT: Declare 'd_limitecreditosol' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
   Dim d_limitecreditosol, d_limitecreditodolar As Double
   
   On Error GoTo CONTROLERRORES
   '''''''''
            If RTrim$(txt(2)) = "" Then
                d_limitecreditosol = 0
            Else
                d_limitecreditosol = txt(2)
            End If
            If RTrim$(txt(3)) = "" Then
                d_limitecreditodolar = 0
            Else
                d_limitecreditodolar = CDbl(txt(3))
            End If
   
            SQL = "UPDATE vt_cliente SET " & _
               "clientelimitecredsoles=" & d_limitecreditosol & "," & _
               "clientelimitecreddolar=" & d_limitecreditodolar & "," & _
               "clientesuspendido=" & chk(0).Value & "," & _
               "usuariocodigo='" & g_usuario & "'," & _
               "fechaact='" & Date & "' " & _
               "WHERE clientecodigo='" & i_codigocliente & "'"
    
            VGCNx.Execute SQL
            
 '******************************************************************************************
        
 DGrid1.Refresh
 MostrarOcultar_Botones (True)
 '''''''''''''''''''''''''''' Busqueda
 Call fncBusqueda(VGCNx, DGrid1)
 ''''''''''''''''''''''''''''
 SSTab1.Tab = 0
 SSTab1.TabEnabled(0) = True

 
Exit Sub
CONTROLERRORES:
   If Err Then
       MsgBox Err.Number & "-" & Err.Description, vbInformation, "ERROR"
       Err = 0
       'VGGeneral.RollbackTrans
       Resume Next
    End If
       
End Sub

Private Sub cCancela_Click()
    SSTab1.TabEnabled(0) = True
    SSTab1.Tab = 0
    SSTab1.SetFocus
    MostrarOcultar_Botones (True)
End Sub

Private Sub cmbBusqueda_Click()
    If i_indexComboBusqueda <> cmbBusqueda.ListIndex Then
        txtBusqueda = ""
    End If
End Sub

Private Sub cmbBusqueda_DropDown()
    i_indexComboBusqueda = cmbBusqueda.ListIndex
End Sub

Private Sub cmbBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdAceptar_Click()
    Call fncBusqueda(VGCNx, DGrid1)
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim J As Integer
  Dim SQL As String
'FIXIT: Declare 'OBJ' con un tipo de datos de enlace en tiempo de compilación              FixIT90210ae-R1672-R1B8ZE
  Dim OBJ As Object
  
  On Error GoTo CONTROLERRORES
  
  SSTab1.TabEnabled(1) = True
  '''''
  Select Case Index
        
     Case 1   'modificar
     
         If DGrid1.Row < 0 Then
            Exit Sub
         End If
         i_codigocliente = RTrim$(DGrid1.Columns(0).Text)
         txt(0) = RTrim$(DGrid1.Columns(0).Text)
         txt(1) = RTrim$(DGrid1.Columns(1).Text)
         txt(2) = RTrim$(DGrid1.Columns(2).Text)
         txt(3) = RTrim$(DGrid1.Columns(3).Text)
         'Activo
         If DGrid1.Columns(4).Value = False Then
            chk(0).Value = 0
         ElseIf DGrid1.Columns(4).Value = True Then
            chk(0).Value = 1
         End If
         SSTab1.Tab = 1
         MostrarOcultar_Botones (False)
         txt(2).SetFocus
      
     Case 2   'eliminar
        If DGrid1.Row < 0 Then
            Exit Sub
        End If
        If MsgBox("Desea eliminar el registro?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          SQL = "DELETE FROM vt_cliente WHERE " & _
                "clientecodigo = '" & DGrid1.Columns(0).Text & "'"
          VGCNx.Execute SQL
          '''''''''''''''''''''''''''' Busqueda
          Call fncBusqueda(VGCNx, DGrid1)
          ''''''''''''''''''''''''''''
        End If
        
     Case 3   'imprimir
         Call Imprimir("RepLimiteCredito.rpt")
     Case 4  ' salir
       Unload Me
  End Select
Exit Sub
CONTROLERRORES:
   If Err Then
       MsgBox Err.Number & "-" & Err.Description, vbInformation, "ERROR"
       Err = 0
       'VGGeneral.RollbackTrans
       Resume Next
    End If

End Sub

Private Sub Form_Load()
 MostrarForm Me, "C2"
 'cAcepta.Enabled = False
 SSTab1.TabEnabled(1) = False
 ''' Busqueda:
 Call fncCargaArregloComboBusqueda(ArregloBusqueda, cmbBusqueda)
 
 Call fncBusqueda(VGCNx, DGrid1)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    SSTab1.TabEnabled(PreviousTab) = False
    'cAcepta.Enabled = False
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)  ' Salta con Enter
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

'FIXIT: Declare 'MostrarOcultar_Botones' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Private Function MostrarOcultar_Botones(Valor As Boolean)
    frmbotones.Visible = Valor
End Function

Private Sub txt_LostFocus(Index As Integer)
If txt(Index) <> "" Then
    If Index = 2 Or Index = 3 Then
        txt(Index).Text = Format(txt(Index).Text, "###,##0.00")
    End If
End If
End Sub

'FIXIT: Declare 'fncBusqueda' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
Private Function fncBusqueda(conexion As Connection, grid As TDBGrid)
    Dim SQL As String
    Dim where As String
    Dim condicion As String
    Dim rs As Recordset
    Dim i As Integer
    
    where = ""
    condicion = ""
    
    SQL = "SELECT " & _
         "clientecodigo as 'Cód.Cliente'," & _
         "clienterazonsocial as 'Razón Social'," & _
         "clientelimitecredsoles as 'Limit.Créd.Soles'," & _
         "clientelimitecreddolar as 'Limit.Créd.Dolar'," & _
         "clientesuspendido as Suspendido " & _
         "FROM vt_cliente "
    
    If cmbBusqueda.ListIndex <> -1 Then
       where = " WHERE " & _
              RTrim$(ArregloBusqueda(0, cmbBusqueda.ListIndex))
       Select Case ArregloBusqueda(2, cmbBusqueda.ListIndex)
         Case "C"
            condicion = " LIKE '" & RTrim$(txtBusqueda) & "%'"
         Case "N"
            condicion = " = " & RTrim$(txtBusqueda)
         Case "B"
            If Left$(txtBusqueda, 1) = "S" Then
                condicion = " = 1"
            ElseIf Left$(txtBusqueda, 1) = "N" Then
                condicion = " = 0"
            End If
       End Select
    End If
       
    SQL = SQL & where & condicion
     
    Set rs = conexion.Execute(SQL)
    Set grid.DataSource = rs
    
 ''''''''''''''''''''''''''''''''''' Tipo Columna
      'For i = 0 To grid.Columns.Count - 1
         'grid.Columns(i).Width = i_width * (Len(a_Arreglo(1, i)) / i_total)
         'If ArregloBusqueda(2, i) = "B" Then
         '   grid.Columns(i).ValueItems.Presentation = dbgCheckBox
         'Else
         '   grid.Columns(i).ValueItems.Presentation = dbgNormal
         'End If
      'Next i
      
      grid.Columns(4).ValueItems.Presentation = dbgCheckBox
      grid.Refresh
    
End Function
'FIXIT: Declare 'fncCargaArregloComboBusqueda' and 'ArrayBusqueda' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Private Function fncCargaArregloComboBusqueda(ArrayBusqueda As Variant, cmb As ComboBox)
Dim i As Integer
    ReDim ArrayBusqueda(0 To 2, 0 To 7)

   'Nombre Campos:
   ArrayBusqueda(0, 0) = "clientecodigo"
   ArrayBusqueda(0, 1) = "clienterazonsocial"
   ArrayBusqueda(0, 2) = "clienteruc"
   ArrayBusqueda(0, 3) = "clientedistrito"
   ArrayBusqueda(0, 4) = "clienteprovincia"
   ArrayBusqueda(0, 5) = "clientedepartamento"
   ArrayBusqueda(0, 6) = "clientetelefono"
   ArrayBusqueda(0, 7) = "estadoreg"
   'Nombres de Campo(Combo Busqueda):
   ArrayBusqueda(1, 0) = "Código"
   ArrayBusqueda(1, 1) = "Razon Social"
   ArrayBusqueda(1, 2) = "RUC"
   ArrayBusqueda(1, 3) = "Distrito"
   ArrayBusqueda(1, 4) = "Provincia"
   ArrayBusqueda(1, 5) = "Departamento"
   ArrayBusqueda(1, 6) = "Telefono"
   ArrayBusqueda(1, 7) = "Activo"
   'Tipo de Dato:
   ArrayBusqueda(2, 0) = "C"
   ArrayBusqueda(2, 1) = "C"
   ArrayBusqueda(2, 2) = "C"
   ArrayBusqueda(2, 3) = "C"
   ArrayBusqueda(2, 4) = "C"
   ArrayBusqueda(2, 5) = "C"
   ArrayBusqueda(2, 6) = "C"
   ArrayBusqueda(2, 7) = "B"
   
   cmb.Clear
   For i = 0 To UBound(ArrayBusqueda, 2)
    cmb.AddItem (RTrim$(ArrayBusqueda(1, i)))
   Next i
    
End Function

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
'Ingresar Mayusculas:
    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

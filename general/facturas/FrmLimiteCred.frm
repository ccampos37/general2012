VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmLimiteCred 
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmbotones 
      Height          =   1050
      Left            =   1800
      TabIndex        =   17
      Top             =   6240
      Width           =   5655
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
         Left            =   3510
         Picture         =   "FrmLimiteCred.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   4590
         Picture         =   "FrmLimiteCred.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Left            =   2440
         Picture         =   "FrmLimiteCred.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   180
         Width           =   870
      End
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
         Left            =   1320
         Picture         =   "FrmLimiteCred.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
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
         Index           =   0
         Left            =   225
         Picture         =   "FrmLimiteCred.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   180
         Width           =   915
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
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
      TabPicture(0)   =   "FrmLimiteCred.frx":154A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBGridProducto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "FrmLimiteCred.frx":1566
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cCancela"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cAcepta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
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
         Left            =   -72120
         TabIndex        =   15
         Top             =   5280
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
         Left            =   -70200
         TabIndex        =   14
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   4575
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   8625
         Begin VB.ComboBox cmbDocumento 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   960
            Width           =   3255
         End
         Begin VB.ComboBox cmbPuntoVta 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   360
            Width           =   3255
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
            Left            =   3960
            MaxLength       =   2
            TabIndex        =   5
            Top             =   3000
            Width           =   1065
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
            Index           =   1
            Left            =   3960
            MaxLength       =   8
            TabIndex        =   4
            Top             =   2280
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
            Index           =   0
            Left            =   3960
            MaxLength       =   3
            TabIndex        =   3
            Top             =   1680
            Width           =   1065
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
            Left            =   3960
            MaxLength       =   2
            TabIndex        =   2
            Top             =   3720
            Width           =   1065
         End
         Begin VB.Label lbl 
            Caption         =   "Documento"
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
            Left            =   2160
            TabIndex        =   13
            Top             =   1080
            Width           =   1605
         End
         Begin VB.Label lbl 
            Caption         =   "Guia Remisión 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   2160
            TabIndex        =   12
            Top             =   3840
            Width           =   1785
         End
         Begin VB.Label lbl 
            Caption         =   "Punto Venta"
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
            Left            =   2160
            TabIndex        =   11
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lbl 
            Caption         =   "Guia Remisión 1"
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
            Left            =   2160
            TabIndex        =   10
            Top             =   3120
            Width           =   1845
         End
         Begin VB.Label lbl 
            Caption         =   "Correlativo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   2160
            TabIndex        =   9
            Top             =   2400
            Width           =   1275
         End
         Begin VB.Label lbl 
            Caption         =   "Serie"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   2160
            TabIndex        =   8
            Top             =   1680
            Width           =   1080
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGridProducto 
         Height          =   5295
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   9340
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
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2752"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
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
   End
End
Attribute VB_Name = "FrmLimiteCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim modoinsert, modoedit As Boolean
Dim i_filaorigen As Integer
Dim i_codigoptovta, i_codigoptovtadoc, i_ptovtaserie As String
Dim i_indexcombo As Integer
''''''''''''''''''''''''
Dim ArregloPuntovta()
Dim ArregloDocumento()

Private Sub cAcepta_Click()
    
   Dim rs As New ADODB.Recordset
   Dim SQL As String
   Dim J As Integer
 
   Dim s_codigopuntoventa As String
   Dim s_codigodocumento As String
   
   On Error GoTo nerror
   ''''''''
   
     If cmbPuntoVta.ListIndex <> -1 Then
        s_codigopuntoventa = ArregloPuntovta(0, cmbPuntoVta.ListIndex)
     Else
        s_codigopuntoventa = ""
     End If
     If cmbDocumento.ListIndex <> -1 Then
        s_codigodocumento = ArregloDocumento(0, cmbDocumento.ListIndex)
     Else
        s_codigodocumento = ""
     End If
        
   If modoinsert = True Then
   
         If Validar_CodigosDuplicados("INSERT") = True Then
            MsgBox "Registro Duplicado", vbCritical, "Error"
            cAcepta.Enabled = False
            Exit Sub
          End If
               
          SQL = "INSERT INTO vt_puntovtadocumento " & _
               "(puntovtacodigo,documentocodigo," & _
               "puntovtadocserie,puntovtadoccorr," & _
               "puntovtadoccorr1,puntovtaguia1,usuariocodigo " & _
               ") VALUES " & _
               "('" & s_codigopuntoventa & "','" & s_codigodocumento & "'," & _
               "'" & txt(0) & "','" & txt(1) & "','" & txt(2) & "'," & _
               "'" & txt(3) & "'," & _
               "'" & g_usuario & "')"

          VGCNx.Execute SQL
                   
    ElseIf modoedit = True Then
   
             If Validar_CodigosDuplicados("UPDATE", i_filaorigen) = True Then
               MsgBox "Registro Duplicado", vbCritical, "Error"
               cAcepta.Enabled = False
               Exit Sub
             End If
                                 
            SQL = "UPDATE vt_puntovtadocumento SET " & _
               "puntovtacodigo='" & s_codigopuntoventa & "'," & _
               "documentocodigo='" & s_codigodocumento & "'," & _
               "puntovtadocserie='" & txt(0) & "'," & _
               "puntovtadoccorr='" & txt(1) & "'," & _
               "puntovtadoccorr1='" & txt(2) & "'," & _
               "puntovtaguia1='" & txt(3) & "'," & _
               "usuariocodigo='" & g_usuario & "' " & _
               "WHERE puntovtacodigo='" & i_codigoptovta & "' " & _
               "AND documentocodigo='" & i_codigoptovtadoc & "' " & _
               "AND puntovtadocserie='" & i_ptovtaserie & "'"
    
            VGCNx.Execute SQL
            
  End If
 '******************************************************************************************
        
 TDBGridProducto.Refresh
      
 Mostrar_Data
 MostrarOcultar_Botones (True)
 '''''''''
 modoinsert = False
 modoedit = False
 '''''''''
 SSTab1.TabEnabled(0) = True
 
Exit Sub
nerror:
   If Err Then
      Err = 0
      Resume Next
   End If
       
End Sub

Private Sub cCancela_Click()
    SSTab1.TabEnabled(0) = True
    SSTab1.Tab = 0
    SSTab1.SetFocus
    MostrarOcultar_Botones (True)
    modoinsert = False
    modoedit = False
End Sub

Private Sub cmbDocumento_Click()
  cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub cmbpuntoVta_Click()
    cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub cmbpuntovta_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim J As Integer
  Dim SQL As String
  Dim OBJ As Object
  
  On Error GoTo nerror
  
  SSTab1.TabEnabled(1) = True
  '''''
  Select Case Index
     Case 0   'nuevo
            For Each OBJ In Me.Controls
               If TypeOf OBJ Is TextBox Then
                    OBJ.Text = ""
                End If
                If TypeOf OBJ Is ComboBox Then
                    OBJ.ListIndex = -1
                End If
            Next
            SSTab1.Tab = 1
            modoinsert = True
            MostrarOcultar_Botones (False)
            cmbPuntoVta.SetFocus
        
     Case 1   'modificar
     
         If TDBGridProducto.Row < 0 Then
            Exit Sub
         End If
         
             Call fncSeleccionaCombo(Trim(TDBGridProducto.Columns(2).Text), cmbDocumento, ArregloDocumento)
             Call fncSeleccionaCombo(Trim(TDBGridProducto.Columns(0).Text), cmbPuntoVta, ArregloPuntovta)
             
             i_codigoptovta = TDBGridProducto.Columns(0).Text
             i_codigoptovtadoc = TDBGridProducto.Columns(2).Text
             i_ptovtaserie = TDBGridProducto.Columns(4).Text
            
             txt(0) = Trim(TDBGridProducto.Columns(4).Text)
             txt(1) = Trim(TDBGridProducto.Columns(5).Text)
             txt(2) = Trim(TDBGridProducto.Columns(6).Text)
             txt(3) = Trim(TDBGridProducto.Columns(7).Text)
                 
        modoedit = True
        SSTab1.Tab = 1
        MostrarOcultar_Botones (False)
        i_filaorigen = TDBGridProducto.Row
        cmbPuntoVta.SetFocus
      
        '''''''''
      
     Case 2   'eliminar
     If TDBGridProducto.Row < 0 Then
            Exit Sub
     End If
         
       If MsgBox("Desea eliminar el registro?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          SQL = "DELETE FROM vt_puntovtadocumento WHERE puntovtacodigo = '" & TDBGridProducto.Columns(0).Text & _
                "' AND documentocodigo = '" & TDBGridProducto.Columns(2).Text & _
                "' AND puntovtadocserie = '" & TDBGridProducto.Columns(4).Text & "'"
          VGCNx.Execute SQL
          Mostrar_Data
       End If
        
     Case 3   'imprimir
         oCrystalReport.DiscardSavedData = True
         oCrystalReport.Action = 1
     Case 4  ' salir
       Unload Me
  End Select
   
nerror:
   If Err Then
      Err = 0
      Resume Next
   End If

End Sub

Private Sub Form_Load()
 MostrarFormVentas Me, "C"
 Mostrar_Data
 cAcepta.Enabled = False
 SSTab1.TabEnabled(1) = False
End Sub

Public Function Mostrar_Data()
  Dim SQL As String
  Dim rs As New ADODB.Recordset
  Dim i As Integer
    
      SQL = "SELECT a.puntovtacodigo as Código," & _
      "b.puntovtadescripcion as 'Desc.Pto.Vta'," & _
      "a.documentocodigo as Documento," & _
      "c.documentodescripcion as 'Desc.Docum.'," & _
      "a.puntovtadocserie as 'Serie'," & _
      "a.puntovtadoccorr as 'Correlativo'," & _
      "a.puntovtadoccorr1 as 'Guia Rem.1'," & _
      "a.puntovtaguia1 as 'Guia Rem.2'" & _
      " " & _
      "FROM  vt_puntovtadocumento a " & _
      "      JOIN  vt_puntoventa b ON a.puntovtacodigo = b.puntovtacodigo" & _
      "      JOIN  vt_documento c ON a.documentocodigo = c.documentocodigo " & _
      "ORDER BY a.puntovtacodigo"
      
      Set rs = VGCNx.Execute(SQL)
      Set TDBGridProducto.DataSource = rs
    
      ' COMBO DOCUMENTO:
      SQL = "SELECT documentocodigo,documentodescripcion " & _
      "FROM vt_documento "
      
      Set rs = VGCNx.Execute(SQL)
      If rs.RecordCount > 0 Then
        ReDim ArregloDocumento(0 To 1, 0 To rs.RecordCount - 1)
        Call fncLlenarArreglo_Combo(rs, cmbDocumento, ArregloDocumento, 1)
      End If
    
      '' COMBO GRUPO VENTA:
      SQL = "SELECT puntovtacodigo,puntovtadescripcion " & _
      "FROM vt_puntoventa "
      
      Set rs = VGCNx.Execute(SQL)
      If rs.RecordCount > 0 Then
        ReDim ArregloPuntovta(0 To 1, 0 To rs.RecordCount - 1)
        Call fncLlenarArreglo_Combo(rs, cmbPuntoVta, ArregloPuntovta, 1)
      End If
    
  '    oCrystalReport.ReportFileName = VGParamSistem.Rutareport & "MantPuntoVtaDoc.rpt"
    
 TDBGridProducto.Refresh
 Set rs = Nothing
 SSTab1.Tab = 0
  
End Function


Private Function Validar_DatosNulos() As Boolean

Validar_Ingreso = False

                If Trim(txt(0)) <> "" And cmbPuntoVta.ListIndex <> -1 _
                  And cmbDocumento.ListIndex <> -1 Then
                    Validar_DatosNulos = True
                    Exit Function
                End If

End Function


Private Sub SSTab1_Click(PreviousTab As Integer)
    SSTab1.TabEnabled(PreviousTab) = False
    cAcepta.Enabled = False
End Sub


Private Sub txt_Change(Index As Integer)
cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)  ' Salta con Enter
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    cAcepta.Enabled = Validar_DatosNulos()
    
    'Ingresar Mayusculas:
    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = KeyAscii - 32
    End If

End Sub

Private Function Validar_CodigosDuplicados(operacion As String, Optional filaorigen As Integer) As Boolean
Dim i As Integer
               
Validar_CodigosDuplicados = False
                    
 TDBGridProducto.MoveFirst
   Do Until TDBGridProducto.EOF
      If operacion = "INSERT" Then
         If Trim(ArregloPuntovta(0, cmbPuntoVta.ListIndex)) = _
            Trim(TDBGridProducto.Columns(0).Text) Then
                 If Trim(ArregloDocumento(0, cmbDocumento.ListIndex)) = _
                    Trim(TDBGridProducto.Columns(2).Text) Then
                        If Trim(txt(0)) = _
                           Trim(TDBGridProducto.Columns(4).Text) Then
                             Validar_CodigosDuplicados = True
                             Exit Function
                        End If
                 End If
         End If
         
      ElseIf operacion = "UPDATE" Then
         If Trim(ArregloPuntovta(0, cmbPuntoVta.ListIndex)) = _
            Trim(TDBGridProducto.Columns(0).Text) Then
              If Trim(ArregloDocumento(0, cmbDocumento.ListIndex)) = _
                 Trim(TDBGridProducto.Columns(2).Text) Then
                    If Trim(txt(0)) = _
                       Trim(TDBGridProducto.Columns(4).Text) _
                    And TDBGridProducto.Row <> filaorigen Then
                           Validar_CodigosDuplicados = True
                           Exit Function
                    End If
               End If
         End If
      End If
      TDBGridProducto.MoveNext
  Loop
    
End Function

Private Function MostrarOcultar_Botones(valor As Boolean)
    frmbotones.Visible = valor
End Function

Private Function fncSeleccionaCombo(ValorCodigo As String, Cbo As ComboBox, Arreglo As Variant)
Dim i As Integer
    For i = 0 To UBound(Arreglo, 2)
       If ValorCodigo = Arreglo(0, i) Then
         Cbo.ListIndex = i
         Exit Function
       End If
    Next i
End Function

Private Function fncLlenarArreglo_Combo(rs As Recordset, Cbo As ComboBox, Arreglo As Variant, dimensiones As Integer)
Dim i As Integer
Dim J As Integer

    i = 0
    Cbo.Clear
    Do Until rs.EOF
        Cbo.AddItem (Trim(rs(1)))
        For J = 0 To dimensiones
            Arreglo(J, i) = Trim(rs(J))
        Next J
        rs.MoveNext
        i = i + 1
    Loop
End Function

Public Function Formatear_Codigo(indice As Integer) As String
Dim cadena As String
Dim i As Integer

cadena = ""
For i = 0 To txt(indice).MaxLength
    cadena = cadena & "0"
Next i

txt(indice) = Right(cadena & Trim(txt(indice)), txt(indice).MaxLength)

End Function

Private Sub txt_LostFocus(Index As Integer)
If txt(Index) <> "" Then
    If Index = 0 Or Index = 1 Then
        Call Formatear_Codigo(Index)
    End If
End If
End Sub





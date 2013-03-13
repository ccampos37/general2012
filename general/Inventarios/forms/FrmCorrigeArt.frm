VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmCorrigeart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Correción de Artículos"
   ClientHeight    =   8910
   ClientLeft      =   915
   ClientTop       =   1710
   ClientWidth     =   9345
   Icon            =   "FrmCorrigeArt.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Exportar"
      TabPicture(0)   =   "FrmCorrigeArt.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Command7"
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(4)=   "Command2"
      Tab(0).Control(5)=   "CommonDialog1"
      Tab(0).Control(6)=   "Frame1"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Importar"
      TabPicture(1)   =   "FrmCorrigeArt.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label24"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label25"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DataGrid2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text6"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin VB.Frame Frame2 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   6375
         Left            =   -74880
         TabIndex        =   1
         Top             =   1800
         Visible         =   0   'False
         Width           =   8535
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4815
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   8493
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label21 
            Caption         =   "Cantidad de registros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5160
            TabIndex        =   4
            Top             =   5400
            Width           =   1815
         End
         Begin VB.Label Label22 
            Caption         =   "11"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   7320
            TabIndex        =   3
            Top             =   5400
            Width           =   735
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   7440
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Importar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   7440
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1920
         TabIndex        =   44
         Top             =   840
         Width           =   5415
      End
      Begin VB.CommandButton Command7 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   735
         Left            =   -71400
         Picture         =   "FrmCorrigeArt.frx":0902
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7185
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Corregir"
         Height          =   735
         Left            =   -72960
         Picture         =   "FrmCorrigeArt.frx":0D44
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7200
         Width           =   855
      End
      Begin VB.Frame Frame4 
         Height          =   1185
         Left            =   -74760
         TabIndex        =   6
         Top             =   360
         Width           =   7125
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAlmacen 
            Height          =   375
            Left            =   1560
            TabIndex        =   7
            Top             =   360
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   661
            XcodMaxLongitud =   0
            xcodwith        =   100
            NomTabla        =   "tabalm"
            TituloAyuda     =   "Almacenes"
            ListaCampos     =   "TAALMA(1),TADESCRI(1),empresacodigo(1)"
            XcodCampo       =   "TAALMA"
            XListCampo      =   "TADESCRI"
            ListaCamposDescrip=   "Codigo,Descripcion,empresa"
            ListaCamposText =   "TAALMA,TADESCRI,empresacodigo"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuTransaccion 
            Height          =   375
            Left            =   1590
            TabIndex        =   8
            Top             =   720
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   661
            XcodMaxLongitud =   0
            xcodwith        =   100
            NomTabla        =   "tabtransa"
            TituloAyuda     =   "Transaciones"
            ListaCampos     =   "tt_codmov(1),tt_descri(1)"
            XcodCampo       =   "tt_codmov"
            XListCampo      =   "tt_descri"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "tt_codmov,tt_descri"
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Almacen"
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
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Transaccion :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   9
            Top             =   720
            Width           =   1110
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exportar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -67320
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -70680
         Top             =   4440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   5655
         Left            =   360
         TabIndex        =   47
         Top             =   1560
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   9975
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "Correción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   4905
         Left            =   -74850
         TabIndex        =   13
         Top             =   1860
         Visible         =   0   'False
         Width           =   8130
         Begin VB.ComboBox Combo4 
            Height          =   315
            ItemData        =   "FrmCorrigeArt.frx":1186
            Left            =   1920
            List            =   "FrmCorrigeArt.frx":1193
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   3240
            Width           =   1455
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            ItemData        =   "FrmCorrigeArt.frx":11BE
            Left            =   5280
            List            =   "FrmCorrigeArt.frx":11C8
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1920
            TabIndex        =   18
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   5280
            TabIndex        =   17
            Top             =   3240
            Width           =   1575
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1920
            TabIndex        =   16
            Top             =   3720
            Width           =   1455
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   6165
            TabIndex        =   15
            Top             =   1755
            Width           =   1575
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   1920
            TabIndex        =   14
            Top             =   4200
            Width           =   1455
         End
         Begin VB.Label Label20 
            Caption         =   "Label1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2640
            TabIndex        =   43
            Top             =   720
            Width           =   3975
         End
         Begin VB.Label Label14 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1920
            TabIndex        =   42
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label13 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1920
            TabIndex        =   41
            Top             =   1080
            Width           =   4695
         End
         Begin VB.Label Label12 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1920
            TabIndex        =   40
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label11 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1920
            TabIndex        =   39
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label10 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1920
            TabIndex        =   38
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo de Cambio"
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   37
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Conversion"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   480
            TabIndex        =   36
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Factura"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   480
            TabIndex        =   35
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Moneda"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   3960
            TabIndex        =   34
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Serie"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   480
            TabIndex        =   33
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Doc Referencial"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   480
            TabIndex        =   32
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Proveedor"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   480
            TabIndex        =   31
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Transacción"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   480
            TabIndex        =   30
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Documento"
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   29
            Top             =   375
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Costo Unitario"
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   28
            Top             =   3720
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Cantidad"
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   2
            Left            =   3960
            TabIndex        =   27
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Total"
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   26
            Top             =   4200
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "Código Artículo"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   495
            TabIndex        =   25
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label16 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1920
            TabIndex        =   24
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label17 
            Caption         =   "Descripción"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   480
            TabIndex        =   23
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label Label18 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1920
            TabIndex        =   22
            Top             =   2520
            Width           =   4935
         End
         Begin VB.Label Label19 
            Caption         =   "Label1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2640
            TabIndex        =   21
            Top             =   1440
            Width           =   2055
         End
      End
      Begin VB.Label Label25 
         Caption         =   "Cantidad de registros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5280
         TabIndex        =   49
         Top             =   7560
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   7560
         TabIndex        =   48
         Top             =   7560
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Archivo a importar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   46
         Top             =   840
         Width           =   1590
      End
   End
End
Attribute VB_Name = "FrmCorrigeart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''Dim db As Database
Dim objExcel As Excel.Application
Dim rs As ADODB.Recordset
Dim rssql As New ADODB.Recordset
Dim RSQL As String
Dim precioant As Double  'costo anterior
Dim sCodMon As String 'El codigo de moneda
Dim Fecha As Date     'Graba la Fecha del Documento
Dim rsSTKART As New ADODB.Recordset
Private Sub Command1_Click()
  Dim precio As Double  ' corregir la
  Dim CANTIDAD As Double
  Dim uSql As String
  Dim RSQL As String
  Dim cant As String
  Dim Serie As String
  Dim Lote As String
  Dim codmon As String
  Dim cacodmov As String
  
  If Frame1.Visible Then
        'Text3.Text = "0"
        If Not IsNumeric(Text3) Then
                MsgBox "Ingrese el Precio unitario !", vbOKOnly + vbExclamation, "Error"
                Text3.SetFocus
                Exit Sub
        End If
        If Not IsNumeric(Text4) Then
                MsgBox "Ingrese la cantidad !", vbOKOnly + vbExclamation, "Error"
                Text4.SetFocus
                Exit Sub
        End If
        If Combo3.ListIndex <> 0 Then
                If Val(Text2) = 0 Then
                    MsgBox "Ingrese el tipo de cambio !", vbOKOnly + vbExclamation, "Error"
                    Text2.SetFocus
                    Exit Sub
                End If
        End If
        If Combo3.ListIndex = 0 Then
            codmon = "01"
        Else
            codmon = "02"
        End If
        If sCodMon <> codmon Then
            If MsgBox("Desea Ud. cambiar el Tipo de moneda declarado inicialmente?", vbYesNo, "Aviso") = vbNo Then
                Exit Sub
            End If
        End If
        If codmon = "01" Then
            precio = Val(Text3.text)
        Else
            precio = Val(Text3.text) '* Val(Text2)
        End If
        CANTIDAD = Val(Text4.text)
        uSql = "Update MovAlmCab set CACODMON = '" & codmon & "', CATIPCAM = " & Val(Text2) & " where CANUMDOC='" & rssql!denumdoc & "' and CAALMA = '" & VGAlma & "'    AND CATD='" & rssql!detd & "' "
        VGCNx.Execute uSql
        uSql = "Update MovAlmDet set DEPRECIO = " & precio & ",DETIPCAM = " & Val(Text2) & " ,DECODMON = '" & codmon & "' where  DEALMA ='" & VGAlma & "'  and DENUMDOC='" & rssql!denumdoc & "' and DECODIGO ='" & rssql!decodigo & "' and DeTD='" & rssql!detd & "'"
        VGCNx.Execute uSql
        Frame1.Visible = False
        Text4 = ""
        Text5 = ""
   Else
        Text2 = "0"
        Text3 = "0"
        Frame1.Visible = True
        Command1.Caption = "&Aceptar"
        RSQL = "select  cacodmov,cacodmon,cafecdoc, canompro from  MovAlmCab  where   CAALMA ='" & VGAlma & "'  and CATD='" & rssql!detd & "' AND CANUMDOC= '" & rssql!denumdoc & "'"
                Set rs = VGCNx.Execute(RSQL)
        If Not rs.EOF Then
            If rs("CACODMON") = "01" Then
                Combo3.ListIndex = 0
                sCodMon = "01"
            Else
                Combo3.ListIndex = 1
                sCodMon = "02"
            End If
         Fecha = rs!CAFECDOC
         cacodmov = rs!cacodmov
         End If
         rs.Close
         Set rs = Nothing
         RSQL = "select   deprecio=isnull(n.DEPRECIO,0),detipcam=isnull(n.DETIPCAM,0),decantid from  MovAlmDet n where   n.DEALMA ='" & VGAlma & "' and DETD='" & rssql!detd & "' AND n.DECODIGO='" & rssql!decodigo & "' and n.DENUMDOC= '" & rssql!denumdoc & "'"
          Set rs = VGCNx.Execute(RSQL)
         If rs.EOF Then
                Exit Sub
         End If
        Text4.text = rs!DECANTID

         If sCodMon = "01" Then
             precioant = rs(0)
         Else
             If IsNull(rs(1)) Then
                MsgBox "Ud no ha ingresado el tipo de cambio", vbInformation, "Aviso"
             End If
             If rs(1) <> 0 Then
                precioant = rs(0)
             Else
                precioant = rs(0)
             End If
         End If
         Text2 = IIf(Not IsNull(rs(1)), rs(1), 0)
         Text3 = Round(precioant, 4)
          Label10 = rssql!denumdoc
          Label11 = cacodmov
          Label16 = rssql!decodigo
          Label18 = rssql!codigodescripcion
          Label13 = rssql!proveedor     ' proveedor
          Label12 = rssql!CARFTDOC
          Text1 = rssql!CARFNDOC
          If Label12 <> "" Then
              Label19 = tipref(Label12)
          End If
          If Label11 <> "" Then Label20 = transa(Label11)
          Call cantidad_art(cant, Serie, Lote)
          Text4.Enabled = True
          Text4 = cant
          If Lote = "" Then
                Label14 = Serie
          Else
                Label14 = Lote
          End If
          Text4.Enabled = False
          Text3.SetFocus
 End If
End Sub

Private Sub Command2_Click()
Call exportarExcel(rssql, " Movimientos Valorizados")
End Sub

Private Sub Command3_Click()
  Dim xrsql As New ADODB.Recordset
  Dim mesproceso As String
  If ExisteElem(0, VGCNx, "##xx") Then VGCNx.Execute "DROP TABLE ##xx"
  mesproceso = Format(VGParamSistem.FechaTrabajo, "yyyy") + Format(VGParamSistem.FechaTrabajo, "mm")
  RSQL = "Select top 0 DECODIGO, codigodescripcion,dealma, DETD,DENUMDOC,deitem, decantid,deprecio,transacciondescripcion, "
  RSQL = RSQL & "proveedor=CANOMPRO,CARFTDOC,CARFNDOC into ##xx from v_kardexvalorizado "
  Set xrsql = VGCNx.Execute(RSQL)
  xrsql.Open "select * from ##xx", VGCNx, adOpenDynamic, adLockBatchOptimistic
    
 Call importarExcel(xrsql, Text6.text, DataGrid2)
 Label24.Caption = xrsql.RecordCount
End Sub

Private Sub Command4_Click()
  CommonDialog1.InitDir = App.Path
  CommonDialog1.Filter = "Excel (*.xlsx)|*.xlsx"
  CommonDialog1.FilterIndex = 1
  CommonDialog1.ShowOpen
  Text6 = CommonDialog1.FileName
  Command3.Enabled = True
End Sub

Private Sub command5_Click()
Call grabacion(rssql)
End Sub

Private Sub Command7_Click()
  If Frame1.Visible Then
        Frame1.Visible = False
        Command1.Caption = "&Corregir"
        Text4 = ""
        Text5 = ""
  Else
        Unload Me
  End If
End Sub

Private Sub Form_Load()
Call Ctr_AyuAlmacen.conexion(VGCNx)
Call Ctr_AyuTransaccion.conexion(VGCNx)
Ctr_AyuTransaccion.filtro = " estadocosto=1 and tipodecosto='C'"
Command3.Enabled = False
End Sub
Private Sub Ctr_ayuAlmacen_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
VGAlma = Ctr_AyuAlmacen.xclave
End Sub
Private Sub Ctr_AyuTransaccion_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Frame2.Visible = True
cargar
End Sub
Private Sub cargar()
  Dim mesproceso As String
  'Text1.SetFocus
  Label14 = ""
  Label19 = ""
  Combo3.ListIndex = 0
  Combo4.ListIndex = 0

  mesproceso = Format(VGParamSistem.FechaTrabajo, "yyyy") + Format(VGParamSistem.FechaTrabajo, "mm")
  RSQL = "Select  DECODIGO, codigodescripcion,dealma, DETD,DENUMDOC,deitem, decantid,deprecio,transacciondescripcion, "
  RSQL = RSQL & "proveedor=CANOMPRO,CARFTDOC,CARFNDOC from v_kardexvalorizado WHERE   CaALMA ='" & VGAlma & "'  AND isnull(CACIERRE,0)=0  "
  RSQL = RSQL & " and catd='NI' and estadocosto=1 and mesproceso='" & mesproceso & "'"
  RSQL = RSQL & " and cacodmov='" & Ctr_AyuTransaccion.xclave & "' oRDER BY DECODIGO, DENUMDOC"
   
  Set rssql = VGCNx.Execute(RSQL)
  
If rssql.RecordCount > 0 Then Label22 = rssql.RecordCount
  If rssql.EOF Then
     MsgBox "No hay articulo valorizados para corregir", vbInformation, mensaje1
      central Me
     Exit Sub
  End If
  Set DataGrid1.DataSource = rssql
  DataGrid1.Refresh
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
       SendKeys "{tab}"
  End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(Text2) Then
      SendKeys "{tab}"
    Else
      If Chr$(KeyAscii) = "." Then Exit Sub
      If ((Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0")) And KeyAscii <> 8 Then KeyAscii = 0
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 And IsNumeric(Text3) Then
                Command1.SetFocus
                Exit Sub
      End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If IsNumeric(Text4) And KeyAscii = 13 And IsNumeric(Text3) Then
        If Not IsNumeric(Text3) Then Exit Sub
        Text5 = Val(Text3) * Val(Text4)
   ElseIf KeyAscii = 13 And IsNumeric(Text5) <> 0 And IsNumeric(Text4) <> 0 Then
         Text3 = Format(Val(Text5) / Val(Text4), "##0.0000")
   Else
        If Chr$(KeyAscii) = "." Then Exit Sub
        If ((Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0")) And KeyAscii <> 8 Then KeyAscii = 0
   End If
End Sub

Private Sub Text5_Change()
If Trim(Text4) <> "" Then
   Text3 = Format(Val(Text5) / Val(Text4), "###0.0000")
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And IsNumeric(Text5) And IsNumeric(Text4) Then
      Text3 = Val(Text5) / Val(Text4)
      Text3 = Format(Text3, "##0.0000")
      Command1.SetFocus
   Else
    If Chr$(KeyAscii) = "." Then Exit Sub
    If ((Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0")) And KeyAscii <> 8 Then KeyAscii = 0
   End If
End Sub

Function tipref(text As Label) As String
 Dim rs As Recordset
 Dim RSQL As String
  RSQL = "select  TDO_DESCRI  FROM TIPO_DOCU  where TDO_TIPDOC = '" & text & "'" '
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  
  Set rs = VGCNx.Execute(RSQL)
  tipref = IIf(Not rs.EOF, rs(0), "")
  rs.Close
End Function

Function transa(text As Label) As String
 Dim rs As Recordset
 Dim RSQL As String
 Dim dato As String
  dato = "I"
  RSQL = "select  TT_DESCRI FROM TabTransa where TT_CODMOV= '" & text & "' and TT_TIPMOV ='" & dato & "'" '
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  Set rs = VGCNx.Execute(RSQL)
  transa = IIf(Not rs.EOF, rs(0), "")
  rs.Close
End Function

Private Sub cantidad_art(pcantidad As String, pserie As String, plote As String)
 Dim Adoreg1 As ADODB.Recordset
 Dim RSQL As String
 RSQL = "select decantid,delote,deserie from MovAlmdet where DENUMDOC='" & rssql!denumdoc & "' and DECODIGO ='" & rssql!decodigo & "' and DEALMA = '" & VGAlma & "'  AND DETD='" & rssql!detd & "' "
 Set Adoreg1 = New ADODB.Recordset
Adoreg1.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
 If Adoreg1.RecordCount = 0 Then
    pcantidad = ""
    pserie = ""
    plote = ""
  Else
    pcantidad = Str(Adoreg1(0))
    pserie = IIf(Not IsNull(Adoreg1(2)), Adoreg1(2), "")
    plote = IIf(Not IsNull(Adoreg1(1)), Adoreg1(1), "")
  End If
End Sub



Private Sub grabacion(rrss As Recordset)
Dim act As New ADODB.Recordset
Dim rr As New ADODB.Recordset
SQL = " select * from ##xx"
Set rr = VGCNx.Execute(SQL)
If rr.RecordCount = 0 Then Exit Sub
rr.MoveFirst
Do While Not rr.EOF()
   SQL = " update movalmdet set deprecio=" & rr!DEPRECIO & " where dealma='" & rr!dealma & "' and detd='" & rr!detd & "'"
   SQL = SQL & " and denumdoc='" & Format(rr!denumdoc, "00000000000") & "' and deitem=" & rr!DEITEM & ""
   SQL = SQL & " and deprecio<> " & rr!DEPRECIO & ""
   Set act = VGCNx.Execute(SQL)
   rr.MoveNext
Loop
Command5.Enabled = False
End Sub

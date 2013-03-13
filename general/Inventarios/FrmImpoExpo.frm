VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmImpoExpo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asistente para exportar"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "FrmImpoExpo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar PB2 
      Height          =   150
      Left            =   -15
      TabIndex        =   30
      Top             =   5625
      Visible         =   0   'False
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   210
      Left            =   -30
      TabIndex        =   29
      Top             =   5205
      Visible         =   0   'False
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   4335
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpoExpo.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpoExpo.frx":2B4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cmdcancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   5655
      TabIndex        =   3
      Top             =   4515
      Width           =   1170
   End
   Begin VB.CommandButton cmdsigue 
      Caption         =   "Siguiente >"
      Height          =   345
      Left            =   4335
      TabIndex        =   2
      Top             =   4515
      Width           =   1170
   End
   Begin VB.CommandButton cmdAtras 
      Caption         =   "< &Atras"
      Height          =   345
      Left            =   3165
      TabIndex        =   1
      Top             =   4515
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4260
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7514
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmImpoExpo.frx":3DCE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "FrmImpoExpo.frx":3DEA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "FrmImpoExpo.frx":3E06
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "FrmImpoExpo.frx":3E22
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame7 
         BackColor       =   &H00808080&
         Height          =   4350
         Left            =   -75015
         TabIndex        =   35
         Top             =   -90
         Width           =   7035
         Begin MSDataGridLib.DataGrid DG1 
            Height          =   1605
            Left            =   300
            TabIndex        =   48
            Top             =   2235
            Width           =   6600
            _ExtentX        =   11642
            _ExtentY        =   2831
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   0
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
               MarqueeStyle    =   2
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   15
            TabIndex        =   36
            Top             =   0
            Width           =   7020
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Elementos relacionados a la guias"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   225
               TabIndex        =   37
               Top             =   150
               Width           =   3180
            End
         End
         Begin MSComctlLib.ListView LvRela 
            Height          =   1080
            Left            =   270
            TabIndex        =   38
            Top             =   825
            Width           =   6645
            _ExtentX        =   11721
            _ExtentY        =   1905
            View            =   2
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   64
            BackColor       =   15925247
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Tablas"
               Object.Width           =   5009
            EndProperty
         End
         Begin VB.Label LbregExp 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0 "
            Height          =   270
            Left            =   6000
            TabIndex        =   49
            Top             =   3945
            Width           =   900
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Tablas relacionadas a las guias"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   315
            TabIndex        =   40
            Top             =   540
            Width           =   2535
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Datos"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   285
            TabIndex        =   39
            Top             =   1950
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00808080&
         Height          =   4335
         Left            =   -75015
         TabIndex        =   5
         Top             =   -90
         Width           =   7035
         Begin VB.TextBox txOrigen 
            BackColor       =   &H00F2FFFF&
            ForeColor       =   &H00404040&
            Height          =   300
            Left            =   1305
            TabIndex        =   15
            Top             =   1260
            Width           =   5130
         End
         Begin VB.CommandButton CmdDestino 
            Caption         =   "..."
            Height          =   315
            Left            =   6030
            TabIndex        =   14
            Top             =   3810
            Width           =   375
         End
         Begin VB.TextBox TxDestino 
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   1395
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   3795
            Width           =   4605
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00808080&
            ForeColor       =   &H00FFFFFF&
            Height          =   1965
            Left            =   555
            TabIndex        =   9
            Top             =   1605
            Width           =   5865
            Begin VB.TextBox TxAlmacen 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1935
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   1005
               Width           =   3180
            End
            Begin VB.CommandButton CmdAlmacen 
               Caption         =   "..."
               Height          =   270
               Left            =   5130
               TabIndex        =   24
               Top             =   1035
               Width           =   330
            End
            Begin MSComCtl2.DTPicker DpFechini 
               Height          =   300
               Left            =   1920
               TabIndex        =   21
               Top             =   1515
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               Format          =   24641537
               CurrentDate     =   37131
            End
            Begin VB.CommandButton CmdClie 
               Caption         =   "..."
               Height          =   270
               Left            =   5130
               TabIndex        =   19
               Top             =   735
               Width           =   330
            End
            Begin VB.CommandButton CmdTipdoc 
               Caption         =   "..."
               Enabled         =   0   'False
               Height          =   285
               Left            =   5130
               TabIndex        =   18
               Top             =   420
               Width           =   330
            End
            Begin VB.TextBox TxCli 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1935
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   705
               Width           =   3180
            End
            Begin VB.TextBox TxTipdoc 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1935
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   405
               Width           =   3180
            End
            Begin MSComCtl2.DTPicker DpFechfin 
               Height          =   300
               Left            =   3780
               TabIndex        =   23
               Top             =   1515
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               Format          =   24641537
               CurrentDate     =   37131
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Almacen :"
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   345
               TabIndex        =   26
               Top             =   1065
               Width           =   1605
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "al"
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   3450
               TabIndex        =   22
               Top             =   1560
               Width           =   390
            End
            Begin VB.Label xFechIni 
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha inicio :"
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   345
               TabIndex        =   20
               Top             =   1575
               Width           =   1440
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente :"
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   345
               TabIndex        =   11
               Top             =   765
               Width           =   1605
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Documento :"
               ForeColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   345
               TabIndex        =   10
               Top             =   435
               Width           =   1680
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Height          =   1140
            Left            =   0
            TabIndex        =   6
            Top             =   -15
            Width           =   7035
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Echo por Fernando Cossio"
               Height          =   210
               Left            =   375
               TabIndex        =   50
               Top             =   825
               Width           =   4125
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Eligir origen y tipo de documento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   720
               TabIndex        =   7
               Top             =   270
               Width           =   3045
            End
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Destino :"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   570
            TabIndex        =   12
            Top             =   3855
            Width           =   1290
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Origen :"
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   600
            TabIndex        =   8
            Top             =   1305
            Width           =   1410
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4365
         Left            =   -15
         TabIndex        =   4
         Top             =   -105
         Width           =   7065
         Begin VB.CheckBox ChkRelacion 
            Caption         =   "Generar tablas relacionadas"
            Height          =   285
            Left            =   2490
            TabIndex        =   34
            Top             =   3975
            Width           =   2685
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Echo por Fernando Cossio"
            Height          =   210
            Left            =   150
            TabIndex        =   51
            Top             =   1590
            Width           =   4125
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00B3BEC1&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   3465
            Left            =   465
            Top             =   495
            Width           =   1320
         End
         Begin VB.Label Label12 
            Caption         =   "Nota. El archivo a generarse es de tipo (Acces-97 *. MDB ) , su nomenclatura es el Ruc del cliente y las fechas de consulta"
            Height          =   705
            Left            =   2460
            TabIndex        =   33
            Top             =   2550
            Width           =   4350
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   2475
            Picture         =   "FrmImpoExpo.frx":3E3E
            Top             =   3240
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   2445
            Picture         =   "FrmImpoExpo.frx":4C80
            Top             =   1320
            Width           =   480
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Asistente para Exportar documentos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3075
            TabIndex        =   28
            Top             =   450
            Width           =   3390
         End
         Begin VB.Label Label8 
            Caption         =   $"FrmImpoExpo.frx":5EF2
            Height          =   990
            Left            =   2445
            TabIndex        =   27
            Top             =   1875
            Width           =   4395
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   945
            Left            =   2175
            Top             =   150
            Width           =   4860
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00808080&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   4185
            Left            =   75
            Top             =   150
            Width           =   2100
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Height          =   1140
         Left            =   -75000
         TabIndex        =   41
         Top             =   -105
         Width           =   7035
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Seleccionar los documentos de origen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   720
            TabIndex        =   42
            Top             =   270
            Width           =   3765
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00808080&
         Height          =   4350
         Left            =   -75015
         TabIndex        =   43
         Top             =   -105
         Width           =   7035
         Begin VB.CheckBox ChKMarca 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "Marcar Todos"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   435
            TabIndex        =   44
            Top             =   1275
            Width           =   1485
         End
         Begin MSComctlLib.ListView LvDocu 
            Height          =   2355
            Left            =   435
            TabIndex        =   45
            Top             =   1560
            Width           =   6390
            _ExtentX        =   11271
            _ExtentY        =   4154
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FlatScrollBar   =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   64
            BackColor       =   15925247
            Appearance      =   0
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "TD"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Serie"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Número"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Almacen"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro Doc :"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   5055
            TabIndex        =   47
            Top             =   4020
            Width           =   810
         End
         Begin VB.Label LbReg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0 "
            Height          =   255
            Left            =   6060
            TabIndex        =   46
            Top             =   3990
            Width           =   765
         End
      End
   End
   Begin VB.Label LBDet 
      Caption         =   "Registrando los detalles"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   -15
      TabIndex        =   32
      Top             =   5415
      Visible         =   0   'False
      Width           =   6045
   End
   Begin VB.Label lbcab 
      Caption         =   "Exportando Cabeceras"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   0
      TabIndex        =   31
      Top             =   4995
      Visible         =   0   'False
      Width           =   5730
   End
End
Attribute VB_Name = "FrmImpoExpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BuscarCarpeta As WindowsUtility
Dim Cnx As New adodb.Connection 'Conexion general de la nueva base de
Dim RSAUX As adodb.Recordset
                                'datos creada
Private Sub ChKMarca_Click()
    If ChKMarca.Value = 1 Then
        Call MarcarList(True)
     Else
        Call MarcarList(False)
    End If
End Sub

Private Sub ChkRelacion_Click()
    If ChkRelacion.Value = 1 Then
       SSTab1.TabVisible(2) = True
      Else
       SSTab1.TabVisible(2) = False
    End If
End Sub

Private Sub CmdAlmacen_Click()
    Dim RSALMA As adodb.Recordset
    Set RSALMA = New adodb.Recordset
    RSALMA.Open "SELECT TAALMA,TADESCRI FROM TABALM ", cConexCom, adOpenKeyset, adLockReadOnly
    frmref.Conectar RSALMA
    frmref.Show 1
    If vGUtil(1) <> "" Then
        TxAlmacen.Tag = vGUtil(1): TxAlmacen = vGUtil(2)
      Else: TxAlmacen.Tag = "": TxAlmacen = ""
    End If
End Sub

Private Sub cmdAtras_Click()
    Dim Cont As Integer
    Cont = 1
    Do While Not SSTab1.TabVisible(SSTab1.Tab - Cont)
        Cont = Cont + 1
    Loop
    SSTab1.Tab = SSTab1.Tab - Cont
End Sub

Private Sub CmdCancelar_Click()
    Set Cnx = Nothing
    Unload Me
End Sub

Private Sub CmdClie_Click()
    Dim RSCLI As adodb.Recordset
    Set RSCLI = New adodb.Recordset
    RSCLI.Open "SELECT CNUMRUC,CNOMCLI FROM MAECLI WHERE CNUMRUC <>''", cConexCom, adOpenKeyset, adLockReadOnly
    frmref.Conectar RSCLI
    frmref.Show 1
    If vGUtil(1) <> "" Then
        TxCli.Tag = vGUtil(1): TxCli = vGUtil(2)
      Else: TxCli.Tag = "": TxCli = ""
    End If
End Sub

Private Sub cmdsigue_Click()
Dim Cont As Integer

    If SSTab1.Tab + 1 = 2 Then
        If Not Validar Then Exit Sub
    End If
    Screen.MousePointer = 11
    If UCase(cmdsigue.Caption) = "&FINALIZAR" Then
        Call Finalizar
    End If
    Cont = 1
    If SSTab1.Tab < 3 Then
        Do While Not SSTab1.TabVisible(SSTab1.Tab + Cont)
            Cont = Cont + 1
        Loop
        SSTab1.Tab = SSTab1.Tab + Cont
    End If
    Screen.MousePointer = 1
End Sub
Private Sub VistaTablasRela()
    Dim rsVista As adodb.Recordset
    Dim itmX As ListItem
    LvRela.ListItems.Clear
    Set itmX = LvRela.ListItems.Add(, "MAEART", "Tabla de Articulos", , 2): itmX.Tag = "ACODIGO-ACODIGO"
    Set rsVista = New adodb.Recordset
    'Verificando si existen unidades
    rsVista.Open "SELECT DISTINCT AUNIDAD FROM TMPSELECT WHERE TRIM(AUNIDAD)<>''", cConexCom
    If rsVista.RecordCount > 0 Then Set itmX = LvRela.ListItems.Add(, "TABUNIMED", "Tabla de Unidades", , 2): itmX.Tag = "AUNIDAD-UM_ABREV"
    Set rsVista = New adodb.Recordset
    'Verificando si existen Familias
    rsVista.Open "SELECT DISTINCT AFAMILIA FROM TMPSELECT WHERE TRIM(AFAMILIA)<>''", cConexCom
    If rsVista.RecordCount > 0 Then Set itmX = LvRela.ListItems.Add(, "FAMILIA", "Tabla de Familias", , 2): itmX.Tag = "AFAMILIA-FAM_CODIGO"
    Set rsVista = New adodb.Recordset
    'Verificando si existen lineas
    rsVista.Open "SELECT DISTINCT AMODELO FROM TMPSELECT WHERE TRIM(AMODELO)<>''", cConexCom
    If rsVista.RecordCount > 0 Then Set itmX = LvRela.ListItems.Add(, "LINEAS", "Tabla de Lineas", , 2): itmX.Tag = "AMODELO-LIN_CODIGO"
    Set rsVista = New adodb.Recordset
    'Verificando si existen Grupos
    rsVista.Open "SELECT DISTINCT AGRUPO FROM TMPSELECT WHERE TRIM(AGRUPO)<>''", cConexCom
    If rsVista.RecordCount > 0 Then Set itmX = LvRela.ListItems.Add(, "GRUPO", "Tabla de Grupos", , 2): itmX.Tag = "AGRUPO-GRU_CODIGO"
    Set rsVista = New adodb.Recordset
    'Verificando si existen Tipo_Articulo
    rsVista.Open "SELECT DISTINCT ATIPO FROM TMPSELECT WHERE TRIM(ATIPO)<>''", cConexCom
    If rsVista.RecordCount > 0 Then Set itmX = LvRela.ListItems.Add(, "TIPO_ARTICULO", "Tabla Tipos de Articulos", , 2): itmX.Tag = "ATIPO-COD_TIPO"
    Set rsVista = New adodb.Recordset
    'Verificando si existen Tallas
    rsVista.Open "SELECT DISTINCT TALLA FROM TMPSELECT WHERE TRIM(TALLA)<>''", cConexCom
    If rsVista.RecordCount > 0 Then Set itmX = LvRela.ListItems.Add(, "TALLA", "Tabla de Tallas", , 2): itmX.Tag = "TALLA-CODIGO"
    Set rsVista = New adodb.Recordset
    'Verificando si existen Clases
    rsVista.Open "SELECT DISTINCT AMARCA FROM TMPSELECT WHERE TRIM(AMARCA)<>''", cConexCom
    If rsVista.RecordCount > 0 Then Set itmX = LvRela.ListItems.Add(, "MAEMARCA", "Tabla de Clases", , 2): itmX.Tag = "AMARCA-COD_MARCA"
    Set rsVista = New adodb.Recordset
    'Verificando si existen Colores
    rsVista.Open "SELECT DISTINCT ACOLOR FROM TMPSELECT WHERE TRIM(ACOLOR)<>''", cConexCom
    If rsVista.RecordCount > 0 Then Set itmX = LvRela.ListItems.Add(, "MAECOLOR", "Tabla de Colores", , 2): itmX.Tag = "ACOLOR-COD_COLOR"
    
End Sub
Private Sub Habilitar_Mov_Bot(Atras As Boolean, Siguiente As Boolean, Optional Cap As String)
    cmdAtras.Enabled = Atras
    cmdsigue.Enabled = Siguiente
    cmdsigue.Caption = Cap
    
End Sub

Private Sub CmdTipdoc_Click()
    Dim RSTIPDOC As adodb.Recordset
    Set RSTIPDOC = New adodb.Recordset
    RSTIPDOC.Open "SELECT TDO_TIPDOC,TDO_DESCRI FROM TIPO_DOCU", cConexCom, adOpenKeyset, adLockReadOnly
    frmref.Conectar RSTIPDOC
    frmref.Show 1
    If vGUtil(1) <> "" Then
        TxTipdoc.Tag = vGUtil(1): TxTipdoc = vGUtil(2)
      Else: TxTipdoc.Tag = "": TxTipdoc = ""
    End If
End Sub

Private Sub CmdDestino_Click()
    Set BuscarCarpeta = New WindowsUtility
    TxDestino = BuscarCarpeta.BrowseForShares(Me.hwnd, "Seleccione la carpeta destino")
End Sub

Private Sub Form_Load()
    Me.Width = 7080: Me.Height = 5355
    Me.txOrigen.Tag = VGCOMP
    Me.txOrigen = VGNemp
    SSTab1.Tab = 0
    ChKMarca.Value = 0
    TxTipdoc = "Guia de Remision"
    TxTipdoc.Tag = "GS"
End Sub

Private Sub LvDocu_ItemCheck(ByVal item As MSComctlLib.ListItem)
    item.Checked = True
End Sub

Private Sub LvRela_ItemClick(ByVal item As MSComctlLib.ListItem)
    Set RSAUX = New adodb.Recordset
    Dim CODIGOTEM As String, CODIGOTABLA As String
    CODIGOTEMP = Mid(item.Tag, 1, InStr(item.Tag, "-") - 1)
    CODIGOTABLA = Mid(item.Tag, InStr(item.Tag, "-") + 1)
    RSAUX.Open "Select * from " & item.key & " Where " & CODIGOTABLA & " IN(" & "SELECT DISTINCT " & CODIGOTEMP & " FROM TMPSELECT)", cConexCom
    LbregExp.Caption = Format(RSAUX.RecordCount, "0 ")
    Set DG1.DataSource = RSAUX
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
        Case 0: Call Habilitar_Mov_Bot(False, True, "&Siguiente >")
        Case 1: Call Habilitar_Mov_Bot(True, True, "&Siguiente >")
        Case 2
            Call Llenar_Lista
            If ChkRelacion.Value = 0 Then
                Call Habilitar_Mov_Bot(True, True, "&Finalizar")
              Else
              Call VistaTablasRela
              Call Habilitar_Mov_Bot(True, True, "&Siguiente >")
            End If
        Case 3: Call Habilitar_Mov_Bot(True, True, "&Finalizar")
    End Select
End Sub
Private Function Seleccionardocumentos(doc As String, RucCliente As String, _
FechaIni As Date, FechaFin As Date, Optional almacen As String = "") As adodb.Recordset
Dim sqlcad As String, sqlvar As String
    If ExisteElem(0, cConexCom, "TMPSELECT") Then cConexCom.Execute "DROP TABLE TMPSELECT"

    Set Seleccionardocumentos = New adodb.Recordset
    If almacen <> "" Then sqlvar = " AND CAALMA='" & Trim(almacen) & "'"
    sqlcad = "Select * from MOVALMCAB WHERE CATD='" & doc & "' AND CARUC='" & RucCliente & "' AND CAFECDOC Between " & DateSQL(DpFechini) & " AND " & DateSQL(DpFechfin)
    Seleccionardocumentos.Open sqlcad & sqlvar, cConexCom, adOpenKeyset, adLockReadOnly
    'Generando el temporal
    sqlcad = "SELECT * INTO TMPSELECT FROM MAEART WHERE ACODIGO IN (Select DECODIGO from MOVALMCAB,MOVALMDET WHERE DEALMA=CAALMA and DETD=CATD AND DENUMDOC=CANUMDOC AND CATD='" & doc & "' AND CARUC='" & RucCliente & "' AND CAFECDOC Between " & DateSQL(DpFechini) & " AND " & DateSQL(DpFechfin) & ")"
    cConexCom.Execute sqlcad & sqlvar
End Function
Private Sub Llenar_Lista()
    Dim rslista As New adodb.Recordset
    Dim itmX As ListItem
    Set rslista = Seleccionardocumentos(TxTipdoc.Tag, TxCli.Tag, DpFechini, DpFechfin)
    LvDocu.ListItems.Clear
    Do While Not rslista.EOF
        Set itmX = LvDocu.ListItems.Add(, ESNULO(rslista!CATD, "") & ESNULO(rslista!CANUMDOC, ""), ESNULO(rslista!CATD, ""), , 1)
        itmX.SubItems(1) = Mid(ESNULO(rslista!CANUMDOC, ""), 1, 3)
        itmX.SubItems(2) = Mid(ESNULO(rslista!CANUMDOC, ""), 4)
        itmX.SubItems(3) = ESNULO(rslista!CAALMA, "")
        rslista.MoveNext
    Loop
    LbReg.Caption = Format(rslista.RecordCount, "0 ")
    ChKMarca.Value = 1
    ChKMarca_Click
    ChKMarca.Enabled = False
End Sub
Private Sub CrearBaseDatos(Nombre As String, RUTA As String)
'    'Dim wrkPredeterminado As Workspace
'    Dim dbsNueva As Database
'    Dim prpBucle As Property
'    Dim RutaNombre As String
'    RutaNombre = Trim(RUTA & "\" & Nombre)
'    ' Obtiene el Workspace predeterminado.
'    Set wrkPredeterminado = DBEngine.Workspaces(0)
'    ' Asegúrese de que no existe un archivo con el
'    ' nombre de la base de datos nueva.
'    If Dir(RutaNombre) <> "" Then Kill RutaNombre
'    ' Crea a una base de datos nueva encriptada con la
'    ' secuencia de intercalación especificada.
'    Set dbsNueva = wrkPredeterminado.CreateDatabase(RutaNombre, _
'    dbLangGeneral, dbEncrypt)
'    dbsNueva.Close
End Sub
Private Sub CrearTablasBase(Nombre As String, RUTA As String)
'Dim ERwinWorkspace As Workspace
'Dim ERwinDatabase As Database
'Dim ERwinTableDef As TableDef
'Dim ERwinQueryDef As QueryDef
'Dim ERwinIndex As Index
'Dim ERwinField As Field
'Dim ERwinRelation As Relation
'
'Dim NombreRuta As String
'NombreRuta = RUTA & "\" & Nombre
'
'    Set ERwinWorkspace = DBEngine.Workspaces(0)
'
'    Set ERwinDatabase = ERwinWorkspace.OpenDatabase(NombreRuta)
'
'    '  CREATE TABLE "MOVALMCAB"
'    Set ERwinTableDef = ERwinDatabase.CreateTableDef("MOVALMCAB")
'    Set ERwinField = ERwinTableDef.CreateField("CAALMA", dbText, 2)
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CATD", dbText, 2)
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CANUMDOC", dbText, 10)
'    ERwinField.Required = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAFECDOC", dbDate)
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CATIPMOV", dbText, 1)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CACODMOV", dbText, 2)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CASITUA", dbText, 1)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CARFTDOC", dbText, 2)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CARFNDOC", dbText, 10)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CASOLI", dbText, 3)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAFECDEV", dbText, 8)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CACODPRO", dbText, 11)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CACENCOS", dbText, 6)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CARFALMA", dbText, 2)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAGLOSA", dbText, 80)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAFECACT", dbDate)
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAHORA", dbText, 8)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAUSUARI", dbText, 8)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CACODCLI", dbText, 11)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CARUC", dbText, 11)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CANOMCLI", dbText, 70)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAFORVEN", dbText, 4)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CACODMON", dbText, 2)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAVENDE", dbText, 2)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CATIPCAM", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CATIPGUI", dbText, 2)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CASITGUI", dbText, 1)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAGUIFAC", dbText, 1)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CADIRENV", dbText, 70)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CACODTRAN", dbText, 11)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CANUMORD", dbText, 20)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAGUIDEV", dbText, 1)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CANOMPRO", dbText, 50)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CANROPED", dbText, 10)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CACOTIZA", dbText, 10)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAPORDESCL", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAPORDESES", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAIMPORTE", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CANOMTRA", dbText, 40)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CADIRTRA", dbText, 50)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CARUCTRA", dbText, 11)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAPLATRA", dbText, 10)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CANROIMP", dbText, 10)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CACODLIQ", dbText, 20)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAESTIMP", dbText, 1)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CACIERRE", dbBoolean)
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CATIPDEP", dbText, 2)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("CAZONAF", dbText, 2)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("FLAGGS", dbBoolean)
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("Asiento", dbBoolean)
'    ERwinTableDef.Fields.Append ERwinField
'    ERwinDatabase.TableDefs.Append ERwinTableDef
'    'Set ERwinField = ERwinTableDef.Fields("CACIERRE")
'    'Call SetFieldProp(ERwinField, "Format", dbtext, "Yes/No")
'    'Set ERwinField = ERwinTableDef.Fields("FLAGGS")
'    'Call SetFieldProp(ERwinField, "Format", dbtext, "Yes/No")
'    'Set ERwinField = ERwinTableDef.Fields("Asiento")
'    'Call SetFieldProp(ERwinField, "Format", dbtext, "Yes/No")
'    'Call SetFieldProp(ERwinField, "Caption", dbtext, "ASIENTO:")
'
'    '  CREATE INDEX "PrimaryKey"
'
'    Set ERwinTableDef = ERwinDatabase.TableDefs("MOVALMCAB")
'    Set ERwinIndex = ERwinTableDef.CreateIndex("PrimaryKey")
'    Set ERwinField = ERwinIndex.CreateField("CAALMA")
'    ERwinIndex.Fields.Append ERwinField
'    Set ERwinField = ERwinIndex.CreateField("CATD")
'    ERwinIndex.Fields.Append ERwinField
'    Set ERwinField = ERwinIndex.CreateField("CANUMDOC")
'    ERwinIndex.Fields.Append ERwinField
'    ERwinIndex.Primary = True
'    ERwinIndex.Clustered = True
'    ERwinTableDef.Indexes.Append ERwinIndex
'
'    '  CREATE TABLE "MOVALMDET"
'    Set ERwinTableDef = ERwinDatabase.CreateTableDef("MOVALMDET")
'    Set ERwinField = ERwinTableDef.CreateField("DEALMA", dbText, 2)
'    ERwinField.Required = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DETD", dbText, 2)
'    ERwinField.Required = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DENUMDOC", dbText, 10)
'    ERwinField.Required = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEITEM", dbInteger)
'    ERwinField.Required = True
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DECODIGO", dbText, 20)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DECODREF", dbText, 40)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DECANTID", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DECANTENT", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DECANREF", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DECANFAC", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEORDEN", dbText, 6)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEPREUNI", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEPRECIO", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEPRECI1", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEDESCTO", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DESTOCK", dbText, 50)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEIGV", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEIMPMN", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEIMPUS", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DESERIE", dbText, 20)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DESITUA", dbText, 1)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEFECDOC", dbDate)
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DECENCOS", dbText, 6)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DERFALMA", dbText, 2)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DETR", dbText, 1)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEESTADO", dbText, 1)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DECODMOV", dbText, 2)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEVALTOT", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DECOMPRO", dbText, 6)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DECODMON", dbText, 2)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DETIPO", dbText, 1)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DETIPCAM", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEPREVTA", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEMONVTA", dbText, 2)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEFECVEN", dbDate)
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEDEVOL", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DESOLI", dbText, 3)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEDESCRI", dbText, 65)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEPORDES", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEIGVPOR", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEDESCLI", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEDESESP", dbDouble)
'    ERwinField.DefaultValue = "0"
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DENUMFAC", dbText, 10)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DELOTE", dbText, 20)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEUNIDAD", dbText, 6)
'    ERwinField.AllowZeroLength = True
'    ERwinTableDef.Fields.Append ERwinField
'    Set ERwinField = ERwinTableDef.CreateField("DEEPQ", dbText, 50)
'    ERwinTableDef.Fields.Append ERwinField
'    ERwinDatabase.TableDefs.Append ERwinTableDef
'
'    '  CREATE INDEX "PrimaryKey"
'
'    Set ERwinTableDef = ERwinDatabase.TableDefs("MOVALMDET")
'    Set ERwinIndex = ERwinTableDef.CreateIndex("PrimaryKey")
'    Set ERwinField = ERwinIndex.CreateField("DEALMA")
'    ERwinIndex.Fields.Append ERwinField
'    Set ERwinField = ERwinIndex.CreateField("DETD")
'    ERwinIndex.Fields.Append ERwinField
'    Set ERwinField = ERwinIndex.CreateField("DENUMDOC")
'    ERwinIndex.Fields.Append ERwinField
'    Set ERwinField = ERwinIndex.CreateField("DEITEM")
'    ERwinIndex.Fields.Append ERwinField
'    ERwinIndex.Primary = True
'    ERwinIndex.Clustered = True
'    ERwinTableDef.Indexes.Append ERwinIndex
'
'    '  CREATE INDEX "DECANTID"
'
'    Set ERwinTableDef = ERwinDatabase.TableDefs("MOVALMDET")
'    Set ERwinIndex = ERwinTableDef.CreateIndex("DECANTID")
'    Set ERwinField = ERwinIndex.CreateField("DECANTID")
'    ERwinIndex.Fields.Append ERwinField
'    ERwinTableDef.Indexes.Append ERwinIndex
'
'    '  CREATE RELATIONSHIP "{2BE1D185-CD34-11D4-859D-00E07D7FA23B}"
'    Set ERwinRelation = ERwinDatabase.CreateRelation("{2BE1D185-CD34-11D4-859D-00E07D7FA23B}", "MOVALMCAB", "MOVALMDET")
'    Set ERwinField = ERwinRelation.CreateField("CAALMA")
'    ERwinField.ForeignName = "DEALMA"
'    ERwinRelation.Fields.Append ERwinField
'    Set ERwinField = ERwinRelation.CreateField("CATD")
'    ERwinField.ForeignName = "DETD"
'    ERwinRelation.Fields.Append ERwinField
'    Set ERwinField = ERwinRelation.CreateField("CANUMDOC")
'    ERwinField.ForeignName = "DENUMDOC"
'    ERwinRelation.Fields.Append ERwinField
'    ERwinRelation.Attributes = ERwinRelation.Attributes + DB_RELATIONDONTENFORCE
'    ERwinDatabase.Relations.Append ERwinRelation
'
'    ERwinDatabase.Close
'    ERwinWorkspace.Close
End Sub
Private Sub Inicilizar(flag As Boolean)
    'Inicializa las barras de progreso y las etiquetas
    PB1.Min = 0: PB1.Max = MarcarList: PB1.Value = 0
    lbcab.Visible = flag: PB1.Visible = flag
    LBDet.Visible = flag: PB2.Visible = flag
End Sub
Private Sub CrearTablasrelacionadas()
''Dim CNMAESTROS As ADODB.Connection
'Dim xitem As ListItem
'Dim CODIGOTEM As String, CODIGOTABLA As String
'    Me.Width = 7080: Me.Height = 6180
''    Set CNMAESTROS = New ADODB.Connection
''    CNMAESTROS.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & txDestino & "\Maestro.mdb"
'    PB1.Min = 0: PB1.Max = LvRela.ListItems.count: PB1.Value = 0
'    lbcab.Visible = flag: PB1.Visible = flag
'    For Each xitem In LvRela.ListItems
'        lbcab.Caption = "Generando las tablas relacionadas ..TT" & Trim(xitem.key)
'        lbcab.Refresh
'        PB1.Value = PB1.Value + 1
'        CODIGOTEMP = Mid(xitem.Tag, 1, InStr(xitem.Tag, "-") - 1)
'        CODIGOTABLA = Mid(xitem.Tag, InStr(xitem.Tag, "-") + 1)
'        Cnx.Execute "Select * INTO TT" & Trim(xitem.key) & " from  [" & cRuta2 & "]." & xitem.key & " Where " & CODIGOTABLA & " IN(" & "SELECT DISTINCT " & CODIGOTEMP & " FROM [" & cRuta2 & "].TMPSELECT)"
'    Next
End Sub
Private Sub Finalizar()
    Dim NombreRuta As String
    Dim xitem As ListItem
On Error GoTo errores:
    Me.Width = 7080: Me.Height = 6180
        
    NombreRuta = VGRUCEMP & "-" & Trim(Format(DpFechini, "dd.mm.yyyy")) & "-" & Trim(Format(DpFechfin, "dd.mm.yyyy")) & ".mdb"
    Call CrearBaseDatos(NombreRuta, TxDestino)
    Call CrearTablasBase(NombreRuta, TxDestino)
    Cnx.CursorLocation = adUseClient
    Cnx.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & TxDestino & "\" & NombreRuta
    'Generando tablas relacionadas a las guias de ventas
    If ChkRelacion.Value = 1 Then
        'Call CrearBaseDatos("Maestro.mdb", txDestino)
        Call CrearTablasrelacionadas
    End If
    '****
    
    Call Inicilizar(True)
    For Each xitem In LvDocu.ListItems
        If xitem.Checked Then
            lbcab.Caption = "Exportando Cabeceras  Nro.Doc : " & xitem.text & " - " & Trim(xitem.SubItems(1)) & Trim(xitem.SubItems(2))
            lbcab.Refresh
            PB1.Value = PB1.Value + 1
            Call grabarCab(xitem.SubItems(3), xitem.text, Trim(xitem.SubItems(1)) & Trim(xitem.SubItems(2)))
            Call grabarDet(xitem.SubItems(3), xitem.text, Trim(xitem.SubItems(1)) & Trim(xitem.SubItems(2)))
        End If
    Next
    MsgBox "Se creo satisfactoriamente", vbInformation
    Call Inicilizar(False)
    Me.Width = 7080: Me.Height = 5355
    cmdAtras.Enabled = False: cmdsigue.Enabled = False
    LvDocu.ListItems.Clear: ChKMarca.Enabled = False
    LbReg.Caption = "0 "
    Exit Sub
errores:
    Select Case Err.Number
      Case 70
            MsgBox "La Base de datos a generarse para la migracion esta siendo utilizada por otro usuario " & Chr(13) & _
                   "Base de datos : " & TxDestino & "\" & NombreRuta & Chr(13) & _
                   "Por favor cierre esta base de datos o intente otro criterio de consulta"
            Call Inicilizar(False)
            Me.Width = 7080: Me.Height = 5355
      Case Else: MsgBox Err.Description
    End Select
    Exit Sub
    Resume
End Sub
Private Sub grabarCab(alma As String, tD As String, numero As String)
Dim RsCab As adodb.Recordset, RsCoCab As adodb.Recordset
Dim i As Integer
    Set RsCab = New adodb.Recordset
    Set RsCoCab = New adodb.Recordset
    On Error Resume Next
    RsCab.Open "Select * from MOVALMCAB WHERE CAALMA='" & Trim(alma) & "'" & _
               " AND CATD='" & Trim(tD) & "' AND CANUMDOC='" & numero & "'", cConexCom, adOpenKeyset, adLockReadOnly
    RsCoCab.Open "SELECT * FROM MOVALMCAB", Cnx, adOpenKeyset, adLockOptimistic
    Do While Not RsCab.EOF
        RsCoCab.AddNew
        For i = 0 To RsCoCab.Fields.count - 1
            RsCoCab.Fields(i) = RsCab.Fields(Trim(RsCoCab.Fields(i).name)).Value
        Next
        RsCoCab.Update
        RsCab.MoveNext
    Loop
End Sub
Private Sub grabarDet(alma As String, tD As String, numero As String)
Dim RsDet As adodb.Recordset, RsCoDet As adodb.Recordset
Dim i As Integer
On Error GoTo ERRGRABDET
    Set RsDet = New adodb.Recordset
    Set RsCoDet = New adodb.Recordset

    RsDet.Open "Select * from MOVALMDET WHERE DEALMA='" & Trim(alma) & "'" & _
               " AND DETD='" & Trim(tD) & "' AND DENUMDOC='" & numero & "'", cConexCom, adOpenKeyset, adLockReadOnly
    RsCoDet.Open "SELECT * FROM MOVALMDET", Cnx, adOpenKeyset, adLockOptimistic
    PB2.Min = 0: PB2.Max = RsDet.RecordCount: PB2.Value = 0
    Do While Not RsDet.EOF
        RsCoDet.AddNew
        For i = 0 To RsCoDet.Fields.count - 1
            RsCoDet.Fields(i) = RsDet.Fields(Trim(RsCoDet.Fields(i).name)).Value
        Next
        PB2.Value = PB2.Value + 1
        RsCoDet.Fields("DEITEM") = PB2.Value
        LBDet.Caption = "Registrando los detalles  Items Nro:" & PB2.Value
        RsCoDet.Update
        RsDet.MoveNext
    Loop
Exit Sub
ERRGRABDET:
    Resume Next
End Sub

Private Function Validar() As Boolean
    Validar = False
    If TxTipdoc.text = "" Then
        MsgBox "Tiene que escoger un tipo de documento", vbExclamation
        CmdTipdoc.SetFocus: Exit Function
    End If
    If TxCli.text = "" Then
        MsgBox "Tiene que escoger un cliente", vbExclamation
        CmdClie.SetFocus: Exit Function
    End If
    If TxDestino.text = "" Then
        MsgBox "Tiene que seleccionar el destino", vbExclamation
        CmdDestino.SetFocus: Exit Function
    End If
    Validar = True
End Function
Private Function MarcarList(Optional Check As Variant) As Long
Dim Cont As Long
Dim xitem As ListItem
    For Each xitem In LvDocu.ListItems
        If Not IsMissing(Check) Then xitem.Checked = Check
        If xitem.Checked Then
            Cont = Cont + 1
        End If
    Next
    MarcarList = Cont
End Function


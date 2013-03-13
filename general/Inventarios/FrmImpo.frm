VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmImpo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asistente para importar documentos"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   Icon            =   "FrmImpo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdRegistro 
      Caption         =   "&Registro"
      Height          =   345
      Left            =   705
      TabIndex        =   50
      Top             =   4515
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSComctlLib.ProgressBar PB2 
      Height          =   150
      Left            =   -15
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   5205
      Visible         =   0   'False
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   30
      Top             =   4320
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
            Picture         =   "FrmImpo.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpo.frx":1C94
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
      BackColor       =   &H00C0C0C0&
      Caption         =   "< &Atras"
      Height          =   345
      Left            =   3165
      TabIndex        =   1
      Top             =   4515
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4260
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7514
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
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
      TabPicture(0)   =   "FrmImpo.frx":2F16
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "FrmImpo.frx":2F32
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Tab 4"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame9"
      Tab(4).Control(1)=   "Frame10"
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   -75000
         TabIndex        =   45
         Top             =   -75
         Width           =   7035
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Registro de errores"
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
            Left            =   300
            TabIndex        =   46
            Top             =   195
            Width           =   3765
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00808080&
         Height          =   4350
         Left            =   -75015
         TabIndex        =   38
         Top             =   -105
         Width           =   7035
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   15
            TabIndex        =   40
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
               TabIndex        =   41
               Top             =   150
               Width           =   3180
            End
         End
         Begin MSDataGridLib.DataGrid DG1 
            Height          =   1635
            Left            =   300
            TabIndex        =   39
            Top             =   2235
            Width           =   6600
            _ExtentX        =   11642
            _ExtentY        =   2884
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
         Begin MSComctlLib.ListView LvRela 
            Height          =   1080
            Left            =   270
            TabIndex        =   42
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
         Begin VB.Label LbregRela 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0 "
            Height          =   270
            Left            =   5955
            TabIndex        =   51
            Top             =   3960
            Width           =   930
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Datos"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   285
            TabIndex        =   44
            Top             =   1950
            Width           =   2535
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Tablas relacionadas a las guias"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   315
            TabIndex        =   43
            Top             =   540
            Width           =   2535
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Height          =   1140
         Left            =   -75000
         TabIndex        =   36
         Top             =   -90
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
            TabIndex        =   37
            Top             =   270
            Width           =   3765
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00808080&
         Height          =   4350
         Left            =   -75015
         TabIndex        =   31
         Top             =   -105
         Width           =   7035
         Begin VB.CheckBox ChKMarca 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   "Marcar Todos"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   435
            TabIndex        =   32
            Top             =   1275
            Width           =   1485
         End
         Begin MSComctlLib.ListView LvDocu 
            Height          =   2355
            Left            =   435
            TabIndex        =   33
            Top             =   1560
            Width           =   6390
            _ExtentX        =   11271
            _ExtentY        =   4154
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
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
            TabIndex        =   35
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
            TabIndex        =   34
            Top             =   3990
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00808080&
         Height          =   4350
         Left            =   -75015
         TabIndex        =   5
         Top             =   -90
         Width           =   7035
         Begin MSComDlg.CommonDialog Cmd1 
            Left            =   6465
            Top             =   1905
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox txDestino 
            BackColor       =   &H00F2FFFF&
            ForeColor       =   &H00404040&
            Height          =   300
            Left            =   1260
            TabIndex        =   14
            Top             =   3795
            Width           =   4980
         End
         Begin VB.CommandButton CmdOrigen 
            Caption         =   "..."
            Height          =   315
            Left            =   5850
            TabIndex        =   13
            Top             =   1260
            Width           =   375
         End
         Begin VB.TextBox TxOrigen 
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Left            =   1215
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1245
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
            Begin MSComCtl2.DTPicker DPFecha 
               Height          =   300
               Left            =   3000
               TabIndex        =   29
               Top             =   1515
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               Format          =   60489729
               CurrentDate     =   37132
            End
            Begin VB.TextBox TxAlmacen 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1935
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   1065
               Width           =   3180
            End
            Begin VB.CommandButton CmdAlmacen 
               Caption         =   "..."
               Height          =   270
               Left            =   5130
               TabIndex        =   17
               Top             =   1095
               Width           =   330
            End
            Begin VB.CommandButton CmdTipdoc 
               Caption         =   "..."
               Enabled         =   0   'False
               Height          =   285
               Left            =   5130
               TabIndex        =   16
               Top             =   720
               Width           =   330
            End
            Begin VB.TextBox TxTipdoc 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1935
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   705
               Width           =   3180
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha :"
               ForeColor       =   &H00C0FFFF&
               Height          =   240
               Left            =   2265
               TabIndex        =   28
               Top             =   1575
               Width           =   600
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00CBD1D3&
               BackStyle       =   0  'Transparent
               Caption         =   "Seleccione el tipo de documento a generar y el almacen a abastecer"
               ForeColor       =   &H00C0FFFF&
               Height          =   330
               Left            =   660
               TabIndex        =   27
               Top             =   345
               Width           =   4935
            End
            Begin VB.Image Image3 
               Height          =   360
               Left            =   210
               Picture         =   "FrmImpo.frx":2F4E
               Stretch         =   -1  'True
               Top             =   180
               Width           =   360
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Almacen :"
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   345
               TabIndex        =   19
               Top             =   1125
               Width           =   1605
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Documento :"
               ForeColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   345
               TabIndex        =   10
               Top             =   750
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
               Left            =   1020
               TabIndex        =   7
               Top             =   435
               Width           =   3045
            End
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Destino :"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   570
            TabIndex        =   11
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
            Caption         =   "Importar tablas relacionadas"
            Height          =   285
            Left            =   2475
            TabIndex        =   30
            Top             =   3960
            Width           =   2685
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Echo por Fernando Cossio"
            Height          =   210
            Left            =   225
            TabIndex        =   52
            Top             =   1620
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
            Caption         =   "Nota.- Este proceso debe ser realizado por un usuario administrador, conocedor del sistema."
            Height          =   705
            Left            =   2475
            TabIndex        =   26
            Top             =   3255
            Width           =   4035
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   2490
            Picture         =   "FrmImpo.frx":41C0
            Top             =   2670
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   2445
            Picture         =   "FrmImpo.frx":5002
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Asistente para Importar Documentos"
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
            TabIndex        =   21
            Top             =   450
            Width           =   3390
         End
         Begin VB.Label Label8 
            Caption         =   $"FrmImpo.frx":6274
            Height          =   990
            Left            =   2445
            TabIndex        =   20
            Top             =   1875
            Width           =   4275
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
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   4020
         Left            =   -74985
         TabIndex        =   47
         Top             =   240
         Width           =   7020
         Begin RichTextLib.RichTextBox RTB1 
            Height          =   3300
            Left            =   120
            TabIndex        =   48
            Top             =   540
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   5821
            _Version        =   393217
            BackColor       =   15925247
            ReadOnly        =   -1  'True
            Appearance      =   0
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"FrmImpo.frx":6315
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Errores en guias"
            ForeColor       =   &H00C0FFFF&
            Height          =   225
            Left            =   165
            TabIndex        =   49
            Top             =   315
            Width           =   1620
         End
      End
   End
   Begin VB.Label LBDet 
      Caption         =   "Registrando los detalles"
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   -15
      TabIndex        =   25
      Top             =   5415
      Visible         =   0   'False
      Width           =   6045
   End
   Begin VB.Label lbcab 
      Caption         =   "Importando Cabeceras"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   0
      TabIndex        =   24
      Top             =   4995
      Visible         =   0   'False
      Width           =   5730
   End
End
Attribute VB_Name = "FrmImpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cnx As New ADODB.Connection 'Conexion general de la data a importar
Dim NumGuia As String
Dim RUCPRO As String
Dim NOMPRO As String
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
    Dim RSALMA As ADODB.Recordset
    Set RSALMA = New ADODB.Recordset
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
    'Set CNX = Nothing
    Unload Me
End Sub

Private Sub CmdOrigen_Click()
    Cmd1.DialogTitle = "Seleccione el archivo a importar"
    Cmd1.Filter = "Archivo Acces 97(*.mdb)|*.mdb"
    Cmd1.ShowOpen
    txOrigen = Cmd1.FileName
End Sub

Private Sub CmdRegistro_Click()
    SSTab1.Tab = 4
    CmdRegistro.Enabled = False
End Sub

Private Sub cmdsigue_Click()
    Dim Cont As Integer
    Dim cad As String
    If SSTab1.Tab + 1 = 2 Then
        If Trim(txOrigen) = "" Then
            MsgBox "Seleccione el archivo de origen", vbExclamation
            CmdOrigen.SetFocus: Exit Sub
        End If
        cad = StrReverse(Mid(StrReverse(txOrigen), 1, InStr(StrReverse(txOrigen), "\") - 1))
        If Not VerifiArchivo(cad) Then
            MsgBox "El archivo a importar no corresponde a un Proveedor registrado en el sistema", vbExclamation, "Advertencia"
            CmdOrigen.SetFocus
            Exit Sub
        End If
        If Not Validar Then Exit Sub
        
        Set Cnx = New ADODB.Connection
        Cnx.CursorLocation = adUseClient
        Cnx.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & txOrigen
    End If
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
End Sub
Private Sub Habilitar_Mov_Bot(Atras As Boolean, Siguiente As Boolean, Optional Cap As String)
    cmdAtras.Enabled = Atras
    cmdsigue.Enabled = Siguiente
    cmdsigue.Caption = Cap
End Sub

Private Sub CmdTipdoc_Click()
    Dim RSTIPDOC As ADODB.Recordset
    Set RSTIPDOC = New ADODB.Recordset
    RSTIPDOC.Open "SELECT TDO_TIPDOC,TDO_DESCRI FROM TIPO_DOCU", cConexCom, adOpenKeyset, adLockReadOnly
    frmref.Conectar RSTIPDOC
    frmref.Show 1
    If vGUtil(1) <> "" Then
        TxTipdoc.Tag = vGUtil(1): TxTipdoc = vGUtil(2)
      Else: TxTipdoc.Tag = "": TxTipdoc = ""
    End If
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub Form_Load()
    Me.Width = 7080: Me.Height = 5355
    Me.TxDestino.Tag = VGCOMP
    Me.TxDestino = VGNemp
    SSTab1.Tab = 0
    ChKMarca.Value = 0
    SSTab1.TabVisible(2) = False
    TxTipdoc.Tag = "NI": TxTipdoc.text = "Nota de Ingreso"
    DPFecha.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Cnx = Nothing
End Sub

Private Sub LvRela_ItemClick(ByVal item As MSComctlLib.ListItem)
Dim RSAUX As ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    Dim CODIGOTEM As String, CODIGOTABLA As String
    CODIGOTEMP = Mid(item.Tag, 1, InStr(item.Tag, "-") - 1)
    CODIGOTABLA = Mid(item.Tag, InStr(item.Tag, "-") + 1)
    RSAUX.Open "Select * from " & item.key, Cnx
    LbregRela.Caption = Format(RSAUX.RecordCount, "0 ")
    Set DG1.DataSource = RSAUX
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
        Case 0
            Call Habilitar_Mov_Bot(False, True, "&Siguiente >")
        Case 1: Call Habilitar_Mov_Bot(True, True, "&Siguiente >")
        Case 2
            Call Habilitar_Mov_Bot(True, True, "&Siguiente >")
            Call VistaTablasRela
        Case 3
            Call Habilitar_Mov_Bot(True, True, "&Finalizar")
            Call Llenar_Lista
    End Select
End Sub
Private Sub VistaTablasRela()
    Dim itmX As ListItem
    LvRela.ListItems.Clear
       
    If ExisteElem(0, Cnx, "TTTABUNIMED") Then _
        Set itmX = LvRela.ListItems.Add(, "TTTABUNIMED", "Tabla de Unidades", , 2): itmX.Tag = "AUNIDAD-UM_ABREV"
    If ExisteElem(0, Cnx, "TTFAMILIA") Then _
        Set itmX = LvRela.ListItems.Add(, "TTFAMILIA", "Tabla de Familias", , 2): itmX.Tag = "AFAMILIA-FAM_CODIGO"
    If ExisteElem(0, Cnx, "TTTIPO_ARTICULO") Then _
        Set itmX = LvRela.ListItems.Add(, "TTTIPO_ARTICULO", "Tabla Tipos de Articulos", , 2): itmX.Tag = "ATIPO-COD_TIPO"
    If ExisteElem(0, Cnx, "TTTALLA") Then _
        Set itmX = LvRela.ListItems.Add(, "TTTALLA", "Tabla de Tallas", , 2): itmX.Tag = "TALLA-CODIGO"
    If ExisteElem(0, Cnx, "TTMAEMARCA") Then _
        Set itmX = LvRela.ListItems.Add(, "TTMAEMARCA", "Tabla de Clases", , 2): itmX.Tag = "AMARCA-COD_MARCA"
    If ExisteElem(0, Cnx, "TTMAECOLOR") Then _
        Set itmX = LvRela.ListItems.Add(, "TTMAECOLOR", "Tabla de Colores", , 2): itmX.Tag = "ACOLOR-COD_COLOR"
    If ExisteElem(0, Cnx, "TTLINEAS") Then _
        Set itmX = LvRela.ListItems.Add(, "TTLINEAS", "Tabla de Lineas", , 2): itmX.Tag = "AMODELO-LIN_CODIGO"
    If ExisteElem(0, Cnx, "TTGRUPO") Then _
        Set itmX = LvRela.ListItems.Add(, "TTGRUPO", "Tabla de Grupos", , 2): itmX.Tag = "AGRUPO-GRU_CODIGO"
    Set itmX = LvRela.ListItems.Add(, "TTMAEART", "Tabla de Articulos", , 2): itmX.Tag = "ACODIGO-ACODIGO"
End Sub
Private Function Seleccionardocumentos() As ADODB.Recordset
Dim sqlcad As String
    Set Seleccionardocumentos = New ADODB.Recordset
    sqlcad = "Select * from MOVALMCAB WHERE CATD='GS'"
    Seleccionardocumentos.Open sqlcad, Cnx, adOpenKeyset, adLockReadOnly
End Function
Private Sub Llenar_Lista()
    If LvDocu.ListItems.count > 0 Then Exit Sub
    Dim rslista As New ADODB.Recordset
    Dim itmX As ListItem
    Set rslista = Seleccionardocumentos
    
    Do While Not rslista.EOF
        Set itmX = LvDocu.ListItems.Add(, ESNULO(rslista!CATD, "") & ESNULO(rslista!CANUMDOC, ""), ESNULO(rslista!CATD, ""), , 1)
        itmX.SubItems(1) = Mid(ESNULO(rslista!CANUMDOC, ""), 1, 3)
        itmX.SubItems(2) = Mid(ESNULO(rslista!CANUMDOC, ""), 4)
        itmX.SubItems(3) = ESNULO(rslista!CAALMA, "")
        rslista.MoveNext
    Loop
    LbReg.Caption = Format(rslista.RecordCount, "0 ")
End Sub
Private Sub Inicilizar(flag As Boolean)
    PB1.Min = 0: PB1.Max = MarcarList: PB1.Value = 0
    lbcab.Visible = flag: PB1.Visible = flag
    LBDet.Visible = flag: PB2.Visible = flag
End Sub

Private Sub Finalizar()
    Dim NombreRuta As String
    Dim xitem As ListItem
    Dim RSAUX As New ADODB.Recordset
    Dim UTLMAXIMO As Double
    Dim flag As Boolean 'Variable que controla el primer titulo del rich
    Dim flagnum As Boolean 'Variable que controla que termine en el ultimo
                           'numero de secuencia en nota de ingreso
On Error GoTo errores:
    Me.Width = 7080: Me.Height = 6180
    If ChkRelacion.Value = 1 Then
        Call GrabaRelaciones
    End If
    
    Call Inicilizar(True)
    PB2.Min = 0: PB2.Value = 0
    
    For Each xitem In LvDocu.ListItems
        If xitem.Checked Then
            lbcab.Caption = " Importando Cabeceras  Nro.Doc : " & xitem.text & " - " & Trim(xitem.SubItems(1)) & Trim(xitem.SubItems(2))
            lbcab.Refresh
            PB1.Value = PB1.Value + 1
                       
            Set RSAUX = New ADODB.Recordset
            RSAUX.Open "SELECT MAX(CANUMDOC) AS MAXIMO FROM MOVALMCAB WHERE CATD='NI' AND CAALMA='" & Trim(TxAlmacen.Tag) & "'", cConexCom
            UTLMAXIMO = CDbl(ESNULO(RSAUX!MAXIMO, 0)) + 1
            NumGuia = Format(UTLMAXIMO, "0000000000")
            If Not Existe(1, TxAlmacen.Tag, "MOVALMCAB", "CAALMA", False, xitem.text, "CARFTDOC", Trim(xitem.SubItems(1)) & Trim(xitem.SubItems(2)), "CARFNDOC") Then
                Call grabarCab(xitem.SubItems(3), xitem.text, Trim(xitem.SubItems(1)) & Trim(xitem.SubItems(2)), NumGuia)
                Call grabarDet(xitem.SubItems(3), xitem.text, Trim(xitem.SubItems(1)) & Trim(xitem.SubItems(2)), NumGuia)
                flagnum = False
              Else
                If Not flag Then
                    RTB1.text = "Guias ya ingresadas como referencias en Notas de ingreso :" & Chr(13) & Chr(10)
                    'RTB1.SelLength = 0
                    RTB1.text = RTB1.text & xitem.text & "-" & Trim(xitem.SubItems(1)) & Trim(xitem.SubItems(2)) & Chr(13) & Chr(10)
                  Else
                    RTB1.text = RTB1.text & xitem.text & "-" & Trim(xitem.SubItems(1)) & Trim(xitem.SubItems(2)) & Chr(13) & Chr(10)
                End If
                flag = True
                flagnum = True
            End If
            
        End If
    Next
    
    'Se actualiza el ultimo numero de nota de ingreso en almacen
    If Not flagnum Then
        cConexCom.Execute "UPDATE TABALM SET TANUMENT=" & UTLMAXIMO & " WHERE TAALMA='" & Trim(TxAlmacen.Tag) & "'"
    End If
    If Not flag Then
        MsgBox "Se Importo satisfactoriamente", vbInformation
     Else
        RTB1.SelStart = 0: RTB1.SelLength = 58
        RTB1.SelColor = &H80&
        RTB1.SelLength = 0
        MsgBox "Se Genero algunos errores por favor dele click en registro", vbExclamation
        CmdRegistro.Visible = True
    End If
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
Private Sub grabarCab(alma As String, tD As String, numero As String, numeroguia As String)
Dim RsCab As ADODB.Recordset, RsCoCab As ADODB.Recordset
Dim i As Integer
    Set RsCab = New ADODB.Recordset
    Set RsCoCab = New ADODB.Recordset
    RsCab.Open "Select * from MOVALMCAB WHERE CAALMA='" & Trim(alma) & "'" & _
               " AND CATD='" & Trim(tD) & "' AND CANUMDOC='" & numero & "'", Cnx, adOpenKeyset, adLockReadOnly
    RsCoCab.Open "MOVALMCAB", cConexCom, adOpenKeyset, adLockOptimistic
    Do While Not RsCab.EOF
        RsCoCab.AddNew
        RsCoCab("CAALMA") = TxAlmacen.Tag
        RsCoCab("CACODMOV") = "CL"
        RsCoCab("CATIPMOV") = "I"
        RsCoCab("CATD") = "NI"
        RsCoCab("CARFTDOC") = "GS"
        RsCoCab("CARFNDOC") = RsCab("CANUMDOC")
        RsCoCab("CANUMDOC") = numeroguia
        RsCoCab("CAFECDOC") = DPFecha.Value
        RsCoCab("CAUSUARI") = VGUSU_CODIGO
        RsCoCab("CAHORA") = Format(Time, "HH:MM:SS")
        RsCoCab("CACODPRO") = RUCPRO
        RsCoCab("CANOMPRO") = NOMPRO
        RsCoCab("CACODMON") = RsCab("CACODMON")
        RsCoCab("CAFECACT") = Date
        RsCoCab("CASITGUI") = "V"
'            For i = 0 To RsCoCab.Fields.count - 1
'                RsCoCab.Fields(i) = RsCab.Fields(Trim(RsCoCab.Fields(i).name)).Value
'            Next
        RsCoCab.Update
        RsCab.MoveNext
    Loop
End Sub
Private Sub GrabaRelaciones()
Dim xitem As ListItem
    'Grabando tablas relacionadas
    PB1.Min = 0: PB1.Max = LvRela.ListItems.count: PB1.Value = 0
    lbcab.Visible = True: PB1.Visible = True
    LBDet.Visible = True: PB2.Visible = True
    For Each xitem In LvRela.ListItems
        PB1.Value = PB1.Value + 1
        lbcab.Caption = "Generando tablas relacionadas :..." & xitem.text
        lbcab.Refresh
        Call GrabartablaS(xitem)
    Next
End Sub
Private Sub GrabartablaS(SITEM As ListItem)
On Error GoTo ERRGRABDET
    Dim RSTABORIGEN As ADODB.Recordset
    Dim RSTABDESTINO As ADODB.Recordset
    Dim Cont As Long
    
    Set RSTABORIGEN = New ADODB.Recordset
    Set RSTABDESTINO = New ADODB.Recordset
    
    RSTABORIGEN.Open SITEM.key, Cnx, adOpenKeyset, adLockReadOnly
    RSTABDESTINO.Open Mid(SITEM.key, 3), cConexCom, adOpenKeyset, adLockOptimistic
    PB2.Value = 0: PB2.Min = 0: PB2.Max = RSTABORIGEN.RecordCount
    Cont = 0
    Do While Not RSTABORIGEN.EOF
        Cont = Cont + 1
        PB2.Value = PB2.Value + 1
        LBDet.Caption = "Generando los registros para cada tabla .... " & Cont
        LBDet.Refresh
        RSTABDESTINO.AddNew
        For i = 0 To RSTABDESTINO.Fields.count - 1
            RSTABDESTINO.Fields(i) = RSTABORIGEN.Fields(Trim(RSTABDESTINO.Fields(i).name)).Value
        Next
        RSTABDESTINO.Update
        RSTABORIGEN.MoveNext
    Loop
    Exit Sub
ERRGRABDET:
    Resume Next
End Sub


Private Sub grabarDet(alma As String, tD As String, numero As String, numeroguia As String)
Dim RsDet As ADODB.Recordset, RsCoDet As ADODB.Recordset
Dim i As Integer
'On Error GoTo ERRP
    Dim Cont As Long
    Set RsDet = New ADODB.Recordset
    Set RsCoDet = New ADODB.Recordset
    RsDet.Open "Select * from MOVALMDET WHERE DEALMA='" & Trim(alma) & "'" & _
               " AND DETD='" & Trim(tD) & "' AND DENUMDOC='" & numero & "'", Cnx, adOpenKeyset, adLockReadOnly
    RsCoDet.Open "MOVALMDET", cConexCom, adOpenKeyset, adLockOptimistic
    PB2.Min = 0: PB2.Max = RsDet.RecordCount: PB2.Value = 0
    Cont = 1
    Do While Not RsDet.EOF
        RsCoDet.AddNew
'        For i = 0 To RsCoDet.Fields.count - 1
'            RsCoDet.Fields(i) = RsDet.Fields(Trim(RsCoDet.Fields(i).name)).Value
'        Next
        PB2.Value = PB2.Value + 1
        RsCoDet("DEALMA") = TxAlmacen.Tag
        RsCoDet("DETD") = "NI"
        RsCoDet("DENUMDOC") = numeroguia
        RsCoDet("DEITEM") = Cont
        RsCoDet("DECODIGO") = RsDet("DECODIGO")
        RsCoDet("DECODREF") = RsDet("DECODREF")
        RsCoDet("DECANTID") = RsDet("DECANTID")
        RsCoDet("DECANTENT") = RsDet("DECANTENT")
        RsCoDet("DECANREF") = RsDet("DECANREF")
        RsCoDet("DECANFAC") = RsDet("DECANFAC")
        RsCoDet("DESERIE") = RsDet("DESERIE")
        RsCoDet("DECODMON") = RsDet("DECODMON")
        RsCoDet("DEDESCRI") = RsDet("DEDESCRI")
        RsCoDet("DELOTE") = RsDet("DELOTE")
        RsCoDet("DEUNIDAD") = RsDet("DEUNIDAD")
        LBDet.Caption = "Registrando los detalles  Items Nro:" & PB2.Value
        LBDet.Refresh
        RsCoDet.Update
        Cont = Cont + 1
        'Actualizando el stock de articulos
        Call ActualizarStock(RsDet("DECODIGO"), RsDet("DECANTID"), ESNULO(RsDet("DESERIE"), ""), ESNULO(RsDet("DELOTE"), ""))
        
        RsDet.MoveNext
    Loop
    Exit Sub
End Sub
Private Function Validar() As Boolean
    Validar = False
    If TxTipdoc.text = "" Then
        MsgBox "Tiene que escoger un tipo de documento", vbExclamation
        CmdTipdoc.SetFocus: Exit Function
    End If
    If TxAlmacen.text = "" Then
        MsgBox "Tiene que seleccionar el Almacen destino", vbExclamation
        CmdAlmacen.SetFocus: Exit Function
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
Private Function VerifiArchivo(ARCHIVO As String) As Boolean
On Error GoTo ERRVER
Dim Codigo As String
    Codigo = Mid(ARCHIVO, 1, InStr(ARCHIVO, "-") - 1)
    If Existe(1, Codigo, "MAEPROV", "PRVCRUC", False) Then
        VerifiArchivo = True
        RUCPRO = Codigo
        NOMPRO = Devolver_Dato(1, Codigo, "MAEPROV", "PRVCRUC", False, "PRVCNOMBRE")
      Else
        VerifiArchivo = False
    End If
Exit Function
ERRVER:
    VerifiArchivo = False
End Function
Private Sub ActualizarStock(Articulo As String, CANTIDAD As Long, Serie As String, Lote As String)
    'Calculando el stock del articulo
    'On Error GoTo ERRA
    If Not Existe(1, Articulo, "STKART", "STCODIGO", False, TxAlmacen.Tag, "STALMA") Then
        cConexCom.Execute "INSERT INTO STKART(STALMA,STCODIGO,STSKDIS) VALUES('" & TxAlmacen.Tag & "','" & Articulo & "'," & CANTIDAD & ")"
      Else
       cConexCom.Execute "UPDATE STKART SET STSKDIS=STSKDIS+" & CANTIDAD & " WHERE STALMA='" & Trim(TxAlmacen.Tag) & "' AND STCODIGO='" & Trim(Articulo) & "'"
    End If
    'SI EL ARTICULO TIENE SERIE
    If ESNULO(Devolver_Dato(1, Articulo, "MAEART", "ACODIGO", False, "AFSERIE"), "") = "S" Then
        If Not Existe(1, Articulo, "STKSERI", "STSCODIGO", False, TxAlmacen.Tag, "STSALMA", Serie, "STSSERIE") Then
            cConexCom.Execute "INSERT INTO STKSERI(STSALMA,STSCODIGO,STSSERIE,STSSKDIS) VALUES('" & TxAlmacen.Tag & "','" & Articulo & "','" & Serie & "',1)"
           Else
          cConexCom.Execute "UPDATE STKSERI SET STSSKDIS=1 WHERE STSALMA='" & Trim(TxAlmacen.Tag) & "' AND STSCODIGO='" & Trim(Articulo) & "' AND STSSERIE='" & Serie & "'"
        End If
    End If
    'SI EL ARTICULO TINE LOTE
    If ESNULO(Devolver_Dato(1, Articulo, "MAEART", "ACODIGO", False, "AFLOTE"), "") = "S" Then
        If Not Existe(1, Articulo, "STKLOTE", "STSCODIGO", False, TxAlmacen.Tag, "STSALMA", Lote, "STSLOTE") Then
            cConexCom.Execute "INSERT INTO STKLOTE(STSALMA,STSCODIGO,STSLOTE,STSLKDIS) VALUES('" & TxAlmacen.Tag & "','" & Articulo & "','" & Lote & "'," & CANTIDAD & ")"
           Else
          cConexCom.Execute "UPDATE STKLOTE SET STSLKDIS=STSLKDIS+" & CANTIDAD & "  WHERE STSALMA='" & Trim(TxAlmacen.Tag) & "' AND STSCODIGO='" & Trim(Articulo) & "' AND STSLOTE='" & Lote & "'"
        End If
    End If
    Exit Sub
'ERRA:
'    Stop
End Sub



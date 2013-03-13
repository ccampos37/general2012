VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form xxxFrmNotaFisico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota Abono/Cargo Fisico"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   12105
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   8160
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   7980
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   14076
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "                                            "
      TabPicture(0)   =   "FrmNotaFisico.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "MBox1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TDBGrid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Fr2(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Fr2(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "oCrystalReport"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin Crystal.CrystalReport oCrystalReport 
         Left            =   2910
         Top             =   8070
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Frame Fr2 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   2
         Left            =   210
         TabIndex        =   62
         Top             =   6180
         Width           =   11535
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   6
            Left            =   300
            TabIndex        =   63
            Top             =   60
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            ForeColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   7
            Left            =   2400
            TabIndex        =   64
            Top             =   60
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            ForeColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   8
            Left            =   4800
            TabIndex        =   65
            Top             =   60
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            ForeColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   9
            Left            =   7290
            TabIndex        =   66
            Top             =   60
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            ForeColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   10
            Left            =   9540
            TabIndex        =   67
            Top             =   60
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            ForeColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   7
            X1              =   9340
            X2              =   9340
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   6
            X1              =   6940
            X2              =   6940
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   5
            X1              =   4420
            X2              =   4420
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   2160
            X2              =   2160
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   3
            X1              =   9360
            X2              =   9360
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   2
            X1              =   6960
            X2              =   6960
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   1
            X1              =   4440
            X2              =   4440
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   0
            X1              =   2175
            X2              =   2175
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Neto Factura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   4
            Left            =   9840
            TabIndex        =   72
            Top             =   435
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total I.G.V."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   3
            Left            =   7680
            TabIndex        =   71
            Top             =   435
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Dctos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   2
            Left            =   5160
            TabIndex        =   70
            Top             =   435
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bruto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   69
            Top             =   435
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   68
            Top             =   435
            Width           =   1335
         End
      End
      Begin VB.Frame Fr2 
         Height          =   885
         Index           =   1
         Left            =   60
         TabIndex        =   46
         Top             =   2790
         Width           =   11775
         Begin VB.CommandButton cAyuda 
            Caption         =   "..."
            Height          =   375
            Index           =   3
            Left            =   3660
            TabIndex        =   42
            Top             =   420
            Width           =   285
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   0
            Left            =   1530
            TabIndex        =   40
            Top             =   420
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   1
            Left            =   2370
            TabIndex        =   41
            Top             =   420
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   2
            Left            =   7710
            TabIndex        =   47
            Top             =   420
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   -2147483648
            Enabled         =   0   'False
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   3
            Left            =   8610
            TabIndex        =   43
            Top             =   420
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   4
            Left            =   9810
            TabIndex        =   44
            Top             =   420
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   5
            Left            =   10740
            TabIndex        =   45
            Top             =   420
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   11
            Left            =   90
            TabIndex        =   48
            Top             =   420
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            BackColor       =   -2147483644
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   12
            Left            =   720
            TabIndex        =   39
            Top             =   420
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   13
            Left            =   90
            TabIndex        =   49
            Top             =   420
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   -2147483648
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   14
            Left            =   90
            TabIndex        =   50
            Top             =   420
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            BackColor       =   -2147483644
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Item"
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
            Index           =   7
            Left            =   120
            TabIndex        =   60
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3960
            TabIndex        =   59
            Top             =   420
            Width           =   3675
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Cant.UM"
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
            Index           =   6
            Left            =   1590
            TabIndex        =   58
            Top             =   180
            Width           =   795
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "%Com"
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
            Index           =   5
            Left            =   10740
            TabIndex        =   57
            Top             =   180
            Width           =   975
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Dscto"
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
            Index           =   4
            Left            =   9900
            TabIndex        =   56
            Top             =   180
            Width           =   735
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Precio Vta"
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
            Index           =   3
            Left            =   8670
            TabIndex        =   55
            Top             =   180
            Width           =   1005
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "U.M."
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
            Left            =   7800
            TabIndex        =   54
            Top             =   180
            Width           =   675
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Descripción"
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
            Index           =   1
            Left            =   3870
            TabIndex        =   53
            Top             =   180
            Width           =   3885
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Codigo"
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
            Index           =   0
            Left            =   2430
            TabIndex        =   52
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Cnt. Ref"
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
            Index           =   8
            Left            =   750
            TabIndex        =   51
            Top             =   180
            Width           =   765
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Nota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   90
         TabIndex        =   34
         Top             =   6870
         Width           =   2835
         Begin VB.Label Label4 
            Caption         =   "Eliminar Item"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   0
            Left            =   1470
            TabIndex        =   38
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Editar Item"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   1
            Left            =   1470
            TabIndex        =   37
            Top             =   570
            Width           =   1125
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   150
            Picture         =   "FrmNotaFisico.frx":001C
            Top             =   330
            Width           =   480
         End
         Begin VB.Label Label5 
            Caption         =   "[DEL]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   780
            TabIndex        =   36
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "[ENTER]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   780
            TabIndex        =   35
            Top             =   570
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2445
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   330
         Width           =   11775
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1260
            Width           =   1485
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   10200
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1260
            Width           =   1425
         End
         Begin VB.CommandButton cAyuda 
            Caption         =   "..."
            Height          =   285
            Index           =   0
            Left            =   3600
            TabIndex        =   21
            Top             =   2070
            Width           =   285
         End
         Begin MSMask.MaskEdBox MBox1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "d/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   5070
            TabIndex        =   6
            Top             =   210
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   8370
            TabIndex        =   7
            Top             =   210
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
            Height          =   345
            Left            =   1320
            TabIndex        =   8
            Top             =   840
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   609
            XcodMaxLongitud =   11
            xcodwith        =   800
            NomTabla        =   "vt_Cliente"
            TituloAyuda     =   "Ayuda de Clientes"
            ListaCampos     =   $"FrmNotaFisico.frx":045E
            XcodCampo       =   "clientecodigo"
            XListCampo      =   "clienterazonsocial"
            ListaCamposDescrip=   "Codigo,Descripcion,Ruc,Direccion,Distrito,LimiteCred,Saldo,T,P,M,D"
            ListaCamposText =   $"FrmNotaFisico.frx":0544
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
            Height          =   315
            Left            =   8040
            TabIndex        =   17
            Top             =   1680
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   556
            XcodMaxLongitud =   3
            xcodwith        =   200
            NomTabla        =   "vt_vendedor"
            TituloAyuda     =   "Ayuda de Vendedores"
            ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
            XcodCampo       =   "vendedorcodigo"
            XListCampo      =   "vendedornombres"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "vendedorcodigo,vendedornombres"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda3 
            Height          =   315
            Left            =   5940
            TabIndex        =   22
            Top             =   2040
            Width           =   5715
            _ExtentX        =   10081
            _ExtentY        =   556
            XcodMaxLongitud =   2
            xcodwith        =   100
            NomTabla        =   "cc_conceptos"
            TituloAyuda     =   "Ayuda de Conceptos"
            ListaCampos     =   "conceptocodigo(1),conceptodescripcion(1)"
            XcodCampo       =   "conceptocodigo"
            XListCampo      =   "conceptodescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "conceptocodigo,conceptodescripcion"
         End
         Begin MSMask.MaskEdBox MBox 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   6870
            TabIndex        =   12
            Top             =   1320
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
            _Version        =   393216
            ClipMode        =   1
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   5010
            TabIndex        =   16
            Top             =   1680
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
            _Version        =   393216
            ClipMode        =   1
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   1
            Left            =   2850
            TabIndex        =   10
            Top             =   1290
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   11
            Top             =   1290
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   4
            Left            =   1320
            TabIndex        =   15
            Top             =   1680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   6
            Left            =   1320
            TabIndex        =   18
            Top             =   2070
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   7
            Left            =   1800
            TabIndex        =   19
            Top             =   2070
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   8
            Left            =   2340
            TabIndex        =   20
            Top             =   2070
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            Caption         =   "Importe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   180
            TabIndex        =   33
            Top             =   1680
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Registro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   3540
            TabIndex        =   32
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label Label5 
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   180
            TabIndex        =   31
            Top             =   1320
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Cambio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   7110
            TabIndex        =   30
            Top             =   240
            Width           =   1215
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000009&
            Index           =   0
            X1              =   30
            X2              =   11790
            Y1              =   570
            Y2              =   570
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   30
            X2              =   11730
            Y1              =   540
            Y2              =   540
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "DETALLE DOCUMENTO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   6
            Left            =   60
            TabIndex        =   29
            Top             =   600
            Width           =   11685
         End
         Begin VB.Label Label4 
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   210
            TabIndex        =   28
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha Emision"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   5430
            TabIndex        =   27
            Top             =   1290
            Width           =   1395
         End
         Begin VB.Label Label5 
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   9390
            TabIndex        =   13
            Top             =   1290
            Width           =   915
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha Vencimiento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   3270
            TabIndex        =   26
            Top             =   1710
            Width           =   1665
         End
         Begin VB.Label Label5 
            Caption         =   "Vendedor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   7110
            TabIndex        =   25
            Top             =   1710
            Width           =   1605
         End
         Begin VB.Label Label5 
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   8
            Left            =   180
            TabIndex        =   24
            Top             =   2070
            Width           =   1305
         End
         Begin VB.Label Label5 
            Caption         =   "Concepto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   9
            Left            =   4950
            TabIndex        =   23
            Top             =   2100
            Width           =   1305
         End
      End
      Begin VB.Frame Frame4 
         Height          =   930
         Left            =   5970
         TabIndex        =   2
         Top             =   6870
         Width           =   1980
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Cancelar"
            Height          =   690
            Index           =   12
            Left            =   1050
            Picture         =   "FrmNotaFisico.frx":0609
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   180
            Width           =   855
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Acepta"
            Height          =   690
            Index           =   11
            Left            =   90
            Picture         =   "FrmNotaFisico.frx":0A4B
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   180
            Width           =   870
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   2265
         Left            =   90
         TabIndex        =   61
         Top             =   3840
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   3995
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
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
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
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(34)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
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
         InsertMode      =   0   'False
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=18,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HC0C0C0&"
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
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Named:id=33:Normal"
         _StyleDefs(65)  =   ":id=33,.parent=0"
         _StyleDefs(66)  =   "Named:id=34:Heading"
         _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=34,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=36:Selected"
         _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=37:Caption"
         _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(75)  =   "Named:id=38:HighlightRow"
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
      Begin MSMask.MaskEdBox MBox1 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   73
         Top             =   390
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         Enabled         =   0   'False
         MaxLength       =   6
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Nro. Planilla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   255
         TabIndex        =   74
         Top             =   390
         Width           =   1215
      End
   End
End
Attribute VB_Name = "xxxFrmNotaFisico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                    
Option Explicit

Dim nLongicampo(6) As Integer
Dim rsdeta As New ADODB.Recordset
Dim wCabe(40)

Dim apedido As String
Dim aalmacen As String
Dim alista As String * 2

'****** Totales de Pedidos***
Dim Tbruto As Double
Dim Tigv As Double
Dim Tdscto As Double
Dim TSub As Double
Dim TImporte As Double
Dim TNeto As Double
Dim TCant As Double
Dim flag As Integer

'***Total Descuentos  ***
Dim DTGlobal As Double
Dim DTCliente As Double
Dim DTPPago As Double
Dim DTOficina As Double
Dim DTItem As Double
Dim DTLinea As Double
Dim DTPromo As Double

'*****************

Dim dllgeneral As New dllgeneral.dll_general

'Mensajes de Pedidos

Const W1TXT1 = "El Proveedor No Existe en el Maestro de Proveedores"
Const W1TXT2 = "El Proveedor No Tiene Número de R.U.C. en el Maestro"
Const W1TXT3 = "El Cliente Esta Suspendido No Atender"
Const W1TXT4 = "El Cliente Ya No Tiene Credito. No Atender"

Const W1TXT6 = "Codigo del Vendedor No Existe en Tabla de Vendedores"
Const W1TXT7 = "El Codigo del Almacen No Existe en Tabla de Almacenes"

Const W1TXT9 = "El Monto de Otros Gastos debe ser un Valor Positivo"

Const W1TXT12 = "El Descuento General debe ser un Valor Positivo"
Const W1TXT13 = "El Descuento de Promoci¢n debe ser un Valor Positivo"
Const W1TXT14 = "El Descuento Pronto Pago debe ser un Valor Positivo"
Const W1TXT17 = "Codigo de la Lista de Precios No Existe"
Const W1TXT18 = "Archivo Maestro de la Lista de Precios No Existe"
Const W1TXT19 = "Codigo del Artículo No Existe en Maestro de Artículos "
Const W1TXT20 = "El Codigo del Articulo No Existe en Maestro de Precios"
Const W1TXT21 = "El Codigo del Articulo Ya Existe en el Proceso de Ventas"
Const W1TXT22 = "La Cantidad a Vender debe ser un Valor Mayor que Cero"
Const W1TXT23 = "La Cantidad a Vender es Mayor que el Actual en Almacén"
Const W1TXT24 = "El Precio de Venta debe de ser un Valor Mayor que Cero"
Const W1TXT25 = "El Descuento por Item debe ser un Valor Positivo"
Const W1TXT28 = "Debe de Ingresar el Nro. de R.U.C. del Proveedor"
Const W1TXT30 = "El Importe debe ser mayor a cero"
Const W1TXT31 = ""
Const W1TXT32 = ""
Const W1TXT33 = ""


Private Sub cAyuda_Click(Index As Integer)
  nAyuda = "": nDetalle = ""
  MBox(7) = Right("000000000000" & MBox(7), MBox(7).MaxLength)
  MBox(8) = Right("000000000000" & MBox(8), MBox(8).MaxLength)
  If Index = 3 Then
    SendKeys "{tab}"
    Exit Sub
  End If
  
 If Trim(MBox(6)) <> "" And CDbl(Trim(MBox(7))) > 0 And CDbl(Trim(MBox(8))) > 0 Then
    SendKeys "{tab}"
    Exit Sub
 End If
 
 If dllgeneral.VerificaDatoExistente(cn, "select * from vt_pedido where clientecodigo='" & Trim(Ctr_Ayuda1.xclave) & "' ") = 1 Then
       Dim gfiltra(1, 2) As String
       gfiltra(1, 1) = "Documento": gfiltra(1, 2) = "pedidonrofact"
       'gfiltra(2, 1) = g_tipobol: gfiltra(2, 2) = "pedidonroboleta"
       FrmAyuda.TipoForma = "1"
       'FrmAyuda.Bdato = Escadena(MBox2(1).Text)
       FrmAyuda.BConexion = cn
       FrmAyuda.BTabla = "vt_pedido"
       FrmAyuda.BCampos = "pedidotipofac as Tipo,pedidonrofact as Documento,pedidofecha as Fecha,pedidomoneda as Moneda,pedidototneto as Total"
       FrmAyuda.BOrden = "pedidofecha"
       FrmAyuda.BCondi = "clientecodigo='" & Trim(Ctr_Ayuda1.xclave) & "'"  ' and pedidocondicionfactura='0' and len(ltrim(pedidonrofact))>1"
       FrmAyuda.BFiltro = gfiltra

 Else
        nAyuda = "": nDetalle = ""
        MsgBox "No existen documentos pendientes...", vbInformation, MsgTitle
        Exit Sub
 End If
 FrmAyuda.Show 1
 If Len(Escadena(nAyuda)) > 0 Then
    MBox(6) = Escadena(nAyuda): MBox(7) = Left(Escadena(nDetalle), 3): MBox(8) = Right(Escadena(nDetalle), 8)
    nAyuda = "": nDetalle = ""
    DoEvents
    Call Carga_Pedido
    Exit Sub
 End If
 nAyuda = "": nDetalle = ""

End Sub




Private Sub cmdBotones_Click(Index As Integer)
   Dim asql As String
   Dim acmd As New ADODB.Command
   Dim j, nl As Integer
   
   On Error GoTo vererror
   
   Select Case Index
    Case 11
        If IsNull(Ctr_Ayuda1.xclave) Or Len(Trim(Ctr_Ayuda1.xclave)) = 0 Then
           MsgBox W1TXT1, vbInformation, MsgTitle
           Ctr_Ayuda1.SetFocus
           Exit Sub
        End If
        If IsNull(Ctr_Ayuda2.xclave) Or Len(Trim(Ctr_Ayuda2.xclave)) = 0 Then
           MsgBox W1TXT6, vbInformation, MsgTitle
           Ctr_Ayuda2.SetFocus
           Exit Sub
        End If
        If IsNull(Ctr_Ayuda3.xclave) Or Len(Trim(Ctr_Ayuda3.xclave)) = 0 Then
           MsgBox "Codigo de conceptos no existe,vbInformation, MsgTitle"
           Ctr_Ayuda3.SetFocus
           Exit Sub
        End If
        If IsNull(MBox1(2).ClipText) Or Len(Trim(MBox1(2).ClipText)) = 0 Or CDbl(MBox1(2).ClipText) <= 0 Then
           MsgBox "Falta Tipo de Cambio", vbInformation, MsgTitle
           Exit Sub
        End If
        If dllgeneral.VerificaDatoExistente(cn, "select * from cp_proveedor where clientecodigo='" & Ctr_Ayuda1.xclave & "' and clientesuspendido='1'") = 1 And Ctr_Ayuda1.xclave <> g_Eventual Then
           MsgBox W1TXT3, vbInformation, MsgTitle
           Exit Sub
        End If
        If dllgeneral.VerificaDatoExistente(cn, "select * from cp_proveedor where clientecodigo='" & Ctr_Ayuda1.xclave & "' and ((clientelimitecreddolar-clientesaldodolares)*" & MBox(8) & "+ (clientelimitecredsoles-clientesaldosoles))-" & TNeto & " <=0") = 1 And Ctr_Ayuda1.xclave <> g_Eventual Then
           MsgBox W1TXT4, vbInformation, MsgTitle
           Exit Sub
        End If
        If CDbl(MBox(4)) <> CDbl(MBox2(10)) Then
           MsgBox "Los Totales no son iguales...Verifique!!!", vbInformation, MsgTitle
           Exit Sub
        End If
        
        cn.BeginTrans
        If GrabarData() = 1 Then
          cn.CommitTrans
          g_TipoMovi = 0
'          If modoventa.emitehoja = "1" Then
'             nl = IIf(modoventa.copiashoja > 0, modoventa.copiashoja, 0)
'             If nl > 0 Then
'                 For J = 1 To nl
'                    Call DocImprimir
'                 Next J
'             End If
'          End If
          Activa 2
          Exit Sub
        Else
           cn.RollbackTrans
           g_TipoMovi = 0
           Activa 2
           Exit Sub
        End If
       g_TipoMovi = 0
    Case 12
       g_TipoMovi = 0
       Unload Me
   End Select
   
vererror:
    If Err Then
       MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
       Err = 0
       'cn.RollbackTrans
       Exit Sub
    End If
End Sub

Public Function Activa(ntipo As Integer)
    If ntipo = 1 Then
        SSTab1.TabEnabled(0) = False
'        SSTab1.TabEnabled(1) = True
        SSTab1.Tab = 1
    ElseIf ntipo = 2 Then
        SSTab1.TabEnabled(0) = True
'        SSTab1.TabEnabled(1) = False
        SSTab1.Tab = 0
    End If
End Function

Private Sub Combo1_Click()
'   MBox(8) = Numero(0) ' Numero(TraeDataSerie("select * from ct_tipocambio where tipocambiofecha=GETDATE()", cnconta))
 '  MBox(8) = TraeDataSerie("select * from ct_tipocambio where tipocambiofecha=GETDATE()", cn)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
  Seguir Combo1, KeyAscii
End Sub

Private Sub Combo2_Click()
  Dim rs As New ADODB.Recordset
  
  If Combo2.ListCount > 0 Then
     Set rs = cn.Execute("select * from vt_puntovtadocumento where puntovtacodigo='" & g_ptoventa & "' and documentocodigo='" & dllgeneral.ComboDato(Combo2.Text) & "'")
     If rs.RecordCount > 0 Then
        MBox(1) = Escadena(rs!puntovtadocserie)
        MBox(2) = Escadena(rs!puntovtadoccorr)
     End If
     rs.Close
     
     Set rs = Nothing
  Else
     MsgBox "No tiene Serie ...Verifique!!", vbInformation, MsgTitle
     Combo2.SetFocus
  End If
  
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
  
  
  Seguir Combo2, KeyAscii
End Sub



Private Sub Ctr_Ayuda3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    
     MBox2(0).SetFocus
  End If
End Sub

Private Sub Form_Load()

   MostrarForm Me, "C"
   MBox(1).Enabled = False
   MBox(2).Enabled = False
   
   MBox1(1) = Format(Date, "dd/mm/yyyy")
   MBox1(2) = Numero(TraeTipoCambio(Date, cn))
   
   Call Ctr_Ayuda1.conexion(cn)
   Call Ctr_Ayuda3.conexion(cn)
   Call Ctr_Ayuda2.conexion(cn)
   
   Call dllgeneral.llenacombo(Combo1, "SELECT * from gr_moneda", cn)
   Call dllgeneral.llenacombo(Combo2, "SELECT * from vt_documento where documentotipo='A'", cn)
   flag = 0
   
End Sub



Public Function CargaGrilla()

   Call rsdeta.Fields.Append("Item", adInteger)
   Call rsdeta.Fields.Append("Codigo", adChar, 8)
   Call rsdeta.Fields.Append("Descripcion", adChar, 40)
   Call rsdeta.Fields.Append("UM", adChar, 3)
   Call rsdeta.Fields.Append("Cant", adDouble)
   Call rsdeta.Fields.Append("Precio_Vta", adDouble)
   Call rsdeta.Fields.Append("Dscto(%)", adDouble)
   Call rsdeta.Fields.Append("Total", adDouble)
   Call rsdeta.Fields.Append("%", adDouble)
   Call rsdeta.Fields.Append("CantRef", adDouble)
   Call rsdeta.Fields.Append("Factor", adDouble)
   Call rsdeta.Fields.Append("%P", adDouble)
   
   rsdeta.Open
   If rsdeta.RecordCount > 0 Then
     Totales
   End If
   ConfigGrid

End Function

Public Function ConfigGrid()
   Set TDBGrid1.DataSource = Nothing
   
   Set TDBGrid1.DataSource = rsdeta
   With TDBGrid1
      .Columns(0).Width = 600
      .Columns(0).Caption = "Item"
      .Columns(1).Width = 1100
      .Columns(1).Caption = "Codigo"
      .Columns(2).Width = 3000
      .Columns(2).Caption = "Descripcion"
      .Columns(3).Width = 600
      .Columns(3).Caption = "UM"
      .Columns(4).Width = 1000
      .Columns(4).Caption = "Cant"
      .Columns(5).Width = 1000
      .Columns(5).Caption = "Precio_Vta"
      .Columns(6).Width = 1000
      .Columns(6).Caption = "Dscto(%)"
      .Columns(7).Width = 800
      .Columns(7).Caption = "Total"
      .Columns(8).Width = 1000
      .Columns(8).Caption = "%"
      .Columns(5).NumberFormat = "###,##0.0000"
      .Columns(6).NumberFormat = "###,##0.00"
      .Columns(7).NumberFormat = "###,##0.0000"
      .Columns(8).NumberFormat = "###,##0.00"
      .Columns(9).Width = 800
      .Columns(9).Caption = "Cant.Ref"
      .Columns(9).NumberFormat = "###,##0"
      .Columns(10).Width = 600
      .Columns(10).Caption = "Factor"
      .Columns(10).NumberFormat = "###,##0.00"
      .Columns(11).Width = 0
      .Columns(11).NumberFormat = "###,##0.00"
   End With
   TDBGrid1.Refresh
End Function

Private Sub Form_Unload(Cancel As Integer)
  Set rsdeta = Nothing
End Sub

Private Sub MBox_GotFocus(Index As Integer)
  Call dllgeneral.Enfoquetexto(MBox(Index))
End Sub


Private Sub MBox_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Index = 8 Then
     DoEvents
     Call Carga_Pedido
    ElseIf Index = 4 Then
      If Not dllgeneral.ValidaCadena(MBox(Index), "N") Then
         MsgBox Msg29, vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox(Index))
         Exit Sub
      End If
      MBox(Index) = Format(MBox(Index), "##,###,##0.00")
    End If
  End If
  Seguir MBox(Index), KeyAscii
End Sub

Private Sub MBox_LostFocus(Index As Integer)
  On Error Resume Next
  Select Case Index
   Case 4
      If Not dllgeneral.ValidaCadena(MBox(Index), "N") Then
         MsgBox Msg29, vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox(Index))
         Exit Sub
      End If
      MBox(Index) = Format(MBox(Index), "##,###,##0.00")
   Case 3, 5
      If Not dllgeneral.ValidaCadena(MBox(Index), "F") Then
         MsgBox "Fecha No Valida", vbInformation, "AVISO"
         'Call dllgeneral.Enfoquetexto(MBox(Index))
         Exit Sub
      End If
   Case 7, 8
        MBox(Index) = Right("000000000000" & MBox(Index), MBox(Index).MaxLength)
        
'      If Not dllgeneral.ValidaCadena(MBox(Index), "D") Then
'         MsgBox Msg29, vbInformation, "AVISO"
'         Call dllgeneral.Enfoquetexto(MBox(Index))
'         Exit Sub
'      End If
 '     MBox(Index) = Right("000000000000" & MBox(Index), MBox(Index).MaxLength)
      Exit Sub
  End Select
End Sub

Private Sub MBox2_GotFocus(Index As Integer)
  On Error Resume Next
'  If Index = 3 Then
'     Call TraerProducto
 ' End If
  Call dllgeneral.Enfoquetexto(MBox2(Index))
End Sub

Private Sub MBox2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  If KeyCode = 13 Then
    If Index = 12 Then
      MBox2(Index) = Format(MBox2(Index), "##,###,##0")
 '   ElseIf Index = 1 Then
 '     Call TraerProducto
    End If
    SendKeys "{tab}"
  ElseIf Index = 1 Then
      If dllgeneral.ValidaCadena(Trim(MBox2(1).ClipText), "N") = False Then
        MBox2(1).MaxLength = 64
      Else
        MBox2(1).MaxLength = 8
      End If
  End If
End Sub

Private Sub MBox2_LostFocus(Index As Integer)
  Dim nregi As Long
  Dim wposi, posi As Integer
  Dim ntabla As String
  Dim wflag As Integer
  
  On Error Resume Next
  
  Select Case Index
   Case 0
      If Not dllgeneral.ValidaCadena(MBox2(Index), "N") And Len(Trim(MBox2(Index))) > 0 Then
         MsgBox Msg29, vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox2(Index))
         Exit Sub
      End If
   Case 1
      ntabla = IIf(Combo2.ListCount > 0, "listapre" & alista, "vt_producto")
      If dllgeneral.VerificaDatoExistente(cn, "select * from " & ntabla & " where productocodigo='" & MBox2(Index).Text & "' and almacencodigo='" & aalmacen & "'") = 0 And Len(Trim(MBox2(Index))) > 0 Then
         'MsgBox W1TXT20, vbInformation, "AVISO"
         'Call dllgeneral.Enfoquetexto(MBox2(Index))
         Call cAyuda_Click(3)
         MBox2(1).MaxLength = 8
         Exit Sub
      Else
        wflag = verificaproducto()
        If wflag = 1 Then
            Label2 = ""
            MsgBox "Ya ingreso el producto...Verifique!!!", vbInformation, MsgTitle
            MBox2(1).SetFocus
            Exit Sub
         End If
            
      End If
   Case 3, 4, 5
      'If Index = 3 Then ' And dllgeneral.ComboDato(Combo5.Text) = "N" Then
      '    Call TraerProducto
      'End If
      If Not dllgeneral.ValidaCadena(MBox2(Index), "N") And Len(Trim(MBox2(Index))) > 0 Then
         MsgBox Msg29, vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox2(Index))
         Exit Sub
      End If
      If Not dllgeneral.ValidaCadena(MBox2(0), "N") Then
         MsgBox Msg29, vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox2(0))
         Exit Sub
      End If
      wflag = verificaproducto()
      If wflag = 1 Then
        Label2 = ""
        MsgBox "Ya ingreso el producto...Verifique!!!", vbInformation, MsgTitle
        MBox2(1).SetFocus
        Exit Sub
      End If
      If Index = 5 Then
         If Trim(MBox2(3)) = "" Or Trim(MBox2(4)) = "" Or Trim(MBox2(5)) = "" Then
           MsgBox Msg29, vbInformation, "AVISO"
           Call dllgeneral.Enfoquetexto(MBox2(1))
           Exit Sub
         End If
      End If
      If Index Like "[45]" Then
         MBox2(Index) = Numero(MBox2(Index))
       Else
         MBox2(Index) = Format(MBox2(Index), "######0.000")
       End If
       If Index = 5 And Len(Trim(MBox2(Index))) > 0 Then
'        If modoventa.nroitem < TDBGrid1.ApproxCount Then
'           MsgBox "Excede el Numero de Items del Documento..!!", vbInformation, MsgTitle
'           Exit Sub
'        End If
        nregi = 0
        wposi = 0
        If rsdeta.RecordCount > 0 Then
            rsdeta.MoveLast
            wposi = rsdeta.Fields(0)
            posi = rsdeta.Fields(0)
            rsdeta.MoveFirst
            Do Until rsdeta.EOF
               If rsdeta.Fields(0) = MBox2(11) Then
                  posi = rsdeta.Fields(0)
                  Exit Do
               End If
               nregi = nregi + 1
               rsdeta.MoveNext
            Loop
        End If
        If rsdeta.RecordCount = nregi Then
          wposi = wposi + 1
          posi = wposi
          rsdeta.AddNew
        End If
        rsdeta.Fields(0) = posi
        rsdeta.Fields(1) = Escadena(MBox2(1))
        rsdeta.Fields(2) = Left(Escadena(Label2) & Space(40), 40)
        rsdeta.Fields(3) = Trim(MBox2(2))
        rsdeta.Fields(4) = Escadena(MBox2(0))
        If parametro.tieneigv = "1" Then
           rsdeta.Fields(5) = (MBox2(3) / (1 + parametro.igv))
        Else
           If modoventa.impuestos = "1" Then
              rsdeta.Fields(5) = (MBox2(3) / (1 + parametro.igv))
           Else
              rsdeta.Fields(5) = MBox2(3)
           End If
        End If
        rsdeta.Fields(6) = Numero(MBox2(4))
        rsdeta.Fields(7) = Numero(MBox2(0) * MBox2(3))   ' IIf(parametro.tieneigv = "1", (MBox2(3) / (1 + (parametro.igv / 100))), MBox2(3)))
        rsdeta.Fields(8) = Numero(MBox2(5))
        rsdeta.Fields(9) = IIf(Len(Trim(MBox2(12))) = 0, 0, Format(MBox2(12), "##,###,##0"))
        rsdeta.Fields(10) = Numero(MBox2(13))
        rsdeta.Fields(11) = IIf(IsNull(MBox2(14)) Or Len(Trim(MBox2(14))) = 0, 0, MBox2(14))
        rsdeta.Update
        Label2 = ""
        TDBGrid1.Row = rsdeta.RecordCount - 1
        
        ConfigGrid
        Totales
        MBox2(11) = wposi + 1
        If MBox2(12).Enabled = True Then
          MBox2(12).SetFocus
        Else
          MBox2(0).SetFocus
        End If
        flag = 0
        Exit Sub
    End If
  End Select

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
  If SSTab1.Tab = 2 Then
     MBox2(0).SetFocus
  ElseIf SSTab1.Tab = 1 Then
     If MBox(0).Enabled = True Then
        MBox(5).SetFocus
     Else
        MBox(5).SetFocus
     End If
  End If
End Sub

 
 
Public Function Totales()
  Dim j As Double
  Dim Previo As Double
  Dim dct02, dct03, dct04, dct05, dct06 As Double
  
  Tbruto = 0: Tigv = 0: Tdscto = 0: TNeto = 0: TCant = 0
  TImporte = 0: TSub = 0
  '--Totales de Descuentos
  DTGlobal = 0: DTCliente = 0: DTPPago = 0: DTOficina = 0: DTItem = 0
  DTLinea = 0: DTPromo = 0
  
  If rsdeta.RecordCount > 0 Then
    rsdeta.MoveFirst
    For j = 0 To rsdeta.RecordCount - 1
       'IMPORTE DE MONTO BRUTO SIN IGV, ES DECIR PRECIO X CANTIDAD
       
       Tbruto = Tbruto + (rsdeta.Fields(4) * rsdeta.Fields(5))
       TCant = TCant + rsdeta.Fields(4)
       TImporte = (rsdeta.Fields(4) * rsdeta.Fields(5))
       
       'DESCUENTO DE CIA O EMPRESA
       'If parametro.tienedscto = "1" Then
       '     dct06 = TImporte * (1 + parametro.descuento)
       'Else
       dct06 = 0
       'End If
         'dct06 = TImporte * (1 + parametro.descuento)
  '        dct06 = TImporte * (CDbl(Text1))
  '     End If
       
       DTCliente = DTCliente + dct06
       
       'DESCUENTO POR ITEM
       dct02 = 0
       'dct02 = (TImporte * (rsdeta.Fields(6) / 100))
       
       DTItem = DTItem + dct02
       
       'DESCUENTO ESPECIAL  :w8dct03 =(w8bruto - w8dct02-w8dct06)*w2dctpp/100
        'dct03 = (TImporte - dct02 - dct06) * (MBox(7) / 100)            '(Tbruto-dct02-dct06)
        dct03 = 0
        
        DTPPago = DTPPago + dct03
        
       'DESCUENTO POR PROMOCION  : w8dct04 =(w8bruto - w8dct02-w8dct03-w8dct06)*w2dctpr/100
        'dct04 = (TImporte - dct02 - dct03 - dct06) * (MBox(6) / 100)
        
        dct04 = 0
        
        DTPromo = DTPromo + dct04
        
       'DESCUENTO GENERAL : w8dct05 =(w8bruto - w8dct02-w8dct03-w8dct04-w8dct06)*w2dctgl/100
        'dct05 = (TImporte - dct02 - dct03 - dct04 - dct06) * (MBox(5) / 100)
        dct05 = 0
        
        DTGlobal = DTGlobal + dct05
       
       'ACUMULADO DE TOTAL DESCUENTOS  :w8dctos = w8dct02 + w8dct03+w8dct04+w8dct05+w8dct06
        Tdscto = Tdscto + (dct02 + dct03 + dct04 + dct05 + dct06)
        
       'ACUMULADO DE SUBTOTAL DE VENTA : w8subto = w8bruto - w8dctos
        TSub = TSub + (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
                
       If parametro.tieneigv = "1" Then
            'CALCULAMOS EL IMPORTE :=  TOTAL IMPORTE SIN IGV - DESCTOS + IGV
            Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
            Previo = (Previo * parametro.igv)
            Tigv = Tigv + Previo
            
            'GRABAMOS EL TOTAL DE IMPORTE EN LA TABLA TEMPORAL PARA MOSTRAR
            Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
            Previo = (Previo * (1 + parametro.igv))
            rsdeta.Fields(7) = Previo
       Else                    'If parametro.tieneigv = "0" Then
             If rsdeta.Fields(11) > 0 Then
                  Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
                  rsdeta.Fields(7) = Previo * (1 + rsdeta(11))
                  Tigv = Tigv + (Previo * rsdeta(11))
             Else
                  Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
                  rsdeta.Fields(7) = Previo
                  'Tigv = Tigv
            End If
       End If
       rsdeta.Update
      
       rsdeta.MoveNext
    Next j
  Else
    Exit Function
  End If
  
  ' w2imp = IIf(w2ciaimp, w2timp, pro_pctimp)
  ' w2imp = IIf(vtmod.mod_imp, w2imp, 0)
  ' w2prepac = IIf(w2dctofe > 0, roun(w2prepac * (100 - w2dctofe) / 100, 4), w2prepac
   
   'set deci to 12
   'w8bruto = w2cant  * w2prepac                                          && Total Bruto
   'w2dctofe = 0
   'If w2fchatn>=pro_fchini and w2fchatn<=pro_fchfin                      && Precio de Oferta en una lista de precios
   '   w2dctofe =pro_dctofi      && Descuentos Ofertas
   '*   w8dct01 = w2cant  * Abs(IIF(w2prelis>w2prepac,w2prelis-w2prepac,0))   && Dcto.Oferta
   'w8dct06 = w8bruto * w0dcto/100                                        && Dcto. por Default
   'w8dct02 = (w8bruto-w8dct06)*w2dctlin/100                              && Dcto.Por Item
  
  'IMPORTE TOTAL NETO DE FACTURA   w8tneto = w8subto + w8impto
  TNeto = Tbruto - Tdscto + Tigv
  MBox2(7) = Format(Tbruto, "#,###,##0.0000")
  MBox2(6) = Numero(TCant)
  MBox2(9) = Numero(Tigv)
  MBox2(8) = Numero(Tdscto)
  MBox2(10) = Numero(TNeto)
  Limpiartexto MBox2, 12, 12
  Limpiartexto MBox2, 13, 13
  Limpiartexto MBox2, 14, 14
  Limpiartexto MBox2, 0, 5
  
End Function

Private Sub TDBGrid1_Click()
   If rsdeta.RecordCount > 0 Then
      TDBGrid1.SetFocus
   End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim nvalor As String
  If KeyCode = 46 Then
     If rsdeta.RecordCount <= 0 Then
        Limpiartexto MBox2, 6, 10
        Exit Sub
     End If
     nvalor = TDBGrid1.Columns(0).Text
     If rsdeta.RecordCount > 0 Then
        rsdeta.MoveFirst
        Do Until rsdeta.EOF
          If rsdeta.Fields(0) = nvalor Then
            rsdeta.Delete adAffectCurrent
            rsdeta.Update
            Exit Do
          End If
          rsdeta.MoveNext
        Loop
     End If
     ConfigGrid
     Totales
     Exit Sub
  ElseIf KeyCode = 13 Then
    Limpiartexto MBox2, 0, 5
    MBox2(11) = TDBGrid1.Columns(0).Text
    MBox2(0) = TDBGrid1.Columns(4).Text
    MBox2(1) = TDBGrid1.Columns(1).Text
    Label2 = TDBGrid1.Columns(2).Text
    MBox2(2) = Escadena(TDBGrid1.Columns(3).Text)
    MBox2(12) = Escadena(TDBGrid1.Columns(9).Text)
    
    If parametro.tieneigv = "1" Then
       If TDBGrid1.Columns(5).Text <> "" Then
         MBox2(3) = Format(TDBGrid1.Columns(5).Text * (1 + (parametro.igv)), "######0.000")
       End If
       
    Else
       If modoventa.impuestos = "1" Then
           MBox2(3) = Format(IIf(IsNull(TDBGrid1.Columns(5).Text) Or Len(Trim(TDBGrid1.Columns(5).Text)) = 0, 0, TDBGrid1.Columns(5).Text) * (1 + (parametro.igv)), "######0.000")
       Else
           MBox2(3) = Format(TDBGrid1.Columns(5).Text, "######0.000")
       End If
    End If
    MBox2(4) = Numero(TDBGrid1.Columns(6).Text)
    MBox2(5) = Numero(TDBGrid1.Columns(8).Text)
    If MBox2(12).Enabled = True Then
      MBox2(12).SetFocus
    Else
      MBox2(0).SetFocus
    End If
    flag = 1
  End If
  
End Sub


Public Function Carga_Pedido()
    Dim csql As New ADODB.Recordset
    Dim acliente As New ADODB.Recordset
    Dim j As Integer
    
    MBox(7) = Right("000000000000" & Trim(MBox(7)), MBox(7).MaxLength)
    MBox(8) = Right("000000000000" & Trim(MBox(8)), MBox(8).MaxLength)
    
    Set csql = cn.Execute("select * from vt_pedido where pedidotipofac='" & MBox(6) & "' and pedidonrofact='" & MBox(7) & MBox(8) & "'")
    
    If csql.RecordCount > 0 Then
        apedido = Escadena(csql!pedidonumero)                      'nro pedido
        alista = Val(Trim(csql!pedidolistaprec))                         'lista precios
        aalmacen = Escadena(csql!almacencodigo)                     'almacen
    Else
       MsgBox "No existe documentos....!!", vbInformation, MsgTitle
       Exit Function
    End If
    csql.Close

    Set csql = cn.Execute("select detpeditem,A.productocodigo,b.productodescripcion,a.unidadcodigo," & _
                          "detpedcantpedida,detpedmontoprecvta,detpeddsctoxitem,detpedimpbruto," & _
                          " detpedporccomis,detpedcantpedidaref,detpedfactorconv " & _
                          "from vt_detallepedido  A " & _
                          "inner Join " & "listapre" & Trim(alista) & " B " & _
                          "on A.productocodigo=B.productocodigo " & _
                          "where pedidonumero='" & apedido & "' and B.almacencodigo='" & aalmacen & "'")
    
    Set rsdeta = Nothing
    Call CargaGrilla
   
    Do Until csql.EOF
       rsdeta.AddNew
       rsdeta.Fields(0) = Escadena(csql!detpeditem)
       rsdeta.Fields(1) = Escadena(csql!productocodigo)
       rsdeta.Fields(2) = Escadena(csql!productodescripcion)
       rsdeta.Fields(3) = Escadena(csql!unidadcodigo)
       rsdeta.Fields(4) = Numero(csql!detpedcantpedida)
       rsdeta.Fields(5) = Numero(IIf(IsNull(csql!detpedmontoprecvta), 0, csql!detpedmontoprecvta))
       rsdeta.Fields(6) = Numero(csql!detpeddsctoxitem)
       rsdeta.Fields(7) = Numero(csql!detpedimpbruto)
       rsdeta.Fields(8) = Numero(csql!detpedporccomis)
       rsdeta.Fields(9) = Numero(IIf(IsNull(csql!detpedcantpedidaref), 0, csql!detpedcantpedidaref))
       rsdeta.Fields(10) = Numero(IIf(IsNull(csql!detpedfactorconv), 0, csql!detpedfactorconv))
       rsdeta.Update
       csql.MoveNext
    Loop
    csql.Close
    Call Totales
    Call ConfigGrid
    Set csql = Nothing

End Function



Public Function GrabarData() As Integer
    Dim j As Integer
    Dim regi As Long
    Dim nsql As String
    Dim ltipo As String
    Dim lzona As String
    Dim Previo As Double
    Dim dct02, dct03, dct04, dct05, dct06 As Double
    Dim tinafecto As Double
    Dim xserie As String * 3
    Dim xfactu As String * 5
    Dim xtipofac As String * 2
    Dim fechasunat As Date
    Dim tcargo As String
    
    Dim acmd As New ADODB.Command
    Dim asql As New ADODB.Recordset
    Dim arbusca As New ADODB.Recordset

    On Error GoTo vererror
    
    GrabarData = 0
    
    '******** CABECERA DE MOVIMIENTO *****************
    
    If CDbl(MBox(4)) <> CDbl(MBox2(10)) Then
       MsgBox "Los Totales no son iguales...Verifique!!!", vbInformation, MsgTitle
       Exit Function
    End If
    If rsdeta.RecordCount = 0 Then
      MsgBox W1TXT30, vbInformation, MsgTitle
      GrabarData = 0
      Exit Function
    End If
    
    Call Totales
    For j = 1 To 29
        wCabe(j) = ""
    Next j
    
    fechasunat = Date
    
    Set asql = cn.Execute("select * from vt_pedido where pedidonumero='" & apedido & "'")
    If asql.RecordCount > 0 Then
        wCabe(1) = Escadena(asql!puntovtacodigo)     'Escadena(asql!p)                       'Pto Venta
        wCabe(2) = Escadena(asql!pedidonumero)      'Trim(MBox(1))                       'nro pedido
        wCabe(3) = Escadena(asql!pedidonrofact)      'Trim(MBox(2))                        'nro factura
        wCabe(4) = Escadena(asql!pedidonrofact)      'Trim(MBox(3))                         'nro boleta
        wCabe(5) = Escadena(asql!pedidonrofact)      'Trim(MBox(4))                         'nro guia
        wCabe(6) = 0      'MBox(5)                       'dscto gral
        wCabe(7) = 0      'MBox(6)                       'dscto promocional
        wCabe(8) = 0      'MBox(7)                       'dscto especial
        wCabe(9) = dllgeneral.ComboDato(Combo1.Text)        'moneda
        wCabe(10) = CDbl(MBox1(2))                      'tipo de cambio
        wCabe(11) = CDbl(Escadena(asql!pedidolistaprec))    'dllgeneral.ComboDato(Combo2.Text)       'lista de precios
        wCabe(12) = " "                                'MBox(9)                      'mensajes
        wCabe(13) = Escadena(asql!modovtacodigo)     'dllgeneral.ComboDato(Combo3.Text)       'modo de venta
        wCabe(14) = MBox(3)                         'MBox(10)                     'fecha de atencion
        wCabe(15) = Escadena(asql!formapagocodigo)     'dllgeneral.ComboDato(Combo4.Text)       'forma de pago
        wCabe(16) = Ctr_Ayuda1.xclave         ' MBox(11)                     'cliente
        wCabe(17) = Ctr_Ayuda2.xclave        'MBox(12)                     'vendedor
        wCabe(18) = 0    'MBox(13)                  'comision
        wCabe(19) = Escadena(asql!almacencodigo)    'Ctr_Ayuda3.xclave        'MBox(14)                     'almacen
        wCabe(20) = 0      'MBox(15)                     'otros gastos
        wCabe(21) = "0"      'MBox(16)                     'nota pedido
        wCabe(22) = "0"      'MBox(17)                     'orden de compra
        wCabe(23) = Escadena(asql!pedidoautorizacion)      'dllgeneral.ComboDato(Combo5.Text)       'autorizacion
        wCabe(24) = 0       'MBox(18)                     'dias pago
        wCabe(25) = MBox2(6)                    'Total Cantidad
        wCabe(26) = Round(MBox2(7), 2)          'Total Bruto
        wCabe(27) = 0    'MBox2(8)              'total fletes --T.D.
        wCabe(28) = Round(MBox2(9), 2)          'Total Igv
        wCabe(29) = Round(MBox2(10), 2)         'Neto a Facturar
        wCabe(30) = Escadena(asql!pedidoentrega)    'MBox(19)                    'entrega pedido
        wCabe(31) = Escadena(asql!clienterazonsocial)  'MBox3(1)                    'nombre cliente
        wCabe(32) = Escadena(asql!clientedireccion)    'MBox3(3)                    'direccion
        wCabe(33) = Escadena(asql!clienteruc)  'MBox3(2)                    'ruc
        wCabe(34) = MBox(3)    'Date                           'fechafactura
        wCabe(35) = DTGlobal                     'Total Descuentos Globales
        wCabe(36) = DTCliente                    'Total Descuentos Cliente
        wCabe(37) = DTOficina                    'Total Descuentos Oficina
        wCabe(38) = DTItem                       'Total Descuentos Item
        wCabe(39) = DTLinea                      'Total Descuentos Linea
        wCabe(40) = DTPromo                      'Total Descuentos x Promocion
        fechasunat = IIf(IsNull(asql!pedidofechasunat), MBox(3), asql!pedidofechasunat)
    Else
        MsgBox "Datos Incompletos del Pedido : " & apedido, vbInformation, MsgTitle
        Exit Function
    End If
    asql.Close
    
    Set asql = Nothing
    
'    nsql = "Insert Into vt_Pedido ("

'    cn.CommitTrans
    
'    cn.BeginTrans
    ' ** Verificando Numeracion de Documentos *****
     wCabe(2) = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoped & "' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "'", cn), 8)
     wCabe(34) = Date                       'fechafactura
     wCabe(3) = g_facserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & dllgeneral.ComboDato(Combo2.Text) & "' and puntovtadocserie='" & g_facserie & "' and puntovtacodigo='" & g_ptoventa & "'", cn), 8)
     wCabe(4) = dllgeneral.ComboDato(Combo2.Text)
     wCabe(5) = "0"
     If dllgeneral.VerificaDatoExistente(cn, "select * from vt_pedido where pedidonrofact='" & MBox(1) & MBox(2) & "' and pedidotipofac='" & dllgeneral.ComboDato(Combo2.Text) & "'") = 1 Then
        MsgBox "Ya existe el Documento " & dllgeneral.ComboDato(Combo2.Text) & "-" & MBox(1) & MBox(2), vbInformation, MsgTitle
        GrabarData = 0
        Exit Function
     End If
    
    '*** Verifica Serie Documentos *****
    
    Set asql = cn.Execute("select puntovtadoccorr from vt_puntovtadocumento Where documentocodigo='" & g_tipoped & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_pedserie & "'")
    If asql.RecordCount > 0 Then
       wCabe(2) = asql!puntovtadoccorr
    End If
    asql.Close
    Set asql = Nothing
    
    nsql = "Update vt_puntovtadocumento " & _
            " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(wCabe(2) + 1)), 8) & "'" & _
            " Where documentocodigo='" & g_tipoped & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_pedserie & "'"
            
     cn.Execute nsql
    
     wCabe(3) = MBox(2).Text
     
    nsql = "Update vt_puntovtadocumento " & _
           " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(wCabe(3) + 1)), 8) & "'" & _
           " Where documentocodigo='" & dllgeneral.ComboDato(Combo2.Text) & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & MBox(1).Text & "'"
        
    cn.Execute nsql
                   
    DoEvents
    '**cambio de documentacion
    wCabe(5) = 0
    
    DoEvents
    '************
    
    Set acmd.ActiveConnection = cg
    acmd.CommandType = adCmdStoredProc
    acmd.CommandText = "cp_ingresanota_pro"
    acmd.CommandTimeout = 0
    acmd.Prepared = True
    With acmd
        .Parameters("@base") = cn.DefaultDatabase
        .Parameters("@tabla") = "vt_pedido"
        .Parameters("@tipo") = IIf(dllgeneral.VerificaDatoExistente(cn, "select * from vt_pedido where pedidonumero='" & wCabe(2) & "'") = 0, "1", "2")
        .Parameters("@puntovta") = wCabe(1)
        .Parameters("@numero") = wCabe(2)
        .Parameters("@factura") = wCabe(3)
        .Parameters("@boleta") = wCabe(4)
        .Parameters("@guia") = wCabe(5)
        .Parameters("@dsctoglobal") = wCabe(6)
        .Parameters("@dsctoppago") = wCabe(7)
        .Parameters("@dsctovtaofi") = wCabe(8)
        .Parameters("@moneda") = wCabe(9)
        .Parameters("@tipocambio") = wCabe(10)
        .Parameters("@listaprecio") = wCabe(11)
        .Parameters("@mensaje") = wCabe(12)
        .Parameters("@modoventa") = wCabe(13)
        .Parameters("@fecha") = wCabe(14)
        .Parameters("@formapago") = wCabe(15)
        .Parameters("@cliente") = wCabe(16)
        .Parameters("@vendedor") = wCabe(17)
        .Parameters("@porcomision") = wCabe(18)
        .Parameters("@almacen") = wCabe(19)
        .Parameters("@totalotros") = wCabe(20)
        .Parameters("@notaped") = wCabe(21)
        .Parameters("@ordencompra") = wCabe(22)
        .Parameters("@autoriza") = wCabe(23)
        .Parameters("@diaspago") = wCabe(24)
        .Parameters("@totalitem") = wCabe(25)
        .Parameters("@totalbruto") = wCabe(26)
        .Parameters("@totalflete") = wCabe(27)
        .Parameters("@totalimpuesto") = wCabe(28)
        .Parameters("@totalneto") = wCabe(29)
        .Parameters("@usuario") = g_usuario
        .Parameters("@fechaactual") = Date
        .Parameters("@totaldsctoxlinea") = wCabe(39)
        .Parameters("@montodsctoppago") = DTPPago
        .Parameters("@entregapedido") = wCabe(30)
        .Parameters("@razon") = wCabe(31)
        .Parameters("@direccion") = wCabe(32)
        .Parameters("@ruc") = wCabe(33)
        .Parameters("@fechafactura") = wCabe(34)
        .Parameters("@TDGlobal") = wCabe(35)
        .Parameters("@TDCliente") = wCabe(36)
        .Parameters("@TDOficina") = wCabe(37)
        .Parameters("@TDItem") = wCabe(38)
        .Parameters("@TDPromo") = wCabe(40)
        .Parameters("@tiporefe") = MBox(6)
        .Parameters("@nrorefe") = Trim(MBox(7) & MBox(8))
        .Parameters("@fsunat") = fechasunat
    End With
    acmd.Execute
    Set acmd = Nothing
    DoEvents
    
'    If wCabe(9) = g_TipoSol Then
'         cn.Execute "Update cp_proveedor " & _
'                    " Set clientesaldosoles=ISNULL(clientesaldosoles,0)-" & CDbl(wCabe(29)) & _
'                    "      Where clientecodigo='" & wCabe(16) & "'"
'    ElseIf wCabe(9) = g_TipoDolar Then
'         cn.Execute "Update cp_proveedor " & _
'                    " Set clientesaldodolares=ISNULL(clientesaldodolares,0)-" & CDbl(wCabe(29)) & _
'                    "      Where clientecodigo='" & wCabe(16) & "'"
'    End If
    
    '********** DETALLE DE MOVIMIENTOS *****************
    rsdeta.MoveFirst
    regi = 0
    tinafecto = 0
    Do Until rsdeta.EOF
        
           'IMPORTE DE MONTO BRUTO SIN IGV, ES DECIR PRECIO X CANTIDAD
           Tbruto = Tbruto + (rsdeta.Fields(4) * rsdeta.Fields(5))
           TCant = TCant + rsdeta.Fields(4)
           TImporte = (rsdeta.Fields(4) * rsdeta.Fields(5))
           
           'DESCUENTO DE CIA O EMPRESA
            dct06 = 0
          
           'DESCUENTO POR ITEM
            dct02 = 0
           
           'DESCUENTO ESPECIAL  :w8dct03 =(w8bruto - w8dct02-w8dct06)*w2dctpp/100
            dct03 = 0
            
           'DESCUENTO POR PROMOCION  : w8dct04 =(w8bruto - w8dct02-w8dct03-w8dct06)*w2dctpr/100
            dct04 = 0
            
           'DESCUENTO GENERAL : w8dct05 =(w8bruto - w8dct02-w8dct03-w8dct04-w8dct06)*w2dctgl/100
            dct05 = 0
           
           'ACUMULADO DE TOTAL DESCUENTOS  :w8dctos = w8dct02 + w8dct03+w8dct04+w8dct05+w8dct06
            Tdscto = dct02 + dct03 + dct04 + dct05 + dct06
            
           'ACUMULADO DE SUBTOTAL DE VENTA : w8subto = w8bruto - w8dctos
           TSub = 0
           TSub = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
           Previo = TSub
           If parametro.tieneigv = "1" Then
              'CALCULAMOS EL IGV
              Previo = (TSub * parametro.igv)
           Else
             If rsdeta.Fields(11) > 0 Then
                  Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
                  Previo = (Previo * rsdeta.Fields(11))
             Else
                 Previo = TSub '
                 tinafecto = tinafecto + TSub
            End If
          End If
        '*********
        
        nsql = "vt_detallepedido"
        
        Set acmd.ActiveConnection = cg
        acmd.CommandType = adCmdStoredProc
        acmd.CommandTimeout = 0
        acmd.CommandText = "vt_ingresodetallepedido_pro"
        acmd.Prepared = True
        With acmd
            '.Parameters("@base") = "VENTAS_PRUEBA"
            .Parameters("@base") = cn.DefaultDatabase
            .Parameters("@tabla") = nsql
            .Parameters("@tipo") = "1"
            .Parameters("@item") = rsdeta.Fields(0)
            .Parameters("@numero") = wCabe(2)
            .Parameters("@producto") = rsdeta.Fields(1)
            .Parameters("@unidad") = rsdeta.Fields(3)
            .Parameters("@cantidad") = rsdeta.Fields(4)
            .Parameters("@preciopacto") = rsdeta.Fields(5)
            .Parameters("@dsctoxitem") = rsdeta.Fields(6)
            .Parameters("@importebruto") = rsdeta.Fields(7)
            .Parameters("@porcomision") = rsdeta.Fields(8)
            .Parameters("@mdsctoitem") = Tdscto
            .Parameters("@mdsctoxlinea") = 0
            .Parameters("@mdsctoxprom") = Previo     '0
            .Parameters("@mimpor") = rsdeta.Fields(7)       'Previo
            .Parameters("@unidadref") = IIf(IsNull(rsdeta.Fields(9)) Or Len(Trim(rsdeta.Fields(9))) = 0, 0, CDbl(rsdeta.Fields(9)))
        End With
        acmd.Execute
        Set acmd = Nothing
            
            '******Actualizamos Saldos en Almacen *********
         xserie = Escadena(MBox(1).Text)
         xfactu = Escadena(MBox(2).Text)
         xtipofac = dllgeneral.ComboDato(Combo2.Text)
    
         Set acmd.ActiveConnection = cg
         acmd.CommandType = adCmdStoredProc
         acmd.CommandTimeout = 0
         acmd.CommandText = "vt_actualizofacart_pro"
         acmd.Prepared = True
         With acmd
             .Parameters("@base") = CStr(cbdatos.DefaultDatabase)
             .Parameters("@tipo") = "1"
             .Parameters("@codcia") = wCabe(19)
             .Parameters("@serie") = xserie
             .Parameters("@factura") = Trim(xfactu)
             .Parameters("@secuencia") = rsdeta.Fields(0)
             .Parameters("@fecha") = wCabe(14)
             .Parameters("@numeope") = 1
             .Parameters("@cliente") = Trim(wCabe(16))
             .Parameters("@articulo") = Trim(rsdeta.Fields(1))
             .Parameters("@DIAS") = wCabe(24)
             .Parameters("@PRECIO") = rsdeta.Fields(5)
             Set arbusca = cbdatos.Execute("select arm_stock from articulo where ARM_CODART='" & Trim(rsdeta.Fields(1)) & "' and ARM_CODCIA='" & wCabe(19) & "'")
             If arbusca.RecordCount > 0 Then
                 .Parameters("@stock") = (arbusca.Fields(0))        'stock articulo
             Else
                 .Parameters("@stock") = 0          'stock articulo
             End If
             arbusca.Close
             Set arbusca = Nothing
             .Parameters("@totaligv") = Previo
             .Parameters("@totalimporte") = TSub
             .Parameters("@totalneto") = rsdeta.Fields(7)
             .Parameters("@cantidad") = rsdeta.Fields(4)
             .Parameters("@serieped") = Left(MBox(1), 3)
             .Parameters("@pedido") = Right(wCabe(2), 5)
             .Parameters("@usuario") = "ADMIN"   'g_usuario
             .Parameters("@moneda") = wCabe(9)
             .Parameters("@serieguia") = 0
             .Parameters("@tipofac") = xtipofac
             .Parameters("@nopera2") = rsdeta.Fields(0)
        End With
        acmd.Execute
        Set acmd = Nothing
     
        Set acmd.ActiveConnection = cg
        acmd.CommandType = adCmdStoredProc
        acmd.CommandTimeout = 0
        acmd.CommandText = "vt_actualizoalmacen_pro"
        acmd.Prepared = True
        With acmd
            .Parameters("@basedes") = CStr(cbdatos.DefaultDatabase)
            .Parameters("@almacen") = wCabe(19)
            .Parameters("@tipo") = "1"
            .Parameters("@articulo") = rsdeta.Fields(1)
            .Parameters("@cantidad") = (-1 * rsdeta.Fields(4))
        End With
        acmd.Execute
        Set acmd = Nothing
    
        '************
        rsdeta.MoveNext
        regi = regi + 1
    Loop
    '*****Actualizamos el Valor de Inafecto**********
    cn.Execute "UPDATE vt_pedido " & _
               " Set Pedidototinafecto=" & tinafecto & _
               " Where pedidonumero='" & wCabe(2) & "'"
    
   '*Grabar en los cargos ***ctacte ***
    
    'If (cOpc2(0).Value Or cOpc2(1).Value) And modoventa.ctacte = "1" Then
        lzona = "00"
        Set asql = cn.Execute("select * from vt_zonavendedor where vendedorcodigo='" & wCabe(17) & "'")
        If asql.RecordCount > 0 Then
            lzona = Escadena(asql!zonacodigo)
        End If
        asql.Close
        Set asql = Nothing
           
        ltipo = "1"
        If dllgeneral.VerificaDatoExistente(cn, "select * from vt_cargo where documentocargo='" & dllgeneral.ComboDato(Combo2.Text) & "' and cargonumdoc='" & Trim(MBox(1) & MBox(2)) & "'") = 0 Then
          ltipo = "1"
        Else
          ltipo = "2"
        End If
        
        If dllgeneral.VerificaDatoExistente(cn, "select * from cp_tipodocumento where tdocumentocodigo='" & dllgeneral.ComboDato(Combo2.Text) & "' and tdocumentotipo='A'") = 1 Then
          tcargo = "A"
        Else
          tcargo = "C"
        End If
        
        If wCabe(9) = g_TipoSol Then
            If tcargo = "A" Then
                 cn.Execute "Update cp_proveedor " & _
                            " Set clientesaldosoles=ISNULL(clientesaldosoles,0)-" & CDbl(wCabe(29)) & _
                            "      Where clientecodigo='" & wCabe(16) & "'"
            Else
                 cn.Execute "Update cp_proveedor " & _
                            " Set clientesaldosoles=ISNULL(clientesaldosoles,0)+" & CDbl(wCabe(29)) & _
                            "      Where clientecodigo='" & wCabe(16) & "'"
            
            End If
         ElseIf wCabe(9) = g_TipoDolar Then
            If tcargo = "A" Then
                 cn.Execute "Update cp_proveedor " & _
                            " Set clientesaldodolares=ISNULL(clientesaldodolares,0)-" & CDbl(wCabe(29)) & _
                            "      Where clientecodigo='" & wCabe(16) & "'"
             Else
                 cn.Execute "Update cp_proveedor " & _
                            " Set clientesaldodolares=ISNULL(clientesaldodolares,0)+" & CDbl(wCabe(29)) & _
                            "      Where clientecodigo='" & wCabe(16) & "'"
             End If
         End If
        
        
        Set acmd.ActiveConnection = cg
        acmd.CommandType = adCmdStoredProc
        acmd.CommandTimeout = 0
        acmd.CommandText = "cp_ingresacargo_pro"
        acmd.Prepared = True
        With acmd
            .Parameters("@base") = cn.DefaultDatabase
            .Parameters("@tipo") = ltipo
            .Parameters("@tabla") = "vt_cargo"
            .Parameters("@tipodocu") = dllgeneral.ComboDato(Combo2.Text)
            .Parameters("@numero") = Trim(MBox(1) & MBox(2))
            .Parameters("@cliente") = Escadena(wCabe(16))
            .Parameters("@vendedor") = Escadena(wCabe(17))
            .Parameters("@zona") = lzona
            .Parameters("@apefecemi") = wCabe(14)
            .Parameters("@moneda") = Escadena(wCabe(9))
            .Parameters("@apeimppag") = wCabe(29)
            .Parameters("@usuario") = g_usuario
            .Parameters("@tipocambio") = wCabe(10)
            .Parameters("@fechaact") = Date
            .Parameters("@flagcancel") = "0"
            .Parameters("@cargoabono") = tcargo
            .Parameters("@concepto") = Escadena(Ctr_Ayuda3.xclave)
        End With
        acmd.Execute
        Set acmd = Nothing
        
        
       '--Actualizamos Cartera de Clientes
       Set acmd.ActiveConnection = cg
       acmd.CommandType = adCmdStoredProc
       acmd.CommandText = "vt_actualizocarteractacte_pro"
       acmd.CommandTimeout = 0
       With acmd
           .Parameters("@base") = CStr(cbdatos.DefaultDatabase)
           .Parameters("@tipo") = "1"
           .Parameters("@cliente") = Trim(wCabe(16))
           .Parameters("@codcia") = wCabe(19)
           .Parameters("@serie") = MBox(1).Text
           .Parameters("@numerofac") = MBox(2).Text
           .Parameters("@tipofac") = dllgeneral.ComboDato(Combo2.Text)
           .Parameters("@fecha") = wCabe(14)
           .Parameters("@fechavence") = wCabe(14)
           .Parameters("@rendi") = 1
           .Parameters("@importe") = (-1 * wCabe(29))
           .Parameters("@guia") = "0"
           .Parameters("@moneda") = wCabe(9)
           .Parameters("@opera") = 1
       End With
       acmd.Execute
       Set acmd = Nothing
       
        ' eliminamos los registros de vt_pedido y vt_detallepedido
'        Set asql = cn.Execute("select * from vt_detallepedido where pedidonumero='" & wCabe(2) & "'")
'        If asql.RecordCount > 0 Then
'           cn.Execute "Delete From vt_detallepedido where pedidonumero='" & wCabe(2) & "'"
'           cn.Execute "Delete From vt_pedido where pedidonumero='" & wCabe(2) & "'"
'        End If
'        asql.Close
'        Set asql = Nothing

                    
    'End If
    MsgBox "Se Grabo Satisfactoriamente el Documento." & Chr(13) & Chr(10) & dllgeneral.ComboDato(Combo2.Text) & " >= " & MBox(1) & MBox(2), vbInformation, MsgTitle
    GrabarData = 1
    
    
vererror:
   If Err Then
      MsgBox Err.Number & "-" & Err.Description
      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & cn.Errors(0).Number & "-" & cn.Errors(0).Description
      Exit Function
   End If
End Function


Public Function verificaproducto() As Integer
   On Error Resume Next
    verificaproducto = 0
    If rsdeta.RecordCount > 0 Then
       rsdeta.MoveFirst
       Do Until rsdeta.EOF
           If Escadena(rsdeta.Fields(1)) = MBox2(1) And flag = 0 Then
              verificaproducto = 1
              Exit Do
           End If
           rsdeta.MoveNext
       Loop
    End If
End Function




Public Sub TraerProducto()
  Dim rabusca As New ADODB.Recordset
  Dim nsql As String
  
  On Error Resume Next

    'If Combo2.ListCount > 0 Then
       nsql = "select * from " & "listapre" & Trim(alista) & " where productocodigo='" & MBox2(1) & "' and almacencodigo='" & aalmacen & "' "
   ' Else
   '    nsql = "select * from vt_producto where productocodigo='" & MBox2(1) & "'"
   ' End If
    Set rabusca = cn.Execute(nsql)
    If rabusca.RecordCount > 0 Then
      Label2 = Escadena(rabusca!productodescripcion)
      MBox2(2) = Escadena(rabusca!unidadcodigo)
      If rabusca!monedacodigo <> dllgeneral.ComboDato(Combo1.Text) Then
         If dllgeneral.ComboDato(Combo1.Text) = g_TipoSol Then
            If rabusca!productoprecvta > 0 Then
               MBox2(3) = Numero(rabusca!productoprecvta * CDbl(MBox(8)))
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = Numero(rabusca!unidadfactorconv)
                  MBox2(13) = Numero(rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = Numero(rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = Numero(rabusca!productoporcimpto)
            Else
               MBox2(3) = Numero(rabusca!productoprecvta)
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = Numero(rabusca!unidadfactorconv)
                  MBox2(13) = Numero(rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = Numero(rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = Numero(rabusca!productoporcimpto)
            End If
         ElseIf dllgeneral.ComboDato(Combo1.Text) = g_TipoDolar Then
            If rabusca!productoprecvta > 0 Then
               MBox2(3) = Numero(rabusca!productoprecvta / CDbl(MBox1(2)))
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = Numero(rabusca!unidadfactorconv)
                  MBox2(13) = Numero(rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = Numero(rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = Numero(rabusca!productoporcimpto)
            Else
               MBox2(3) = Numero(rabusca!productoprecvta)
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = Numero(rabusca!unidadfactorconv)
                  MBox2(13) = Numero(rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = Numero(rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = Numero(rabusca!productoporcimpto)
            End If
         End If
      Else
         MBox2(3) = Numero(rabusca!productoprecvta)
         If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
            MBox2(0) = Numero(rabusca!unidadfactorconv)
            MBox2(13) = Numero(rabusca!unidadfactorconv)
         ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
            MBox2(13) = Numero(rabusca!unidadfactorconv)
         Else
            MBox2(13) = 1
         End If
         MBox2(14) = Numero(rabusca!productoporcimpto)
      End If
    Else
      Label2 = "":    MBox2(2) = ""
    End If
    rabusca.Close
    Set rabusca = Nothing
End Sub




Public Function DocImprimir()
'   Dim rf As New ADODB.Command
'   Dim puntero As String
'   Dim cuenta As Double
'   Dim J As Integer
'   Dim busca As New dll_apisgen.dll_apis
'
'   On Error Resume Next
'
'   Set rf.ActiveConnection = cg
'   rf.CommandText = "vt_impresion"
'   rf.CommandType = adCmdStoredProc
'   rf.Prepared = True
'   rf.Parameters("@base") = CStr(cn.DefaultDatabase)
'   If cOpc2(0).Value Then
'     rf.Parameters("@tabla") = "vt_detallepedido"
'   ElseIf cOpc2(1).Value Then
'     rf.Parameters("@tabla") = "vt_detallepedido"
'   ElseIf cOpc2(2).Value Then
'     rf.Parameters("@tabla") = "vt_detallepedido"
'   Else
'     rf.Parameters("@tabla") = g_DetallePuntoVta
'   End If
'   rf.Parameters("@lista") = "listapre" & Trim(dllgeneral.ComboDato(Combo2.Text))
'   rf.Parameters("@almacen") = Ctr_Ayuda3.xclave
'   rf.Parameters("@numero") = CStr(MBox(1))
'   rf.Parameters("@items") = CStr(modoventa.nroitem)
'   rf.Execute
'
'   Set rf = Nothing
'
'   If cOpc2(0).Value Then
'     oCrystalReport.ReportFileName = RutaRep & "Repfactuimpresa.rpt"
'   ElseIf cOpc2(1).Value Then
'     oCrystalReport.ReportFileName = RutaRep & "Repboletaimpresa.rpt"
'   ElseIf cOpc2(2).Value Then
'     oCrystalReport.ReportFileName = RutaRep & "Repboimpresa.rpt"
'   Else
'     oCrystalReport.ReportFileName = RutaRep & "Reppedido.rpt"
'   End If
'   oCrystalReport.LogOnServer "pdssql.dll", _
'         busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dserver", ""), _
'         busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dbase", ""), _
'         busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "duser", ""), _
'         busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dpass", "")
'   oCrystalReport.Connect = _
'        "DSN=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dserver", "") & ";" & _
'        "DSQ=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dbase", "") & ";" & _
'        "UID=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "duser", "") & ";" & _
'        "PWD=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dpass", "")
'
'   oCrystalReport.Destination = crptToWindow
'   oCrystalReport.WindowState = crptMaximized
'   oCrystalReport.DiscardSavedData = True
'   With oCrystalReport
'       If cOpc2(0).Value Then
'          .Formulas(0) = "nro='" & MBox(2) & "'"
'       ElseIf cOpc2(1).Value Then
'          .Formulas(0) = "nro='" & MBox(3) & "'"
'       ElseIf cOpc2(2).Value Then
'          .Formulas(0) = "nro='" & MBox(4) & "'"
'       Else
'          .Formulas(0) = "nro='" & MBox(1) & "'"
'       End If
'       .Formulas(1) = "cliente='" & MBox3(1) & "'"
'       .Formulas(2) = "fecha='" & Str(Day(MBox(10))) & Space(3) & dllgeneral.DESMES(Month(MBox(10))) & Space(4) & Right(Str(Year(MBox(10))), 1) & "'"
'       .Formulas(3) = "direccion='" & MBox3(3) & "'"
'       .Formulas(4) = "dni='" & MBox3(2) & "'"
'       If cOpc2(0).Value Or cOpc2(1).Value Then
'         .Formulas(5) = "letras= '" & "SON : " & dllgeneral.NUMLET(Round(CDbl(MBox2(10)), 2)) & IIf(dllgeneral.ComboDato(Combo1.Text) = g_TipoSol, "Nuevos Soles", "Dolares Americanos") & "'"
'       Else
'         .Formulas(5) = "letras= '" & "SON : " & dllgeneral.NUMLET(Round(CDbl(MBox2(10)), 2)) & IIf(dllgeneral.ComboDato(Combo1.Text) = g_TipoSol, "Nuevos Soles", "Dolares Americanos") & "'"
'         .Formulas(6) = "forma='" & Combo4.Text & "'"
'         .Formulas(7) = "moneda='" & Combo1.Text & "'"
'       End If
'   End With
'   oCrystalReport.Action = 1
'
   
End Function
   
   

Public Sub CargarModo()
'     Dim rs As New ADODB.Recordset
'     On Error Resume Next
'     Set rs = cn.Execute("select * from vt_modoventa where modovtacodigo='" & dllgeneral.ComboDato(Combo3.Text) & "'")
'     If rs.RecordCount > 0 Then
'        modoventa.descuento = Escadena(rs!modovtadscto)
'        modoventa.impuestos = Escadena(IIf(IsNull(rs!modovtaimpuestos) Or rs!modovtaimpuestos = 0, "0", "1"))
'        modoventa.nroitem = IIf(IsNull(rs!modovtaitemxdoc), 10, rs!modovtaitemxdoc)
'        modoventa.copiashoja = IIf(IsNull(rs!modovtacopiashojatrab), 1, rs!modovtacopiashojatrab)
'        modoventa.copiasbol = IIf(IsNull(rs!modovtacopiasboleta), 1, rs!modovtacopiasboleta)
'        modoventa.copiasfac = IIf(IsNull(rs!modovtacopiasfact), 1, rs!modovtacopiasfact)
'        modoventa.ctacte = Escadena(IIf(IsNull(rs!modovtaactctacte) Or rs!modovtaactctacte = 0, "0", "1"))
'        modoventa.ctrlinventario = Escadena(IIf(IsNull(rs!modovtactrlinventario) Or rs!modovtactrlinventario = 0, "0", "1"))
'        modoventa.emitehoja = Escadena(IIf(IsNull(rs!modovtaemitehoja) Or rs!modovtaemitehoja = 0, "0", "1"))
'        modoventa.emitefact = Escadena(IIf(IsNull(rs!modovtasolemitfact) Or rs!modovtasolemitfact = 0, "0", "1"))
'        modoventa.emiteguia = Escadena(IIf(IsNull(rs!modovtaemiteguia) Or rs!modovtaemiteguia = 0, "0", "1"))
'        modoventa.ingcliente = Escadena(IIf(IsNull(rs!modovtaingcodclie) Or rs!modovtaingcodclie = 0, "0", "1"))
'        modoventa.ingforma = Escadena(IIf(IsNull(rs!modovtaingformapag) Or rs!modovtaingformapag = 0, "0", "1"))
'        modoventa.ingguia = Escadena(IIf(IsNull(rs!modovtaingguiarem) Or rs!modovtaingguiarem = 0, "0", "1"))
'        modoventa.inghoja = Escadena(IIf(IsNull(rs!modovtainghojatrab) Or rs!modovtainghojatrab = 0, "0", "1"))
'        modoventa.ingpedido = Escadena(IIf(IsNull(rs!modovtaingpedido) Or rs!modovtaingpedido = 0, "0", "1"))
'        modoventa.modificaguia = Escadena(IIf(IsNull(rs!modovtacorrguiarem) Or rs!modovtacorrguiarem = 0, "0", "1"))
'        modoventa.unidadmedida = Escadena(IIf(IsNull(rs!modovtaunidadmedida) Or rs!modovtaunidadmedida = "V", "V", Escadena(rs!modovtaunidadmedida)))
'        modoventa.unidadmedida = Left(modoventa.unidadmedida, 1)
'        modoventa.usafactor = Escadena(IIf(IsNull(rs!modovtausafactconv) Or rs!modovtausafactconv = 0, "0", "1"))
'
'        MBox(1).Enabled = IIf(modoventa.documento = g_tipoped And modoventa.numeraauto <> "1" And modoventa.ingpedido = "1", True, False) 'Modo de pedido
'        MBox(2).Enabled = IIf(modoventa.documento = g_tipofac And modoventa.numeraauto <> "1", True, False) 'Modo de factura
'        MBox(3).Enabled = IIf(modoventa.documento = g_tipobol And modoventa.numeraauto <> "1", True, False) 'Modo de boleta
'        MBox(4).Enabled = IIf(modoventa.documento = g_tipoguia And modoventa.numeraauto <> "1" And modoventa.ingguia = "1", True, False)  'Modo de Modifica
'
'        modoventa.numeraauto = Escadena(IIf(IsNull(rs!modovtanumautom) Or rs!modovtanumautom = 0, "0", "1"))
'        modoventa.documento = Escadena(IIf(IsNull(rs!documentocodigo), "", rs!documentocodigo))
'
'        MBox2(0).Enabled = IIf(modoventa.usafactor = 0 Or (modoventa.usafactor = "1" And modoventa.unidadmedida = "V"), True, False)
'        MBox2(12).Enabled = IIf(modoventa.usafactor = 0 Or (modoventa.usafactor = "1" And modoventa.unidadmedida = "R"), True, False)
'     End If
'     rs.Close
'     Set rs = Nothing

End Sub

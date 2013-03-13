VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form PrcEliminadocu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eliminar Documentos"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fr2 
      Height          =   645
      Index           =   0
      Left            =   60
      TabIndex        =   39
      Top             =   -30
      Width           =   5625
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   180
         Width           =   1425
      End
      Begin VB.CommandButton cBusca 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   4380
         TabIndex        =   36
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox aBusca 
         Height          =   285
         Index           =   0
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   34
         Top             =   210
         Width           =   435
      End
      Begin VB.TextBox aBusca 
         Height          =   285
         Index           =   1
         Left            =   3180
         MaxLength       =   8
         TabIndex        =   35
         Top             =   210
         Width           =   885
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Doc."
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
         Height          =   225
         Left            =   150
         TabIndex        =   41
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3060
         TabIndex        =   40
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1635
      Left            =   30
      TabIndex        =   1
      Top             =   570
      Width           =   11325
      Begin VB.Label Label1 
         Caption         =   "Fecha"
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
         Height          =   225
         Index           =   0
         Left            =   330
         TabIndex        =   19
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "No.Documento"
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
         Height          =   225
         Index           =   1
         Left            =   4320
         TabIndex        =   18
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
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
         Height          =   225
         Index           =   3
         Left            =   330
         TabIndex        =   16
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Vendedor"
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
         Height          =   225
         Index           =   4
         Left            =   330
         TabIndex        =   15
         Top             =   930
         Width           =   855
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   5
         Left            =   4530
         TabIndex        =   14
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Lista Precios"
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
         Height          =   225
         Index           =   6
         Left            =   8820
         TabIndex        =   13
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda"
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
         Height          =   225
         Index           =   7
         Left            =   360
         TabIndex        =   12
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   1
         Left            =   5700
         TabIndex        =   9
         Top             =   240
         Width           =   2475
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Index           =   2
         Left            =   9780
         TabIndex        =   8
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   7
         Top             =   570
         Width           =   9735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   6
         Top             =   900
         Width           =   2715
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   5
         Left            =   5640
         TabIndex        =   5
         Top             =   900
         Width           =   2805
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   6
         Left            =   10080
         TabIndex        =   4
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   7
         Left            =   1440
         TabIndex        =   3
         Top             =   1230
         Width           =   2745
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   8
         Left            =   5640
         TabIndex        =   2
         Top             =   1230
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "No. Pedido"
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
         Height          =   225
         Index           =   2
         Left            =   8760
         TabIndex        =   17
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Cambio"
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
         Height          =   225
         Index           =   8
         Left            =   4470
         TabIndex        =   11
         Top             =   1260
         Width           =   1245
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   7305
      Width           =   11400
      _ExtentX        =   20108
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
   Begin VB.Frame Frame1 
      Height          =   4125
      Left            =   30
      TabIndex        =   20
      Top             =   2190
      Width           =   11355
      Begin VB.Frame Fr2 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   885
         Index           =   2
         Left            =   90
         TabIndex        =   21
         Top             =   3120
         Width           =   11205
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   6
            Left            =   150
            TabIndex        =   22
            Top             =   240
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
            Left            =   2250
            TabIndex        =   23
            Top             =   240
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
            Left            =   4650
            TabIndex        =   24
            Top             =   240
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
            Left            =   7020
            TabIndex        =   25
            Top             =   240
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
            Left            =   9270
            TabIndex        =   26
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
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
            Left            =   210
            TabIndex        =   31
            Top             =   675
            Width           =   1335
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
            Left            =   2610
            TabIndex        =   30
            Top             =   675
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
            Left            =   5010
            TabIndex        =   29
            Top             =   675
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
            Left            =   7530
            TabIndex        =   28
            Top             =   675
            Width           =   1215
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
            Left            =   9690
            TabIndex        =   27
            Top             =   675
            Width           =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   0
            X1              =   1995
            X2              =   1995
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
            Index           =   2
            X1              =   6780
            X2              =   6780
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   3
            X1              =   9090
            X2              =   9090
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   1980
            X2              =   1980
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   5
            X1              =   4425
            X2              =   4425
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   6
            X1              =   6765
            X2              =   6765
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   7
            X1              =   9105
            X2              =   9105
            Y1              =   120
            Y2              =   1215
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   2865
         Left            =   90
         TabIndex        =   32
         Top             =   210
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   5054
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
   End
   Begin VB.Frame Frame4 
      Height          =   930
      Left            =   5430
      TabIndex        =   42
      Top             =   6300
      Width           =   2010
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Cancelar"
         Height          =   690
         Index           =   12
         Left            =   1140
         Picture         =   "PrcEliminadocu.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   180
         Width           =   825
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Acepta"
         Height          =   690
         Index           =   11
         Left            =   90
         Picture         =   "PrcEliminadocu.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   180
         Width           =   870
      End
   End
End
Attribute VB_Name = "PrcEliminadocu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsdeta As New ADODB.Recordset
Dim xcliente As String
Dim xmoneda As String * 2
Dim xDocu As String
Dim xmonto As Double
Dim adll As New dllgeneral.dll_general




Private Sub aBusca_Change(Index As Integer)
  If Len(Trim(aBusca(Index))) = 0 Then
     If Index = 0 Then
        aBusca(1) = ""
     End If
     Call Limpiartexto(MBox2, 6, 10)
     Call Limpiartexto(Label2, 0, 8)
     Call CargaGrilla
  End If
  
End Sub

Private Sub aBusca_GotFocus(Index As Integer)
     Call Limpiartexto(MBox2, 6, 10)
     Call Limpiartexto(Label2, 0, 8)
     Call CargaGrilla
  
End Sub

Private Sub aBusca_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim nsql As String
  
  If KeyCode = 112 Then  ' Ayuda de Productos
       If adll.ComboDato(Combo2.Text) = g_tipobol Then
            nsql = "CASE pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofecha as Fecha,pedidonrofact as Boleta,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
       ElseIf adll.ComboDato(Combo2.Text) = g_tipofac Then
            nsql = "CASE pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofecha as Fecha,pedidonrofact as Factura,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
       ElseIf adll.ComboDato(Combo2.Text) = g_tipoped Then
            nsql = "CASE pedidoestado WHEN '1' THEN '*' ELSE '' END,pedidofecha as Fecha,pedidonumero as Pedido,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
       Else
            nsql = "CASE pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofecha as Fecha,pedidonrofact as Documento,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
       End If
       Dim sfiltra(1 To 2, 1 To 2) As String
       sfiltra(1, 1) = "Cliente": sfiltra(1, 2) = "clienterazonsocial"
       sfiltra(2, 1) = "Ruc": sfiltra(2, 2) = "clienteruc"
       FrmAyudaPedidos.TipoForma = 2
       FrmAyudaPedidos.Bdata = "0"
       FrmAyudaPedidos.BConexion = VGCNx
       FrmAyudaPedidos.BTabla = "vt_pedido"
       FrmAyudaPedidos.BCampos = nsql
       If adll.ComboDato(Combo2.Text) = g_tipobol Then
            FrmAyudaPedidos.BCondi = "pedidotipofac='" & g_tipobol & "'"
            FrmAyudaPedidos.BOrden = "pedidonrofact"
       ElseIf adll.ComboDato(Combo2.Text) = g_tipofac Then
            FrmAyudaPedidos.BCondi = "pedidotipofac='" & g_tipofac & "'"
            FrmAyudaPedidos.BOrden = "pedidonrofact"
       ElseIf adll.ComboDato(Combo2.Text) = g_tipoped Then
            FrmAyudaPedidos.BCondi = ""
            FrmAyudaPedidos.BOrden = "pedidonumero"
       Else
            FrmAyudaPedidos.BCondi = "pedidotipofac='" & adll.ComboDato(Combo2.Text) & "'"
            FrmAyudaPedidos.BOrden = "pedidonrofact"
       End If
       FrmAyudaPedidos.BFiltro = sfiltra
       FrmAyudaPedidos.Show 1
       aBusca(0) = Left(nAyuda, aBusca(0).MaxLength)
       aBusca(1) = Right(nAyuda, aBusca(1).MaxLength)
       nAyuda = "": nDetalle = ""
   ElseIf KeyCode = 13 Then
       SendKeys "{tab}"
   End If
End Sub

Private Sub aBusca_LostFocus(Index As Integer)
    If Index = 0 Then
       aBusca(Index) = Right("000000000000" & aBusca(Index), aBusca(Index).MaxLength)
    ElseIf Index = 1 Then
       aBusca(Index) = Right("0000000000000" & aBusca(Index), aBusca(Index).MaxLength)
    End If
    
End Sub

Private Sub cBusca_Click()
    Dim csql As New ADODB.Recordset
    Dim acliente As New ADODB.Recordset
    Dim nvalor As String
    Dim nsql As String
    Dim J As Integer
    
    Call Limpiartexto(MBox2, 6, 10)
    Call Limpiartexto(Label2, 0, 8)
    Call CargaGrilla
    
    If adll.ComboDato(Combo2.Text) = g_tipobol Then
       nsql = "select * from vt_pedido where pedidonrofact='" & Trim(aBusca(0) & aBusca(1)) & "' and pedidotipofac='" & g_tipobol & "'"
    ElseIf adll.ComboDato(Combo2.Text) = g_tipofac Then
        nsql = "select * from vt_pedido where pedidonrofact='" & Trim(aBusca(0) & aBusca(1)) & "' and pedidotipofac='" & g_tipofac & "'"
    ElseIf adll.ComboDato(Combo2.Text) = g_tipoped Then
        nsql = "select * from vt_pedido where pedidonumero='" & Trim(aBusca(0) & aBusca(1)) & "'"
    Else
        nsql = "select * from vt_pedido where pedidonrofact='" & Trim(aBusca(0) & aBusca(1)) & "' and pedidotipofac='" & adll.ComboDato(Combo2.Text) & "'"
    End If
    nvalor = ""
    Set csql = VGCNx.Execute(nsql)
    If csql.RecordCount > 0 Then
        nvalor = Escadena(csql!pedidonumero)
        If adll.ComboDato(Combo2.Text) = g_tipobol Then
            Label2(0) = Format(csql!pedidofechafact, "dd/mm/yyyy")
            Label2(1) = g_tipobol & "-" & Escadena(csql!pedidonrofact)
        ElseIf adll.ComboDato(Combo2.Text) = g_tipofac Then
            Label2(0) = Format(csql!pedidofechafact, "dd/mm/yyyy")
            Label2(1) = g_tipofac & "-" & Escadena(csql!pedidonrofact)
        ElseIf adll.ComboDato(Combo2.Text) = g_tipoped Then
            Label2(0) = Format(csql!pedidofecha, "dd/mm/yyyy")
            Label2(1) = g_tipoped & "-" & Escadena(csql!pedidonumero)
        Else
            Label2(0) = Format(csql!pedidofechafact, "dd/mm/yyyy")
            Label2(1) = adll.ComboDato(Combo2.Text) & "-" & Escadena(csql!pedidonrofact)
        End If
        Label2(2) = Escadena(csql!pedidonumero)
        Label2(3) = Escadena(csql!clientecodigo) & "-" & Escadena(csql!clienterazonsocial)
        Label2(4) = Escadena(csql!vendedorcodigo)
        xcliente = csql!clientecodigo: xmonto = csql!pedidototneto: xDocu = csql!pedidonumero
        xmoneda = csql!pedidomoneda
        Set acliente = VGCNx.Execute("select * from vt_vendedor where vendedorcodigo='" & Escadena(csql!vendedorcodigo) & "'")
        If acliente.RecordCount > 0 Then
           Label2(4) = Label2(4) & "-" & Escadena(acliente!vendedornombres)
        Else
            Label2(4) = Label2(4)
        End If
        acliente.Close
        Set acliente = Nothing
        Label2(5) = Escadena(csql!almacencodigo)
        Set acliente = VGCNx.Execute("select * from vt_almacen where almacencodigo='" & Escadena(csql!almacencodigo) & "'")
        If acliente.RecordCount > 0 Then
           Label2(5) = Label2(5) & "-" & Escadena(acliente!almacendescripcion)
        Else
            Label2(5) = Label2(5)
        End If
        acliente.Close
        Set acliente = Nothing
        
        Label2(6) = Escadena(csql!pedidolistaprec)
        Label2(7) = Escadena(csql!pedidomoneda)
        Set acliente = VGCNx.Execute("select * from gr_moneda where monedacodigo='" & Escadena(csql!pedidomoneda) & "'")
        If acliente.RecordCount > 0 Then
           Label2(7) = Label2(7) & "-" & Escadena(acliente!monedadescripcion)
        Else
           Label2(7) = Label2(7)
        End If
        acliente.Close
        Set acliente = Nothing
        Label2(8) = numero(csql!pedidotipcambio)
        MBox2(6) = numero(csql!pedidototitem)
        MBox2(7) = Format(csql!pedidototbruto, "##,###,##0.0000")
        MBox2(8) = numero(csql!pedidomontodsctoglobal + csql!pedidomontodsctocliente + csql!pedidomontodsctoppago + csql!pedidomontodsctovtaoficina + csql!pedidototaldsctoxitem + csql!pedidototaldsctoxlinea + csql!pedidototaldsctoxprom)
        MBox2(9) = numero(csql!pedidototimpuesto)
        MBox2(10) = numero(csql!pedidototneto)
        
    Else
        MsgBox "No existe Informacion del Documento...Verifique!!", vbInformation, MsgTitle
        csql.Close
        Set csql = Nothing
        Exit Sub
    End If
    csql.Close
       
    Set csql = VGCNx.Execute("select detpeditem,A.productocodigo,b.adescri,a.unidadcodigo," & _
                          "detpedcantpedida,detpedmontoprecvta,detpeddsctoxitem,detpedimpbruto," & _
                          " detpedporccomis " & _
                          "from vt_detallepedido A " & _
                          "inner Join " & _
                          "[" & VGCNx.DefaultDatabase & "].dbo.maeart B" & _
                          " ON A.productocodigo=b.acodigo COLLATE Modern_Spanish_CI_AI " & _
                          "where pedidonumero='" & nvalor & "'")
    
    Set rsdeta = Nothing
    Call CargaGrilla

    Do Until csql.EOF
       rsdeta.AddNew
       rsdeta.Fields(0) = Escadena(csql!detpeditem)
       rsdeta.Fields(1) = Escadena(csql!productocodigo)
       rsdeta.Fields(2) = Escadena(csql!adescri)
       rsdeta.Fields(3) = Escadena(csql!unidadcodigo)
       rsdeta.Fields(4) = numero(csql!detpedcantpedida)
       rsdeta.Fields(5) = numero(IIf(IsNull(csql!detpedmontoprecvta), 0, csql!detpedmontoprecvta))
       rsdeta.Fields(6) = numero(csql!detpeddsctoxitem)
       rsdeta.Fields(7) = numero(csql!detpedimpbruto)
       rsdeta.Fields(8) = numero(csql!detpedporccomis)
       rsdeta.Update
       csql.MoveNext
    Loop
    csql.Close
    
    Call ConfigGrid
    Set csql = Nothing
End Sub


Public Function CargaGrilla()

   Set rsdeta = Nothing
   
   Call rsdeta.Fields.Append("Item", adInteger)
   Call rsdeta.Fields.Append("Codigo", adChar, 20)
   Call rsdeta.Fields.Append("Descripcion", adChar, 100)
   Call rsdeta.Fields.Append("UM", adChar, 3)
   Call rsdeta.Fields.Append("Cant", adDouble)
   Call rsdeta.Fields.Append("Precio_Vta", adDouble)
   Call rsdeta.Fields.Append("Dscto(%)", adDouble)
   Call rsdeta.Fields.Append("Total", adDouble)
   Call rsdeta.Fields.Append("%", adDouble)
   
   rsdeta.Open
   ConfigGrid

End Function

Public Function ConfigGrid()

   Set TDBGrid1.DataSource = rsdeta
   With TDBGrid1
      .Columns(0).Width = 600
      .Columns(0).Caption = "Item"
      .Columns(1).Width = 1100
      .Columns(1).Caption = "Codigo"
      .Columns(2).Width = 3500
      .Columns(2).Caption = "Descripcion"
      .Columns(3).Width = 600
      .Columns(3).Caption = "UM"
      .Columns(4).Width = 1000
      .Columns(4).Caption = "Cant"
      .Columns(5).Width = 1000
      .Columns(5).Caption = "Precio_Vta"
      .Columns(6).Width = 1000
      .Columns(6).Caption = "Dscto(%)"
      .Columns(7).Width = 1000
      .Columns(7).Caption = "Total"
      .Columns(8).Width = 1000
      .Columns(8).Caption = "%"
      .Columns(5).NumberFormat = "###,##0.0000"
      .Columns(6).NumberFormat = "###,##0.00"
      .Columns(7).NumberFormat = "###,##0.0000"
      .Columns(8).NumberFormat = "###,##0.00"
   End With
   TDBGrid1.Refresh
End Function


Private Sub cmdBotones_Click(Index As Integer)
  Dim ntipo, nnume As String
  Dim rs As New ADODB.Recordset
  Dim acmd As New ADODB.Command
  On Error GoTo nerror
  
  Select Case Index
    Case 11
    
      If adll.VerificaDatoExistente(VGCNx, "select * from vt_abono where documentoabono='" & adll.ComboDato(Label2(1)) & "' and abononumdoc='" & Mid(Label2(1), Len(g_tipofac) + 2, Len(Trim(Label2(1)))) & "'") = 1 Then
        MsgBox "No se puede eliminar el documento existen abonos...verifique!!!", vbInformation, MsgTitle
        Exit Sub
      End If
    
      If MsgBox("Desea Eliminar el Documento?", vbYesNo, MsgTitle) = vbNo Then
         Exit Sub
      End If
      VGCNx.BeginTrans
      If adll.ComboDato(Label2(1)) = g_tipofac Then
          ntipo = g_tipofac
          nnume = Mid(Label2(1), Len(g_tipofac) + 2, Len(Trim(Label2(1))))
      ElseIf adll.ComboDato(Label2(1)) = g_tipobol Then
          ntipo = g_tipobol
          nnume = Mid(Label2(1), Len(g_tipobol) + 2, Len(Trim(Label2(1))))
      ElseIf adll.ComboDato(Label2(1)) = g_tipoguia Then
          ntipo = g_tipoguia
          nnume = Mid(Label2(1), Len(g_tipoguia) + 2, Len(Trim(Label2(1))))
      End If
      If Len(Trim(nnume)) > 0 Then
         Set acmd.ActiveConnection = VGgeneral
         acmd.CommandType = adCmdStoredProc
         acmd.CommandText = "vt_anulafactura_pro"
         acmd.CommandTimeout = 0
         With acmd
             .Parameters("@base") = VGCNx.DefaultDatabase
             .Parameters("@docu") = xDocu
             .Parameters("@tipo") = ntipo
             .Parameters("@nume") = nnume
             .Parameters("@fanula") = Date
             .Parameters("@empresa") = VGParametros.empresacodigo
             .Parameters("@puntovta") = VGParametros.puntovta
             
         End With
         acmd.Execute
         Set acmd = Nothing
        If adll.ComboDato(Label2(1)) = g_tipofac Then
            VGCNx.Execute "delete from vt_cargo where documentocargo='" & g_tipofac & "' and cargonumdoc='" & nnume & "'"
         ElseIf adll.ComboDato(Label2(1)) = g_tipobol Then
            VGCNx.Execute "delete from vt_cargo where documentocargo='" & g_tipobol & "' and cargonumdoc='" & nnume & "'"
         ElseIf adll.ComboDato(Label2(1)) = g_tipoguia Then
            VGCNx.Execute "delete from vt_cargo where documentocargo='" & g_tipoguia & "' and cargonumdoc='" & nnume & "'"
         Else
            VGCNx.Execute "delete from vt_cargo where documentocargo='" & adll.ComboDato(Label2(1)) & "' and cargonumdoc='" & nnume & "'"
         End If
         
         '--Actualizamos en clientes
         
         If xmoneda = g_tiposol Then
              VGCNx.Execute "Update vt_cliente " & _
                         " Set clientesaldosoles=clientesaldosoles-" & xmonto & _
                         "      Where clientecodigo='" & xcliente & "'"
         ElseIf xmoneda = g_tipodolar Then
              VGCNx.Execute "Update vt_cliente " & _
                         " Set clientesaldodolares=clientesaldodolares-" & xmonto & _
                         "      Where clientecodigo='" & xcliente & "'"
         End If
         
         '--Actualizamos en Cartera de Clientes
         Set rs = VGCNx.Execute("select * from vt_pedido where pedidonumero='" & Label2(2) & "'")
              
         '******Actualizamos Saldos en Almacen *********
          If rsdeta.RecordCount > 0 Then
'            If adll.VerificaDatoExistente(VGcnx, "select * from vt_modoventa where  modovtactrlinventario='1' and modovtacodigo='" & rs!modovtacodigo & "'") = 1 Then
               VGCNx.Execute "delete from movalmdet " & _
                          " where detd='GS' and denumdoc='" & Trim(xDocu) & "' and dealma='" & rs!almacencodigo & "'"
               
                VGCNx.Execute "delete from movalmcab " & _
                           " where catd='GS' and canumdoc='" & Trim(xDocu) & "' and caalma='" & rs!almacencodigo & "'"
    
                 rsdeta.MoveFirst
                 Do Until rsdeta.EOF
                    Set acmd.ActiveConnection = VGgeneral
                    acmd.CommandType = adCmdStoredProc
                    acmd.CommandTimeout = 0
                    acmd.CommandText = "vt_actualizoalma_pro"
                    acmd.Prepared = True
                    With acmd
                          .Parameters("@basedes") = CStr(VGCNx.DefaultDatabase)
                          .Parameters("@almacen") = adll.ComboDato(Label2(5))
                          .Parameters("@tipo") = "2"
                          .Parameters("@articulo") = Trim(rsdeta.Fields(1))
                          .Parameters("@cantidad") = rsdeta.Fields(4)
                    End With
                    acmd.Execute
                   
                    Set acmd = Nothing
                    rsdeta.MoveNext
                  Loop
'               End If
           End If
           rs.Close
           Set rs = Nothing
              
          '*** Eliminar el documento *********
          VGCNx.Execute "Delete From vt_detallepedido where pedidonumero='" & xDocu & "'"
          VGCNx.Execute "Delete From vt_pedido where pedidonumero='" & xDocu & "'"
                        
          VGCNx.CommitTrans
          '**********
          
          Open "c:\asystem2.dat" For Append As #99
          Print #99, "Usuario: " & g_usuario; Tab(30); "Fecha Hora : " & Now; Tab(45); "Doc. : " & ntipo & "-" & nnume
          Close #99
          
          MsgBox "El documento ha sido eliminado satisfactoriamente...!!!", vbInformation, MsgTitle
       End If
    Case 12
       Unload Me
  End Select
  
nerror:
    If Err Then
        MsgBox "Error : No se pudo completar la Transaccion..." & Err.Number & "-" & Err.Description, vbInformation, MsgTitle
        Err = 0
        VGCNx.RollbackTrans
        Exit Sub
    End If
    
  
End Sub

Private Sub Combo2_Click()
     aBusca(0) = ""
     aBusca(1) = ""
     Call Limpiartexto(MBox2, 6, 10)
     Call Limpiartexto(Label2, 0, 8)
     Call CargaGrilla
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   Seguir Combo2, KeyAscii
End Sub

Private Sub Form_Load()
    MostrarFormVentas Me, "C"
    
    Call Limpiartexto(MBox2, 6, 10)
    Call CargarTipoVentas(Combo2, 6)    'Todos los documentos
    Call CargaGrilla
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set rsdeta = Nothing
End Sub


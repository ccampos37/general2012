VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmExportarDataCoa 
   Caption         =   "Exportación Datos al COA"
   ClientHeight    =   6672
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5532
   LinkTopic       =   "Form1"
   ScaleHeight     =   6672
   ScaleWidth      =   5532
   Begin VB.Frame Frame1 
      Caption         =   "Criterios de la Exportación"
      Height          =   5445
      Left            =   150
      TabIndex        =   3
      Top             =   105
      Width           =   5190
      Begin VB.CheckBox chk 
         Caption         =   "Notas de Crédito y Débito"
         Height          =   240
         Index           =   2
         Left            =   165
         TabIndex        =   15
         Top             =   795
         Width           =   2190
      End
      Begin VB.CheckBox chk 
         Caption         =   "Comprobantes de Pago"
         Height          =   240
         Index           =   1
         Left            =   165
         TabIndex        =   14
         Top             =   510
         Width           =   2190
      End
      Begin VB.CheckBox chk 
         Caption         =   "Exportar Proveedores"
         Height          =   240
         Index           =   0
         Left            =   165
         TabIndex        =   13
         Top             =   240
         Width           =   2190
      End
      Begin VB.TextBox txDir 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1830
         TabIndex        =   6
         Top             =   1305
         Width           =   2865
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   4725
         TabIndex        =   5
         Top             =   1290
         Width           =   330
      End
      Begin VB.CommandButton CmdProceso 
         Caption         =   "Exportar informacion"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1830
         TabIndex        =   4
         Top             =   1650
         Width           =   3225
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   135
         Top             =   4650
         _ExtentX        =   995
         _ExtentY        =   995
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExportarDataCoa.frx":0000
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView LvwTablas 
         Height          =   3150
         Left            =   105
         TabIndex        =   7
         Top             =   2175
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   5546
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre Tabla"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nº Reg"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Reg Procesados"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   11805
         Top             =   1050
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   315
         Left            =   3795
         TabIndex        =   8
         Top             =   690
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   550
         _Version        =   393216
         Format          =   19660801
         CurrentDate     =   37655
      End
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   315
         Left            =   3810
         TabIndex        =   9
         Top             =   285
         Width           =   1290
         _ExtentX        =   2265
         _ExtentY        =   550
         _Version        =   393216
         Format          =   19660801
         CurrentDate     =   37655
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha inicio :"
         Height          =   300
         Left            =   2820
         TabIndex        =   12
         Top             =   315
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Fin  :"
         Height          =   300
         Left            =   2835
         TabIndex        =   11
         Top             =   735
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Ruta de Archivo a Generar para el Directorio COA"
         Height          =   615
         Left            =   105
         TabIndex        =   10
         Top             =   1275
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   1995
      TabIndex        =   0
      Top             =   6120
      Width           =   1620
   End
   Begin Crystal.CrystalReport CryRepo1 
      Left            =   4995
      Top             =   6045
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar PGBreg 
      Height          =   180
      Left            =   0
      TabIndex        =   1
      Top             =   6465
      Visible         =   0   'False
      Width           =   5430
      _ExtentX        =   9589
      _ExtentY        =   318
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lbpg 
      Caption         =   "Procesando Registros"
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   15
      TabIndex        =   2
      Top             =   6240
      Visible         =   0   'False
      Width           =   1950
   End
End
Attribute VB_Name = "frmExportarDataCoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


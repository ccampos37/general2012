VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frWiese 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos Teleméticos - Banco Wiese Sudameris"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "frWiese.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   4650
      TabIndex        =   17
      Top             =   5370
      Width           =   1515
   End
   Begin VB.CommandButton cmGenerar 
      Caption         =   "&Generar Archivo"
      Height          =   375
      Left            =   2940
      TabIndex        =   16
      Top             =   5370
      Width           =   1515
   End
   Begin MSDataGridLib.DataGrid dgLista 
      Height          =   2880
      Left            =   90
      TabIndex        =   12
      Top             =   2085
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   5080
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
      Caption         =   "Configuración"
      Height          =   1950
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   6105
      Begin AplisetControlText.Aplitext xOrden 
         Height          =   300
         Left            =   4350
         TabIndex        =   19
         Top             =   1110
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Text            =   ""
      End
      Begin VB.CommandButton cmBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   5205
         TabIndex        =   11
         Top             =   1440
         Width           =   795
      End
      Begin AplisetControlText.Aplitext xRuta 
         Height          =   285
         Left            =   1890
         TabIndex        =   10
         Top             =   1455
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   503
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xConvenio 
         Height          =   300
         Left            =   4350
         TabIndex        =   8
         Top             =   735
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xCodigo 
         Height          =   300
         Left            =   1890
         TabIndex        =   6
         Top             =   1110
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xOficina 
         Height          =   300
         Left            =   1890
         TabIndex        =   4
         Top             =   735
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xCtaCte 
         Height          =   285
         Left            =   1890
         TabIndex        =   2
         Top             =   390
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   503
         Text            =   ""
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Número de Orden"
         Height          =   195
         Left            =   2880
         TabIndex        =   18
         Top             =   1163
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ruta de Archivo"
         Height          =   195
         Left            =   255
         TabIndex        =   9
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro. de Convenio"
         Height          =   195
         Left            =   2880
         TabIndex        =   7
         Top             =   810
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "Oficina de la Cta. Cte."
         Height          =   240
         Left            =   255
         TabIndex        =   5
         Top             =   795
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Registro Bco. Wiese"
         Height          =   195
         Left            =   255
         TabIndex        =   3
         Top             =   1155
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número de Cta.Cte."
         Height          =   195
         Left            =   255
         TabIndex        =   1
         Top             =   450
         Width           =   1395
      End
   End
   Begin VB.Label xTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   270
      Left            =   4635
      TabIndex        =   15
      Top             =   5010
      Width           =   1515
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Abonos"
      Height          =   270
      Left            =   2790
      TabIndex        =   14
      Top             =   5010
      Width           =   1830
   End
   Begin VB.Label xNumTrab 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 0 Trabajadores"
      Height          =   270
      Left            =   105
      TabIndex        =   13
      Top             =   5010
      Width           =   2670
   End
End
Attribute VB_Name = "frWiese"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMBUSCAR_Click()
    frSelDir.Show 1
    xRuta.Text = VPTAREA
End Sub

Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    MsgBox "No encontro el archivo Wiese.Exe. No se tiene autorizado la versión 6.0 del Sistema Telewiese Empresarial", vbCritical
End Sub


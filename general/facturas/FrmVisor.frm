VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmVisor 
   Caption         =   "Reporte Ventas por Factura"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTHasta 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   64487425
      CurrentDate     =   37489
   End
   Begin MSComCtl2.DTPicker DTDesde 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   64487425
      CurrentDate     =   37489
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   240
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
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
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "Fecha Hasta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lbl 
      Caption         =   "Fecha Desde"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "Cód.Almacen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "FrmVisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
'Dim cmd As New Command
   
   ' EJECUTA STORE PROCEDURE
                
'                Set cmd.ActiveConnection = cn
'                cmd.CommandType = adCmdStoredProc
'                cmd.CommandText = "vt_Edg_RepVtasxFact"
'                With cmd
'                    .Parameters("@codalmacen") = IIf(Trim(txt(0)) = "", Null, Trim(txt(0)))
'                    .Parameters("@fecdesde") = Trim(txt(1))
'                    .Parameters("@fechasta") = Trim(txt(2))
'                End With
'                cmd.Execute
'                oCrystalReport.Connect = "DSN=DESARROLLO;DSQ=Ventas_Prueba;UID=pirata"

                If DTDesde > DTHasta Then
                    MsgBox "Feche Desde debe ser mayor a Fecha Hasta"
                    Exit Sub
                End If
                 
                oCrystalReport.ReportFileName = RutaRep & "Procesos\RepVtasxFactura.rpt"
                oCrystalReport.DiscardSavedData = True
                oCrystalReport.StoredProcParam(0) = IIf(Trim(txt(0)) = "", "%", Trim(txt(0)))
                oCrystalReport.StoredProcParam(1) = DTDesde
                oCrystalReport.StoredProcParam(2) = DTHasta
                oCrystalReport.Action = 1
                
End Sub

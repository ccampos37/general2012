VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmFv 
   Caption         =   "Formato de Venta"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1920
   ScaleWidth      =   3645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCan 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1920
      TabIndex        =   1
      Top             =   1140
      Width           =   1245
   End
   Begin VB.CommandButton CmdImp 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   540
      TabIndex        =   0
      Top             =   1140
      Width           =   1245
   End
   Begin MSComCtl2.DTPicker DtHasta 
      Height          =   285
      Left            =   1500
      TabIndex        =   2
      Top             =   540
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   503
      _Version        =   393216
      Format          =   49086465
      CurrentDate     =   39689
   End
   Begin MSComCtl2.DTPicker DtDesde 
      Height          =   285
      Left            =   1500
      TabIndex        =   3
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   503
      _Version        =   393216
      Format          =   49086465
      CurrentDate     =   39689
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   90
      Top             =   900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicial :"
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
      Left            =   150
      TabIndex        =   5
      Top             =   240
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Final :"
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
      Left            =   150
      TabIndex        =   4
      Top             =   570
      Width           =   1125
   End
End
Attribute VB_Name = "FrmFv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCan_Click()
Unload Me
End Sub

Private Sub CmdImp_Click()
On Error GoTo Errores

If DtDesde > DtHasta Then
   MsgBox "Fecha Desde debe ser mayor a Fecha Hasta", vbInformation, "AVISO"
   Exit Sub
End If
                                   
Screen.MousePointer = 11
                                   
With oCrystalReport
     .Reset
     If Me.Caption = "Formato de Venta" Then
        .ReportFileName = VGParamSistem.Rutareport & "vt_formato_venta.rpt"
     Else
        .ReportFileName = VGParamSistem.Rutareport & "vt_tiposdecontacto.rpt"
     End If
     
     If VGsql = 1 Then
         .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
        Else
         .Connect = VGcadenareport2
      End If

      .DiscardSavedData = True
      .Destination = crptToWindow
      .WindowState = crptMaximized
      .WindowShowPrintSetupBtn = True
      .WindowShowExportBtn = True
      .WindowShowZoomCtl = True
      .WindowShowNavigationCtls = True
      .WindowShowPrintBtn = True
      .WindowTitle = "Formato Pedido Produccion"
      .StoredProcParam(0) = VGParamSistem.BDEmpresa
      .StoredProcParam(1) = DtDesde
      .StoredProcParam(2) = DtHasta
      .Action = 1
      
End With
  
Screen.MousePointer = 1

Exit Sub

Errores:
Screen.MousePointer = 1
MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
Err = 0

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
DtDesde.Value = Date
DtHasta.Value = Date
End Sub

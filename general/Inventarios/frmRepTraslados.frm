VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RepTraslados 
   Caption         =   "Informe de Traslados"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3585
   LinkTopic       =   "Form2"
   ScaleHeight     =   2340
   ScaleWidth      =   3585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   645
      Left            =   1965
      Picture         =   "frmRepTraslados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1530
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   645
      Left            =   690
      Picture         =   "frmRepTraslados.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1530
      Width           =   795
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   3135
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1290
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   52822017
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   1290
         TabIndex        =   2
         Top             =   660
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   52822017
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Desde"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   270
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   1470
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "RepTraslados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim adll As New dllgeneral.dll_general
    
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Inv036 -- Documentos Emitidos Detallados"
    CrystalReport1.ReportFileName = cRutP & "inv036.rpt"
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
       
   CrystalReport1.LogOnServer "pdssql.dll", _
                VGServer, _
                VGBase3, _
                VGBUsuario, _
                VGPassw
                        
    CrystalReport1.Connect = "DSN=" & VGServer & ";DSQ=" & VGBase3 & ";UID=" & VGUsuario & ";PWD=" & VGPassw
       
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.StoredProcParam(0) = Trim(cConexCom.DefaultDatabase)
    CrystalReport1.StoredProcParam(1) = Format(DTPicker1.Value, "dd/mm/yyyy")
    CrystalReport1.StoredProcParam(2) = Format(DTPicker2.Value, "dd/mm/yyyy")
    
    CrystalReport1.formulas(0) = "fechainicio ='" & DTPicker1 & "'"
    CrystalReport1.formulas(1) = "fechafin ='" & DTPicker2 & "'"
    CrystalReport1.formulas(2) = "emp ='" & VGNemp & "'"
    
    If CrystalReport1.Status <> 2 Then
       
       CrystalReport1.Action = 1
    End If
    Screen.MousePointer = 1
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  DTPicker1.Value = Date
  DTPicker2.Value = Date
  
End Sub



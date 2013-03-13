VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRepClie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Clientes"
   ClientHeight    =   3030
   ClientLeft      =   3585
   ClientTop       =   2445
   ClientWidth     =   5100
   Icon            =   "frmrepclie.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5100
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   360
      Picture         =   "frmrepclie.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2052
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   1224
      Picture         =   "frmrepclie.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2052
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   2820
      Left            =   108
      TabIndex        =   0
      Top             =   72
      Width           =   4728
      Begin VB.OptionButton Option3 
         Caption         =   "Cátalogo por R.U.C."
         Height          =   375
         Left            =   252
         TabIndex        =   5
         Top             =   1152
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cátalogo por Nombre"
         Height          =   375
         Left            =   264
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cátalogo por Código"
         Height          =   255
         Left            =   264
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   2364
         Left            =   2772
         Picture         =   "frmrepclie.frx":114E
         Stretch         =   -1  'True
         Top             =   288
         Width           =   1788
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmRepClie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Dim Codigo2 As String
   Codigo2 = VGparametros.RucEmpresa
   If Option2.Value Then
     If VGRclie Then
        CrystalReport1.WindowTitle = "Inv010 -- Control de Inventarios"
        CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv010.rpt"
     Else
        CrystalReport1.WindowTitle = "Inv005 -- Control de Inventarios"
        CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv005.rpt"
     End If
   End If
   If Option1.Value Then
     If VGRclie Then
         CrystalReport1.WindowTitle = "Inv010 -- Control de Inventarios"
         CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv010.rpt"
     Else
         CrystalReport1.WindowTitle = "Inv006 -- Control de Inventarios"
         CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv006.rpt"
     End If
   End If
   
   If Option3.Value Then
     If VGRclie Then
         CrystalReport1.WindowTitle = "Inv511 -- Control de Inventarios"
         CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv010.rpt"
     Else
         CrystalReport1.WindowTitle = "Inv510 -- Control de Inventarios"
         CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv510.rpt"
     End If
   End If
   
   Ubi_Tab CrystalReport1
   CrystalReport1.DiscardSavedData = True
   CrystalReport1.WindowShowPrintBtn = True
   CrystalReport1.WindowShowRefreshBtn = True
   CrystalReport1.WindowShowSearchBtn = True
   CrystalReport1.WindowShowPrintSetupBtn = True
   CrystalReport1.formulas(0) = "EMPRESA='" & VGparametros.NomEmpresa & "'"
   CrystalReport1.formulas(1) = "HORA='" & Format(Time, "hh:mm:ss") & "'"
   CrystalReport1.Destination = crptToWindow
   If VGsql = 1 Then
     CrystalReport1.Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
     Else
     CrystalReport1.Connect = VGcadenareport2
   End If
   CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
   CrystalReport1.StoredProcParam(1) = "maeprov"
   
   If CrystalReport1.Status <> 2 Then
    CrystalReport1.Action = 1
   End If
End Sub

Private Sub Command8_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  central Me
  Option1.Value = True
  If VGRclie Then
     Me.Caption = "Listado de Clientes"
 Else
    Me.Caption = "Listado de Proveedores"
End If
End Sub


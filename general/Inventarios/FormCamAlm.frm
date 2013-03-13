VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormCamAlm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Almacen"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   Icon            =   "FormCamAlm.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command20 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   2145
      Picture         =   "FormCamAlm.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   840
   End
   Begin VB.CommandButton Command21 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3825
      Picture         =   "FormCamAlm.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   840
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Inventarios.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   5625
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TabAlm"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormCamAlm.frx":114E
      Height          =   2040
      Left            =   240
      OleObjectBlob   =   "FormCamAlm.frx":1162
      TabIndex        =   0
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "FormCamAlm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command20_Click()
If Data1.Recordset.RecordCount > 0 Then
        VGAlma = Data1.Recordset("TAALMA")
        VGNomAlm = Data1.Recordset("TADESCRI")
        MDIPrincipal.Caption = "Sistema de Inventario" & "     " & VGNomAlm & "    " & VGparametros.RucEmpresa
        Unload Me
End If
End Sub

Private Sub Command21_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Data1.DatabaseName = cRuta2
   dbGrid1.Visible = True
   central FormCamAlm
End Sub

VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FrmGraf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gráfico Comparativo"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10125
   Icon            =   "FrmGraf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6420
      Left            =   45
      OleObjectBlob   =   "FrmGraf.frx":08CA
      TabIndex        =   0
      Top             =   75
      Width           =   10050
   End
End
Attribute VB_Name = "FrmGraf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim RsAux As New ADODB.Recordset
    RsAux.Open "TmpPlanGroup", DBAuxCom
   With MSChart1
      .ShowLegend = True
      Set .DataSource = RsAux
   End With
End Sub

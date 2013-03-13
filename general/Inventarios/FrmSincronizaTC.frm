VERSION 5.00
Begin VB.Form FrmSincronizaTC 
   Caption         =   "Sincronización del Tipo de Cambio"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form2"
   ScaleHeight     =   1395
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Salir"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Sinronizar Ahora"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Este proceso sincronizará el tipo de cambio de cada documento del almacén con los registrados en contabilidad para cada fecha."
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "FrmSincronizaTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents EvConex As ADODB.Connection
Attribute EvConex.VB_VarHelpID = -1
Dim Finalizo As Boolean
Private Sub CmdAceptar_Click()
Dim SUBICA As String
Dim SQL As String

 Set EvConex = New ADODB.Connection
 
 EvConex.ConnectionString = VGcnx.ConnectionString
 EvConex.CursorLocation = VGcnx.CursorLocation
 EvConex.Provider = VGcnx.Provider
 EvConex.Open
 
 CmdAceptar.Enabled = False
 CmdExit.Enabled = False
 Screen.MousePointer = 11
 
 If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
    SUBICA = "[" & cRuta4 & "].TIPO_CAMBIO T"
 Else
    If UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
       SUBICA = "[" & cRuta2 & "].TIPO_CAMBIO T"
    End If
 End If

 SQL = "Update ( MOVALMDET  D INNER JOIN MOVALMCAB C ON D.DEALMA=C.CAALMA AND D.DETD=C.CATD AND D.DENUMDOC=C.CANUMDOC) "
 SQL = SQL & "INNER JOIN  " & SUBICA & " ON C.CAFECDOC=T.TIPOCAMB_FECHA  "
 SQL = SQL & " SET C.CATIPCAM=T.TIPOCAMB_VENTA,D.DETIPCAM=T.TIPOCAMB_VENTA "
 
 Finalizo = False
 
 EvConex.Execute SQL

 While Not Finalizo
  DoEvents
 Wend
 
 Screen.MousePointer = 1
 CmdExit.Enabled = True
 EvConex.Close
 Set EvConex = Nothing

End Sub
Private Sub CmdExit_Click()
  Unload Me
End Sub
Private Sub EvConex_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
 Finalizo = True
End Sub


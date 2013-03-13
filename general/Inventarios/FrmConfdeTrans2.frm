VERSION 5.00
Begin VB.Form FrmConfdeTrans2 
   Caption         =   "Confirmación de Transferencia"
   ClientHeight    =   2025
   ClientLeft      =   1815
   ClientTop       =   1890
   ClientWidth     =   5520
   LinkTopic       =   "Form2"
   ScaleHeight     =   2025
   ScaleWidth      =   5520
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   450
      Left            =   3045
      TabIndex        =   1
      Top             =   1275
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   450
      Left            =   1005
      TabIndex        =   0
      Top             =   1275
      Width           =   1425
   End
   Begin VB.Label LblProgress 
      Height          =   390
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   4440
   End
End
Attribute VB_Name = "FrmConfdeTrans2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RSetCab As ADODB.Recordset
Dim RSetCabDest As ADODB.Recordset
Dim RsetDet1 As ADODB.Recordset
Dim RsetDet2 As ADODB.Recordset
Dim V_AOrig As String, V_TDOrig As String, V_NDOrig As String, V_ADest As String, V_TDDest As String, V_NDDest As String

Private Sub Command1_Click()
Dim csql As String
Dim Affec As Long
Set RSetCab = New ADODB.Recordset
Set RsetDet1 = New ADODB.Recordset
Set RsetDet2 = New ADODB.Recordset
Set RSetCabDest = New ADODB.Recordset
Command1.Enabled = False

'LLeno el recordset con todas las NS y GS existentes
csql = "SELECT * FROM MOVALMCAB WHERE (CATD ='NS' OR CATD = 'GS') AND (not CARFALMA is null) And (not CARFALMA = '')"
RSetCab.Open csql, VGcnx, adOpenDynamic, adLockOptimistic
RSetCab.MoveFirst

While Not RSetCab.EOF
 ' *******    UBICO DATOS ORIGEN    *******
 'Cabecera
 'TxtAOrig = RSetCab!CAALMA
 'TxtTDOrig = RSetCab!CATD
 'TxtNDOrig = RSetCab!CANUMDOC
 'csql = "select * from movalmcab where CAALMA ='" & TxtAOrig & "' and CATD ='" & TxtTDOrig & "' and CANUMDOC ='" & TxtNDOrig & "'"
 'RSetCab.Open csql, Vgcnx, adOpenDynamic, adLockOptimistic
 LblProgress.Caption = "Confirmando todas las transferencias..."
 V_AOrig = RSetCab!CAALMA
 V_TDOrig = RSetCab!CATD
 V_NDOrig = RSetCab!CANUMDOC

 'Cargo detalles Origen
 csql = "SELECT * FROM MOVALMDET WHERE DEALMA ='" & V_AOrig & "' AND DETD ='" & V_TDOrig & "' AND DENUMDOC ='" & V_NDOrig & "'"
 RsetDet1.Open csql, VGcnx, adOpenDynamic, adLockOptimistic

 ' *******      UBICO DATOS DESTINO    *******
 'Cabecera
 V_ADest = RSetCab!CARFALMA 'Tomo Almacen Destino para ubicar la cabecera destino
 'RSetCab.Close
 'Set RSetCab = New ADODB.Recordset
 csql = "SELECT * FROM MOVALMCAB WHERE CAALMA ='" & V_ADest & "' AND CARFTDOC ='" & V_TDOrig & "' AND CARFNDOC = '" & V_NDOrig & "'"
 RSetCabDest.Open csql, VGcnx, adOpenDynamic, adLockOptimistic
 
 If RSetCabDest.RecordCount <= 0 Then
   'MsgBox "No hay registro destino, cabecera destino"
   
 Else
   V_TDDest = RSetCabDest!CATD 'Tipo Documento Registro Destino (Cabecera)
   V_NDDest = RSetCabDest!CANUMDOC 'Numero Documento Registro Destino (Cabecera)

   'Recorro detalles origen ubicando su detalle destino
   'Si algun detalle origen no tiene destino creo y lleno el destino correspondiente
   If RsetDet1.RecordCount > 0 Then 'Si hay detalles de origen
     RsetDet1.MoveFirst
   End If
   While Not RsetDet1.EOF
    csql = "SELECT * FROM MOVALMDET WHERE DEALMA = '" & V_ADest & "' AND DETD = '" & V_TDDest & "' AND DENUMDOC = '" & V_NDDest & "' AND DEITEM = " & RsetDet1!DEITEM & ""  'se retiro  por aveces el cursor no graba de forma consecutiva  AND DECODIGO = '" & RsetDet1!DECODIGO & "'
    RsetDet2.Open csql, VGcnx, adOpenDynamic, adLockOptimistic
    If RsetDet2.RecordCount <= 0 Then
      With RsetDet1
      VGcnx.Execute "INSERT INTO MOVALMDET VALUES ('" & V_ADest & "','" & V_TDDest & "','" & V_NDDest & "'," & _
      !DEITEM & ",'" & !decodigo & "','" & !DECODREF & "'," & !DECANTID & "," & !DECANTENT & _
      "," & !DECANREF & "," & !DECANFAC & ",'" & !DEORDEN & "'," & !DEPREUNI & "," & !DEPRECIO & _
      "," & !DEPRECI1 & "," & !DEDESCTO & ",'" & !DESTOCK & "'," & !DEIGV & "," & !DEIMPMN & "," & _
      !DEIMPUS & ",'" & !DESERIE & "','" & !DESITUA & "'," & FechS(!DEFECDOC, Sqlf) & ",'" & !DECENCOS & "','" & !DERFALMA & "','" & !DETR & "','" & !DEESTADO & "','" & !DECODMOV & "'," & !DEVALTOT & ",'" & !DECOMPRO & "','" & !DECODMON & "','" & !DETIPO & "'," & !DETIPCAM & "," & !DEPREVTA & ",'" & !DEMONVTA & "'," & FechS(!DEFECVEN, Sqlf) & "," & !DEDEVOL & ",'" & !DESOLI & "','" & !DEDESCRI & "'," & !DEPORDES & "," & !DEIGVPOR & "," & !DEDESCLI & "," & !DEDESESP & ",'" & !DENUMFAC & "','" & !DELOTE & "','" & !DEUNIDAD & "')"
      
      End With
    End If
    RsetDet2.Close
    RsetDet1.MoveNext
   Wend
   RsetDet1.Close
 End If
 
 RSetCabDest.Close
 RSetCab.MoveNext
Wend
RSetCab.Close
LblProgress.Caption = "Todas las transferencias han sido confirmadas.  Proceso finalizado."
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}"
End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub


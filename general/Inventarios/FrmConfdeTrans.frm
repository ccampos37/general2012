VERSION 5.00
Begin VB.Form FrmConfdeTrans 
   Caption         =   "Confirmación de Transferencia"
   ClientHeight    =   3420
   ClientLeft      =   1815
   ClientTop       =   1890
   ClientWidth     =   5835
   LinkTopic       =   "Form2"
   ScaleHeight     =   3420
   ScaleWidth      =   5835
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   450
      Left            =   3045
      TabIndex        =   2
      Top             =   2670
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   450
      Left            =   1050
      TabIndex        =   1
      Top             =   2655
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   285
      TabIndex        =   0
      Top             =   150
      Width           =   5160
      Begin VB.TextBox TxtAOrig 
         Height          =   285
         Left            =   1695
         MaxLength       =   2
         TabIndex        =   9
         Top             =   825
         Width           =   480
      End
      Begin VB.TextBox TxtNDOrig 
         Height          =   315
         Left            =   1695
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1650
         Width           =   2115
      End
      Begin VB.TextBox TxtTDOrig 
         Height          =   285
         Left            =   1695
         MaxLength       =   2
         TabIndex        =   6
         Top             =   1245
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Almacen Origen"
         Height          =   300
         Left            =   315
         TabIndex        =   8
         Top             =   885
         Width           =   1410
      End
      Begin VB.Label Label3 
         Caption         =   "Nro de Dcmto"
         Height          =   270
         Left            =   345
         TabIndex        =   5
         Top             =   1665
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Dcmto"
         Height          =   300
         Left            =   300
         TabIndex        =   4
         Top             =   1260
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "Esta opción permite confirmar la transferencia entre almacenes"
         Height          =   510
         Left            =   345
         TabIndex        =   3
         Top             =   375
         Width           =   4185
      End
   End
End
Attribute VB_Name = "FrmConfdeTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSetCab As ADODB.Recordset
Dim RsetDet1 As ADODB.Recordset
Dim RsetDet2 As ADODB.Recordset
Dim V_AOrig As String, V_TDOrig As String, V_NDOrig As String, V_ADest As String, V_TDDest As String, V_NDDest As String

Private Sub Command1_Click()
Dim csql As String
Set RSetCab = New ADODB.Recordset
Set RsetDet1 = New ADODB.Recordset
Set RsetDet2 = New ADODB.Recordset

If TxtTDOrig <> "GS" And TxtTDOrig <> "NS" Then
  MsgBox "Documento no valido para una transferencia", vbInformation, "Aviso"
  Exit Sub
End If
' *******    UBICO DATOS ORIGEN    *******
'Cabecera
csql = "select * from movalmcab where CAALMA ='" & TxtAOrig & "' and CATD ='" & TxtTDOrig & "' and CANUMDOC ='" & TxtNDOrig & "'"
RSetCab.Open csql, VGcnx, adOpenDynamic, adLockOptimistic
If RSetCab.RecordCount <= 0 Then
  MsgBox "NO existe ese documento"
  Exit Sub
End If
V_AOrig = RSetCab!CAALMA
V_TDOrig = RSetCab!CATD
V_NDOrig = RSetCab!CANUMDOC

'Cargo detalles Origen
csql = "SELECT * FROM MOVALMDET WHERE DEALMA ='" & V_AOrig & "' AND DETD ='" & V_TDOrig & "' AND DENUMDOC ='" & V_NDOrig & "'"
RsetDet1.Open csql, VGcnx, adOpenDynamic, adLockOptimistic

' *******      UBICO DATOS DESTINO    *******
'Cabecera
V_ADest = RSetCab!CARFALMA 'Tomo Almacen Destino para ubicar la cabecera destino
RSetCab.Close
'Set RSetCab = New ADODB.Recordset
csql = "SELECT * FROM MOVALMCAB WHERE CAALMA ='" & V_ADest & "' AND CARFTDOC ='" & V_TDOrig & "' AND CARFNDOC = '" & V_NDOrig & "'"
RSetCab.Open csql, VGcnx, adOpenDynamic, adLockOptimistic
If RSetCab.RecordCount <= 0 Then
  MsgBox "No hay registro destino, cabecera destino"
  Exit Sub
End If
V_TDDest = RSetCab!CATD 'Tipo Documento Registro Destino (Cabecera)
V_NDDest = RSetCab!CANUMDOC 'Numero Documento Registro Destino (Cabecera)

'Recorro detalles origen ubicando su detalle destino
'Si algun detalle origen no tiene destino creo y lleno el destino correspondiente
RsetDet1.MoveFirst
While Not RsetDet1.EOF
 csql = "SELECT * FROM MOVALMDET WHERE DEALMA = '" & V_ADest & "' AND DETD = '" & V_TDDest & "' AND DENUMDOC = '" & V_NDDest & "' AND DEITEM = " & RsetDet1!DEITEM & " AND DECODIGO = '" & RsetDet1!decodigo & "'"
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


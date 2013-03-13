VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFormAyuda"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub mostrar(cnx As ADODB.Connection, ByRef campos As Fields, ByRef Nregistros As Long)
On Error GoTo errnoesta
   
    Set m_RecordSet = New ADODB.Recordset
    m_RecordSet.Open m_SQLCadena, cnx, adOpenKeyset, adLockReadOnly
    
     If Not ExisteCampoEnRecordSet(m_RecordSet, m_primercampo) Then
        MsgBox "No se tipeo bien el nombre  del primer campo a mostrar  en la Propiedad=>¡PrimerCampo!=" & Chr(13) & _
               m_primercampo, vbCritical
        Exit Sub
    End If
    
    If Not ArmarRSdeCampos Then Exit Sub
    
    FrmAyuda.Show vbModal
    Set campos = m_fileld
    Nregistros = m_registros
    Exit Sub
errnoesta:
    MsgBox Err.Description, vbCritical, "DllAyuda"
End Sub
Private Function ArmarRSdeCampos() As Boolean
Dim i As Integer
Dim resulcaddesc As String
Dim pos2 As Integer, descrip As String
Dim tipcontrol As Integer
    ArmarRSdeCampos = True
    resulcaddesc = Trim(m_ListaCamposDescrip)
    Set rscampos = New ADODB.Recordset
    Call rscampos.Fields.Append("campo", adVarChar, 100)
    Call rscampos.Fields.Append("descripcion1", adVarChar, 100)
    Call rscampos.Fields.Append("tipocontrol", adInteger)
    rscampos.Open
    For i = 0 To m_RecordSet.Fields.Count - 1
        rscampos.AddNew
        rscampos!campo = m_RecordSet.Fields(i).Name
        
        pos2 = InStr(1, resulcaddesc, ",", vbTextCompare)
        descrip = Trim(Mid(resulcaddesc, 1, IIf(pos2 = 0, Len(resulcaddesc), pos2 - 1)))
        rscampos!descripcion1 = IIf(descrip = "", rscampos!campo, descrip)
        resulcaddesc = Trim(Mid(resulcaddesc, pos2 + 1, Len(resulcaddesc) - IIf(pos2 = 0, Len(resulcaddesc), pos2)))
        
        'Identificar El tipo de Control
        Select Case m_RecordSet.Fields(i).Type
            Case adChar, adVarChar, 130
                tipcontrol = 1
            Case adBoolean
                tipcontrol = 2
            Case adDate, 135
                tipcontrol = 3
        End Select
        rscampos!tipocontrol = tipcontrol
        rscampos.Update
    Next
End Function
Public Function ExisteCampoEnRecordSet(rs As ADODB.Recordset, Optional ByVal campo As String) As Boolean
Dim cad As String
On Error GoTo Errexistetabla
    ExisteCampoEnRecordSet = True
    cad = rs.Fields(campo).Name
    Exit Function
Errexistetabla:
    ExisteCampoEnRecordSet = False
    Exit Function
End Function
Public Property Let TituloAyuda(ByVal vNewValue As String)
    m_TituloAyuda = vNewValue
End Property
Public Property Let PrimerCampo(ByVal vNewValue As String)
    m_primercampo = vNewValue
End Property
Public Property Let ListaCamposDescrip(ByVal vNewValue As Variant)
    m_ListaCamposDescrip = vNewValue
End Property
Public Property Let SQLCadena(ByVal vNewValue As Variant)
    m_SQLCadena = vNewValue
End Property

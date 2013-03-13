VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAsiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de Asiento  Contable"
   ClientHeight    =   3555
   ClientLeft      =   2865
   ClientTop       =   1755
   ClientWidth     =   4440
   ControlBox      =   0   'False
   Icon            =   "frmAsiento.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4440
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   300
      TabIndex        =   5
      Top             =   135
      Width           =   3840
      Begin VB.Image Image1 
         Height          =   480
         Left            =   225
         Picture         =   "frmAsiento.frx":08CA
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000018&
         Caption         =   "Antes de realizar el envio, verifique  que  todas las familias de  Articulos esten relacionadas  con su cuenta contable."
         Height          =   810
         Left            =   1140
         TabIndex        =   6
         Top             =   285
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   2475
      TabIndex        =   3
      Top             =   2940
      Width           =   1620
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   345
      TabIndex        =   2
      Top             =   2940
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1845
      TabIndex        =   0
      Top             =   1710
      Width           =   1140
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   1860
      TabIndex        =   1
      Top             =   2205
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "MMMM 'del' yyyy"
      Format          =   50462723
      CurrentDate     =   36437
      MaxDate         =   401768
      MinDate         =   36161
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmAsiento.frx":0D0C
      Left            =   1860
      List            =   "frmAsiento.frx":0D34
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1020
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   270
      Top             =   1485
      Width           =   3885
   End
   Begin VB.Label Label9 
      Caption         =   "Mes"
      Height          =   255
      Left            =   525
      TabIndex        =   8
      Top             =   2220
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de cambio"
      Height          =   300
      Left            =   465
      TabIndex        =   4
      Top             =   1800
      Width           =   1320
   End
End
Attribute VB_Name = "FrmAsiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cVGDBT As ADODB.Connection
Dim cConexAux As ADODB.Connection
Dim adoreg As ADODB.Recordset
Dim adofam As ADODB.Recordset
Dim adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim Adodc3 As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim cRt As String
Dim almacen As String
Dim nTra As Integer
Dim Adodc4 As New ADODB.Recordset
Public Mes1 As Integer
Dim entro As Boolean

Private Sub Carga_RepoVal()
Dim rAdo As ADODB.Recordset
Dim Aux, CADENA As String
Dim cAnoMes As String, cCod As String
Dim cSql1 As String, CSQL2 As String
Dim Dia1 As Integer
Dim nSaldo As Double, nCosPro As Double
Dim nPrecio As Double, nCantid As Double
Dim I As Integer
Dim Codi As String
'No debe manejar por almacen
On Error GoTo ErrCarga
Set adodc1 = New ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
Set rAdo = New ADODB.Recordset
'DTPicker1 = Date
Mes1 = Month(DTPicker1) 'Combo3.ListIndex + 1 '

cAnoMes = Year(DTPicker1) & Format(Mes1, "00")
almacen = VGAlma
nTra = 1
cConexAux.BeginTrans
cConexAux.Execute "Delete From al_Kardex_Val"
cConexAux.CommitTrans
nTra = 0


 'familia
cSql1 = "Select A.*,B.*,AFamilia,Amodelo,Agrupo From"
cSql1 = cSql1 & " ((MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC )"
cSql1 = cSql1 & " Left Join MAEART C on A.Decodigo=C.Acodigo)"
cSql1 = cSql1 & " Where Month(CAFECDOC) = " & Mes1 & " and Year(CAFECDOC) = " & Year(DTPicker1) & " AND CAALMA = '" & almacen & "' And CASITGUI<>'A' and asiento=false" 'AND AFAMILIA >= '" & Text1 & "'  And  AFAMILIA <= '" & Text2 & "'"
cSql1 = cSql1 & " Order By AFamilia,DECODIGO,CAFECDOC,CAHORA"


adodc1.Open cSql1, VGCNx, adOpenStatic
Adodc2.Open "Select * From MoresMes Where SMALMA = '" & almacen & "'  and SMUSPREUNI <> 0 and SMMNPREUNI <> 0 " & _
                       " Order By SMMESPRO", VGCNx, adOpenStatic

 If Adodc2.RecordCount > 0 Then
        If Val(cAnoMes) - Val(Adodc2("SMMESPRO")) > 1 Then
               ' MsgBox "El Costo utilizado es del Mes de " & DesMes(Right(Adodc2("SMMESPRO"), 2)) & Chr(13) & ", porque no se ha hecho el Cierre en los meses anteriores", vbInformation, "Información"
        End If
Else
        MsgBox "No se ha hecho Cierre en los meses anteriores y su Costo Inicial será Cero", vbInformation, "Información"
End If
cCod = ""
'*************************
  Carga
'*************************
'
'For I = 0 To List1.ListCount - 1
'  List1.ListIndex = I
'  If List1.Selected(I) = True Then
'    If Adodc1.RecordCount > 0 Then
'      Adodc1.MoveFirst
'      Carga
'    End If
'  End If
'Next

Exit Sub
ErrCarga:
        MsgBox Err.Description
        If nTra = 1 Then cConexAux.RollbackTrans
End Sub


Private Sub Carga()
Dim Li As Integer
Dim cCod As String, Codi As String
Dim Aux, CADENA As String
Dim cAnoMes As String
Dim cSql1 As String, CSQL2 As String
Dim Dia1 As Integer
Dim nSaldo As Double, nCosPro As Double
Dim nPrecio As Double, nCantid As Double

'Mes1 = Month(DTPicker1)
cAnoMes = Year(DTPicker1) & Format(Mes1, "00")
cCod = "": Codi = "": Li = 0
Do While Not adodc1.EOF
 Codi = IIf(IsNull(adodc1("Afamilia")), "", adodc1("Afamilia")): Li = 4

       If (adodc1("CATD") = "GS" And adodc1("CACODMOV") = "GV" And adodc1("CASITGUI") = "F") Then
            adodc1.MoveNext
       Else
            If cCod = adodc1("DECODIGO") Then
                    nCantid = adodc1("DECANTID")
                    If adodc1("DEPRECIO") <> 0 Then
                        nPrecio = adodc1("DEPRECIO")
                    Else
                        nPrecio = nCosPro
                    End If
                    If adodc1("CATD") = "NI" Then
                        If IIf(adodc1("CATD") = "NI", (nSaldo + adodc1("DECANTID")), (nSaldo - adodc1("DECANTID"))) <> 0 Then
                            If nCosPro <> 0 Then
                                nCosPro = ((nSaldo * nCosPro) + (adodc1("DECANTID") * nPrecio)) / IIf(adodc1("CATD") = "NI", (nSaldo + adodc1("DECANTID")), (nSaldo - adodc1("DECANTID")))
                            Else
                                nCosPro = nPrecio
                            End If
                        Else
                                nCosPro = ((nSaldo * nCosPro) + (adodc1("DECANTID") * nPrecio)) / 1
                        End If
                    End If
                    If adodc1("CATD") = "NI" Then
                            nSaldo = nSaldo + nCantid
                    Else
                            nSaldo = nSaldo - nCantid
                    End If
                    'debe ser un update de articulo par obtener uno solo
                    CSQL2 = "Update al_Kardex_Val  set   COS_PRO=" & nCosPro & "  ,SAL_STOCK=   " & nSaldo & "  where COD_ART= '" & adodc1("DECODIGO") & "'"
                    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,FEC_DOC,HOR_DOC,COD_MOV,TIP_TRANSA,NUM_DOC,CAN_ART,PRE_UNIT,COS_PRO,SAL_STOCK,Cod_Fam)  "
                    CSQL2 = CSQL2 & " Values ('" & adodc1("DECODIGO") & "',#" & Format(adodc1("CAFECDOC"), "mm/dd/yyyy") & "#,'" & adodc1("CAHORA") & "','" & adodc1("CACODMOV") & "',"
                    CSQL2 = CSQL2 & "'" & adodc1("CATD") & "','" & adodc1("CANUMDOC") & "', " & adodc1("DECANTID") & ","
                    If adodc1("CATD") = "NI" Then
                        CSQL2 = CSQL2 & "" & nPrecio & "," & nCosPro & "," & nSaldo & ",'" & adodc1("AFAMILIA") & "' )"
                    Else
                        CSQL2 = CSQL2 & "" & nCosPro & "," & nCosPro & "," & nSaldo & ",'" & adodc1("AFAMILIA") & "')"
                    End If
                    nTra = 1
                    cConexAux.BeginTrans
                    cConexAux.Execute CSQL2
                    cConexAux.CommitTrans
                    nTra = 0

                   cCod = adodc1("DECODIGO")
            Else
                  'Cuando el codigo es diferente
                    nSaldo = 0: nCosPro = 0
                    If Adodc2.RecordCount > 0 Then
                            Adodc2.MoveFirst
                            Adodc2.Filter = "SMCODIGO = '" & adodc1("DECODIGO") & "'"
                           ' busca el moresmes si hay saldo
                            If Not Adodc2.EOF Then
                                    Adodc2.MoveLast
                                    If Adodc2("SMMESPRO") = cAnoMes Then
                                            Adodc2.MovePrevious
                                            If Adodc2.BOF Then
                                                    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC,COD_FAM) Values ('" & adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL','" & adodc1("AFAMILIA") & "')"
                                            Else
                                                    nSaldo = Adodc2("SMCANENT") - Adodc2("SMCANSAL")
                                                    nCosPro = Adodc2("SMMNPREUNI")
                                                     Set Adodc3 = New ADODB.Recordset
                                                    Adodc3.Open "Select sum(SMCANENT - SMCANSAL)  as saldo From MoresMes Where SMALMA = '" & almacen & "' And SMMESPRO < '" & cAnoMes & "' And SMCODIGO = '" & adodc1("DECODIGO") & "'", VGCNx, adOpenStatic
                                                    nSaldo = 0
                                                    nCosPro = 0
                                                    If Adodc3.RecordCount > 0 Then
                                                            nSaldo = IIf(IsNull(Adodc3("Saldo")), 0, Adodc3("Saldo"))
                                                    End If
                                                    Set Adodc3 = New ADODB.Recordset
                                                    Adodc3.Open "Select  SMMNPREUNI as promedio From MoresMes Where SMALMA = '" & almacen & "' And SMMESPRO < '" & cAnoMes & "' And SMCODIGO = '" & adodc1("DECODIGO") & "'  order by SMMESPRO DESC ", VGCNx, adOpenStatic
                                                     If Adodc3.RecordCount > 0 Then
                                                            nCosPro = IIf(IsNull(Adodc3("promedio")), 0, Adodc3("promedio"))
                                                     End If
                                                    Adodc3.Close
                                                    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC,COD_FAM) Values ('" & adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL','" & adodc1("AFAMILIA") & "')"
                                            End If
                                    Else
                                            If Val(cAnoMes) - Val(Adodc2("SMMESPRO")) = 1 Then
                                                    nSaldo = Adodc2("SMCANENT") - Adodc2("SMCANSAL")
                                                    nCosPro = Adodc2("SMMNPREUNI")
                                                    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC,COD_FAM) Values ('" & adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL','" & adodc1("AFAMILIA") & "')"
                                            Else
                                                    cAnoMes = Year(DTPicker1) & Format(Mes1, "00")
                                                    Set Adodc3 = New ADODB.Recordset
                                                    Adodc3.Open "Select sum(SMCANENT - SMCANSAL)  as saldo From MoresMes Where SMALMA = '" & almacen & "' and SMMESPRO >= '" & Adodc2("SMMESPRO") & _
                                                                            "' And SMMESPRO < '" & cAnoMes & "' And SMCODIGO = '" & adodc1("DECODIGO") & "'", VGCNx, adOpenStatic
                                                    nSaldo = 0
                                                    nCosPro = 0
                                                    If Adodc3.RecordCount > 0 Then
                                                            nSaldo = IIf(IsNull(Adodc3("Saldo")), 0, Adodc3("Saldo"))
                                                    End If
                                                    Set Adodc3 = New ADODB.Recordset
                                                    Adodc3.Open "Select  SMMNPREUNI as promedio From MoresMes Where SMALMA = '" & almacen & "' And SMMESPRO < '" & cAnoMes & "' And SMCODIGO = '" & adodc1("DECODIGO") & "'  order by SMMESPRO ", VGCNx, adOpenStatic
                                                     If Adodc3.RecordCount > 0 Then
                                                            nCosPro = IIf(IsNull(Adodc3("promedio")), 0, Adodc3("promedio"))
                                                     End If
                                                    Adodc3.Close
                                                    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC,COD_FAM) Values ('" & adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL','" & adodc1("AFAMILIA") & "')"
                                            End If
                                    End If
                            Else
                                    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC,COD_FAM) Values ('" & adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL','" & adodc1("AFAMILIA") & "')"
                           End If
                   Else
                            Set Adodc3 = New ADODB.Recordset
                            Adodc3.Open "Select sum(SMCANENT - SMCANSAL)  as saldo From MoresMes Where SMALMA = '" & almacen & _
                                                    "' And SMMESPRO < '" & cAnoMes & "' And SMCODIGO = '" & adodc1("DECODIGO") & "'", VGCNx, adOpenStatic
                            nSaldo = 0
                            If Adodc3.RecordCount > 0 Then
                                    nSaldo = IIf(IsNull(Adodc3("Saldo")), 0, Adodc3("Saldo"))
                            End If
                            Adodc3.Close
                            nCosPro = 0
                            CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC,COD_FAM) Values ('" & adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL','" & adodc1("AFAMILIA") & "')"
                   End If
                   If Trim(CSQL2) <> "" Then
                        nTra = 1
                        cConexAux.BeginTrans
                        cConexAux.Execute CSQL2
                        cConexAux.CommitTrans
                        nTra = 0
                        CSQL2 = ""
                    End If
                   Adodc2.Filter = ""
                   cCod = adodc1("DECODIGO")
                    nCantid = adodc1("DECANTID")
                    If adodc1("DEPRECIO") <> 0 Then
                        nPrecio = adodc1("DEPRECIO")
                    Else
                        nPrecio = nCosPro
                    End If
                    If adodc1("CATD") = "NI" Then
                        If IIf(adodc1("CATD") = "NI", (nSaldo + adodc1("DECANTID")), (nSaldo - adodc1("DECANTID"))) <> 0 Then
                            If nCosPro <> 0 Then
                                nCosPro = ((nSaldo * nCosPro) + (adodc1("DECANTID") * nPrecio)) / IIf(adodc1("CATD") = "NI", (nSaldo + adodc1("DECANTID")), (nSaldo - adodc1("DECANTID")))
                            Else
                                nCosPro = nPrecio
                            End If
                        Else
                                nCosPro = ((nSaldo * nCosPro) + (adodc1("DECANTID") * nPrecio)) / 1
                        End If
                    End If
                    If adodc1("CATD") = "NI" Then
                            nSaldo = nSaldo + nCantid
                    Else
                            nSaldo = nSaldo - nCantid
                    End If
                    CSQL2 = "Update al_Kardex_Val  set   COS_PRO =" & nCosPro & "  ,SAL_STOCK =   " & nSaldo & "  where COD_ART= '" & adodc1("DECODIGO") & "'"
                    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,FEC_DOC,HOR_DOC,COD_MOV,TIP_TRANSA,NUM_DOC,CAN_ART,PRE_UNIT,COS_PRO,SAL_STOCK,COD_FAM)  "
                    CSQL2 = CSQL2 & " Values ('" & adodc1("DECODIGO") & "',#" & Format(adodc1("CAFECDOC"), "mm/dd/yyyy") & "#,'" & IIf(IsNull(adodc1("CAHORA")) Or Trim(adodc1("CAHORA")) = "", " ", adodc1("CAHORA")) & "','" & adodc1("CACODMOV") & "',"
                    CSQL2 = CSQL2 & "'" & adodc1("CATD") & "','" & adodc1("CANUMDOC") & "', " & adodc1("DECANTID") & ","

                    If adodc1("CATD") = "NI" Then
                         CSQL2 = CSQL2 & "" & nPrecio & "," & nCosPro & "," & nSaldo & ","
                    Else
                         CSQL2 = CSQL2 & "" & nCosPro & "," & nCosPro & "," & nSaldo & ","
                    End If
                        CSQL2 = CSQL2 & "'" & adodc1("AFAMILIA") & "' )"
                   nTra = 1
                  cConexAux.BeginTrans
                  cConexAux.Execute CSQL2
                  cConexAux.CommitTrans
                   nTra = 0

                  cCod = adodc1("DECODIGO")
            End If
            adodc1.MoveNext
        End If
        If adodc1.EOF Then Exit Do

Loop
End Sub

Private Sub Command1_Click()
Dim csql As String
  'verifica si tiene subdiario
csql = "select conf_codigo from configuracion "
Set adoreg = New ADODB.Recordset
adoreg.Open csql, VGCNx, adOpenDynamic, adLockOptimistic

If Trim(Text1) = "" Then
   MsgBox "No ha ingresado el Tipo de Cambio.... ", vbInformation, "Inventarios"
   Exit Sub
End If
If Not IsNumeric(Text1) Then
   MsgBox "El Valor Ingresado no es Numerico....!", vbInformation, "Inventarios"
   Exit Sub
End If

If Not adoreg.EOF Then
   If IsNull(adoreg(0)) Or adoreg(0) = "" Then
       MsgBox "No se ha definido el subdiario", vbInformation, "Aviso"
       adoreg.Close
       Exit Sub
   End If
Else
       MsgBox "No se ha definido el subdiario", vbInformation, "Aviso"
       adoreg.Close
       Exit Sub
End If
adoreg.Close
 If MsgBox("Estas seguro de realizar este proceso?", vbInformation + vbOKCancel, "Aviso") = vbOK Then
        If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
            Carga_RepoVal
            familia
        Else
             MsgBox " No tiene enlace con contabilidad", vbInformation, "Aviso"
        End If
End If
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub Form_Activate()
   ADOConectar
   DTPicker1 = Date
End Sub

Private Sub Form_Load()
   central Me
   entro = False
End Sub

Private Sub ADOConectar()
cRt = App.Path & "\BdAuxCom.Mdb"
Set cConexAux = New ADODB.Connection
cConexAux.CursorLocation = adUseClient
cConexAux.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & cRt & ";"
cConexAux.Open
End Sub
Private Sub familia()

Dim cSqlE1 As String
Dim cSqlE2 As String
Dim cSqlE3 As String
Dim csql As String
Dim secue As String
Dim suma As Double
Dim subdiar As String
Dim cNUMDOC As String
Dim tipca As Double
Dim ya_inserto As Long

On Error GoTo Err
Mes1 = DTPicker1.Month ' Se tiene que boorar solo para prueba
tipca = CDbl(Text1)
'subdiar = "02"
Set adoreg = New ADODB.Recordset
Set adofam = New ADODB.Recordset
secue = "0000"
suma = 0
ADOConectar1

csql = "select conf_codigo from configuracion "
adoreg.Open csql, VGCNx, adOpenDynamic, adLockOptimistic
If adoreg.EOF Or IsNull(adoreg(0)) Or adoreg(0) = "" Then
       MsgBox "No se ha definido el subdiario", vbInformation, "Aviso"
       adoreg.Close
       Exit Sub
End If
subdiar = adoreg(0)
cNUMDOC = ANumeracion(subdiar)
adofam.Open "SELECT * from familia  order by fam_haber", VGCNx, adOpenDynamic, adLockOptimistic


If Not adofam.EOF Then

    While Not adofam.EOF
          If IsNull(adofam("fam_debe")) Or IsNull(adofam("fam_haber")) Then
               MsgBox "No ha definido cta contable para la familia  " & adofam("fam_codigo"), vbExclamation, "Aviso"
               Exit Sub
          End If
          If adofam("fam_haber") = " " Then
               MsgBox "No ha definido cta contable para la familia  " & adofam("fam_codigo"), vbExclamation, "Aviso"
               Exit Sub
          End If
          adofam.MoveNext
    Wend
    adofam.MoveFirst
    
    cConexAux.Execute "Delete from CabMov1"
    cConexAux.Execute "Delete from DetMov1"
    While Not adofam.EOF
    
         Dim ado1 As New ADODB.Recordset
         Dim Total As Double
         Dim Cod As String
         Total = 0   'Order By Cod_Art,Fec_Doc,hor_doc
         ado1.Open "Select Cod_Art,Cos_Pro,Sal_Stock from al_Kardex_Val where cod_fam= '" & adofam("fam_codigo") & "'", cConexAux, adOpenDynamic, adLockOptimistic
         If ado1.RecordCount > 0 Then
            ado1.MoveFirst
            Do While Not ado1.EOF
               Cod = ado1("Cod_Art")
               ado1.MoveNext
               If Not ado1.EOF Then
                  If Cod <> ado1("Cod_Art") Then
                     ado1.MovePrevious
                     Total = Total + Format((ado1("Cos_Pro") * ado1("Sal_Stock")), "###,###,###,##0.00")
                  Else
                      ado1.MovePrevious
                  End If
                  ado1.MoveNext
               End If
               If ado1.EOF Then
                  Exit Do
               End If
            Loop
            ado1.MoveLast
            Total = Total + Format((ado1("Cos_Pro") * ado1("Sal_Stock")), "###,###,###,##0.00")
         End If
         ado1.Close
'
         If Total <> 0 Then
                csql = "SELECT sum( cos_pro*sal_stock) as saldo from al_Kardex_Val where (cos_pro*sal_stock)>0 and cod_fam= '" & adofam("fam_codigo") & "'"
                Set adoreg = New ADODB.Recordset
                adoreg.Open csql, cConexAux, adOpenDynamic, adLockOptimistic
                If Not adoreg.EOF Then
                   If IsNull(adoreg("Saldo")) Or adoreg("Saldo") < 0 Then
                   
                   Else
                       ' " & Format(Mes1, "00") & "
                       secue = Format(secue + 1, "000")
                       cSqlE2 = "Insert Into DetMov1 (SUBDIAR_CODIGO,DMOV_SECUE,DMOV_C_COMPR,DMOV_FECHA,DMOV_CUENT,"
                       cSqlE2 = cSqlE2 & "DMOV_HABER,DMOV_HABUS)  values  ('" & subdiar & "','" & secue & "', '" & cNUMDOC & "',#" & Format(DTPicker1, "mm/dd/yyyy") & "#,'" & adofam("fam_haber") & "'," & CDbl(Total) & "," & CDbl(Total) * tipca & " )"
    
                       secue = Format(secue + 1, "000")
                       cSqlE3 = "Insert Into DetMov1 (SUBDIAR_CODIGO,DMOV_SECUE,DMOV_C_COMPR,DMOV_FECHA,DMOV_CUENT,"
                       cSqlE3 = cSqlE3 & "DMOV_DEBE,DMOV_DEBUS)  values ('" & subdiar & "','" & secue & "', '" & cNUMDOC & "',#" & Format(DTPicker1, "mm/dd/yyyy") & "#,'" & adofam("fam_debe") & "'," & CDbl(Total) & "," & CDbl(Total) * tipca & " )"
                    
                       cConexAux.BeginTrans
                       cConexAux.Execute cSqlE2
                       cConexAux.Execute cSqlE3
                       cConexAux.CommitTrans
                       suma = CDbl(Total) + suma
                   End If
                End If
         End If
         adofam.MoveNext
    Wend               'iif('" & cMond & "' ='MN',val(format(1/CMOV_TIPCA,'.000000'))  '  revisarr*************
    
    If CDbl(suma) > 0 Then
    
       cSqlE1 = "Insert Into CabMov1 (SUBDIAR_CODIGO,CMOV_C_COMPR,CMOV_FECHA,CMOV_MONED,CMOV_CONVE,CMOV_TIPCA,CMOV_DEBE,CMOV_HABER,CMOV_DEBUS,CMOV_HABUS)"
       cSqlE1 = cSqlE1 & " values ('" & subdiar & "', '" & cNUMDOC & "',#" & Format(DTPicker1, "mm/dd/yyyy") & "# ,'MN','VTA'," & tipca & "," & suma & "," & suma & "," & suma * tipca & "," & suma * tipca & "   )"
       
       cConexAux.BeginTrans
       cConexAux.Execute cSqlE1
       cConexAux.CommitTrans
       'cVGDBT.Execute cSqlE1
       
    End If
End If
adofam.Close

csql = "Select * From al_Kardex_Val WHERE cos_pro>0 and sal_stock>0"
Set adoreg = New ADODB.Recordset
adoreg.Open csql, cConexAux, adOpenDynamic, adLockOptimistic
If Not adoreg.EOF Then
   MsgBox "Se realizó el asiento previo", vbInformation, "Sistema de Inventario"
ElseIf suma = 0 Then
   MsgBox "No hay registro para incorporar a contabilidad, " & Chr(13) & "quizás ya se incorporo anteriormente", vbInformation, "Sistema de Inventario"
   adoreg.Close
   Exit Sub
Else
   MsgBox "No se realizó el asiento previo revisar Costos y Stock", vbInformation, "Sistema de Inventario"
   adoreg.Close
   Exit Sub
End If
adoreg.Close
FrmAsiento2.Show 1
Exit Sub
Err:

 If Err.Number = -2147467259 Then
    'MsgBox "Ya se realizó el asiento", vbInformation, "Aviso"
    'cVGDBT.RollbackTrans
 Else
     MsgBox Err.Description, vbInformation, "Aviso"
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err

If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") And entro Then
        If cVGDBT.State = 1 Then cVGDBT.Close
End If
Exit Sub
Err:
   MsgBox Err.Description, vbInformation
End Sub


Private Sub ADOConectar1()
If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
   Set cVGDBT = New ADODB.Connection
   With cVGDBT 'para Movimientos
        .CursorLocation = adUseClient
        .Provider = "Microsoft.Jet.OLEDB.3.51"
        .ConnectionString = "Data Source=" & VGParamSistem.RutaReport & VGContTra & Year(DTPicker1) & ".MDB"
        .Open
   End With
   entro = True
Else
   MsgBox "No tiene enlace con El modulo de Contabilidad", vbInformation, "Aviso"
   Command1.Enabled = False
   Command2.SetFocus
End If
End Sub


Public Function ANumeracion(ASUB As String) As String
Dim adoreg As ADODB.Recordset
Set adoreg = New ADODB.Recordset

adoreg.Open "SELECT MAX(CMOV_C_COMPR) FROM CABMOV" & Format(Month(DTPicker1), "00") & " WHERE SUBDIAR_CODIGO='" & ASUB & "' AND MONTH(CMOV_FECHA)=" & Month(DTPicker1), cVGDBT, adOpenStatic
If adoreg.RecordCount <> 0 Then
        ANumeracion = Format(Val(IIf(IsNull(adoreg.Fields(0)), 0, adoreg.Fields(0))) + 1, "0000")
Else
        ANumeracion = 1
End If
adoreg.Close

End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And IsNumeric(Text1) Then
      DTPicker1.SetFocus
   Else
     If Chr(KeyAscii) = "." And IsNumeric(Text1) Then Exit Sub
     If ((Chr$(KeyAscii) < "0" Or Chr(KeyAscii) > "9")) And KeyAscii <> 8 Then KeyAscii = 0
   End If
End Sub


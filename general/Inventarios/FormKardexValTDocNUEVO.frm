VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormKardexValTXDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kardex Valorizado Por Documentos"
   ClientHeight    =   4620
   ClientLeft      =   1470
   ClientTop       =   1590
   ClientWidth     =   5865
   ControlBox      =   0   'False
   Icon            =   "FormKardexValTDocNUEVO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5865
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   390
      Top             =   3825
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   690
      Left            =   2310
      Picture         =   "FormKardexValTDocNUEVO.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3675
      Width           =   735
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3180
      Picture         =   "FormKardexValTDocNUEVO.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3645
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   4485
      Left            =   135
      TabIndex        =   3
      Top             =   60
      Width           =   5655
      Begin VB.TextBox TxTransa 
         Height          =   285
         Index           =   1
         Left            =   1530
         MaxLength       =   2
         TabIndex        =   13
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox TxTransa 
         Height          =   285
         Index           =   0
         Left            =   1530
         MaxLength       =   2
         TabIndex        =   12
         Top             =   1890
         Width           =   495
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "FormKardexValTDocNUEVO.frx":114E
         Left            =   1530
         List            =   "FormKardexValTDocNUEVO.frx":115B
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1470
         Width           =   1380
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FormKardexValTDocNUEVO.frx":1179
         Left            =   1530
         List            =   "FormKardexValTDocNUEVO.frx":1183
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1050
         Width           =   2160
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1530
         TabIndex        =   1
         Top             =   645
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MMMM 'del' yyyy"
         Format          =   24707075
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormKardexValTDocNUEVO.frx":1197
         Left            =   1530
         List            =   "FormKardexValTDocNUEVO.frx":1199
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2760
      End
      Begin VB.Label lbltrans 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2130
         TabIndex        =   16
         Top             =   1965
         Width           =   3225
      End
      Begin VB.Label lbltrans 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   14
         Top             =   2235
         Width           =   3225
      End
      Begin VB.Label Label10 
         Caption         =   "Tip de Mvto:"
         Height          =   255
         Left            =   255
         TabIndex        =   11
         Top             =   1515
         Width           =   1035
      End
      Begin VB.Label Label8 
         Caption         =   "Del Mvmto."
         Height          =   255
         Left            =   255
         TabIndex        =   9
         Top             =   1950
         Width           =   990
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda   :"
         Height          =   195
         Left            =   270
         TabIndex        =   8
         Top             =   1125
         Width           =   765
      End
      Begin VB.Label Label9 
         Caption         =   "Mes"
         Height          =   255
         Left            =   300
         TabIndex        =   6
         Top             =   690
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Al  Mvmto."
         Height          =   255
         Left            =   270
         TabIndex        =   5
         Top             =   2355
         Width           =   900
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   270
         TabIndex        =   4
         Top             =   285
         Width           =   735
      End
   End
End
Attribute VB_Name = "FormKardexValTXDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cConexAux As ADODB.Connection
Dim adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim Adodc3 As ADODB.Recordset
Dim rS As ADODB.Recordset
Dim cRt As String
Dim almacen As String
Dim almacenAnt  As String
Dim nTra As Integer
Dim Adodc4 As ADODB.Recordset
Dim Adostk As ADODB.Recordset
Dim AdoMes As ADODB.Recordset
Dim Total As Double
Dim tipo As String

Private Sub CmdSalir_Click()
If cConexAux.State = 1 Then Set cConexAux = Nothing
Unload Me
End Sub

Private Sub Combo1_Click()
rS.MoveFirst
rS.Move Combo1.ListIndex
almacen = Format(rS(0), "00")
End Sub

Private Sub CmdAceptar_Click()
Dim rSql As String
Dim Va1 As String, Va2 As String
Screen.MousePointer = 11
On Error GoTo err
Set Adodc3 = New ADODB.Recordset  'Para sacar la descripcion del rango elegido

Carga_RepoVal
       CrystalReport1.WindowTitle = "Inv132 -- Control de Inventarios"
       CrystalReport1.ReportFileName = RUTA & "reportes\inv132.rpt"
  
  Dim ccadena As String
  If CrystalReport1.ReportFileName = "" Then
      MsgBox "No hay registros a imprimir", vbInformation, "Aviso"
      Screen.MousePointer = 1
      Exit Sub
  End If
  ccadena = ""
  
  If Mid(Combo4.text, 1, 1) = "T" Then
    If TxTransa(0) <> "" Then
      ccadena = "({KARDEX_VAL.ING_SAL} ='I' AND  "
      ccadena = ccadena & "{KARDEX_VAL.COD_MOV} =>'" & TxTransa(0) & "' AND  {KARDEX_VAL.COD_MOV} <='" & TxTransa(1) & "') OR "
      ccadena = ccadena & "({KARDEX_VAL.ING_SAL} ='S' AND  "
      ccadena = ccadena & "{KARDEX_VAL.COD_MOV} =>'" & TxTransa(0) & "' AND  {KARDEX_VAL.COD_MOV} <='" & TxTransa(1) & "') "
    End If
  Else
    ccadena = "{KARDEX_VAL.ING_SAL} ='" & Mid(Combo4.text, 1, 1) & "'"
    If TxTransa(0) <> "" Then
      ccadena = ccadena & " AND  {KARDEX_VAL.COD_MOV} >='" & TxTransa(0) & "' AND  {KARDEX_VAL.COD_MOV} <='" & TxTransa(1) & "' "
    End If
  End If
  
  Ubi_Tab CrystalReport1
  CrystalReport1.DiscardSavedData = True
  CrystalReport1.Destination = crptToWindow
  CrystalReport1.WindowShowPrintBtn = True
  CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.WindowShowSearchBtn = True
  CrystalReport1.WindowShowPrintSetupBtn = True
  CrystalReport1.SelectionFormula = ccadena
  CrystalReport1.Formulas(0) = "ALMACEN = '" & UCase(Combo1.text) & "'"
  CrystalReport1.Formulas(1) = "Mes= '" & UCase(Format(DTPicker1, "MMMM - yyyy")) & "'"
  CrystalReport1.Formulas(2) = "EMPRESA= '" & UCase(VGNemp) & "'"
  CrystalReport1.Formulas(3) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
  If Combo3.ListIndex <> 0 Then
     CrystalReport1.Formulas(4) = "MONEY= 'DOLAR'"
  Else
      CrystalReport1.Formulas(4) = "MONEY= 'SOLES'"
  End If
  
  If Combo2.ListIndex = 0 Then CrystalReport1.Formulas(5) = "ART1= '" & Text1 & "'"
  If Combo2.ListIndex = 0 Then CrystalReport1.Formulas(6) = "ART2 ='" & Text2 & "'"

  If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
 Screen.MousePointer = 1
 Exit Sub
err:
   MsgBox err.Description, vbInformation, "Aviso"
   Screen.MousePointer = 1
   Resume
End Sub

Private Sub Carga_RepoVal()
Dim rAdo As ADODB.Recordset
Dim Aux, cadena As String
Dim cAnoMes As String, cCod As String
Dim cSql1 As String, CSQL2 As String
Dim Dia1 As Integer, Mes1 As Integer
Dim nSaldo As Double, nCosPro As Double
Dim nPrecio As Double, nCantid As Double
Dim i As Integer
Dim Codi As String
Dim rSql As String
Dim csql As String
Dim uSql As String
Dim saldo As Double
On Error GoTo ErrCarga
Set adodc1 = New ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
Set rAdo = New ADODB.Recordset
    
Mes1 = Month(DTPicker1)
cAnoMes = Year(DTPicker1) & Format(Mes1, "00")

nTra = 1
cConexAux.BeginTrans
cConexAux.Execute "Delete From Kardex_Val"
cConexAux.CommitTrans
nTra = 0

If Text1.Visible And Trim(Text2) = "" Then
    Text2 = Text1
End If
'limpia
'csql = "delete  from MORESMES where SMCANENT = 0 and SMCANSAL =  0"
'cConexCom.Execute csql
       
If Combo2.ListIndex = 0 Then
cSql1 = "Select * From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
cSql1 = cSql1 & "  Where Month(CAFECDOC) = " & Mes1 & " and Year(CAFECDOC) = " & Year(DTPicker1) & " AND CAALMA = '" & almacen & "'  "  'AND DECODIGO >= '" & Text1 & "' AND DECODIGO = 'CH-AB029-E34'
cSql1 = cSql1 & "  And CASITGUI<>'A' Order By DECODIGO,CAFECDOC,CAHORA"  'And  DECODIGO <= '" & Text2 & "'
ElseIf Combo2.ListIndex = 1 Then    'GRUPO
cSql1 = "Select A.*,B.*,AFamilia,Amodelo,Agrupo From"
cSql1 = cSql1 & " ((MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC )"
cSql1 = cSql1 & " Left Join MAEART C on A.Decodigo=C.Acodigo)"
cSql1 = cSql1 & " Where Month(CAFECDOC) = " & Mes1 & " and Year(CAFECDOC) = " & Year(DTPicker1) & " AND CAALMA = '" & almacen & "'" 'AND AGRUPO >= '" & Text1 & "'  And  AGRUPO <= '" & Text2 & "'"
cSql1 = cSql1 & " and AMODELO = '" & Text4 & "'  And  AFAMILIA = '" & Text3 & "' And CASITGUI<>'A' "
cSql1 = cSql1 & " Order By Agrupo,DECODIGO,CAFECDOC,CAHORA"
ElseIf Combo2.ListIndex = 2 Then    'LINEA
cSql1 = "Select A.*,B.*,AFamilia,Amodelo,Agrupo From"
cSql1 = cSql1 & " ((MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC )"
cSql1 = cSql1 & " Left Join MAEART C on A.Decodigo=C.Acodigo)"
cSql1 = cSql1 & " Where Month(CAFECDOC) = " & Mes1 & " and Year(CAFECDOC) = " & Year(DTPicker1) & " AND CAALMA = '" & almacen & "'" ' AND AMODELO >= '" & Text1 & "'  And  AMODELO <= '" & Text2 & "'"
cSql1 = cSql1 & " and AFAMILIA = '" & Text3 & "' And CASITGUI<>'A'"
cSql1 = cSql1 & " Order By Amodelo,DECODIGO,CAFECDOC,CAHORA"
ElseIf Combo2.ListIndex = 3 Then   'FAMILIA
cSql1 = "Select A.*,B.*,AFamilia,Amodelo,Agrupo From"
cSql1 = cSql1 & " ((MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC )"
cSql1 = cSql1 & " Left Join MAEART C on A.Decodigo=C.Acodigo)"
cSql1 = cSql1 & " Where Month(CAFECDOC) = " & Mes1 & " and Year(CAFECDOC) = " & Year(DTPicker1) & " AND CAALMA = '" & almacen & "' And CASITGUI<>'A'   " 'AND AFAMILIA >= '" & Text1 & "'  And  AFAMILIA <= '" & Text2 & "'"
cSql1 = cSql1 & " Order By AFamilia,DECODIGO,CAFECDOC,CAHORA"
End If

adodc1.Open cSql1, cConexCom, adOpenStatic
''Adodc2.Open "Select * From MoresMes Where SMALMA = '" & almacen & "'  and SMUSPREUNI <> 0 and SMMNPREUNI <> 0 " & _
                       " Order By SMMESPRO", cConexCom, adOpenStatic
Adodc2.Open "Select * From MoresMes Where SMALMA = '" & almacen & "' Order By SMMESPRO", cConexCom, adOpenStatic
 
'If Adodc2.RecordCount > 0 Then
 '       If Val(cAnoMes) - Val(Adodc2("SMMESPRO")) > 1 Then
              '  MsgBox "El Costo utilizado es del Mes de " & DesMes(Right(Adodc2("SMMESPRO"), 2)) & Chr(13) & ", porque no se ha hecho el Cierre en los meses anteriores", vbInformation, "Información"
 '       End If
'Else
'        MsgBox "No se ha hecho Cierre en los meses anteriores y su Costo Inicial será Cero", vbInformation, "Información"
'End If
'cCod = ""

'**************************************
'*****ROBERTO*********************************

    Call Carga2
    Exit Sub

'**************************************
'**************************************

If Combo2.ListIndex = 0 Then
    If adodc1.RecordCount > 0 Then
       adodc1.MoveFirst
       Carga
    End If
    
    'Para los articulos que no han  tenido movimiento ,suma=suma+saldotemp
    'uSql = "select s.stcodigo, s.stskdis from stkart s  WHERE S.STCODIGO NOT  In (SELECT COD_ART FROM kardex_val IN '" & App.Path & "\bdauxcom.mdb ')  and  s.STALMA ='" & almacen & "' and  s.stcodigo >= '" & Text1 & "' and s.stcodigo <= '" & Text2 & "'"
    uSql = "SELECT SUM(S.SMCANENT- S.SMCANSAL) AS COL, SMCODIGO  " & _
            "FROM MORESMES AS s " & _
            "WHERE S.SMCODIGO NOT  In (SELECT COD_ART FROM kardex_val IN '" & App.Path & "\bdauxcom.mdb ')  and  s.smalma = '" & almacen & "'  and  s.smcodigo >= '" & Trim(Text1) & "' and s.smcodigo <= '" & Trim(Text2) & "' and s.smmespro < '" & cAnoMes & "' group by smcodigo"
    Set Adostk = New ADODB.Recordset
    Adostk.Open uSql, cConexCom, adOpenStatic
    While Not Adostk.EOF
       csql = "select top 1 (SMMNPREUNI) as costo from moresmes where   smcodigo = '" & Adostk(1) & "' and  smalma = '" & almacen & "' and smmespro < '" & cAnoMes & "' order by   smmespro desc "
       Set AdoMes = New ADODB.Recordset
       AdoMes.Open csql, cConexCom, adOpenStatic
       If Not AdoMes.EOF Then
'             If Adostk(0) <> 0 Then
'                     rSql = "INSERT INTO kardex_val ( COD_ART, NUM_DOC, SAL_STOCK, COS_PRO )" & _
'                                " values ('" & Adostk(1) & "' ,'SALDO INICIAL'," & Adostk(0) & "," & AdoMes("costo") & ")"
'                    cConexAux.Execute rSql
'              End If

       End If
       Adostk.MoveNext
    Wend
   'aqui debo tener la suma acumulada a imprimir
 
  Exit Sub
End If

For i = 0 To List1.ListCount - 1
  List1.ListIndex = i
  If List1.Selected(i) = True Then
    If adodc1.RecordCount > 0 Then
      adodc1.MoveFirst
      Carga2
    End If
  End If
Next

Exit Sub
ErrCarga:
        MsgBox err.Description
        If nTra = 1 Then cConexAux.RollbackTrans
        Resume
End Sub

Private Sub Carga()
Dim TCamb As Double
Dim Li As Integer
Dim cCod As String, Codi As String
Dim Aux, cadena As String
Dim cAnoMes As String
Dim cSql1 As String, CSQL2 As String
Dim Dia1 As Integer, Mes1 As Integer
Dim nSaldo As Double, nCosPro As Double
Dim nPrecio As Double, nCantid As Double
Dim cMesPro As String
Dim Rsql1 As String
Dim nTipCam As String
'**********Roberto
Dim rSAX As New ADODB.Recordset
Dim SqlM As String
Dim Flag0 As Integer
'*************************
Mes1 = Month(DTPicker1)
cAnoMes = Year(DTPicker1) & Format(Mes1, "00")
cCod = "": Codi = "": Li = 0
Flag0 = 0
Do While Not adodc1.EOF
    If Combo2.ListIndex = 1 Then Codi = IIf(IsNull(adodc1("Agrupo")), "", adodc1("Agrupo")): Li = 8 'captura el dato para la comparacion
    If Combo2.ListIndex = 2 Then Codi = IIf(IsNull(adodc1("Amodelo")), "", adodc1("Amodelo")): Li = 4
    If Combo2.ListIndex = 3 Then Codi = IIf(IsNull(adodc1("Afamilia")), "", adodc1("Afamilia")): Li = 4
    
    If adodc1("detipcam") <> 0 And adodc1("cacodmon") = "ME" Then
        TCamb = adodc1("detipcam")
    Else
     If Combo3.ListIndex = 0 Then
        TCamb = 1
     Else
        
        If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
           TCamb = Val(Devolver_Dato(3, adodc1("cafecdoc"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
        ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
               TCamb = Val(Devolver_Dato(1, adodc1("cafecdoc"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
        End If
     End If
    End If
    
    If Trim(Mid(List1.text, 1, Li)) = Codi Or Combo2.ListIndex = 0 Then
    
       If (adodc1("CATD") = "GS" And adodc1("CACODMOV") = "GF" And adodc1("CASITGUI") = "F") Then '(adodc1("CATD") = "GS" And adodc1("CACODMOV") = "GV" And adodc1("CASITGUI") = "F") Or
            adodc1.MoveNext
       Else
'******************************************************
'*********HABILITADO ROBERTO MAZA
'            If adodc1("DECODIGO") = "CH-AB460-E46" Then
'                     MsgBox "NOTA"
'            End If
'******************************************************
            If cCod = adodc1("DECODIGO") Then
                   'Codigo aumentado
                    
                    nCantid = adodc1("DECANTID")
                    If Combo3.ListIndex <> 0 Then
                        If adodc1("DEPRECIO") <> 0 Then
                           If CDbl(TCamb) <> 0 Then
                              nPrecio = CDbl(adodc1("DEPRECIO") / TCamb)
                           Else
                               nPrecio = adodc1("DEPRECIO")
                           End If
                        Else
                            nPrecio = CDbl(nCosPro)
                        End If
                    Else
                        If adodc1("DEPRECIO") <> 0 Then
                           nPrecio = adodc1("DEPRECIO")
                        Else
                            nPrecio = nCosPro
                        End If
                    End If
                    If adodc1("CATD") = "NI" Or adodc1("CATD") = "NC" Then
                        If IIf(adodc1("CATD") = "NI" Or adodc1("CATD") = "NC", (nSaldo + adodc1("DECANTID")), (nSaldo - adodc1("DECANTID"))) <> 0 Then
                            If nPrecio <> 0 Then
                                nCosPro = ((nSaldo * nCosPro) + (adodc1("DECANTID") * nPrecio)) / IIf(adodc1("CATD") = "NI" Or adodc1("CATD") = "NC", (nSaldo + adodc1("DECANTID")), (nSaldo - adodc1("DECANTID")))
                            Else
                               nCosPro = nPrecio
                            End If
'                            If Combo3.ListIndex <> 0 Then
'                                   If CDbl(TCamb) <> 0 Then
'                                      nCosPro = nCosPro / TCamb
'                                   End If
'                            End If
                        Else
                              
                                 nCosPro = ((nSaldo * nCosPro) + (adodc1("DECANTID") * nPrecio)) / 1
                              
'                              If Combo3.ListIndex <> 0 Then
'                                  If CDbl(TCamb) <> 0 Then
'                                      nCosPro = nCosPro / TCamb
'                                 End If
'                               End If
                                
                        End If
                    End If
                    If adodc1("CATD") = "NI" Or adodc1("CATD") = "NC" Then
                            nSaldo = nSaldo + nCantid
                    Else
                            nSaldo = nSaldo - nCantid
                    End If
                    CSQL2 = "Insert Into Kardex_Val (COD_ART,FEC_DOC,HOR_DOC,COD_MOV,TIP_TRANSA,NUM_DOC,CAN_ART,PRE_UNIT,COS_PRO,SAL_STOCK,SER_LOT,ING_SAL)  "
                    CSQL2 = CSQL2 & " Values ('" & adodc1("DECODIGO") & "',#" & Format(adodc1("CAFECDOC"), "mm/dd/yyyy") & "#,'" & adodc1("CAHORA") & "','" & adodc1("CACODMOV") & "',"
                    CSQL2 = CSQL2 & "'" & adodc1("CATD") & "','" & adodc1("CANUMDOC") & "', " & adodc1("DECANTID") & ","
                    '************************roberto
                    If nCosPro = 0 And nPrecio = 0 Then Flag0 = 1
                    '************************************
                    If adodc1("CATD") = "NI" Or adodc1("CATD") = "NC" Then
                        '*************************************Roberto
                        If nCosPro = 0 Then nCosPro = nPrecio
                        '***************************************
                        CSQL2 = CSQL2 & "" & nPrecio & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','I')"
                    Else
                        '*************************************Roberto
                        If nCosPro = 0 Then nCosPro = nPrecio
                        '***************************************
                        CSQL2 = CSQL2 & "" & nCosPro & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','S')"
                    End If
                    nTra = 1
                    cConexAux.BeginTrans
                    cConexAux.Execute CSQL2
                    cConexAux.CommitTrans
                    nTra = 0
                   
                   cCod = adodc1("DECODIGO")
            Else
                   'aqui actulizo el moresmes ********************************************
                   If cCod <> "" Then
                         cMesPro = Year(DTPicker1) & Format(Mes1, "00")
                         
                         If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
                             nTipCam = Val(Devolver_Dato(3, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
                         ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
                             nTipCam = Val(Devolver_Dato(1, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
                         End If
                         Rsql1 = "Update  MoResMes  Set  SMMNPREUNI  = " & nCosPro & ", SMUSPREUNI  = " & nCosPro / IIf(nTipCam = 0, 1, nTipCam) & "  where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & cMesPro & "' AND  SMCODIGO= '" & cCod & "'"
                        
                         cConexCom.BeginTrans
                         cConexCom.Execute Rsql1
                         cConexCom.CommitTrans
                         
                   End If
                   
                   'aqui lo nuevos valores
                    nSaldo = 0: nCosPro = 0
                    Flag0 = 0   'RMAZA
                    If Adodc2.RecordCount > 0 Then
                            Adodc2.MoveFirst
                            Adodc2.Filter = "SMCODIGO = '" & adodc1("DECODIGO") & "'"
                            If Not Adodc2.EOF Then
                                    Adodc2.MoveLast
                                    If Adodc2("SMMESPRO") = cAnoMes Then
                                            Adodc2.MovePrevious
                                            If Adodc2.BOF Then
                                                   ' CSQL2 = "Insert Into Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
                                            Else
                                                    nSaldo = Adodc2("SMCANENT") - Adodc2("SMCANSAL")
                                                    nCosPro = Adodc2("SMMNPREUNI")
                                                    
                                                    If Combo3.ListIndex <> 0 Then
                                                       If CDbl(TCamb) <> 0 Then
                                                          nCosPro = nCosPro / TCamb
                                                       End If
                                                    End If
                                                    Set Adodc3 = New ADODB.Recordset
                                                    Adodc3.Open "Select sum(SMCANENT - SMCANSAL)  as saldo From MoresMes Where SMALMA = '" & almacen & "' And SMMESPRO < '" & cAnoMes & "' And SMCODIGO = '" & adodc1("DECODIGO") & "'", cConexCom, adOpenStatic
                                                    nSaldo = 0
                                                    nCosPro = 0
                                                    If Adodc3.RecordCount > 0 Then
                                                            nSaldo = IIf(IsNull(Adodc3("Saldo")), 0, Adodc3("Saldo"))
                                                    End If
                                                    Set Adodc3 = New ADODB.Recordset
                                                    Adodc3.Open "Select  SMMNPREUNI as promedio From MoresMes Where SMALMA = '" & almacen & "' And SMMESPRO < '" & cAnoMes & "' And SMCODIGO = '" & adodc1("DECODIGO") & "'  order by SMMESPRO DESC ", cConexCom, adOpenStatic
                                                     If Adodc3.RecordCount > 0 Then
                                                            nCosPro = IIf(IsNull(Adodc3("promedio")), 0, Adodc3("promedio"))
                                                        If Combo3.ListIndex <> 0 Then
                                                            If CDbl(TCamb) <> 0 Then
                                                               nCosPro = nCosPro / TCamb
                                                            End If
                                                        End If

                                                     End If
                                                    Adodc3.Close
                                                    'CSQL2 = "Insert Into Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
                                            End If
                                    Else
                                            If Val(cAnoMes) - Val(Adodc2("SMMESPRO")) = 1 Then
                                                    nSaldo = Adodc2("SMCANENT") - Adodc2("SMCANSAL")
                                                    nCosPro = Adodc2("SMMNPREUNI")
                                                    If Combo3.ListIndex <> 0 Then
                                                       If CDbl(TCamb) <> 0 Then
                                                          nCosPro = nCosPro / TCamb
                                                       End If
                                                    End If
                                                    
                                                    'CSQL2 = "Insert Into Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
                                            Else
                                                    Set Adodc3 = New ADODB.Recordset
                                                    Adodc3.Open "Select sum(SMCANENT - SMCANSAL)  as saldo From MoresMes Where SMALMA = '" & almacen & "' And SMMESPRO < '" & cAnoMes & "' And SMCODIGO = '" & adodc1("DECODIGO") & "'", cConexCom, adOpenStatic
                                                    nSaldo = 0
                                                    nCosPro = 0
                                                    If Adodc3.RecordCount > 0 Then
                                                            nSaldo = IIf(IsNull(Adodc3("Saldo")), 0, Adodc3("Saldo"))
                                                    End If
                                                    Set Adodc3 = New ADODB.Recordset
                                                    Adodc3.Open "Select  SMMNPREUNI as promedio From MoresMes Where SMALMA = '" & almacen & "' And SMMESPRO < '" & cAnoMes & "' And SMCODIGO = '" & adodc1("DECODIGO") & "'  order by SMMESPRO  DESC ", cConexCom, adOpenStatic
                                                     If Adodc3.RecordCount > 0 Then
                                                            nCosPro = IIf(IsNull(Adodc3("promedio")), 0, Adodc3("promedio"))
                                                     
                                                             If Combo3.ListIndex <> 0 Then
                                                                If CDbl(TCamb) <> 0 Then
                                                                    nCosPro = nCosPro / TCamb
                                                                End If
                                                             End If
                                            End If
                                                     
                                                    Adodc3.Close
                                                    'CSQL2 = "Insert Into Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
                                            End If
                                    End If
                            Else
                                    'CSQL2 = "Insert Into Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
                           End If
                   Else
                            'Aqui se realizohastael mes de proceso
                            Set Adodc3 = New ADODB.Recordset
                            Adodc3.Open "Select sum(SMCANENT - SMCANSAL)  as saldo From MoresMes Where SMALMA = '" & almacen & "'" & _
                                                    " And SMMESPRO < '" & cAnoMes & "' And SMCODIGO = '" & adodc1("DECODIGO") & "'", cConexCom, adOpenStatic
                            nSaldo = 0
                            If Adodc3.RecordCount > 0 Then
                                    nSaldo = IIf(IsNull(Adodc3("Saldo")), 0, Adodc3("Saldo"))
                            End If
                            Adodc3.Close
                            nCosPro = 0
                            'CSQL2 = "Insert Into Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
                   End If
'                   If Trim(CSQL2) <> "" Then
'                        nTra = 1
'                        cConexAux.BeginTrans
'                        cConexAux.Execute CSQL2
'                        cConexAux.CommitTrans
'                        nTra = 0
'                        CSQL2 = ""
'                    End If
                   Adodc2.Filter = ""
                   'cCod = Adodc1("DECODIGO")
                    nCantid = adodc1("DECANTID")
                    If adodc1("DEPRECIO") <> 0 Then
                        
                        nPrecio = adodc1("DEPRECIO")
                        If Combo3.ListIndex <> 0 Then
                           If CDbl(TCamb) <> 0 Then
                              nPrecio = nPrecio / TCamb
                           End If
                        End If
                    Else
                        nPrecio = nCosPro
                    End If
                    

                    If adodc1("CATD") = "NI" Or adodc1("CATD") = "NC" Then
                        If IIf(adodc1("CATD") = "NI" Or adodc1("CATD") = "NC", (nSaldo + adodc1("DECANTID")), (nSaldo - adodc1("DECANTID"))) <> 0 Then
                            If nCosPro <> 0 Then
                               nCosPro = ((nSaldo * nCosPro) + (adodc1("DECANTID") * nPrecio)) / IIf(adodc1("CATD") = "NI" Or adodc1("CATD") = "NC", (nSaldo + adodc1("DECANTID")), (nSaldo - adodc1("DECANTID")))
                               If Combo3.ListIndex <> 0 Then
                                   If CDbl(TCamb) <> 0 Then
                                      nCosPro = nCosPro / TCamb
                                   End If
                               End If
                            Else
                                nCosPro = nPrecio
                            End If
                            

                        Else
                                     nCosPro = ((nSaldo * nCosPro) + (adodc1("DECANTID") * nPrecio)) / 1
                                
                                If Combo3.ListIndex <> 0 Then
                                   If CDbl(TCamb) <> 0 Then
                                      nCosPro = nCosPro / TCamb
                                   End If
                               End If
                        End If
                    End If
                    If adodc1("CATD") = "NI" Or adodc1("CATD") = "NC" Then
                            nSaldo = nSaldo + nCantid
                    Else
                            nSaldo = nSaldo - nCantid
                    End If
                    CSQL2 = "Insert Into Kardex_Val (COD_ART,FEC_DOC,HOR_DOC,COD_MOV,TIP_TRANSA,NUM_DOC,CAN_ART,PRE_UNIT,COS_PRO,SAL_STOCK,SER_LOT,ING_SAL)  "
                    CSQL2 = CSQL2 & " Values ('" & adodc1("DECODIGO") & "',#" & Format(adodc1("CAFECDOC"), "mm/dd/yyyy") & "#,'" & IIf(IsNull(adodc1("CAHORA")) Or Trim(adodc1("CAHORA")) = "", " ", adodc1("CAHORA")) & "','" & adodc1("CACODMOV") & "',"
                    CSQL2 = CSQL2 & "'" & adodc1("CATD") & "','" & adodc1("CANUMDOC") & "', " & adodc1("DECANTID") & ","
                    '*******************************************roberto
                    If nCosPro = 0 And nPrecio = 0 Then Flag0 = 1 '****
                    '**************************************************
                    If adodc1("CATD") = "NI" Or adodc1("CATD") = "NC" Then
                        '*************************************Roberto
                        If nCosPro = 0 Then nCosPro = nPrecio '******
                        '********************************************
                         CSQL2 = CSQL2 & "" & nPrecio & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','I')"
                    Else
                        '*************************************Roberto
                        If nCosPro = 0 Then nCosPro = nPrecio
                        '***************************************
                         CSQL2 = CSQL2 & "" & nCosPro & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','S')"
                    End If
                 
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
    Else
      adodc1.MoveNext
    End If

    If Flag0 = 1 Then
            '*****************************************ROBERTO
            If adodc1("CATD") = "NI" Or adodc1("CATD") = "NC" Then
                        '*************************************Roberto
                        If nPrecio = 0 Then nPrecio = nCosPro
                        '***************************************
                SqlM = "Update kardex_val set PRE_UNIT=" & nPrecio & ",cos_pro=" & nCosPro & "  WHERE COD_ART='" & cCod & "'  and (Pre_unit=0 AND cos_pro=0) and (Tip_Transa='NC' or Tip_Transa='NI') "
            Else
                SqlM = "Update kardex_val set PRE_UNIT=" & nCosPro & ",Cos_Pro=" & nCosPro & " WHERE COD_ART='" & cCod & "'  and (Pre_unit=0 AND cos_pro=0) and (Tip_Transa<>'NC' and Tip_Transa<>'NI') "
            End If
                cConexAux.BeginTrans
                cConexAux.Execute SqlM
                cConexAux.CommitTrans
            '*****************************************
    End If

Loop
 If cCod <> "" Then
    cMesPro = Year(DTPicker1) & Format(Mes1, "00")
                         
    If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
       nTipCam = Val(Devolver_Dato(3, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
    ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
           nTipCam = Val(Devolver_Dato(1, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
    End If
    Rsql1 = "Update  MoResMes  Set  SMMNPREUNI  = " & nCosPro & ", SMUSPREUNI  = " & nCosPro / IIf(nTipCam = 0, 1, nTipCam) & "  where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & cMesPro & "' AND  SMCODIGO= '" & cCod & "'"
    cConexCom.BeginTrans
    cConexCom.Execute Rsql1
    cConexCom.CommitTrans
  End If
End Sub

Private Sub Carga_Almacen()
Dim rSql As String
Dim i As Integer
rSql = "Select TAALMA,TADESCRI FROM TabAlm "
Set rS = New ADODB.Recordset
rS.Open rSql, cConexCom, adOpenStatic
Do While Not rS.EOF
     Combo1.AddItem (rS(1))
     rS.MoveNext
     If rS.EOF Then Exit Do
Loop
rS.MoveFirst
For i = 0 To rS.RecordCount - 1
  If rS(0) = VGAlma Then
    Combo1.ListIndex = i
    Exit For
  Else
    rS.MoveNext
  End If
Next
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Combo2_Click()
Text3 = "": Text4 = "": Option1.Visible = False: Option2.Visible = False
List1.Clear
If Combo2.ListIndex = 0 Then                'ARTICULOS
  Label1.Visible = False: Label2.Visible = False
  Text3.Visible = False: Text4.Visible = False
  Text1.Visible = False: Text2.Visible = False
  Option1.Visible = False: Option2.Visible = False
  'Option1.Top = 2610: Option2.Top = 2610
  List1.Visible = False
ElseIf Combo2.ListIndex = 1 Then            'GRUPO
  Label1.Visible = True: Label2.Visible = True
  Text3.Visible = True: Text4.Visible = True
  Text1.Visible = False: Text2.Visible = False
  List1.Visible = True
ElseIf Combo2.ListIndex = 2 Then             'LINEA
  Label1.Visible = True: Label2.Visible = False
  Text3.Visible = True: Text4.Visible = False
  Text1.Visible = False: Text2.Visible = False
  List1.Visible = True
ElseIf Combo2.ListIndex = 3 Then              'FAMILIA
  Text1 = "": Text2 = ""
  Label1.Visible = False: Label2.Visible = False
  Text3.Visible = False: Text4.Visible = False
  Text1.Visible = False: Text2.Visible = False
  Option1.Visible = True: Option2.Visible = True
  Option1.Top = 1400: Option2.Top = 1400
  List1.Visible = True
  Cargalist
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub


Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim rSql As String
central Me
Carga_Almacen
Combo2.ListIndex = 0
If Combo1.ListIndex = 0 Then VGForm1 = 6
DTPicker1.Value = Date
ADOConectar
End Sub







Private Sub ADOConectar()
cRt = App.Path & "\BdAuxCom.Mdb"
Set cConexAux = New ADODB.Connection
cConexAux.CursorLocation = adUseClient
cConexAux.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & cRt & ";"
cConexAux.Open
End Sub

Private Sub TxTransa_DblClick(Index As Integer)
  Dim Adodc3 As ADODB.Recordset
'        If Combo3.ListIndex <> 0 Then
'           tipo = "S"
'        Else
'           tipo = "I"
'        End If
        Set Adodc3 = New ADODB.Recordset
        Adodc3.Open "SELECT TT_CODMOV,TT_DESCRI FROM Tabtransa where  TT_tipmov = '" & Mid(Combo4.text, 1, 1) & "'", cConexCom, adOpenStatic, adLockOptimistic
        frmReferencia.conectar Adodc3, "SELECT TT_CODMOV,TT_DESCRI FROM Tabtransa where  TT_tipmov = '" & Mid(Combo4.text, 1, 1) & "'"
        frmReferencia.Label1.Caption = "Transacciones"
        frmReferencia.Show vbModal
        Adodc3.Close
        If vGUtil(1) <> "" Then
                TxTransa(Index) = vGUtil(1)
                lbltrans(Index) = Mid(vGUtil(2), 1, 25)
        End If
        If TxTransa(Index).text <> "" Then Call TxTransa_KeyPress(Index, 13)


End Sub

Private Sub TxTransa_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
 SendKeys "{TAB}"
End If
End Sub

Private Sub Carga2()
Dim TCamb As Double
Dim Li As Integer
Dim cCod As String, Codi As String
Dim Aux, cadena As String
Dim cAnoMes As String
Dim cSql1 As String, CSQL2 As String
Dim Dia1 As Integer, Mes1 As Integer
Dim nSaldo As Double, nCosPro As Double
Dim nPrecio As Double, nCantid As Double
Dim cMesPro As String
Dim Rsql1 As String
Dim nTipCam As String
'**********Roberto
Dim rSAX As New ADODB.Recordset
Dim rsMespro As New ADODB.Recordset
Dim SqlM As String
Dim Flag0 As Integer
Dim VALMOV, VALANTE As Double
'*************************
Mes1 = Month(DTPicker1)
cAnoMes = Year(DTPicker1) & Format(Mes1, "00")
cCod = "": Codi = "": Li = 0
Flag0 = 0


Do While Not adodc1.EOF
    
    If Combo2.ListIndex = 1 Then Codi = IIf(IsNull(adodc1("Agrupo")), "", adodc1("Agrupo")): Li = 8 'captura el dato para la comparacion
    If Combo2.ListIndex = 2 Then Codi = IIf(IsNull(adodc1("Amodelo")), "", adodc1("Amodelo")): Li = 4
    If Combo2.ListIndex = 3 Then Codi = IIf(IsNull(adodc1("Afamilia")), "", adodc1("Afamilia")): Li = 4
    
       If adodc1("detipcam") <> 0 And adodc1("cacodmon") = "ME" Then
          TCamb = adodc1("detipcam")
       Else
           If Combo3.ListIndex = 0 Then
              TCamb = 1
           Else
              If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
                  TCamb = Val(Devolver_Dato(3, adodc1("cafecdoc"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
              ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
                 TCamb = Val(Devolver_Dato(1, adodc1("cafecdoc"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
              End If
           End If
      End If
    
    
    
       If (adodc1("CATD") = "GS" And adodc1("CACODMOV") = "GF" And adodc1("CASITGUI") = "F") Then '(adodc1("CATD") = "GS" And adodc1("CACODMOV") = "GV" And adodc1("CASITGUI") = "F") Or
            adodc1.MoveNext
       Else
          If cCod <> adodc1("DECODIGO") Then
             nPrecio = 0: nCantid = 0: nSaldo = 0: nCosPro = 0

             If Not Adodc2.EOF Then
                Adodc2.MoveFirst
                Adodc2.Filter = "SMCODIGO = '" & adodc1("DECODIGO") & "' AND SMMESPRO ='" & AnioMesAnterior(cAnoMes) & "'"
             End If
             
             If Not Adodc2.EOF Then
                nSaldo = Adodc2!SMSALDOINI + (Adodc2!SMCANENT - Adodc2!SMCANSAL)
                nCosPro = Adodc2("SMMNPREUNI")
             Else
                nSaldo = 0: nCosPro = 0
             End If

'             If Not Adodc2.EOF Then
'                Adodc2.MoveFirst
'                Adodc2.Filter = "SMCODIGO = '" & adodc1("DECODIGO") & "' AND SMMESPRO <'" & cAnoMes & "'"
'             End If
'
'             While Not Adodc2.EOF
'                   nSaldo = nSaldo + Adodc2!SMCANENT
'                   nSaldo = nSaldo - Adodc2!SMCANSAL
'                   nCosPro = Adodc2("SMMNPREUNI")
'                   Adodc2.MoveNext
'             Wend
             VALANTE = nCosPro * nSaldo
             Adodc2.Filter = ""
          End If
          
  
          nCantid = adodc1("DECANTID")
              
         'If Combo3.ListIndex <> 0 Then
         '    If adodc1("DEPRECIO") <> 0 Then
         '        If CDbl(TCamb) <> 0 Then
         '           nPrecio = CDbl(adodc1("DEPRECIO") / TCamb)
         '        Else
         '            nPrecio = adodc1("DEPRECIO")
         '        End If
         '    Else
         '        nPrecio = CDbl(nCosPro)
         '    End If
         'Else
         'If adodc1("DEPRECIO") <> 0 Then
                                  
           nPrecio = adodc1("DEPRECIO")
          
          If adodc1("CATIPMOV") = "I" Then
             nSaldo = nSaldo + nCantid
             VALMOV = nCantid * nPrecio
          Else
             nSaldo = nSaldo - nCantid
             VALMOV = nCantid * nCosPro
          End If
      
          
'****************************************
          If adodc1("CATIPMOV") = "I" Then
                If nSaldo <> 0 Then
                   nCosPro = (VALMOV + VALANTE) / nSaldo
                End If
          End If
'****************************************
          VALANTE = nCosPro * nSaldo
                
          CSQL2 = "Insert Into Kardex_Val (COD_ART,FEC_DOC,HOR_DOC,COD_MOV,TIP_TRANSA,NUM_DOC,CAN_ART,PRE_UNIT,COS_PRO,SAL_STOCK,SER_LOT,ING_SAL)  "
          CSQL2 = CSQL2 & " Values ('" & adodc1("DECODIGO") & "',#" & Format(adodc1("CAFECDOC"), "mm/dd/yyyy") & "#,'" & adodc1("CAHORA") & "','" & adodc1("CACODMOV") & "',"
          CSQL2 = CSQL2 & "'" & adodc1("CATD") & "','" & adodc1("CANUMDOC") & "', " & adodc1("DECANTID") & ","
                              
          If adodc1("CATIPMOV") = "I" Then
              CSQL2 = CSQL2 & "" & nPrecio & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','I')"
          Else
              CSQL2 = CSQL2 & "" & nCosPro & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','S')"
          End If
          nTra = 1
          cConexAux.BeginTrans
          cConexAux.Execute CSQL2
          cConexAux.CommitTrans
          nTra = 0
         
         cCod = adodc1("DECODIGO")
             
         adodc1.MoveNext
       End If
               
       If adodc1.EOF Then Exit Do

Loop

End Sub


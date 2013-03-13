VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormKardexValTXDocumento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Kardex Valorizado por Movimiento"
   ClientHeight    =   4605
   ClientLeft      =   1470
   ClientTop       =   1590
   ClientWidth     =   5355
   Icon            =   "FormKardexValTXDocumento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FormKardexValTXDocumento.frx":0442
   ScaleHeight     =   4605
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   576
      Top             =   3888
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   690
      Left            =   1740
      Picture         =   "FormKardexValTXDocumento.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3672
      Width           =   705
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   708
      Left            =   2616
      Picture         =   "FormKardexValTXDocumento.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3672
      Width           =   705
   End
   Begin VB.Frame Frame1 
      Height          =   4464
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   5220
      Begin VB.TextBox TxTransa 
         Height          =   285
         Index           =   0
         Left            =   1704
         MaxLength       =   2
         TabIndex        =   4
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox TxTransa 
         Height          =   285
         Index           =   1
         Left            =   1704
         MaxLength       =   2
         TabIndex        =   5
         Top             =   2556
         Width           =   495
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "FormKardexValTXDocumento.frx":1108
         Left            =   1710
         List            =   "FormKardexValTXDocumento.frx":1115
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1620
         Width           =   2790
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FormKardexValTXDocumento.frx":1133
         Left            =   1710
         List            =   "FormKardexValTXDocumento.frx":113D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   2790
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1710
         TabIndex        =   1
         Top             =   795
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MMMM 'del' yyyy"
         Format          =   52887555
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormKardexValTXDocumento.frx":1151
         Left            =   1710
         List            =   "FormKardexValTXDocumento.frx":1153
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   390
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
         Height          =   252
         Index           =   0
         Left            =   2304
         TabIndex        =   17
         Top             =   2196
         Width           =   2832
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
         Height          =   252
         Index           =   1
         Left            =   2340
         TabIndex        =   16
         Top             =   2580
         Width           =   2796
      End
      Begin VB.Label Label4 
         Caption         =   "Al  Mvmto."
         Height          =   252
         Left            =   444
         TabIndex        =   15
         Top             =   2628
         Width           =   900
      End
      Begin VB.Label Label8 
         Caption         =   "Del Mvmto."
         Height          =   252
         Left            =   432
         TabIndex        =   14
         Top             =   2220
         Width           =   996
      End
      Begin VB.Label lbread 
         Caption         =   "LECTURA DE DATOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   480
         TabIndex        =   13
         Top             =   3192
         Visible         =   0   'False
         Width           =   4008
      End
      Begin VB.Label Label10 
         Caption         =   "Tip de Mvto:"
         Height          =   255
         Left            =   465
         TabIndex        =   12
         Top             =   1665
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda   :"
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   1275
         Width           =   765
      End
      Begin VB.Label Label9 
         Caption         =   "Mes"
         Height          =   255
         Left            =   510
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   435
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
Dim adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim cRt As String
Dim almacen As String
Dim nTra, nConReg, nTotRec As Integer
Dim tipo As String

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
rs.MoveFirst
rs.Move Combo1.ListIndex
almacen = Format(rs(0), "00")
End Sub

Private Sub CmdAceptar_Click()
On Error GoTo Err
'**************Valida el Ingreso**********
If Left(Combo3.text, 1) = "" Then
   MsgBox "Seleccione el Tipo de Moneda al la que desea el Informe de Valorización", vbInformation, "Faltan Datos"
   Exit Sub
End If

If Left(Combo4.text, 1) = "" Then
   MsgBox "Seleccione un Tipo de Movimiento del Kardex para el Informe de Valorización", vbInformation, "Faltan Datos"
   Exit Sub
End If

'*****************************************************************
'**Procedimiento que se Encarga Cargar el Temporal para el Reporte
'*****************************************************************
   If Carga_RepoVal = -1 Then Exit Sub
'*****************************************************************
  Screen.MousePointer = 11

  CrystalReport1.WindowTitle = "Kardex Valorizado por Documento"
  CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "Rep_Mov_ValXDcmto.rpt"
  
  Dim ccadena As String
  If CrystalReport1.ReportFileName = "" Then
      MsgBox "No hay registros a imprimir", vbInformation, "Aviso"
      Screen.MousePointer = 1
      Exit Sub
  End If
  ccadena = ""
  
  If Mid(Combo4.text, 1, 1) = "T" Then
     If TxTransa(0) <> "" Then
       ccadena = "({al_Kardex_Val.ING_SAL} ='I' AND  "
       ccadena = ccadena & "{al_Kardex_Val.COD_MOV} >='" & TxTransa(0) & "' AND  {al_Kardex_Val.COD_MOV} <='" & TxTransa(1) & "') OR "
       ccadena = ccadena & "({al_Kardex_Val.ING_SAL} ='S' AND  "
       ccadena = ccadena & "{al_Kardex_Val.COD_MOV} >='" & TxTransa(0) & "' AND  {al_Kardex_Val.COD_MOV} <='" & TxTransa(1) & "') "
     End If
  Else
     ccadena = "{al_Kardex_Val.ING_SAL} ='" & Mid(Combo4.text, 1, 1) & "'"
     If Mid(Combo4.text, 1, 1) = "I" Then
        If TxTransa(0) <> "" Then ccadena = ccadena & " AND ({al_Kardex_Val.COD_MOV} >='" & TxTransa(0) & "' AND  {al_Kardex_Val.COD_MOV} <='" & TxTransa(1) & "') "
     Else
        If TxTransa(0) <> "" Then ccadena = ccadena & " AND ({al_Kardex_Val.COD_MOV} >='" & TxTransa(0) & "' AND  {al_Kardex_Val.COD_MOV} <='" & TxTransa(1) & "') "
     End If
  End If
  'Ubi_Tab CrystalReport1
  CrystalReport1.DiscardSavedData = True
  CrystalReport1.Destination = crptToWindow
  CrystalReport1.WindowShowPrintBtn = True
  CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.WindowShowSearchBtn = True
  CrystalReport1.WindowShowPrintSetupBtn = True
  CrystalReport1.ReplaceSelectionFormula (ccadena)
  
  CrystalReport1.formulas(0) = "ALMACEN = '" & UCase(Combo1.text) & "'"
  CrystalReport1.formulas(1) = "Mes= '" & UCase(Format(DTPicker1, "MMMM - yyyy")) & "'"
  CrystalReport1.formulas(2) = "EMPRESA= '" & UCase(VGparametros.RucEmpresa) & "'"
  CrystalReport1.formulas(3) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
  If Combo3.ListIndex <> 0 Then
     CrystalReport1.formulas(4) = "MONEY= 'DOLAR'"
  Else
      CrystalReport1.formulas(4) = "MONEY= 'SOLES'"
  End If
  
  If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
 Screen.MousePointer = 1
 Exit Sub
 
Err:
   MsgBox Err.Description, vbInformation, "Aviso"
   Screen.MousePointer = 1
   
End Sub

Function Carga_RepoVal() As Long
Dim cAnoMes As String, cCod As String
Dim cSql1 As String, CSQL2 As String
On Error GoTo ErrCarga
Set adodc1 = New ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
   
cAnoMes = Year(DTPicker1.Year) & Format(DTPicker1.Month, "00")

nTra = 1
'------------------
'VGcnx.BeginTrans
VGCNx.Execute ("Delete From al_Kardex_Val")
'VGcnx.CommitTrans
'------------
'cSql1 = "Delete From al_Kardex_Val"
'Adodc1.Open cSql1, VGcnx, adOpenStatic
'Adodc1.Close
'------------

nTra = 0
lbread.Caption = "Espere un Momento...!"
Frame1.Refresh

cSql1 = "Select A.*,B.*,C.adescri,D.TT_descri From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC INNER JOIN MaeArt C ON C.ACODIGO = a.DECODIGO INNER JOIN Tabtransa D ON D.tt_codmov=B.cacodmov and D.tt_tipmov=B.catipmov  "
cSql1 = cSql1 & "  Where Month(CAFECDOC) = " & DTPicker1.Month & " and Year(CAFECDOC) = " & DTPicker1.Year & " AND CAALMA = '" & almacen & "' "
cSql1 = cSql1 & "  And CASITGUI<>'A'  AND not (CATD='GS' And CACODMOV='GF' And CASITGUI='F') Order By DECODIGO,CAFECDOC,CAHORA "

adodc1.Open cSql1, VGCNx, adOpenStatic
nConReg = 0: nTotRec = adodc1.RecordCount

If nTotRec = 0 Then
   MsgBox "No hay Información para Procesar en el mes que usted Selecciono....!", vbInformation, "Verifique....!"
   Carga_RepoVal = -1
   Exit Function
End If

Adodc2.Open "Select * From MoresMes Where SMALMA = '" & almacen & "' Order By SMMESPRO", VGCNx, adOpenStatic
lbread.Visible = True
Call LoadIngresoSalida
lbread.Visible = False

Carga_RepoVal = 1
Exit Function
'**************************************
'**************************************

ErrCarga:
        MsgBox Err.Description
        'Resume
        If nTra = 1 Then VGCNx.RollbackTrans
        lbread.Visible = False
End Function


Private Sub Carga_Almacen()
Dim rsql As String
Dim I As Integer
rsql = "Select TAALMA,TADESCRI FROM TabAlm "
Set rs = New ADODB.Recordset
rs.Open rsql, VGCNx, adOpenStatic
Do While Not rs.EOF
     Combo1.AddItem (rs(1))
     rs.MoveNext
     If rs.EOF Then Exit Do
Loop
rs.MoveFirst
For I = 0 To rs.RecordCount - 1
  If rs(0) = VGAlma Then
    Combo1.ListIndex = I
    Exit For
  Else
    rs.MoveNext
  End If
Next
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub



Private Sub Combo4_Click()
       If Left(Combo4.text, 1) = "I" Or Left(Combo4.text, 1) = "S" Then
          TxTransa(0).text = ""
          TxTransa(1).text = ""
          lbltrans(0).Caption = ""
          lbltrans(1).Caption = ""
       End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim rsql As String
central Me
Carga_Almacen
If Combo1.ListIndex = 0 Then VGForm1 = 6
DTPicker1.Value = Date
'ADOConectar
End Sub

Private Sub ADOConectar()
cRt = App.Path & "\BdAuxCom.Mdb"
Set VGCNx = New ADODB.Connection
VGCNx.CursorLocation = adUseClient
VGCNx.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & cRt & ";"
' vgcnx.open
End Sub

Private Sub LoadIngresoSalida()
Dim TCamb As Double
Dim cCod As String
Dim cAnoMes As String
Dim cSql1 As String, CSQL2 As String
Dim nSaldo As Double, nCosPro As Double
Dim nPrecio As Double, nCantid As Double
Dim cMesPro As String
'**********Roberto
Dim VALMOV, VALANTE As Double
On Local Error GoTo ERRAR
'*************************
cAnoMes = Year(DTPicker1) & Format(DTPicker1.Month, "00")
cCod = ""

Do While Not adodc1.EOF
        
  lbread.Caption = "LECTURA DE DATOS  " & nConReg & " - " & nTotRec
  Frame1.Refresh
  nConReg = nConReg + 1
        
       If adodc1("CAtipcam") <> 0 Then
          TCamb = adodc1("CAtipcam")
       Else
          If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
              TCamb = Val(Devolver_Dato(3, adodc1("cafecdoc"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
          ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
              TCamb = Val(Devolver_Dato(1, adodc1("cafecdoc"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
          End If
       End If
          
          
   If cCod <> adodc1("DECODIGO") Then
      nPrecio = 0: nCantid = 0: nSaldo = 0: nCosPro = 0
      '*************************************************************
      '***Busca Saldo Inicial en el Mes Anterior
      '*************************************************************
      Adodc2.Filter = "SMCODIGO = '" & adodc1("DECODIGO") & "' AND SMMESPRO ='" & AnioMesAnterior(cAnoMes) & "'"
      If Not Adodc2.EOF Then
         nSaldo = IIf(IsNull(Adodc2!SMSALDOINI), 0, Adodc2!SMSALDOINI) + (Adodc2!SMCANENT - Adodc2!SMCANSAL)
         nCosPro = IIf(Combo3.ListIndex <> 0, Adodc2("SMUSPREUNI"), Adodc2("SMMNPREUNI"))
      Else
         nSaldo = 0: nCosPro = 0
      End If
      '*************************************************************
      
      VALANTE = nCosPro * nSaldo   'Valorizacion Anteriror
      
      Adodc2.Filter = ""
   End If
   
   nCantid = adodc1("DECANTID")
   '****************************************
   '***Soles y Dolares
   '****************************************
   If Combo3.ListIndex <> 0 Then
      If Round(TCamb, 3) > 0 Then
          If cNull(adodc1("CACODMON")) = "02" Then
             nPrecio = adodc1("DEPRECIO")
          Else
             nPrecio = (adodc1("DEPRECIO") / TCamb)
          End If
      Else
          nPrecio = 0
      End If
   Else
      If cNull(adodc1("CACODMON")) = "02" Then
         nPrecio = adodc1("DEPRECIO") * TCamb
      Else
         nPrecio = adodc1("DEPRECIO")
      End If
   End If
    '****************************************
       
    If adodc1("CATIPMOV") = "I" Then
       nSaldo = nSaldo + nCantid
       VALMOV = nCantid * nPrecio
    Else
       nSaldo = nSaldo - nCantid
       VALMOV = nCantid * nCosPro
    End If

   
   '****************************************
   '***Calculo del Kardex
   '****************************************
    If adodc1("CATIPMOV") = "I" Then
          If nSaldo <> 0 Then
             nCosPro = (VALMOV + VALANTE) / nSaldo
          End If
    End If
    VALANTE = nCosPro * nSaldo
                    

    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,DESCRIPCION,FEC_DOC,HOR_DOC,COD_MOV,TIP_TRANSA,DES_TRANSA,NUM_DOC,CAN_ART,PRE_UNIT,COS_PRO,SAL_STOCK,SER_LOT,ING_SAL)  "
    CSQL2 = CSQL2 & " Values ('" & adodc1("DECODIGO") & "','" & adodc1("ADESCRI") & "','" & adodc1("CAFECDOC") & "','" & adodc1("CAHORA") & "','" & adodc1("CACODMOV") & "',"
    CSQL2 = CSQL2 & "'" & adodc1("CATD") & "','" & adodc1("CATD") & "','" & adodc1("CANUMDOC") & "', " & adodc1("DECANTID") & ","
    If adodc1("CATIPMOV") = "I" Then
        CSQL2 = CSQL2 & "" & nPrecio & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','I')"
    Else
        CSQL2 = CSQL2 & "" & nCosPro & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','S')"
    End If
    nTra = 1
    'VGcnx.BeginTrans
    VGCNx.Execute CSQL2
    'VGcnx.CommitTrans
    nTra = 0
   
    cCod = adodc1("DECODIGO")
       
    adodc1.MoveNext
        
                
        If adodc1.EOF Then Exit Do

Loop
  
lbread.Caption = "LECTURA DE DATOS  " & nConReg & " - " & nTotRec
Frame1.Refresh

Exit Sub
ERRAR:
     MsgBox Err.Description
     'Resume
     lbread.Visible = False
     
End Sub

Private Sub TxTransa_DblClick(Index As Integer)
  Dim Adodc3 As ADODB.Recordset
'        If Combo3.ListIndex <> 0 Then
'           tipo = "S"
'        Else
'           tipo = "I"
'        End If
        Set Adodc3 = New ADODB.Recordset
        Adodc3.Open "SELECT TT_CODMOV,TT_DESCRI FROM Tabtransa where  TT_tipmov = '" & Mid(Combo4.text, 1, 1) & "'", VGCNx, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc3, "SELECT TT_CODMOV,TT_DESCRI FROM Tabtransa where  TT_tipmov = '" & Mid(Combo4.text, 1, 1) & "'"
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

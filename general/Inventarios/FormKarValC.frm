VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormKarValC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kardex Valorizado - Centro Costos"
   ClientHeight    =   4365
   ClientLeft      =   1380
   ClientTop       =   2520
   ClientWidth     =   5445
   ControlBox      =   0   'False
   Icon            =   "FormKarValC.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5445
   Begin VB.Frame Frame1 
      Height          =   3165
      Left            =   180
      TabIndex        =   6
      Top             =   60
      Width           =   4995
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FormKarValC.frx":08CA
         Left            =   1590
         List            =   "FormKarValC.frx":08D4
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2625
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormKarValC.frx":08E8
         Left            =   1560
         List            =   "FormKarValC.frx":08EA
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   2760
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   2055
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   1530
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1545
         TabIndex        =   11
         Top             =   960
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MMMM 'del' yyyy"
         Format          =   53280771
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   555
         TabIndex        =   13
         Top             =   2670
         Width           =   1035
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   435
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Al"
         Height          =   255
         Left            =   975
         TabIndex        =   9
         Top             =   2055
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "Del"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Mes"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   1035
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3000
      Picture         =   "FormKarValC.frx":08EC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3525
      Width           =   735
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   1260
      Picture         =   "FormKarValC.frx":0D2E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3525
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   195
      Left            =   3330
      TabIndex        =   5
      Top             =   3615
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   344
      _Version        =   393216
      Format          =   53280769
      CurrentDate     =   36710
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   30
      Top             =   3375
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FormKarValC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'la mejora de kardex valorizado por articulo mejorarlo en centro de costos

Option Explicit
Dim cConexAux As ADODB.Connection
Dim adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim Adodc3 As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim cRt As String
Dim almacen As String
Dim nTra As Integer
Dim Adodc4 As New ADODB.Recordset

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex >= 0 Then
 rs.MoveFirst
 rs.Move Combo1.ListIndex
 almacen = Format(rs(0), "00")
End If
End Sub

Private Sub CmdAceptar_Click()
Dim rsql As String
Dim Va1 As String, Va2 As String
On Error GoTo err
Screen.MousePointer = 11
Set Adodc3 = New ADODB.Recordset  'Para sacar la descripcion del rango elegido

If Text1 <> "" And Text2 <> "" Then
 
  Set Adodc3 = New ADODB.Recordset
  Adodc3.Open "Select CENtrocostoDESCRIPCION from ct_CENTROCOSTO Where CENtroCOSToCODIGO='" & Text1 & "'", VGcnxCT, adOpenStatic
  If Adodc3.RecordCount > 0 Then
     Va1 = Adodc3("CENtroCOSTodESCRIPCION")
 Else
     MsgBox "El codigo:" & Text1 & "   de Centro de Costo No existe", vbExclamation, "Error"
     Screen.MousePointer = 1
    Exit Sub
  End If
  Adodc3.Close
  
  Set Adodc3 = New ADODB.Recordset
  Adodc3.Open "Select CENtrocostoDESCRIPCION from ct_CENTROCOSTO Where CENtroCOSToCODIGO='" & Text2 & "'", VGcnxCT, adOpenStatic
  If Adodc3.RecordCount > 0 Then
     Va2 = Adodc3("CENtroCOSToDESCRIPCION")
  Else
      MsgBox "El codigo:" & Text2 & "   de Centro de Costo No existe", vbExclamation, "Error"
      Screen.MousePointer = 1
     Exit Sub
  End If
  Adodc3.Close
  CrystalReport1.Reset
  CrystalReport1.WindowTitle = "Inv022 -- Control de Inventarios"
  CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "al_kardexCentroCostoDetallado.rpt"
  Carga_RepoVal
  
  CrystalReport1.DiscardSavedData = True
  CrystalReport1.Destination = crptToWindow
  CrystalReport1.WindowShowPrintBtn = True
  CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.WindowShowSearchBtn = True
  CrystalReport1.WindowShowPrintSetupBtn = True
  CrystalReport1.formulas(0) = "ALMACEN = '" & UCase(Combo1.text) & "'"
  CrystalReport1.formulas(1) = "Mes= '" & UCase(Format(DTPicker1, "MMMM - yyyy")) & "'"
  CrystalReport1.formulas(2) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
  CrystalReport1.formulas(3) = "ART1= '" & Text1 & "'"
  CrystalReport1.formulas(4) = "ART2 ='" & Text2 & "'"
  CrystalReport1.formulas(5) = "des1= '" & Va1 & "'"
  CrystalReport1.formulas(6) = "des2 ='" & Va2 & "'"
  
  CrystalReport1.StoredProcParam(0) = VGCNx.DefaultDatabase
  CrystalReport1.StoredProcParam(1) = VGcnxCT.DefaultDatabase

  If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End If

Screen.MousePointer = 1
Exit Sub
err:

    MsgBox err.Description, vbInformation, "Aviso"
    Screen.MousePointer = 1
End Sub

Private Sub Carga_RepoVal()
Dim rAdo As ADODB.Recordset
Dim Aux, CADENA As String
Dim cAnoMes As String, cCod As String
Dim cSql1 As String, CSQL2 As String
Dim Dia1 As Integer, Mes1 As Integer
Dim nSaldo As Double, nCosPro As Double
Dim nPrecio As Double, nCantid As Double
Dim I As Integer
Dim Codi As String

On Error GoTo ErrCarga
Set adodc1 = New ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
Set rAdo = New ADODB.Recordset
    
Mes1 = Month(DTPicker1)
cAnoMes = Year(DTPicker1) & Format(Mes1, "00")

nTra = 1
VGCNx.BeginTrans
VGCNx.Execute "Delete From al_Kardex_CC"
VGCNx.CommitTrans
nTra = 0

cSql1 = "Select A.*,B.* From"
cSql1 = cSql1 & " MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD And B.CANUMDOC = A.DENUMDOC"
cSql1 = cSql1 & " Where Month(CAFECDOC) = " & Mes1 & " and Year(CAFECDOC) = " & Year(DTPicker1) & " AND CAALMA = '" & almacen & "'"
cSql1 = cSql1 & " AND CASITGUI<>'A' "
cSql1 = cSql1 & " Order By dealma,DECODIGO,CAFECDOC,catipmov,CAHORA"

adodc1.Open cSql1, VGCNx, adOpenStatic
Adodc2.Open "Select * From MoresMes Where SMALMA = '" & almacen & "'  and SMMNPREUNI <> 0 " & _
                       " Order By SMMESPRO", VGCNx, adOpenStatic

 If Adodc2.RecordCount > 0 Then
        If Val(cAnoMes) - Val(Adodc2("SMMESPRO")) > 1 Then
                'MsgBox "El Costo utilizado es del Mes de " & DesMes(Right(Adodc2("SMESPRO"), 2)) & Chr(13) & ", porque no se ha hecho el Cierre en los meses anteriores", vbInformation, "Información"
        End If
Else
        MsgBox "No se ha hecho Cierre en los meses anteriores y su Costo Inicial será Cero", vbInformation, "Información"
End If

cCod = ""
If adodc1.RecordCount > 0 Then
    adodc1.MoveFirst
    Carga
End If

Exit Sub
ErrCarga:
        MsgBox err.Description
        If nTra = 1 Then VGCNx.RollbackTrans
        Exit Sub
        Resume
End Sub

Private Sub Carga()
Dim Li As Integer
Dim cCod As String, Codi As String
Dim Aux, CADENA As String
Dim cAnoMes As String
Dim cSql1 As String, CSQL2 As String
Dim Dia1 As Integer, Mes1 As Integer
Dim nSaldo As Double, nCosPro As Double
Dim nPrecio As Double, nCantid As Double
On Error GoTo errores

cCod = "": Codi = "": Li = 0
Do While Not adodc1.EOF
       If adodc1("DECODIGO") = "07.05.02" Then
         MsgBox "hola"
       End If
       If (adodc1("CACODMOV") = "GV" And adodc1("CASITGUI") = "F") Then
            adodc1.MoveNext
       Else
            If Trim(cCod) = Trim(adodc1("DECODIGO")) Then
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
                    
                    CSQL2 = "Insert Into AL_KarDEX_CC (COD_ART,FEC_DOC,HOR_DOC,COD_MOV,TIP_TRANSA,NUM_DOC,CAN_ART,CENTROCOSTOCODIGO,PRE_UNIT,COS_PRO,SAL_STOCK)  "
                    CSQL2 = CSQL2 & " Values ('" & Trim(adodc1("DECODIGO")) & "','" & Format(adodc1("CAFECDOC"), "dd/mm/yyyy") & "','" & adodc1("CAHORA") & "','" & adodc1("CACODMOV") & "',"
                    CSQL2 = CSQL2 & "'" & adodc1("CATD") & "','" & adodc1("CANUMDOC") & "', " & adodc1("DECANTID") & ",'" & adodc1("CACENCOS") & "',"
                    If adodc1("CATD") = "NI" Then
                        CSQL2 = CSQL2 & "" & nPrecio & "," & nCosPro & "," & nSaldo & ")"
                    Else
                        CSQL2 = CSQL2 & "" & nCosPro & "," & nCosPro & "," & nSaldo & ")"
                    End If
                    
                    nTra = 1
                    'cConexAux.BeginTrans
                    VGCNx.Execute CSQL2
                    'cConexAux.CommitTrans
                    nTra = 0
                   
                   cCod = adodc1("DECODIGO")
            Else
                    nSaldo = 0: nCosPro = 0
                    If Adodc2.RecordCount > 0 Then
                            Adodc2.MoveFirst
                            Adodc2.Filter = "SMCODIGO = '" & adodc1("DECODIGO") & "'"
                            If Not Adodc2.EOF Then
                                    Adodc2.MoveLast
                                    If Adodc2("SMMESPRO") = cAnoMes Then
                                            Adodc2.MovePrevious
                                            If Adodc2.BOF Then
                                                    CSQL2 = "Insert Into KarVal_CC (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC,KCENCOS) Values ('" & adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL','" & adodc1("CACENCOS") & "')"
                                                    CSQL2 = ""
                                            Else
                                                    nSaldo = Adodc2("SMCANENT") - Adodc2("SMCANSAL")
                                                    nCosPro = Adodc2("SMMNPREUNI")
                                                    CSQL2 = "Insert Into al_KarVal_CC (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC,KCENCOS) Values ('" & adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL','" & adodc1("CACENCOS") & "')"
                                                    CSQL2 = ""
                                            End If
                                    Else
                                            If Val(cAnoMes) - Val(Adodc2("SMMESPRO")) = 1 Then
                                                    nSaldo = Adodc2("SMCANENT") - Adodc2("SMCANSAL")
                                                    nCosPro = IIf(IsNull(Adodc2("SMMNPREUNI")), 0, Adodc2("SMMNPREUNI"))
                                                    'cSql2 = "Insert Into KarVal_CC (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC,KCENCOS) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL','" & Adodc1("CACENCOS") & "')"
                                                    CSQL2 = ""
                                            Else
                                                    Set Adodc3 = New ADODB.Recordset
                                                    Adodc3.Open "Select sum(SMCANENT - SMCANSAL)  as saldo From MoresMes Where SMALMA = '" & almacen & "' and SMMESPRO >= '" & Adodc2("SMMESPRO") & _
                                                                            "' And SMMESPRO < '" & cAnoMes & "' And SMCODIGO = '" & adodc1("DECODIGO") & "'", VGCNx, adOpenStatic
                                                    nSaldo = 0
                                                    If Adodc3.RecordCount > 0 Then
                                                            nSaldo = IIf(IsNull(Adodc3("Saldo")), 0, Adodc3("Saldo"))
                                                    End If
                                                    Adodc3.Close
                                                    nCosPro = IIf(IsNull(Adodc2("SMMNPREUNI")), 0, Adodc2("SMMNPREUNI"))
                                                    CSQL2 = ""
                                                    'cSql2 = "Insert Into KarVal_CC (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC,KCENCOS) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL','" & Adodc1("CACENCOS") & "')"
                                            End If
                                    End If
                            Else
                                   CSQL2 = ""
                                    'cSql2 = "Insert Into KarVal_CC (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC,KCENCOS) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL','" & Adodc1("CACENCOS") & "')"
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
                            CSQL2 = ""
                            'cSql2 = "Insert Into KarVal_CC (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC,KCENCOS) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL','" & Adodc1("CACENCOS") & "')"
                   End If
                   If Trim(CSQL2) <> "" Then
                        nTra = 1
                        'cConexAux.BeginTrans
                        cConexAux.Execute CSQL2
                        'cConexAux.CommitTrans
                        nTra = 0
                        CSQL2 = ""
                    End If
                   Adodc2.Filter = ""
                   'cCod = Adodc1("DECODIGO")
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
                 
                   CSQL2 = "Insert Into al_Kardex_CC (COD_ART,FEC_DOC,HOR_DOC,COD_MOV,TIP_TRANSA,NUM_DOC,CAN_ART,CENTROCOSTOCODIGO,PRE_UNIT,COS_PRO,SAL_STOCK)  "
                   CSQL2 = CSQL2 & " Values ('" & Trim(adodc1("DECODIGO")) & "','" & Format(adodc1("CAFECDOC"), "dd/mm/yyyy") & "','" & IIf(IsNull(adodc1("CAHORA")) Or Trim(adodc1("CAHORA")) = "", " ", adodc1("CAHORA")) & "','" & adodc1("CACODMOV") & "',"
                   CSQL2 = CSQL2 & "'" & adodc1("CATD") & "','" & adodc1("CANUMDOC") & "', " & adodc1("DECANTID") & ",'" & adodc1("CACENCOS") & "',"
                   If adodc1("CATD") = "NI" Then
                        CSQL2 = CSQL2 & "" & nPrecio & "," & nCosPro & "," & nSaldo & ")"
                   Else
                        CSQL2 = CSQL2 & "" & nCosPro & "," & nCosPro & "," & nSaldo & ")"
                   End If
                 
                   nTra = 1
                   'cConexAux.BeginTrans
                   VGCNx.Execute CSQL2
                   'cConexAux.CommitTrans
                   nTra = 0
                
                  cCod = Trim(adodc1("DECODIGO"))
            End If
            adodc1.MoveNext
        End If
        If adodc1.EOF Then Exit Do
Loop
Exit Sub
errores:
        MsgBox err.Description
        If nTra = 1 Then VGCNx.RollbackTrans
        Exit Sub
        Resume

End Sub

Private Sub Carga_Almacen()
Dim rsql As String
rsql = "Select  TAALMA,TADESCRI FROM TabAlm "
Set rs = New ADODB.Recordset
rs.Open rsql, VGCNx, adOpenStatic
Do While Not rs.EOF
     Combo1.AddItem (rs(1))
     rs.MoveNext
     If rs.EOF Then Exit Do
Loop
Combo1.ListIndex = VGAlma - 1
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{Tab}"
End If
End Sub

Private Sub Form_Load()
central Me
Carga_Almacen
VGForm1 = 6
DTPicker1.Value = Date
Combo3.ListIndex = 0
End Sub

Private Sub Text1_DblClick()
Dim Adodc2 As Recordset
Set Adodc3 = New ADODB.Recordset

Adodc3.Open "SELECT centrocostocodigo,centrocostodescripcion FROM ct_centrocosto where  centrocostotipo = 3", VGCNx, adOpenStatic, adLockOptimistic

frmReferencia.Conectar Adodc3, "SELECT centrocostocodigo,centrocostodescripcion FROM ct_centrocosto where  centrocostotipo = 3"
frmReferencia.Label1.Caption = "Centro de Costos"
frmReferencia.Show vbModal
Adodc3.Close
If vGUtil(1) <> "" Then
        Text1 = (vGUtil(1))
End If

If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
      MsgBox "Ingrese un Código menor al Fin ", vbOKOnly, "Error"
      Exit Sub
      
End If
If Text1 <> "" Then
        Text2.Enabled = True
        Text2.SetFocus
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text1_DblClick
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1 <> "" Then
      If Existe(3, Text1, "ct_CENTROCOSTO", "CENtrocostoCODIGO", False) = False Then
               MsgBox "El Código no existe", vbInformation, "Información"
              Text2.Enabled = True
              Text2.SetFocus
      Else
            SendKeys "{Tab}"
      End If
  End If
End Sub

Private Sub Text2_DblClick()
Set Adodc2 = New ADODB.Recordset

VGForm1 = 6
Adodc3.Open "SELECT centrocostocodigo,centrocostodescripcion FROM ct_centrocosto where  centrocostotipo = 3", VGCNx, adOpenStatic, adLockOptimistic

frmReferencia.Conectar Adodc3, "SELECT centrocostocodigo,centrocostodescripcion FROM ct_centrocosto where  centrocostotipo = 3"
frmReferencia.Label1.Caption = "Centro de Costos"
frmReferencia.Show vbModal
Adodc3.Close
If vGUtil(1) <> "" Then
        Text2 = (vGUtil(1))
End If
If Trim(Text2) <> "" Then
        CmdAceptar.SetFocus
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text2_DblClick
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Trim(Text2) <> "" Then

     Text2 = Trim(Text2)
      If Existe(3, Text2, "ct_CENTROCOSTO", "CENtrocostocodigo", False) = False Then
            If Text1 > Text2 Then
                   MsgBox "El Código Fin debe ser Mayor que el Inicio", vbExclamation, "Aviso"
                   Exit Sub
            End If
            CmdAceptar.SetFocus
      Else
        SendKeys "{Tab}"
      End If
End If
End Sub


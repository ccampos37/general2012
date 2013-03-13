VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmKardexValTransaccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kardex Valorizado Por Transaccion"
   ClientHeight    =   5205
   ClientLeft      =   1470
   ClientTop       =   1590
   ClientWidth     =   8250
   ControlBox      =   0   'False
   Icon            =   "FrmKardexValTransaccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8250
   Begin VB.Frame Framemov 
      Caption         =   "Tipo de Movimiento"
      Height          =   735
      Left            =   360
      TabIndex        =   24
      Top             =   4320
      Width           =   4695
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuTipoMovimiento 
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   500
         NomTabla        =   "tabtransa"
         ListaCampos     =   "TT_CODMOV(1),TT_DESCRI(1)"
         XcodCampo       =   "TT_CODMOV"
         XListCampo      =   "TT_DESCRI"
         ListaCamposDescrip=   "codigo,,descripcion"
         ListaCamposText =   "TT_CODMOV,TT_DESCRI"
         Requerido       =   0   'False
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1215
      Left            =   5280
      TabIndex        =   17
      Top             =   3840
      Width           =   2655
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   690
         Left            =   120
         Picture         =   "FrmKardexValTransaccion.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   690
         Left            =   1440
         Picture         =   "FrmKardexValTransaccion.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ordenado Por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   360
      TabIndex        =   16
      Top             =   2280
      Width           =   4695
      Begin VB.OptionButton Option1 
         Caption         =   "Familia"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Articulo"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuFamilia 
         Height          =   375
         Left            =   135
         TabIndex        =   18
         Top             =   1440
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "familia"
         ListaCampos     =   "FAM_CODIGO(1),FAM_NOMBRE(1)"
         XcodCampo       =   "FAM_CODIGO"
         XListCampo      =   "FAM_NOMBRE"
         ListaCamposDescrip=   "ODIGO,descripcion"
         ListaCamposText =   "FAM_CODIGO,FAM_NOMBRE"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuarticulo 
         Height          =   375
         Left            =   135
         TabIndex        =   19
         Top             =   720
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "maeart"
         ListaCampos     =   "acodigo(1),adescri(1)"
         XcodCampo       =   "acodigo"
         XListCampo      =   "adescri"
         ListaCamposDescrip=   "ODIGO,descripcion"
         ListaCamposText =   "acodigo,adescri"
         Requerido       =   0   'False
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Impresion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   5160
      TabIndex        =   12
      Top             =   2400
      Width           =   2775
      Begin VB.OptionButton Option2 
         Caption         =   "Resumen"
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Articulo"
         Height          =   495
         Index           =   2
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Detallado"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   7575
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Todos los almacenes"
         Height          =   375
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "FrmKardexValTransaccion.frx":114E
         Left            =   1395
         List            =   "FrmKardexValTransaccion.frx":115B
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1230
         Width           =   2100
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FrmKardexValTransaccion.frx":1179
         Left            =   1395
         List            =   "FrmKardexValTransaccion.frx":1183
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   690
         Width           =   2160
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmKardexValTransaccion.frx":1197
         Left            =   4755
         List            =   "FrmKardexValTransaccion.frx":1199
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   250
         Width           =   2520
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   4755
         TabIndex        =   7
         Top             =   735
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MMMM 'del' yyyy"
         Format          =   51773443
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.Label Label9 
         Caption         =   "Mes"
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Tip de Mvto:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1275
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda   :"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   765
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7875
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   2235
         Width           =   3225
      End
   End
End
Attribute VB_Name = "FrmKardexValTransaccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim Adodc3 As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim cRt As String
Dim almacen As String
Dim almacenAnt  As String
Dim nTra As Integer
Dim Adodc4 As ADODB.Recordset
Dim Adostk As ADODB.Recordset
Dim AdoMes As ADODB.Recordset
Dim Total As Double
Dim tipo As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Combo1.Enabled = False
 Else
   Combo1.Enabled = True
End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Combo1_Click()
rs.MoveFirst
rs.Move Combo1.ListIndex
almacen = Format(rs(0), "00")
End Sub

Private Sub cmdAceptar_Click()
Dim rsql As String
Dim SQL As New ADODB.Recordset
Dim Va1 As String, Va2 As String
Dim Reporte As String
Dim aparam(2) As Variant
Dim aform(4) As Variant
Dim titulo As String
Dim filtro As String
Screen.MousePointer = 11
On Error GoTo Err
Set Adodc3 = New ADODB.Recordset  'Para sacar la descripcion del rango elegido
If Option2(1).Value = True Then
     Reporte = "al_kardexTransaccionDetallado.rpt"
ElseIf Option2(2).Value = True Then
     Reporte = "al_kardexTransaccionxarticulos.rpt"
  Else
     Reporte = "al_kardexTransaccionresumen.rpt"
End If

If Ctr_Ayuarticulo.xclave <> "" Then filtro = " and decodigo='" & Ctr_Ayuarticulo.xclave & "'"
If Ctr_AyuFamilia.xclave <> "" Then filtro = filtro & " and maeart.afamilia='" & Ctr_AyuFamilia.xclave & "'"
If Ctr_AyuTipoMovimiento.xclave <> "" Then filtro = " and a.cacodmov='" & Ctr_AyuTipoMovimiento.xclave & "'"

VGCNx.Execute "Delete From al_Kardex_Val"

If Check1.Value = 1 Then
   Set SQL = VGCNx.Execute("select * from tabalm where almacenvalorizado='1' ")
   SQL.MoveFirst
   Do Until SQL.EOF
      almacen = Format(SQL(0), "00")
      Carga_RepoVal (filtro)
      SQL.MoveNext
      aform(0) = "ALMACEN ='TODOS'"
   Loop
 Else
   Carga_RepoVal (filtro)
   aform(0) = "ALMACEN = '" & UCase(Combo1.text) & "'"
End If
aform(1) = "Mes= '" & UCase(Format(DTPicker1, "MMMM - yyyy")) & "'"
aform(2) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
If Combo3.ListIndex <> 0 Then
   aform(3) = "MONEY= 'DOLAR'"
 Else
   aform(3) = "MONEY= 'SOLES'"
End If
If Option2(3).Value = True Then
     aparam(0) = VGCNx.DefaultDatabase
     aparam(1) = VGparametros.empresacodigo
     aparam(2) = Format(DTPicker1, "yyyy") + Format(DTPicker1, "mm")
     aparam(3) = "1"
  Else
  aparam(0) = VGCNx.DefaultDatabase
  If Option1(1) = True Then
     aparam(1) = 1
  Else
     aparam(1) = 2
  End If
End If

Call ImpresionRptProc(Reporte, aform, aparam, , titulo + " - " + Reporte)
Screen.MousePointer = 1
Exit Sub
Err:
   MsgBox Err.Description, vbInformation, "Aviso"
   Screen.MousePointer = 1
End Sub

Private Sub Carga_RepoVal(dato)
Dim rAdo As ADODB.Recordset
Dim Aux, cadena As String
Dim cAnoMes As String, cCod As String
Dim cSql1 As String, CSQL2 As String
Dim Dia1 As Integer, Mes1 As Integer
Dim nSaldo As Double, nCosPro As Double
Dim nPrecio As Double, nCantid As Double
Dim I As Integer
Dim Codi As String
Dim rsql As String
Dim csql As String
Dim uSql As String
Dim saldo As Double
Dim xx As Integer
On Error GoTo ErrCarga
Set adodc1 = New ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
Set rAdo = New ADODB.Recordset
    
Mes1 = Month(DTPicker1)
cAnoMes = Year(DTPicker1) & Format(Mes1, "00")

nTra = 0
       
cSql1 = "Select * From MovAlmCab a Inner Join MovAlmDet b On a.CAALMA = b.DEALMA And a.CATD = b.DETD  And a.CANUMDOC = b.DENUMDOC "
cSql1 = cSql1 & "  Where Month(CAFECDOC) = " & Mes1 & " and Year(CAFECDOC) = " & Year(DTPicker1) & " AND CAALMA = '" & almacen & "'  "
cSql1 = cSql1 & "  And CASITGUI<>'A' " & dato
cSql1 = cSql1 & " Order By DECODIGO,CAFECDOC,catipmov,canumdoc"

adodc1.Open cSql1, VGCNx, adOpenStatic

Adodc2.Open "Select * From MoresMes Where SMALMA = '" & almacen & "' Order By SMMESPRO", VGCNx, adOpenStatic
 
    Call Carga2
    Exit Sub

If xx = 0 Then
    If adodc1.RecordCount > 0 Then
       adodc1.MoveFirst
       Carga
    End If
    
    'Para los articulos que no han  tenido movimiento ,suma=suma+saldotemp
    'uSql = "select s.stcodigo, s.stskdis from stkart s  WHERE S.STCODIGO NOT  In (SELECT COD_ART FROM al_Kardex_Val IN '" & App.Path & "\bdauxcom.mdb ')  and  s.STALMA ='" & almacen & "' and  s.stcodigo >= '" & Text1 & "' and s.stcodigo <= '" & Text2 & "'"
    uSql = "SELECT SUM(S.SMCANENT- S.SMCANSAL) AS COL, SMCODIGO  " & _
            "FROM MORESMES AS s " & _
            "WHERE S.SMCODIGO NOT  In (SELECT COD_ART FROM al_Kardex_Val )  and  s.smalma = '" & almacen & "'   s.smmespro < '" & cAnoMes & "' group by smcodigo"
    Set Adostk = New ADODB.Recordset
    Adostk.Open uSql, VGCNx, adOpenStatic
    While Not Adostk.EOF
       csql = "select top 1 (SMMNPREUNI) as costo from moresmes where   smcodigo = '" & Adostk(1) & "' and  smalma = '" & almacen & "' and smmespro < '" & cAnoMes & "' order by   smmespro desc "
       Set AdoMes = New ADODB.Recordset
       AdoMes.Open csql, VGCNx, adOpenStatic
       If Not AdoMes.EOF Then
'             If Adostk(0) <> 0 Then
'                     rSql = "INSERT INTO al_Kardex_Val ( COD_ART, NUM_DOC, SAL_STOCK, COS_PRO )" & _
'                                " values ('" & Adostk(1) & "' ,'SALDO INICIAL'," & Adostk(0) & "," & AdoMes("costo") & ")"
'                    VGcnx.Execute rSql
'              End If

       End If
       Adostk.MoveNext
    Wend
   'aqui debo tener la suma acumulada a imprimir
 
  Exit Sub
End If
adodc1.MoveFirst
Carga2
Exit Sub
ErrCarga:
        MsgBox Err.Description
        If nTra = 1 Then VGCNx.RollbackTrans
      '  Resume
        
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
Dim xx As Integer
'**********Roberto
Dim rSAX As New ADODB.Recordset
Dim SqlM As String
Dim Flag0 As Integer
'*************************
xx = 0
Mes1 = Month(DTPicker1)
cAnoMes = Year(DTPicker1) & Format(Mes1, "00")
cCod = "": Codi = "": Li = 0
Flag0 = 0
Do While Not adodc1.EOF
    
    If adodc1("detipcam") <> 0 And adodc1("cacodmon") = "02" Then
        TCamb = adodc1("detipcam")
    Else
     If Combo3.ListIndex = 0 Then
        TCamb = 1
     Else
        
        If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
           TCamb = Val(Devolver_Dato(3, adodc1("cafecdoc"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
        ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
               TCamb = Val(Devolver_Dato(1, adodc1("cafecdoc"), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
        End If
     End If
    End If
    
    If xx = 0 Then
    
       If adodc1("CACODMOV") = "GF" And adodc1("CASITGUI") = "F" Then '(adodc1("CATD") = "GS" And adodc1("CACODMOV") = "GV" And adodc1("CASITGUI") = "F") Or
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
                    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,FEC_DOC,HOR_DOC,COD_MOV,TIP_TRANSA,NUM_DOC,CAN_ART,PRE_UNIT,COS_PRO,SAL_STOCK,SER_LOT,ING_SAL)  "
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
                    VGCNx.BeginTrans
                    VGCNx.Execute CSQL2
                    VGCNx.CommitTrans
                    nTra = 0
                   
                   cCod = adodc1("DECODIGO")
            Else
                   'aqui actulizo el moresmes ********************************************
                   If cCod <> "" Then
                         cMesPro = Year(DTPicker1) & Format(Mes1, "00")
                         
                         If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
                             nTipCam = Val(Devolver_Dato(3, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
                         ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
                             nTipCam = Val(Devolver_Dato(1, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
                         End If
                         Rsql1 = "Update  MoResMes  Set  SMMNPREUNI  = " & nCosPro & ", SMUSPREUNI  = " & nCosPro / IIf(nTipCam = 0, 1, nTipCam) & "  where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & cMesPro & "' AND  SMCODIGO= '" & cCod & "'"
                        
                         VGCNx.BeginTrans
                         VGCNx.Execute Rsql1
                         VGCNx.CommitTrans
                         
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
                                                   ' CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
                                            Else
                                                    nSaldo = Adodc2("SMCANENT") - Adodc2("SMCANSAL")
                                                    nCosPro = Adodc2("SMMNPREUNI")
                                                    
                                                    If Combo3.ListIndex <> 0 Then
                                                       If CDbl(TCamb) <> 0 Then
                                                          nCosPro = nCosPro / TCamb
                                                       End If
                                                    End If
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
                                                        If Combo3.ListIndex <> 0 Then
                                                            If CDbl(TCamb) <> 0 Then
                                                               nCosPro = nCosPro / TCamb
                                                            End If
                                                        End If

                                                     End If
                                                    Adodc3.Close
                                                    'CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
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
                                                    
                                                    'CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
                                            Else
                                                    Set Adodc3 = New ADODB.Recordset
                                                    Adodc3.Open "Select sum(SMCANENT - SMCANSAL)  as saldo From MoresMes Where SMALMA = '" & almacen & "' And SMMESPRO < '" & cAnoMes & "' And SMCODIGO = '" & adodc1("DECODIGO") & "'", VGCNx, adOpenStatic
                                                    nSaldo = 0
                                                    nCosPro = 0
                                                    If Adodc3.RecordCount > 0 Then
                                                            nSaldo = IIf(IsNull(Adodc3("Saldo")), 0, Adodc3("Saldo"))
                                                    End If
                                                    Set Adodc3 = New ADODB.Recordset
                                                    Adodc3.Open "Select  SMMNPREUNI as promedio From MoresMes Where SMALMA = '" & almacen & "' And SMMESPRO < '" & cAnoMes & "' And SMCODIGO = '" & adodc1("DECODIGO") & "'  order by SMMESPRO  DESC ", VGCNx, adOpenStatic
                                                     If Adodc3.RecordCount > 0 Then
                                                            nCosPro = IIf(IsNull(Adodc3("promedio")), 0, Adodc3("promedio"))
                                                     
                                                             If Combo3.ListIndex <> 0 Then
                                                                If CDbl(TCamb) <> 0 Then
                                                                    nCosPro = nCosPro / TCamb
                                                                End If
                                                             End If
                                            End If
                                                     
                                                    Adodc3.Close
                                                    'CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
                                            End If
                                    End If
                            Else
                                    'CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
                           End If
                   Else
                            'Aqui se realizohastael mes de proceso
                            Set Adodc3 = New ADODB.Recordset
                            Adodc3.Open "Select sum(SMCANENT - SMCANSAL)  as saldo From MoresMes Where SMALMA = '" & almacen & "'" & _
                                                    " And SMMESPRO < '" & cAnoMes & "' And SMCODIGO = '" & adodc1("DECODIGO") & "'", VGCNx, adOpenStatic
                            nSaldo = 0
                            If Adodc3.RecordCount > 0 Then
                                    nSaldo = IIf(IsNull(Adodc3("Saldo")), 0, Adodc3("Saldo"))
                            End If
                            Adodc3.Close
                            nCosPro = 0
                            'CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & Adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
                   End If
'                   If Trim(CSQL2) <> "" Then
'                        nTra = 1
'                        VGcnx.BeginTrans
'                        VGcnx.Execute CSQL2
'                        VGcnx.CommitTrans
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
                    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,FEC_DOC,HOR_DOC,COD_MOV,TIP_TRANSA,NUM_DOC,CAN_ART,PRE_UNIT,COS_PRO,SAL_STOCK,SER_LOT,ING_SAL)  "
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
                  VGCNx.BeginTrans
                  VGCNx.Execute CSQL2
                  VGCNx.CommitTrans
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
                SqlM = "Update al_Kardex_Val set PRE_UNIT=" & nPrecio & ",cos_pro=" & nCosPro & "  WHERE COD_ART='" & cCod & "'  and (Pre_unit=0 AND cos_pro=0) and (Tip_Transa='NC' or Tip_Transa='NI') "
            Else
                SqlM = "Update al_Kardex_Val set PRE_UNIT=" & nCosPro & ",Cos_Pro=" & nCosPro & " WHERE COD_ART='" & cCod & "'  and (Pre_unit=0 AND cos_pro=0) and (Tip_Transa<>'NC' and Tip_Transa<>'NI') "
            End If
                VGCNx.BeginTrans
                VGCNx.Execute SqlM
                VGCNx.CommitTrans
            '*****************************************
    End If

Loop
 If cCod <> "" Then
    cMesPro = Year(DTPicker1) & Format(Mes1, "00")
                         
    If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
       nTipCam = Val(Devolver_Dato(3, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
    ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
           nTipCam = Val(Devolver_Dato(1, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
    End If
    Rsql1 = "Update  MoResMes  Set  SMMNPREUNI  = " & nCosPro & ", SMUSPREUNI  = " & nCosPro / IIf(nTipCam = 0, 1, nTipCam) & "  where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & cMesPro & "' AND  SMCODIGO= '" & cCod & "'"
    VGCNx.BeginTrans
    VGCNx.Execute Rsql1
    VGCNx.CommitTrans
  End If
End Sub

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
Ctr_AyuTipoMovimiento.filtro = "tt_tipmov='" & Left(Combo4.text, 1) & "'"
If Combo4.ListIndex = 2 Then
   Framemov.Visible = False
 Else
   Framemov.Visible = True
End If
End Sub


Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim rsql As String
central Me
Check1.Value = 0
Carga_Almacen
'Combo2.ListIndex = 0
'If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
If Combo1.ListIndex = 0 Then VGForm1 = 6
DTPicker1.Value = VGParamSistem.FechaTrabajo
Call Ctr_AyuFamilia.conexion(VGCNx)
Call Ctr_AyuTipoMovimiento.conexion(VGCNx)
Call Ctr_Ayuarticulo.conexion(VGCNx)
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
On Error GoTo Err

Do While Not adodc1.EOF
    
    
       If adodc1("detipcam") <> 0 And adodc1("cacodmon") = "02" Then
          TCamb = adodc1("detipcam")
       Else
              TCamb = 1
      End If
    
       If (adodc1("CATD") = "GS" And adodc1("CACODMOV") = "GF" And adodc1("CASITGUI") = "F") Then '(adodc1("CATD") = "GS" And adodc1("CACODMOV") = "GV" And adodc1("CASITGUI") = "F") Or
            adodc1.MoveNext
       Else
          If cCod <> adodc1("DECODIGO") Then
             nPrecio = 0: nCantid = 0: nSaldo = 0: nCosPro = 0

             If Not Adodc2.EOF Then
                Adodc2.MoveFirst
                Adodc2.Filter = "SMCODIGO = '" & adodc1("DECODIGO") & "' AND SMMESPRO <'" & cAnoMes & "'"
             End If
             
             While Not Adodc2.EOF
                   nSaldo = nSaldo + Adodc2!SMCANENT
                   nSaldo = nSaldo - Adodc2!SMCANSAL
                   nCosPro = Adodc2("SMMNPREUNI")
                   Adodc2.MoveNext
             Wend
             VALANTE = nCosPro * nSaldo
             Adodc2.Filter = ""
          End If
          
  
          nCantid = adodc1("DECANTID")
              
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
                
          CSQL2 = "Insert Into al_Kardex_Val (COD_ART,FEC_DOC,HOR_DOC,COD_MOV,TIP_TRANSA,NUM_DOC,CAN_ART,PRE_UNIT,COS_PRO,SAL_STOCK,SER_LOT,ING_SAL,alma)  "
          CSQL2 = CSQL2 & " Values ('" & adodc1("DECODIGO") & "','" & Format(adodc1("CAFECDOC"), "dd/mm/yyyy") & "','" & adodc1("CAHORA") & "','" & adodc1("CACODMOV") & "',"
          CSQL2 = CSQL2 & "'" & adodc1("CATD") & "','" & adodc1("CANUMDOC") & "', " & adodc1("DECANTID") & ","
                              
          If adodc1("CATIPMOV") = "I" Then
              CSQL2 = CSQL2 & "" & nPrecio & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','I','" & almacen & "')"
          Else
              CSQL2 = CSQL2 & "" & nCosPro & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','S','" & almacen & "')"
          End If
          nTra = 1
          VGCNx.BeginTrans
          VGCNx.Execute CSQL2
          VGCNx.CommitTrans
          nTra = 0
         
         cCod = adodc1("DECODIGO")
             
         adodc1.MoveNext
       End If
               
       If adodc1.EOF Then Exit Do

Loop
Err:
      '  MsgBox (Err.Description)
      '  Resume
        If nTra = 1 Then VGCNx.RollbackTrans
End Sub


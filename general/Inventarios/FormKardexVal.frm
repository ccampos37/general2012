VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormKardexVal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kardex Valorizado"
   ClientHeight    =   4500
   ClientLeft      =   1530
   ClientTop       =   2430
   ClientWidth     =   4890
   ControlBox      =   0   'False
   Icon            =   "FormKardexVal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   4890
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   195
      Left            =   4020
      TabIndex        =   18
      Top             =   4125
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   344
      _Version        =   393216
      Format          =   52887553
      CurrentDate     =   36710
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   135
      Top             =   2370
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   690
      Left            =   1260
      Picture         =   "FormKardexVal.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3645
      Width           =   735
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   2790
      Picture         =   "FormKardexVal.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3645
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   3420
      Left            =   180
      TabIndex        =   12
      Top             =   72
      Width           =   4440
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FormKardexVal.frx":114E
         Left            =   1335
         List            =   "FormKardexVal.frx":1158
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1395
         Width           =   2160
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1335
         TabIndex        =   1
         Top             =   660
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MMMM 'del' yyyy"
         Format          =   52887555
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Resumido"
         Height          =   270
         Left            =   2400
         TabIndex        =   8
         Top             =   2610
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Detallado"
         Height          =   270
         Left            =   780
         TabIndex        =   7
         Top             =   2610
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   1185
         Left            =   216
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   3990
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1335
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1410
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3135
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1410
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FormKardexVal.frx":116C
         Left            =   1335
         List            =   "FormKardexVal.frx":117C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1035
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1350
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1755
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1335
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2115
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormKardexVal.frx":11A1
         Left            =   1335
         List            =   "FormKardexVal.frx":11A3
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   255
         Width           =   2760
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Incluye Costo de Salidas "
         Height          =   195
         Left            =   570
         TabIndex        =   21
         Top             =   3000
         Width           =   3585
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda   :"
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   1455
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Familia"
         Height          =   225
         Left            =   390
         TabIndex        =   20
         Top             =   1455
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Línea"
         Height          =   225
         Left            =   2550
         TabIndex        =   19
         Top             =   1425
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Mes"
         Height          =   255
         Left            =   375
         TabIndex        =   17
         Top             =   705
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
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
         Left            =   345
         TabIndex        =   16
         Top             =   1065
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Del"
         Height          =   255
         Left            =   735
         TabIndex        =   15
         Top             =   1755
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Al"
         Height          =   255
         Left            =   750
         TabIndex        =   14
         Top             =   2115
         Width           =   135
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   375
         TabIndex        =   13
         Top             =   270
         Width           =   735
      End
   End
End
Attribute VB_Name = "FormKardexVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cConexAux As ADODB.Connection
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

Private Sub CmdSalir_Click()
If cConexAux.State = 1 Then Set cConexAux = Nothing
Unload Me
End Sub

Private Sub Combo1_Click()
rs.MoveFirst
rs.Move Combo1.ListIndex
almacen = Format(rs(0), "00")
End Sub

Private Sub CmdAceptar_Click()
Dim rsql As String
Dim Va1 As String, Va2 As String
Screen.MousePointer = 11
On Error GoTo Err
Set Adodc3 = New ADODB.Recordset  'Para sacar la descripcion del rango elegido
If Combo2.ListIndex = 0 Then
    If Text1.Visible And Trim(Text1) = "" Then
        MsgBox "Ingrese el código a imprimir", vbInformation, "Aviso"
        Text1.SetFocus
        Screen.MousePointer = 1
        Exit Sub
    End If
ElseIf Combo2.ListIndex = 1 Then
    If Text3 = "" Then
        MsgBox "Ingrese el código de Familia", vbInformation, "Aviso"
        Text3.SetFocus
        Screen.MousePointer = 1
        Exit Sub
    ElseIf Text4 = "" Then
        MsgBox "Ingrese el código de Línea", vbInformation, "Aviso"
        Text4.SetFocus
        Screen.MousePointer = 1
        Exit Sub
    End If
ElseIf Combo2.ListIndex = 2 Then
    If Text3 = "" Then
        MsgBox "Ingrese el código de Familia", vbInformation, "Aviso"
        Text3.SetFocus
        Screen.MousePointer = 1
        Exit Sub
    End If
End If
Carga_RepoVal

'Para calcular el total de los detalles
'If Combo2.ListIndex = 0 Then
'    Dim ado1 As New ADODB.Recordset
'    Dim Total As Double
'    Dim Cod As String
'    Total = 0   'Order By Cod_Art,Fec_Doc,hor_doc
'    ado1.Open "Select Cod_Art,Cos_Pro,Sal_Stock from al_Kardex_Val order by cod_art,fec_doc,hor_doc ", cConexAux, adOpenDynamic, adLockOptimistic
'    If ado1.RecordCount > 0 Then
'        ado1.MoveFirst
'        Do While Not ado1.EOF
'            Cod = ado1("Cod_Art")
'            ado1.MoveNext
'            If Not ado1.EOF Then
'                    If Cod <> ado1("Cod_Art") Then
'                        ado1.MovePrevious
'                        Total = Total + Format((Format(ado1("Cos_Pro"), "####.00000") * Format(ado1("Sal_Stock"), "###.00000")), "###,###,###,##0.0000")
'                    Else
'                        ado1.MovePrevious
'                    End If
'                    ado1.MoveNext
'            End If
'            If ado1.EOF Then
'                Exit Do
'            End If
'        Loop
'        ado1.MoveLast
'        Total = Total + Format((Format(ado1("Cos_Pro"), "####.00000") * Format(ado1("Sal_Stock"), "###.00000")), "###,###,###,##0.00")
'    End If
'    ado1.Close
'End If
''''''''''''''''''''''''''''''''''''''''''''''''''''
  If Combo2.ListIndex = 0 Then
      If Text1 <> "" And Text2 <> "" Then
        If Option1.Value = True Then
          If Check1.Value = 1 Then
            CrystalReport1.WindowTitle = "Inv027s -- Control de Inventarios"
            CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv027s.rpt"
          Else
            CrystalReport1.WindowTitle = "Inv027 -- Control de Inventarios"
            CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv027.rpt"
          End If
        Else
           CrystalReport1.WindowTitle = "Inv029 -- Control de Inventarios"
           CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv029.rpt"
'          CrystalReport1.SelectionFormula = "{al_Kardex_Val.COSPRO}={KARDEXAUX.C1}"
        End If
      End If
   ElseIf Combo2.ListIndex = 1 Then
    CrystalReport1.WindowTitle = "Inv026 -- Control de Inventarios"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv026.rpt"
   ElseIf Combo2.ListIndex = 2 Then
    CrystalReport1.WindowTitle = "Inv028 -- Control de Inventarios"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv028.rpt"
   ElseIf Combo2.ListIndex = 3 Then
    If Option1.Value = True Then
      CrystalReport1.WindowTitle = "Inv025 -- Control de Inventarios"
      CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv025.rpt"
   Else
      CrystalReport1.WindowTitle = "Inv030 -- Control de Inventarios"
      CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv030.rpt"
   End If
  End If
  If CrystalReport1.ReportFileName = "" Then
      MsgBox "No hay registros a imprimir", vbInformation, "Aviso"
      Screen.MousePointer = 1
      Exit Sub
  End If
  Ubi_Tab CrystalReport1
  CrystalReport1.DiscardSavedData = True
  CrystalReport1.Destination = crptToWindow
  CrystalReport1.WindowShowPrintBtn = True
  CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.WindowShowSearchBtn = True
  CrystalReport1.WindowShowPrintSetupBtn = True
  CrystalReport1.formulas(0) = "ALMACEN = '" & UCase(Combo1.text) & "'"
  CrystalReport1.formulas(1) = "Mes= '" & UCase(Format(DTPicker1, "MMMM - yyyy")) & "'"
  CrystalReport1.formulas(2) = "EMPRESA= '" & UCase(VGparametros.RucEmpresa) & "'"
  CrystalReport1.formulas(3) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
  If Combo3.ListIndex <> 0 Then
     CrystalReport1.formulas(4) = "MONEY= 'DOLAR'"
  Else
      CrystalReport1.formulas(4) = "MONEY= 'SOLES'"
  End If
  If Combo2.ListIndex = 0 Then CrystalReport1.formulas(5) = "ART1= '" & Text1 & "'"
  If Combo2.ListIndex = 0 Then CrystalReport1.formulas(6) = "ART2 ='" & Text2 & "'"
  'If Combo2.ListIndex = 0 Then CrystalReport1.Formulas(7) = "TotGen ='" & Format(Total, "###,###,###,##0.00") & "'"

  If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
 Screen.MousePointer = 1
 Exit Sub
Err:
   MsgBox Err.Description, vbInformation, "Aviso"
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
Dim rsql As String
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
cConexAux.Execute "Delete From al_Kardex_Val"
cConexAux.CommitTrans
nTra = 0

If Text1.Visible And Trim(Text2) = "" Then
    Text2 = Text1
End If
'limpia
'csql = "delete  from MORESMES where SMCANENT = 0 and SMCANSAL =  0"
'Vgcnx.Execute csql
       
If Combo2.ListIndex = 0 Then
cSql1 = "Select * From MovAlmDet A Inner Join MovAlmCab B On B.CAALMA = A.DEALMA And B.CATD = A.DETD  And B.CANUMDOC = A.DENUMDOC "
cSql1 = cSql1 & "  Where Month(CAFECDOC) = " & Mes1 & " and Year(CAFECDOC) = " & Year(DTPicker1) & " AND CAALMA = '" & almacen & "' AND DECODIGO >= '" & Text1 & "' "
cSql1 = cSql1 & " And  DECODIGO <= '" & Text2 & "' And CASITGUI<>'A' Order By DECODIGO,CAFECDOC,CAHORA"
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
cSql1 = cSql1 & " Where Month(CAFECDOC) = " & Mes1 & " and Year(CAFECDOC) = " & Year(DTPicker1) & " AND CAALMA = '" & almacen & "' And CASITGUI<>'A' " 'AND AFAMILIA >= '" & Text1 & "'  And  AFAMILIA <= '" & Text2 & "'"
cSql1 = cSql1 & " Order By AFamilia,DECODIGO,CAFECDOC,CAHORA"
End If

adodc1.Open cSql1, VGCNx, adOpenStatic
Adodc2.Open "Select * From MoresMes Where SMALMA = '" & almacen & "'  and SMUSPREUNI <> 0 and SMMNPREUNI <> 0 " & _
                       " Order By SMMESPRO", VGCNx, adOpenStatic
 
 If Adodc2.RecordCount > 0 Then
        If Val(cAnoMes) - Val(Adodc2("SMMESPRO")) > 1 Then
              '  MsgBox "El Costo utilizado es del Mes de " & DesMes(Right(Adodc2("SMMESPRO"), 2)) & Chr(13) & ", porque no se ha hecho el Cierre en los meses anteriores", vbInformation, "Información"
        End If
Else
        MsgBox "No se ha hecho Cierre en los meses anteriores y su Costo Inicial será Cero", vbInformation, "Información"
End If
cCod = ""
If Combo2.ListIndex = 0 Then
    If adodc1.RecordCount > 0 Then
       adodc1.MoveFirst
       Carga
    End If
    
    'Para los articulos que no han  tenido movimiento ,suma=suma+saldotemp
    'uSql = "select s.stcodigo, s.stskdis from stkart s  WHERE S.STCODIGO NOT  In (SELECT COD_ART FROM al_Kardex_Val IN '" & App.Path & "\bdauxcom.mdb ')  and  s.STALMA ='" & almacen & "' and  s.stcodigo >= '" & Text1 & "' and s.stcodigo <= '" & Text2 & "'"
    uSql = "SELECT SUM(S.SMCANENT- S.SMCANSAL) AS COL, SMCODIGO  " & _
            "FROM MORESMES AS s " & _
            "WHERE S.SMCODIGO NOT  In (SELECT COD_ART FROM al_Kardex_Val IN '" & App.Path & "\bdauxcom.mdb ')  and  s.smalma = '" & almacen & "'  and  s.smcodigo >= '" & Trim(Text1) & "' and s.smcodigo <= '" & Trim(Text2) & "' and s.smmespro < '" & cAnoMes & "' group by smcodigo"
    Set Adostk = New ADODB.Recordset
    Adostk.Open uSql, VGCNx, adOpenStatic
    While Not Adostk.EOF
       csql = "select top 1 (SMMNPREUNI) as costo from moresmes where   smcodigo = '" & Adostk(1) & "' and  smalma = '" & almacen & "' and smmespro < '" & cAnoMes & "' order by   smmespro desc "
       Set AdoMes = New ADODB.Recordset
       AdoMes.Open csql, VGCNx, adOpenStatic
       If Not AdoMes.EOF Then
             If Adostk(0) <> 0 Then
                     rsql = "INSERT INTO al_Kardex_Val ( COD_ART, NUM_DOC, SAL_STOCK, COS_PRO )" & _
                                " values ('" & Adostk(1) & "' ,'SALDO INICIAL'," & Adostk(0) & "," & AdoMes("costo") & ")"
                    cConexAux.Execute rsql
              End If
       End If
       Adostk.MoveNext
    Wend
   'aqui debo tener la suma acumulada a imprimir
 
  Exit Sub
End If

For I = 0 To List1.ListCount - 1
  List1.ListIndex = I
  If List1.Selected(I) = True Then
    If adodc1.RecordCount > 0 Then
      adodc1.MoveFirst
      Carga
    End If
  End If
Next

Exit Sub
ErrCarga:
        MsgBox Err.Description
        If nTra = 1 Then cConexAux.RollbackTrans
End Sub

Private Sub Carga()
Dim TCamb As Double
Dim Li As Integer
Dim cCod As String, Codi As String
Dim Aux, CADENA As String
Dim cAnoMes As String
Dim cSql1 As String, CSQL2 As String
Dim Dia1 As Integer, Mes1 As Integer
Dim nSaldo As Double, nCosPro As Double
Dim nPrecio As Double, nCantid As Double
Dim cMesPro As String
Dim nPrePro As Double
Dim Rsql1 As String
Dim nTipCam As String

Mes1 = Month(DTPicker1)
cAnoMes = Year(DTPicker1) & Format(Mes1, "00")
cCod = "": Codi = "": Li = 0
Do While Not adodc1.EOF
    If Combo2.ListIndex = 1 Then Codi = IIf(IsNull(adodc1("Agrupo")), "", adodc1("Agrupo")): Li = 8 'captura el dato para la comparacion
    If Combo2.ListIndex = 2 Then Codi = IIf(IsNull(adodc1("Amodelo")), "", adodc1("Amodelo")): Li = 4
    If Combo2.ListIndex = 3 Then Codi = IIf(IsNull(adodc1("Afamilia")), "", adodc1("Afamilia")): Li = 4
    
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
    If Trim(Mid(List1.text, 1, Li)) = Codi Or Combo2.ListIndex = 0 Then
    
       If (adodc1("CATD") = "GS" And adodc1("CACODMOV") = "GF" And adodc1("CASITGUI") = "F") Then '(adodc1("CATD") = "GS" And adodc1("CACODMOV") = "GV" And adodc1("CASITGUI") = "F") Or
            adodc1.MoveNext
       Else
            
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
                            If nCosPro <> 0 Then
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
                    If adodc1("CATD") = "NI" Or adodc1("CATD") = "NC" Then
                        CSQL2 = CSQL2 & "" & nPrecio & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','I')"
                    Else
                        CSQL2 = CSQL2 & "" & nCosPro & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','S')"
                    End If
                    nTra = 1
                    cConexAux.BeginTrans
                    cConexAux.Execute CSQL2
                    'Debe actualizar solo el ultimo mes
'                    CSQL2 = "Update STKART set stskdis= " & Val(nSaldo) & " where  STCODIGO= '" & adodc1("DECODIGO") & "' and STALMA = '" & VGAlma & "'"
'                    Vgcnx.Execute CSQL2
                    
                    cConexAux.CommitTrans
                    nTra = 0
                   
                   cCod = adodc1("DECODIGO")
            Else
                   'aqui actulizo el moresmes ********************************************
                  
                   If cCod <> "" Then
                         cMesPro = Year(DTPicker1) & Format(Mes1, "00")
                         nPrePro = Val(Devolver_Dato(1, cCod, "StkArt", "STCODIGO", False, "STKPREPRO", VGAlma, "STALMA"))
                         
                         If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
                             nTipCam = Val(Devolver_Dato(3, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
                         ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
                             nTipCam = Val(Devolver_Dato(1, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
                         End If
                         'Rsql1 = "Update  MoResMes  Set  SMMNPREUNI  = " & nCosPro & ", SMUSPREUNI  = " & nCosPro / IIf(nTipCam = 0, 1, nTipCam) & "  where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & cMesPro & "' AND  SMCODIGO= '" & cCod & "'"
                        
                         VGCNx.BeginTrans
                         VGCNx.Execute Rsql1
                         VGCNx.CommitTrans
                         
                   End If
                   
                   'aqui lo nuevos valores
                    nSaldo = 0: nCosPro = 0
                    If Adodc2.RecordCount > 0 Then
                            Adodc2.MoveFirst
                            Adodc2.Filter = "SMCODIGO = '" & adodc1("DECODIGO") & "'"
                            If Not Adodc2.EOF Then
                                    Adodc2.MoveLast
                                    If Adodc2("SMMESPRO") = cAnoMes Then
                                            Adodc2.MovePrevious
                                            If Adodc2.BOF Then
                                                    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
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
                                                    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
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
                                                    
                                                    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
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
                                                    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
                                            End If
                                    End If
                            Else
                                    CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
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
                            CSQL2 = "Insert Into al_Kardex_Val (COD_ART,SAL_STOCK,COS_PRO,NUM_DOC) Values ('" & adodc1("DECODIGO") & "'," & nSaldo & "," & nCosPro & ",'SALDO INICIAL')"
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

                    If adodc1("CATD") = "NI" Then
                         CSQL2 = CSQL2 & "" & nPrecio & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','I')"
                    Else
                         CSQL2 = CSQL2 & "" & nCosPro & "," & nCosPro & "," & nSaldo & ",'" & IIf(IsNull(adodc1("DESERIE")) Or adodc1("DESERIE") = "", adodc1("DELOTE"), adodc1("DESERIE")) & "','S')"
                    End If
                 
                   nTra = 1
                  cConexAux.BeginTrans
                  cConexAux.Execute CSQL2
                  cConexAux.CommitTrans
                  'Debe actualizar solo el ultimo mes
'                  CSQL2 = "Update STKART set stskdis= " & Val(nSaldo) & " where  STCODIGO= '" & adodc1("DECODIGO") & "' and STALMA = '" & VGAlma & "'"
'                  Vgcnx.Execute CSQL2
                  
                  nTra = 0
                
                  cCod = adodc1("DECODIGO")
            End If
            adodc1.MoveNext
        End If
        If adodc1.EOF Then Exit Do
    Else
      adodc1.MoveNext
    End If
Loop
 If cCod <> "" Then
    cMesPro = Year(DTPicker1) & Format(Mes1, "00")
    nPrePro = Val(Devolver_Dato(1, cCod, "StkArt", "STCODIGO", False, "STKPREPRO", VGAlma, "STALMA"))
                         
    If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
       nTipCam = Val(Devolver_Dato(3, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
    ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
           nTipCam = Val(Devolver_Dato(1, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
    End If
    'Rsql1 = "Update  MoResMes  Set  SMMNPREUNI  = " & nCosPro & ", SMUSPREUNI  = " & nCosPro / IIf(nTipCam = 0, 1, nTipCam) & "  where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & cMesPro & "' AND  SMCODIGO= '" & cCod & "'"
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

Private Sub Combo2_Click()
Text3 = "": Text4 = "": Option1.Visible = False: Option2.Visible = False
List1.Clear
If Combo2.ListIndex = 0 Then                'ARTICULOS
  Label1.Visible = False: Label2.Visible = False
  Text3.Visible = False: Text4.Visible = False
  Text1.Visible = True: Text2.Visible = True
  Option1.Visible = True: Option2.Visible = True
  Option1.Top = 2610: Option2.Top = 2610
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
  Option1.Top = 5400: Option2.Top = 5400
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
central Me
Carga_Almacen
Combo2.ListIndex = 0
'If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
If Combo1.ListIndex = 0 Then VGForm1 = 6
DTPicker1.Value = Date
ADOConectar
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdAceptar.SetFocus
End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Combo2.ListIndex = 0 Then CmdAceptar.SetFocus Else List1.SetFocus
End If
End Sub

Private Sub Option2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Combo2.ListIndex = 0 Then CmdAceptar.SetFocus Else List1.SetFocus
End If
End Sub

Private Sub Text1_DblClick()
Dim Adodc2 As Recordset
Set Adodc3 = New ADODB.Recordset
If Combo2.ListIndex = 0 Then
        VGForm1 = 6
        almacenAnt = VGAlma
        VGAlma = almacen
        FormAyuArt1.Show 1
        VGAlma = almacenAnt
ElseIf Combo2.ListIndex = 1 Then
        Adodc3.Open "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'", VGCNx, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc3, "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'"
        frmReferencia.Label1.Caption = "Grupos de Artículos"
        frmReferencia.Show vbModal
        Adodc3.Close
        If vGUtil(1) <> "" Then
                Text1 = (vGUtil(1))
        End If
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
      If Existe(1, Text1, "MaeArt", "Acodigo", False) = False Then
               MsgBox "El Código no existe", vbInformation, "Información"
               Text1.Enabled = True
               Text1.SetFocus
      Else
               SendKeys "{Tab}"
      End If
Else
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub Text2_DblClick()
Set Adodc2 = New ADODB.Recordset
If Combo2.ListIndex = 0 Then
        VGForm1 = 6
        almacenAnt = VGAlma
        VGAlma = almacen
        FormAyuArt1.Show 1
        VGAlma = almacenAnt
ElseIf Combo2.ListIndex = 1 Then
        Adodc2.Open "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'", VGCNx, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'"
        frmReferencia.Label1.Caption = "Grupos de Artículos"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text2 = (vGUtil(1))
        End If
ElseIf Combo2.ListIndex = 2 Then
        Adodc2.Open "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'", VGCNx, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'"
        frmReferencia.Label1.Caption = "Líneas de Artículos"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text2 = (vGUtil(1))
        End If
ElseIf Combo2.ListIndex = 3 Then
        Adodc2.Open "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", VGCNx, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
        frmReferencia.Label1.Caption = "Familias de Artículos"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
                Text2 = (vGUtil(1))
        End If
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
      If Existe(1, Text2, "MaeArt", "Acodigo", False) = False Then
                   MsgBox "El codigo no existe", vbExclamation, "Aviso"
                   Text2.SetFocus
                   Exit Sub
      End If
      If Text1 > Text2 Then
                   MsgBox "El Código Fin debe ser Mayor que el Inicio", vbExclamation, "Aviso"
                   Text2.SetFocus
                   Exit Sub
      End If
      SendKeys "{Tab}"
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub ADOConectar()
cRt = App.Path & "\BdAuxCom.Mdb"
Set cConexAux = New ADODB.Connection
cConexAux.CursorLocation = adUseClient
cConexAux.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & cRt & ";"
cConexAux.Open
End Sub

Private Sub Text3_Change()
If Text3 = "" Then List1.Clear: Text4 = ""
End Sub

Private Sub Text3_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
If Combo2.ListIndex = 1 Or Combo2.ListIndex = 2 Then
         Adodc2.Open "Select FAM_CODIGO,FAM_NOMBRE,FAM_CTA from FAMILIA ", VGCNx, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select FAM_CODIGO,FAM_NOMBRE,FAM_CTA from FAMILIA "
         frmReferencia.Label1.Caption = "Familias"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
            Text3 = (vGUtil(1))
         End If
         If Combo2.ListIndex = 1 Then If Text4 <> "" Then Cargalist: Exit Sub
         If Combo2.ListIndex = 2 Then Cargalist
End If
End Sub

Private Sub Cargalist()
Dim Cod As String, Des As String
Set Adodc4 = New ADODB.Recordset
List1.Clear

If Combo2.ListIndex = 1 Then
  Adodc4.Open "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'", VGCNx, adOpenStatic, adLockOptimistic
ElseIf Combo2.ListIndex = 2 Then
  Adodc4.Open "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'", VGCNx, adOpenStatic, adLockOptimistic
ElseIf Combo2.ListIndex = 3 Then
  Adodc4.Open "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", VGCNx, adOpenStatic, adLockOptimistic
End If
If Adodc4.RecordCount > 0 Then
  Do While Not Adodc4.EOF
    Cod = Space(8)
    Des = Space(45)
    If Combo2.ListIndex = 1 Then Cod = Adodc4("Gru_Codigo"): Des = Adodc4("Gru_Nombre")
    If Combo2.ListIndex = 2 Then Cod = Adodc4("Lin_Codigo"): Des = Adodc4("Lin_Nombre")
    If Combo2.ListIndex = 3 Then Cod = Adodc4("Fam_Codigo"): Des = Adodc4("Fam_Nombre")
    List1.AddItem Cod & "  " & Des
    Adodc4.MoveNext
  Loop
End If
Adodc4.Close
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text3_DblClick
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Text3 <> "" Then
     
     If Text4.Visible = True Then Text4.SetFocus: Exit Sub
     If Combo2.ListIndex = 1 Or Combo2.ListIndex = 2 Then Cargalist: List1.SetFocus
  End If
End Sub

Private Sub Text4_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
If Combo2.ListIndex = 1 Then
         Adodc2.Open "Select LIN_CODIGO,LIN_NOMBRE from LINEAS Where FAM_CODIGO='" & Text3 & "'", VGCNx, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select LIN_CODIGO,LIN_NOMBRE from LINEAS Where FAM_CODIGO='" & Text3 & "'"
         frmReferencia.Label1.Caption = "Líneas"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
            Text4 = (vGUtil(1))
         End If
         If Text4 <> "" Then Cargalist
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text4_DblClick
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Text4 <> "" Then Cargalist: List1.SetFocus
End Sub

  

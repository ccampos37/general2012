VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmConsultaNotas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Documento"
   ClientHeight    =   5400
   ClientLeft      =   1500
   ClientTop       =   1650
   ClientWidth     =   9960
   Icon            =   "FrmconsultaNotas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleMode       =   0  'User
   ScaleWidth      =   10454.76
   Begin VB.Frame Frame2 
      Caption         =   "Consulta de Documentos"
      Height          =   4476
      Left            =   30
      TabIndex        =   29
      Top             =   30
      Width           =   9708
      Begin VB.OptionButton Option3 
         Caption         =   "Guias"
         Height          =   225
         Left            =   1440
         TabIndex        =   33
         Top             =   2388
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nota de Salida"
         Height          =   300
         Left            =   1440
         TabIndex        =   32
         Top             =   1968
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nota de Ingreso"
         Height          =   195
         Left            =   1440
         TabIndex        =   31
         Top             =   1668
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Todos"
         Height          =   225
         Left            =   1440
         TabIndex        =   30
         Top             =   2736
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   6165
         TabIndex        =   34
         Top             =   2430
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   97910785
         CurrentDate     =   36704
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   6165
         TabIndex        =   35
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   97910785
         CurrentDate     =   36704
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctrayu_almacen 
         Height          =   348
         Left            =   2664
         TabIndex        =   38
         Top             =   1248
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   609
         XcodMaxLongitud =   0
         xcodwith        =   300
         NomTabla        =   "TABALM"
         ListaCampos     =   "taalma(1),tadescri(1)"
         XcodCampo       =   "taalma"
         XListCampo      =   "tadescri"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "taalma,tadescri"
      End
      Begin VB.Label Label14 
         Caption         =   "Almacen"
         Height          =   300
         Left            =   1560
         TabIndex        =   39
         Top             =   1248
         Width           =   972
      End
      Begin VB.Label Label13 
         Caption         =   "Hasta"
         Height          =   252
         Left            =   5232
         TabIndex        =   37
         Top             =   2424
         Width           =   732
      End
      Begin VB.Label Label12 
         Caption         =   "Desde"
         Height          =   252
         Left            =   5208
         TabIndex        =   36
         Top             =   2064
         Width           =   732
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle del Documento"
      ForeColor       =   &H80000008&
      Height          =   4440
      Left            =   96
      TabIndex        =   4
      Top             =   96
      Width           =   9240
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7365
         TabIndex        =   28
         Top             =   1410
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   1095
         Width           =   1770
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   1440
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid FG2 
         Height          =   2175
         Left            =   210
         TabIndex        =   9
         Top             =   2010
         Width           =   8850
         _ExtentX        =   15610
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         ScrollTrack     =   -1  'True
         FormatString    =   "    CODIGO   |                DESCRIPCION            |      UNIDAD    |   CANTIDAD   |         COSTO      |  UNI_ALM    "
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   40
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Moneda"
         Height          =   210
         Left            =   6480
         TabIndex        =   27
         Top             =   1440
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   240
         X2              =   7920
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   6720
         TabIndex        =   20
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Doc Referencial"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Transaccion"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   3480
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         BorderStyle     =   6  'Inside Solid
         Index           =   0
         X1              =   240
         X2              =   7920
         Y1              =   1820
         Y2              =   1820
      End
      Begin VB.Label Label6 
         Caption         =   "Num"
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Lblndoc 
         Caption         =   "Label1"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   1440
         Width           =   2070
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   3630
         TabIndex        =   10
         Top             =   1095
         Width           =   3735
      End
   End
   Begin VB.CommandButton Cmdconsultar 
      Caption         =   "&Consultar"
      Height          =   675
      Left            =   3000
      Picture         =   "FrmconsultaNotas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4590
      Width           =   775
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   480
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   675
      Left            =   3000
      Picture         =   "FrmconsultaNotas.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4590
      Width           =   775
   End
   Begin VB.CommandButton Command7 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   4860
      Picture         =   "FrmconsultaNotas.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4605
      Width           =   775
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   3372
      Left            =   240
      TabIndex        =   21
      Top             =   1176
      Width           =   9276
      _ExtentX        =   16351
      _ExtentY        =   5953
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   975
      Left            =   1110
      TabIndex        =   23
      Top             =   -12
      Width           =   6735
      Begin VB.TextBox TxtBuscar 
         Height          =   285
         Left            =   1305
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmconsultaNotas.frx":1590
         Left            =   4440
         List            =   "FrmconsultaNotas.frx":15A0
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   97910785
         CurrentDate     =   36679
      End
      Begin VB.Label Label21 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "Indice"
         Height          =   255
         Left            =   3480
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmConsultaNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''Dim db As Database
Dim tipo As String * 2
Dim cImpresora As String, cPuerto As String, cControlador As String

Private Sub CmdImprimir_Click()
  'validar si dichos codigo existen
  imprimir
End Sub

Private Sub Combo1_Click()
 FG.Col = Combo1.ListIndex
 'FG.ColSel = 1
 FG.Sort = 5
End Sub

Private Sub CmdConsultar_Click()
Dim precio As Double
'Dim db1 As Database
'Dim Tipo As String * 2
 Dim CANTIDAD As Double
 Dim contador As Integer
 Dim Rsql1 As String
 Dim RSQL As String
 Dim csql As String
 Dim rs As Recordset
 Dim rs1 As Recordset
 Dim serie_lote As String
 If Frame2.Visible Then
    Text5.Visible = False
    Label11.Visible = False
    If Option3.Value Then
        tipo = "GS"
    ElseIf Option2.Value Then
        tipo = "NS"
    ElseIf Option1.Value Then
        tipo = "NI"
    Else
        tipo = "XX"
        RSQL = "select  m.CATD, m.CANUMDOC, m.CACODMOV ,m.CAFECDOC, m.CACODPRO,m.CACODCLI, m.CARFTDOC ,m.CARFNDOC,m.CASITGUI,m.canomcli,m.cadirenv,caruc from MovAlmCab m where  m.CAALMA ='" & VGAlma & "'   and m.CATD  IN  ('NI','NS','GS')  and  m.cafecdoc  between " & (DTPicker2.Value) & " and " & (DTPicker3.Value) & " ORDER BY m.CANUMDOC"    '
    End If
    If tipo <> "XX" Then
        RSQL = "select  m.CATD, m.CANUMDOC, m.CACODMOV ,m.CAFECDOC, m.CACODPRO,m.CACODCLI, m.CARFTDOC ,m.CARFNDOC,m.CASITGUI,m.canomcli,m.cadirenv,caruc, catipmov  from MovAlmCab m where  m.CAALMA ='" & VGAlma & "' and m.CATD='" & tipo & "' " '
        RSQL = RSQL & " and m.cafecdoc > ='" & DTPicker2 & "' and m.cafecdoc <='" & DTPicker3 & "' ORDER BY m.CANUMDOC"
    End If
    Set rs = VGCNx.Execute(RSQL)
    FG.Rows = 1
    If rs.EOF Then
        MsgBox "no hay documentos registrados" & Chr(13) & "verifique el rango de fecha", vbCritical, "Aviso"
        Exit Sub
    End If
    rs.MoveFirst
    FG.Visible = False
    While Not rs.EOF
       FG.AddItem (rs(0) & vbTab & rs(1) & vbTab & rs(2) & vbTab & rs(3) & vbTab & rs(4) & vbTab & rs(5) & vbTab & rs(6) & vbTab & rs(7) & vbTab & rs(8) & vbTab & rs(9) & vbTab & rs(10) & vbTab & rs(11) & vbTab & rs(12))
       rs.MoveNext
    Wend
    FG.Visible = True
    rs.Close
    Frame2.Visible = False
    Frame1.Visible = False
    CmdConsultar.Caption = "Consultar"
    cmdImprimir.Visible = True
    Exit Sub
End If
If Frame1.Visible Then
    Frame1.Visible = False
Else
    If FG.Row = 0 Then
         MsgBox "No se ha seleccionado ningún" & Chr(13) & " documento", vbInformation, "Consulta de Documentos"
         Exit Sub
    End If
    If CmdConsultar.Visible Then
         cmdImprimir.Visible = True
         CmdConsultar.Visible = False
    Else
         CmdConsultar.Visible = True
    End If
     Frame1.Visible = True
     CmdConsultar.Caption = "&Aceptar"
     Label10 = Format(FG.TextMatrix(FG.Row, 3), "dd/mm/yyyy") 'fecha
     Text2 = FG.TextMatrix(FG.Row, 2)  'tras
     Label7 = transa
     Label9 = " ACTIVO"
     If FG.TextMatrix(FG.Row, 8) = "A" Then
         Label9 = "ANULADO"   ' estado registro
     End If
     Text1 = FG.TextMatrix(FG.Row, 0) ' tipo de doc
     tipo = Text1
     Label19 = FG.TextMatrix(FG.Row, 1) ' cod de doc
     Text3 = FG.TextMatrix(FG.Row, 4)  'proveedor
     Label8 = Mid(prove, 1, 30)
     Text4 = FG.TextMatrix(FG.Row, 6)  'doc ref
     Lblndoc = FG.TextMatrix(FG.Row, 7)  'proveedor
     Rsql1 = "select n.DECODIGO, m.ADESCRI, m.AUNIDAD, n.DECANTID, n.DEPRECIO,DESERIE,DELOTE,DEITEM,DECENCOS  from MovAlmDet n, MaeArt m  where  n.DEALMA ='" & VGAlma & "' AND n.DETD = '" & Text1 & "' AND n.DENUMDOC ='" & Label19 & "' AND m.ACODIGO = n.DECODIGO "  '
     Rsql1 = Rsql1 & "  UNION select n.DECODIGO,N.DEDESCRI,'UNI',' ','0',' ',' ',DEITEM,DECENCOS  from MovAlmDet n  where  n.DEALMA ='" & VGAlma & "' AND n.DETD = '" & Text1 & "' AND n.DENUMDOC ='" & Label19 & "' AND n.DECODIGO ='TEXTO' ORDER BY DEITEM  " '
     Set rs = VGCNx.Execute(Rsql1)
     If rs.EOF Then
            MsgBox "El documento no tiene detalle", vbInformation, "Aviso"
            Exit Sub
     Else
       rs.MoveFirst
     End If
     FG2.Visible = False
     FG2.Rows = 1
     While Not rs.EOF
        If Not IsNull(rs(5)) And rs(5) <> "" Then
            serie_lote = rs(5)
        ElseIf Not IsNull(rs(6)) Then
            serie_lote = rs(6)
        Else
            serie_lote = ""
        End If
        FG2.AddItem (rs(0) & vbTab & rs(1) & vbTab & serie_lote & vbTab & Format(rs(3), "#0.#00") & vbTab & rs(2) & vbTab & Format(rs(4), "###0.000") & vbTab & rs(7) & vbTab & rs!DECENCOS)
        rs.MoveNext
     Wend
     rs.Close
     'FG2.Col = 6
     'FG2.Sort = 5
     FG2.Visible = True
End If
End Sub

Private Sub Command7_Click()
If Frame1.Visible Then
     limpia
     Frame1.Visible = False
     CmdConsultar.Visible = True
     cmdImprimir.Visible = False
Else
     'Db.Close
     Unload Me
End If
End Sub

Private Sub Ctrayu_almacen_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    VGAlma = Ctrayu_almacen.xclave
End Sub

Private Sub Form_Load()
'Dim db As Database
Dim rs As Recordset
Dim RSQL As String

DTPicker3 = Date
 DTPicker2 = DateAdd("m", -1, Date)
 limpia
 Label7 = ""
 Label8 = ""
 DTPicker1.Visible = False
 'Cmdconsultar.Caption = "Aceptar"
 cmdImprimir.Visible = False
FG.FormatString = "Tipo Doc.|Número de Doc| Tr| Fecha |^ Proveedor|^Cliente|Td REF|<Num.Doc Ref.|^Situación|<Nombre del cliente|^Direccion de entrega|^Ruc   |^Estado"
FG.Row = 0
FG.ColWidth(0) = 800
FG.ColWidth(1) = 1500
FG.ColWidth(2) = 800
FG.ColWidth(3) = 1000
FG.ColWidth(4) = 1300
FG.ColWidth(5) = 1300
FG.ColWidth(6) = 800
FG.ColWidth(7) = 1500
FG.ColWidth(8) = 1000
FG.ColWidth(9) = 3500
FG.ColWidth(10) = 3500
FG.ColWidth(11) = 1000
FG.ColWidth(12) = 1000
FG.ColAlignment(1) = 1
Frame2.Visible = True
Combo1.ListIndex = 0
Call Ctrayu_almacen.conexion(VGCNx)
End Sub

Private Sub limpia()
Label10 = ""
Text1 = ""
Text2 = ""
Text3 = ""
Label19 = ""
Text5 = "01"
FG2.Clear
FG2.Rows = 1
FG2.Cols = 8
FG2.TextMatrix(0, 0) = " Codigo"
FG2.TextMatrix(0, 1) = " Descripcion"
FG2.TextMatrix(0, 2) = " Serie \Lote"
FG2.TextMatrix(0, 3) = " Cantidad"
FG2.TextMatrix(0, 4) = " Unidad"
FG2.TextMatrix(0, 5) = " Costo Unit"
FG2.TextMatrix(0, 6) = " Costo Inf"
FG2.TextMatrix(0, 7) = " Centr.Costo"
FG2.Row = 0
FG2.Cols = 8
FG2.ColWidth(0) = 2000
FG2.ColAlignment(0) = 1
FG2.ColAlignment(2) = 1
FG2.ColWidth(1) = 3220
FG2.ColWidth(2) = 1900
FG2.ColWidth(3) = 1000
FG2.ColWidth(4) = 1000
FG2.ColWidth(5) = 1500
FG2.ColWidth(6) = 2
FG2.ColWidth(7) = 1400
End Sub

Function transa() As String
'Dim db As Database
Dim rs As Recordset
Dim RSQL As String
transa = ""
If FG.TextMatrix(FG.Row, 0) = "GS" Or FG.TextMatrix(FG.Row, 0) = "NS" Then
    RSQL = "select  TT_DESCRI FROM TabTransa where TT_CODMOV= '" & FG.TextMatrix(FG.Row, 2) & "' AND TT_TIPMOV = 'S'"
Else
    RSQL = "select  TT_DESCRI FROM TabTransa where TT_CODMOV= '" & FG.TextMatrix(FG.Row, 2) & "' AND TT_TIPMOV = 'I'"
End If
'Set db = Workspaces(0).OpenDatabase(cRuta2)
'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)

Set rs = VGCNx.Execute(RSQL)
If Not rs.EOF Then
    transa = rs(0)
End If
rs.Close
End Function
Function prove() As String

'Dim db As Database
Dim rs As Recordset
Dim RSQL As String

 prove = ""
  RSQL = "select clienterazonsocial FROM cp_proveedor where clientecodigo= '" & FG.TextMatrix(FG.Row, 4) & "'" '
 '[rsql = "select PRVCNOMBRE FROM maeprov where PRVCCODIGO= '" & FG.TextMatrix(FG.Row, 4) & "'" '
' Set db = Workspaces(0).OpenDatabase(cRuta2)
' Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
 Set rs = VGCNx.Execute(RSQL)
 If Not rs.EOF Then
        prove = rs(0)
 End If
 rs.Close
End Function

Private Sub imprimir()
Dim cNomRepor As String
On Error GoTo ErrImp
If FG.TextMatrix(FG.Row, 6) = "GR" And tipo = "NS" Then
   imprimirguias
 Else
    CrystalReport1.Reset
    cNomRepor = "REPNOTAING.rpt"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & cNomRepor
              
    CrystalReport1.Connect = VGCadenaReport2
    CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
    CrystalReport1.StoredProcParam(1) = VGAlma
    CrystalReport1.StoredProcParam(2) = Trim(Text1.text)
    CrystalReport1.StoredProcParam(3) = Trim(Label19.Caption)
                            
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.formulas(0) = "fecha='" & Label10.Caption & "'"
    CrystalReport1.formulas(1) = "xtrans = '" & Label7.Caption & "' "
    CrystalReport1.formulas(2) = "xtd = '" & Trim(Text2.text) & "' "
    CrystalReport1.formulas(3) = "xndoc = '" & Trim(Label19.Caption) & "' "
                            
    If tipo = "NI" Then
       CrystalReport1.WindowTitle = "RepNotaIng -- Impresion de Notas de Ingreso"
       CrystalReport1.formulas(4) = "Xnalma = '" & Ctrayu_almacen.xclave & "' "
       CrystalReport1.formulas(5) = "Dalma = '" & Ctrayu_almacen.xnombre & "' "
       CrystalReport1.formulas(6) = "AlmaDes = '" & VGAlma & "' "
       CrystalReport1.formulas(7) = "Dalmades = '" & Ctrayu_almacen.xnombre & "' "
     ElseIf tipo = "NS" Then
          CrystalReport1.WindowTitle = "RepNotaIng -- Impresion de Notas de Salida"
          CrystalReport1.formulas(4) = "Xnalma = '" & VGAlma & "' "
          CrystalReport1.formulas(5) = "Dalma = '" & Ctrayu_almacen.xnombre & "' "
          CrystalReport1.formulas(6) = "AlmaDes = '" & VGAlma & "' "
          CrystalReport1.formulas(7) = "Dalmades = '" & Ctrayu_almacen.xnombre & "' "
    End If
                            
    CrystalReport1.formulas(8) = "NRef = '" & Lblndoc.Caption & "' "
    CrystalReport1.formulas(9) = "DocRef = '" & Text4.text & "' "
    CrystalReport1.formulas(10) = "TTrans = '" & Label7.Caption & "' "
    CrystalReport1.formulas(11) = "emp = '" & VGParametros.RucEmpresa & "'"
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowState = crptMaximized

    If CrystalReport1.Status <> 2 Then
       CrystalReport1.Action = 1
       VGCNx.Execute "Update MovAlmCab Set CaEstImp = 'I' Where CATD = '" & tipo & "' and CANUMDOC = '" & Text4.text & "'"
    End If
End If
Exit Sub
ErrImp:
     MsgBox Err.Description
     Resume Next
     Exit Sub
End Sub
Private Sub imprimirguias()

Dim nguia As String
Dim ntabla As String
Dim busca As New dll_apisgen.dll_apis
Dim VGDllGeneral As New dll_general
Dim rb As New ADODB.Recordset
Dim rb1 As New ADODB.Recordset
Dim contador As Double
Dim contador1 As Double
Dim numguias As Integer, TCant As Integer, nflag As Integer
Dim SQL As String
Dim ruc As String
Dim inicio As Integer
Dim fin As Integer
Dim J As Integer
Dim numero As String
Dim distrito As String


ntabla = "movalmdet"
contador = 0

VGCNx.Execute "delete from gtempfile"
VGCNx.Execute "delete from tempfile"
SQL = "INSERT into gtempfile Select b.decantid,c.acodigo,c.adescri,0,0,0,b.decantid,c.aunidad "
SQL = SQL & " From movalmcab  A inner join movalmdet b "
SQL = SQL & " ON a.caalma=b.dealma and a.catd=b.detd and a.canumdoc=b.denumdoc  "
SQL = SQL & " inner join maeart C ON b.decodigo=c.acodigo "
SQL = SQL & " Where a.carftdoc='" & FG.TextMatrix(FG.Row, 6) & "' and carfndoc='" & FG.TextMatrix(FG.Row, 7) & "' order by c.afamilia,c.alinea,c.agrupo "

VGCNx.Execute SQL

contador = 0
Set rb = VGCNx.Execute("select * from gtempfile ")
If rb.RecordCount > 0 Then
    If rb.RecordCount Mod 50 > 0 Then
        numguias = Int(rb.RecordCount / 50) + 1
     Else
         numguias = Int(rb.RecordCount / 50)
    End If
     rb.MoveFirst
     Do While contador < numguias
              contador = contador + 1
              inicio = (contador - 1) * 50 + 1
              If contador * 50 > rb.RecordCount Then
                 fin = rb.RecordCount
               Else
                 fin = contador * 50
              End If
              nguia = FG.TextMatrix(FG.Row, 7)
              contador1 = 0
              If fin > rb.RecordCount Then
                 fin = rb.RecordCount - inicio
              End If
              VGCNx.Execute "delete from gtempfile2filas"
          For J = inicio To fin
                 contador1 = contador1 + 1
                 If contador1 <= 25 Then
                     SQL = "INSERT INTO gtempfile2filas(item,producto1,descripcion1,cantidad1,importe1,"
                     SQL = SQL & "cantidad2,importe2) "
                     SQL = SQL & " VALUES ( '" & contador1 & "','" & RTrim(rb!productocodigo) & "','" & RTrim(rb!productodescripcion) & "','" & rb!detpedcantpedida & "','" & rb!detpedimpbruto & "',0,0)"
                  Else
                     TCant = contador1 - 25
                      SQL = "UPDATE gtempfile2filas set producto2 ='" & RTrim(rb!productocodigo) & "',"
                      SQL = SQL & " descripcion2='" & RTrim(rb!productodescripcion) & "',"
                      SQL = SQL & "cantidad2='" & rb!detpedcantpedida & "',"
                        SQL = SQL & "importe2= '" & rb!detpedimpbruto & "'"
                        SQL = SQL & " where item = " & TCant & ""
                 End If
                 VGCNx.Execute SQL
                 rb.MoveNext
          Next J
          CrystalReport1.Reset
          CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "Repguiaimpresa.rpt"
  
          CrystalReport1.Connect = VGCadenaReport2
          
          CrystalReport1.WindowShowPrintSetupBtn = True
          CrystalReport1.WindowShowExportBtn = True
          CrystalReport1.WindowShowZoomCtl = True
          CrystalReport1.WindowShowNavigationCtls = True
          CrystalReport1.WindowShowPrintBtn = True
          CrystalReport1.Destination = crptToWindow
          CrystalReport1.WindowState = crptMaximized
          CrystalReport1.DiscardSavedData = True
          With CrystalReport1
                   .formulas(0) = "nro='" & Text4 & "'"
                   .formulas(1) = "cliente='" & FG.TextMatrix(FG.Row, 9) & "'"
                   .formulas(2) = "fecha='" & CStr(Day(FG.TextMatrix(FG.Row, 3))) & "     " & VGDllGeneral.DesMes(Month(FG.TextMatrix(FG.Row, 3))) & "                       " & Right(CStr(Year(FG.TextMatrix(FG.Row, 3))), 4) & "'"
                   .formulas(3) = "direccion='" & FG.TextMatrix(FG.Row, 10) & "'"
                   .formulas(4) = "dni='" & ruc & "'"
                   .formulas(5) = "opedido=''"
'                   .formulas(6) = "ocompra='" & Text8 & "'"
'                   .formulas(7) = "guia='" & nguia & "'"
'                   .formulas(8) = "distrito='" & distrito & "'"
'                   .formulas(9) = "destino='" & Text7 & "'"
                   Set rb1 = VGCNx.Execute("select * from empresa where emp_codigo='" & VGCodEmpresa & "'")
                   If rb1.RecordCount > 0 Then
                      .formulas(10) = "partida='" & Trim(rb1!emp_direccion) & "'"
                    Else
                      .formulas(10) = "partida=''"
                   End If
                   .StoredProcParam(0) = VGParamSistem.BDEmpresa
                   '.ParameterFields(0) = VGcnx.DefaultDatabase
                    If .Status <> 2 Then .Action = 1
          End With
          SQL = nguia
          MsgBox "Proceda a imprimir la GUIA DE REMISION .", vbInformation, SQL
    Loop
End If
rb.Close

nerror:
   If Err Then
      If nflag = 1 Then
         VGCNx.RollbackTrans
      End If
      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
      Exit Sub
   End If
  
End Sub

Private Sub Txtbuscar_Change()
Dim I As Integer
Dim n As Integer
   n = Combo1.ListIndex
   If TxtBuscar <> "" Then
      For I = 1 To FG.Rows - 1
          If UCase(Left(FG.TextMatrix(I, n), Len(TxtBuscar))) = UCase(Trim(TxtBuscar)) Then
             Exit For
          End If
      Next I
      If I >= FG.Rows Then
            FG.HighLight = flexHighlightNever
      Else
            FG.HighLight = flexHighlightAlways
            FG.TopRow = I
            FG.Row = I
            FG.Col = 0
            FG.ColSel = FG.Cols - 1
      End If
   End If
   
   'Dim i As Integer
   'Dim N As Integer
   'N = Combo1.ListIndex
   'i f TxtBuscar <> "" Then
    '  For i = 1 To FG1.Rows - 1
   '       If UCase(Left(FG1.TextMatrix(i, N), Len(TxtBuscar))) = UCase(Trim(TxtBuscar)) Then
   '          Exit For
   '       End If
   '   Next i
   '   If i >= FG1.Rows Then
   '         FG1.HighLight = flexHighlightNever
   '   Else
   '         FG1.HighLight = flexHighlightAlways
   '         FG1.TopRow = i
   '         FG1.Row = i
   '         FG1.Col = 0
   '         FG1.ColSel = FG1.Cols - 1
   '   End If
   'End If
   
End Sub

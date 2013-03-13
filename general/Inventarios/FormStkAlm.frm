VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FormStkAlm 
   Caption         =   "Stock de Articulos"
   ClientHeight    =   4905
   ClientLeft      =   5325
   ClientTop       =   1215
   ClientWidth     =   5850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   5850
   Begin VB.CheckBox Chkstokcero 
      Caption         =   "Suprimir Stock   0"
      Height          =   255
      Left            =   285
      TabIndex        =   27
      Top             =   3870
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   288
      ItemData        =   "FormStkAlm.frx":0000
      Left            =   3840
      List            =   "FormStkAlm.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   3570
      Width           =   1440
   End
   Begin VB.CheckBox ChkSerie 
      Caption         =   "Con Serie o Lote"
      Height          =   255
      Left            =   285
      TabIndex        =   24
      Top             =   3570
      Width           =   1695
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   300
      Top             =   4140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Consulta"
      Height          =   1215
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   5415
      Begin VB.ComboBox cboMon 
         Height          =   315
         ItemData        =   "FormStkAlm.frx":0023
         Left            =   4425
         List            =   "FormStkAlm.frx":002D
         TabIndex        =   29
         Top             =   780
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CheckBox Chkprecio 
         Caption         =   "Con Costo"
         Height          =   195
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3180
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Consolidado"
         Height          =   225
         Left            =   360
         TabIndex        =   1
         Top             =   825
         Width           =   1335
      End
      Begin VB.Label lblMon 
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
         Height          =   255
         Left            =   3540
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Por Almacen"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros"
      Height          =   1695
      Left            =   240
      TabIndex        =   16
      Top             =   1620
      Width           =   1815
      Begin VB.OptionButton Option1 
         Caption         =   "Cod. Ubicacion"
         Height          =   195
         Left            =   270
         TabIndex        =   23
         Top             =   2085
         Width           =   1425
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Grupos"
         Height          =   255
         Left            =   255
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Líneas"
         Height          =   255
         Left            =   255
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Familias"
         Height          =   255
         Left            =   255
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton OpArt 
         Caption         =   "Artículos"
         Height          =   255
         Left            =   255
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame FrameRep 
      Height          =   2055
      Left            =   2400
      TabIndex        =   15
      Top             =   1440
      Width           =   3255
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   8
         Top             =   510
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   7
         Top             =   210
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.OptionButton OpRango 
         Caption         =   "Rango"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1035
         Width           =   1455
      End
      Begin VB.OptionButton OpTodos 
         Caption         =   "Todos los Artículos"
         Height          =   270
         Left            =   360
         TabIndex        =   9
         Top             =   765
         Width           =   1785
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1320
         Width           =   1470
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1635
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Línea"
         Height          =   225
         Left            =   330
         TabIndex        =   22
         Top             =   525
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Familia"
         Height          =   225
         Left            =   300
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio "
         Height          =   255
         Left            =   765
         TabIndex        =   20
         Top             =   1335
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         Height          =   255
         Left            =   765
         TabIndex        =   19
         Top             =   1635
         Width           =   735
      End
   End
   Begin VB.CommandButton Command7 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   624
      Left            =   3030
      Picture         =   "FormStkAlm.frx":0039
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4176
      Width           =   660
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   624
      Left            =   1800
      Picture         =   "FormStkAlm.frx":047B
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4176
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Ordenado por "
      Height          =   255
      Left            =   2505
      TabIndex        =   25
      Top             =   3585
      Width           =   990
   End
End
Attribute VB_Name = "FormStkAlm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim almacen     As String
Dim Conexion    As String
Dim Adodc3      As ADODB.Recordset
Dim bMon As Boolean

Dim tituloreporte As String
Dim NombreReporte As String
Dim OrdenReporte As String
Private Sub Chkprecio_Click()
If Chkprecio.Value = 1 Then
   lblMon.Visible = True
   cboMon.Visible = True
   cboMon.ListIndex = 0
Else
   lblMon.Visible = False
   cboMon.Visible = False
   cboMon.ListIndex = -1
End If
End Sub

Private Sub Combo1_Click()
almacen = Mid(Combo1, 1, 2)
End Sub

Private Sub Command7_Click()
    MousePointer = vbHourglass
    If ExisteElem(0, cConexCom, TempoTAB) Then cConexCom.Execute "DROP TABLE " & TempoTAB
    If Frame1.Visible And Frame2.Visible Then
        MousePointer = vbDefault
        Unload Me
    Else
        Frame1.Visible = True
        Frame2.Visible = True
        FrameRep.Visible = False
        MousePointer = vbDefault
    End If
End Sub

Private Sub Command1_Click()
    bMon = False
    If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
        MsgBox "Ingrese un código menor al fin ", vbOKOnly, "Error"
        Exit Sub
    End If
    Screen.MousePointer = 11
    If OpArt.Value Then
         imprimir
    ElseIf Option2.Value Then
        Imprimir2
    ElseIf Option3.Value Then
        Imprimir3
    ElseIf Option4.Value Then
        Imprimir4
    End If
    Screen.MousePointer = 1
End Sub

Private Sub Form_Load()
    Me.Height = 5250
    Me.Width = 6015
    '*****************
    Frame1.Visible = True
    Frame2.Visible = True
    '*****************
    TempoTAB = "##" & ComputerName & "STKALM"
    If ExisteElem(0, cConexCom, TempoTAB) Then cConexCom.Execute "DROP TABLE " & TempoTAB
    Carga_Almacen
    central FormStkAlm
    OpArt.Value = True
    OpTodos.Value = True
    FrameRep.Caption = " Por Articulos"
    VGForm1 = 3
    Combo2.ListIndex = 0
End Sub

Private Sub OpRango_Click()
If OpRango.Value Then
    Text1.Enabled = True
    Text2.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub OpArt_Click()
OpArt.Value = True
FrameRep.Caption = " Por Articulos "
OpTodos.Caption = "Todos los Articulos"
limpiar_t1_t2
OpTodos.Top = 300: OpRango.Top = 650
Text1.Top = 1100: Label2.Top = 1100
Text2.Top = 1500: Label3.Top = 1500
End Sub

Private Sub Option2_Click()

Option2.Value = True
FrameRep.Caption = " Por Familias "
OpTodos.Caption = "Todos las Familias"
limpiar_t1_t2
OpTodos.Top = 300: OpRango.Top = 650
Text1.Top = 1100: Label2.Top = 1100
Text2.Top = 1500: Label3.Top = 1500
End Sub

Private Sub Option3_Click()
Option3.Value = True
FrameRep.Caption = " Por Lineas "
OpTodos.Caption = "Todos las Líneas "
limpiar_t1_t2
Label4.Visible = True
Text3.Visible = True
OpTodos.Top = 550: OpRango.Top = 900
Text1.Top = 1200: Label2.Top = 1200
Text2.Top = 1600: Label3.Top = 1600
End Sub

Private Sub Option4_Click()
Option4.Value = True
FrameRep.Caption = " Por Grupos "
OpTodos.Caption = "Todos los Grupos"
limpiar_t1_t2
Label4.Visible = True: Text3.Visible = True
Label5.Visible = True: Text4.Visible = True
OpTodos.Top = 850: OpRango.Top = 1100
Text1.Top = 1400: Label2.Top = 1400
Text2.Top = 1700: Label3.Top = 1700
End Sub

Private Sub Carga_Almacen()

Dim RSQL      As String
Dim rs        As ADODB.Recordset
Dim i         As Integer
 
RSQL = "select TAALMA,TADESCRI FROM TabAlm "
Set rs = New ADODB.Recordset
rs.Open RSQL, cConexCom, adOpenStatic, adLockOptimistic

While Not rs.EOF
  Combo1.AddItem (rs(0)) & "  " & (rs(1))
  rs.MoveNext
Wend

rs.MoveFirst
For i = 0 To rs.RecordCount - 1
  If rs(0) = VGAlma Then
    Combo1.ListIndex = i
    Exit For
  Else
    rs.MoveNext
  End If
Next
rs.Close
End Sub

Private Sub imprimir()
Dim Codigo1       As String
Dim Codigo2       As String
Dim cadena        As String
Dim RSQL          As String
Dim tex1          As String, tex2         As String
Dim Va1           As String, Va2          As String

ReDim arrfor(7)
ReDim arrparam(2)
    
    Codigo1 = UCase(Trim(Text1))
    Set Adodc3 = New ADODB.Recordset
    
    RSQL = "Select ACodigo,Adescri from MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo " & _
                " Where Stalma='" & almacen & "' AND not(A.AFSERIE='N' AND A.AFLOTE='N' AND A.AFSTOCK='N') " & _
                " AND RTRIM(LTRIM(A.ACODIGO))<>'TEXTO' Order by Acodigo"
                
        Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
        If Adodc3.RecordCount > 0 Then
            If Text1 = "" And Text2 = "" Then
                Adodc3.MoveFirst
                'campos para formulas de reporte
                tex1 = IIf(IsNull(Adodc3("ACodigo")), "", Adodc3("ACodigo"))
                Va1 = IIf(IsNull(Adodc3("ADescri")), "", Adodc3("Adescri"))
                Adodc3.MoveLast
                tex2 = IIf(IsNull(Adodc3("ACodigo")), "", Adodc3("ACodigo"))
                Va2 = IIf(IsNull(Adodc3("ADescri")), "", Adodc3("ADescri"))
            End If
        End If
        Adodc3.Close
        
    Cadfiltro = ""
    If ChkSerie = 0 Then 'Si no es con serie.....
        G_SQL = " SELECT STKART.STCODIGO, MAEART.ADESCRI, STKART.STALMA, MAEART.AUNIDAD, " & _
            " STKART.STSKDIS, STKART.STKPREPRO, STKART.STKPREPROUS, TABALM.TADESCRI, MAEART.AFSERIE, " & _
            " MAEART.AFLOTE INTO " & TempoTAB & _
            " FROM MAEART INNER JOIN STKART ON MAEART.ACODIGO = STKART.STCODIGO INNER JOIN " & _
            " TABALM ON STKART.STALMA = TABALM.TAALMA "
        If Me.Check1.Value = 0 Then G_SQL = G_SQL + " WHERE STALMA='" & almacen & "'"
    Else
        G_SQL = "SELECT STKART.STCODIGO, MAEART.ADESCRI, STKART.STALMA, MAEART.AUNIDAD, STKART.STSKDIS, " & _
            " STKLOTE.STSLKDIS, STKSERI.STSSKDIS, STKSERI.STSSERIE, STKLOTE.STSLOTE, STKART.STKPREPRO, " & _
            " STKART.STKPREPROUS, TABALM.TADESCRI , MAEART.AFSERIE, MAEART.AFLOTE INTO " & TempoTAB & _
            " FROM MAEART INNER JOIN STKART ON MAEART.ACODIGO = STKART.STCODIGO INNER JOIN " & _
            " STKSERI ON STKART.STALMA = STKSERI.STSALMA AND STKART.STCODIGO = STKSERI.STSCODIGO " & _
            " INNER JOIN TABALM ON STKART.STALMA = TABALM.TAALMA INNER JOIN STKLOTE " & _
            " ON STKART.STALMA = STKLOTE.STSALMA AND STKART.STCODIGO = STKLOTE.STSCODIGO " & _
            " where not(MAEART.AFSERIE='N' and MAEART.AFLOTE='N' AND MAEART.AFSTOCK='N' AND STKART.STCODIGO<>'TEXTO')"
        If Me.Check1.Value = 0 Then G_SQL = G_SQL + " and STALMA='" & almacen & "'"
    End If
    
    If ChkSerie = 1 Then 'Si es con serie.....
        G_SQL = "SELECT STKART.STCODIGO, MAEART.ADESCRI, STKART.STALMA, MAEART.AUNIDAD, STKART.STSKDIS, " & _
            " STKLOTE.STSLKDIS, STKSERI.STSSKDIS, STKSERI.STSSERIE, STKLOTE.STSLOTE, STKART.STKPREPRO, " & _
            " STKART.STKPREPROUS, TABALM.TADESCRI , MAEART.AFSERIE, MAEART.AFLOTE INTO " & TempoTAB & _
            " FROM MAEART INNER JOIN STKART ON MAEART.ACODIGO = STKART.STCODIGO INNER JOIN " & _
            " STKSERI ON STKART.STALMA = STKSERI.STSALMA AND STKART.STCODIGO = STKSERI.STSCODIGO " & _
            " INNER JOIN TABALM ON STKART.STALMA = TABALM.TAALMA INNER JOIN STKLOTE " & _
            " ON STKART.STALMA = STKLOTE.STSALMA AND STKART.STCODIGO = STKLOTE.STSCODIGO " & _
            " where not(MAEART.AFSERIE='N' and MAEART.AFLOTE='N' AND MAEART.AFSTOCK='N' AND STKART.STCODIGO<>'TEXTO')"
        If Me.Check1.Value = 0 Then G_SQL = G_SQL + " and STALMA='" & almacen & "'"
    End If
    
    If Chkstokcero.Value = 1 Then Cadfiltro = Cadfiltro & " and STKART.STskdis<>0 "
    If OpTodos.Value Then
        If Check1.Value = 0 Then 'si no es consolidado
            If ChkSerie = 1 Then 'Si es con serie.....
                TituloRpt = "Inv136 -- Control de Inventarios"
                NombreRpt = IIf(Combo2.ListIndex = 1, "\inv142.rpt", "\inv136.rpt")
                bMon = True
            ElseIf ChkSerie = 0 Then 'Si NO es con serie.....
                TituloRpt = "Inv078 -- Control de Inventarios"
                NombreRpt = "\inv078.rpt"
            Else
                TituloRpt = "Inv066-- Control de Inventarios"
                NombreRpt = "\inv066.rpt"
                bMon = True
            End If
            
        Else                                     'Consolidado
            If ChkSerie = 1 Then
                    TituloRpt = "Inv137 -- Control de Inventarios"
                    NombreRpt = IIf(Combo2.ListIndex = 1, "\inv143.rpt", "\inv137.rpt")
            ElseIf Chkprecio = 0 Then
                    TituloRpt = "Inv079 -- Control de Inventarios"
                    NombreRpt = "\inv079.rpt"
            Else
                    TituloRpt = "Inv081 -- Control de Inventarios"
                    NombreRpt = "\inv081.rpt"
            End If
            Rem MVV Cadfiltro = Cadfiltro + "AND not(AFSERIE='N' and AFLOTE='N' AND AFSTOCK='N' AND STCODIGO<>'TEXTO')"
            
        End If
            
        If Combo2.ListIndex = 1 Then
            'OrdenReporte = "+{MAEART.ADESCRI}"
        Else
            'OrdenReporte = "+{STKART.STCODIGO}"
        End If
        arrfor(0) = "emp = '" & VGNemp & "'"
        arrfor(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
        arrfor(2) = "campoini = '" & tex1 & "'"
        arrfor(3) = "campofin = '" & tex2 & "'"
        arrfor(4) = "detaini = '" & Va1 & "'"
        arrfor(5) = "detafin = '" & Va2 & "'"
        If bMon Then arrfor(6) = "xMoneda = '" & cboMon.text & "'" Else arrfor(6) = ""
        arrparam(0) = cConexCom.DefaultDatabase
        arrparam(1) = TempoTAB
        
        If ExisteElem(0, cConexCom, TempoTAB) Then cConexCom.Execute "DROP TABLE " & TempoTAB
        cConexCom.Execute G_SQL + Cadfiltro
        Call ImpresionRptCad(CrystalReport1, NombreRpt, arrfor, arrparam, OrdenReporte, NombreRpt)
        Exit Sub
    End If

    If Text1 = "" Then
        MsgBox "Ingrese el codigo", vbExclamation, "Error"
        Text1.SetFocus
        Exit Sub
    End If
    
    Codigo2 = UCase(Trim(Text2))
    RSQL = "Select ADescri from MAEART Where ACodigo='" & Codigo1 & "'"
    Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount = 1 Then
        Va1 = Adodc3("Adescri")
    End If
    Adodc3.Close
  
    RSQL = "Select ADescri from MAEART Where ACodigo='" & Codigo2 & "'"
    Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount = 1 Then
        Va2 = Adodc3("Adescri")
    End If
    Adodc3.Close

    If OpArt.Value Then
        If Check1.Value = 1 Then
            If ChkSerie = 1 Then
                TituloRpt = "Inv137 -- Control de Inventarios"
                NombreRpt = "\inv137.rpt"
                bMon = True
            ElseIf Chkprecio = 0 Then
                TituloRpt = "Inv079 -- Control de Inventarios"
                NombreRpt = "\inv079.rpt"
            Else
                TituloRpt = "Inv081 -- Control de Inventarios"
                NombreRpt = "\inv081.rpt"
                bMon = True
            End If
            
            Rem MVV Cadfiltro = Cadfiltro + " and not(MAEART.AFSERIE='N' and MAEART.AFLOTE='N' AND MAEART.AFSTOCK='N' AND STKART.STCODIGO<>'TEXTO')"
            
            If Text2 <> "" Then
                Codigo2 = Text2
                Cadfiltro = Cadfiltro + "AND (STCODIGO >= '" & Codigo1 & "' AND STCODIGO <= '" & Codigo2 & "')"
                If "\inv079.rpt" <> NombreRpt And "\inv081.rpt" <> NombreRpt Then
                    Cadfiltro = Cadfiltro + " AND STSLKDIS > 0 "
                End If
            Else
                Codigo2 = Codigo1: Va2 = Va1
                Cadfiltro = "AND STKART.STCODIGO = '" & Codigo1 & "' "
                If "\inv079.rpt" <> NombreRpt And "\inv081.rpt" <> NombreRpt Then
                    Cadfiltro = Cadfiltro + " AND STKLOTE.STSLKDIS > 0 "
                End If
            End If
        Else 'CHECK1.VALUE
            If ChkSerie = 1 Then
                TituloRpt = "Inv136 -- Control de Inventarios"
                NombreRpt = "\inv136.rpt"
                bMon = True
            ElseIf Chkprecio = 0 Then
                NombreRpt = "\inv078.rpt"
                TituloRpt = "\Inv078 -- Control de Inventarios"
            Else
                NombreRpt = "\inv066.rpt"
                TituloRpt = "\Inv066 -- Control de Inventarios"
                bMon = True
            End If
            If Text2 <> "" Then
                Codigo2 = Text2
                Cadfiltro = Cadfiltro + " and (STCODIGO >= '" & Codigo1 & "' AND STCODIGO <=  '" & Codigo2 & "')"
            Else
                Codigo2 = Codigo1: Va2 = Va1
                Cadfiltro = Cadfiltro + " and STCODIGO = '" & Codigo1 & "' "
            End If
        End If
        
'        If Chkstokcero.Value = 1 Then
'               cadena = cadena & " and {STKART.STskdis}<>0 "
'               Cadfiltro = Cadfiltro & " and STskdis<>0 "
'            Else
'               cadena = cadena
'            End If
        
        If Combo2.ListIndex = 1 Then
'          CrystalReport1.SortFields(0) = "+{al_stkartCons.ADESCRI}"
        Else
'          CrystalReport1.SortFields(0) = "+{al_stkartCons.STCODIGO}"
        End If
        arrfor(0) = "emp = '" & VGNemp & "'"
        arrfor(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
        arrfor(2) = "campoini = '" & Codigo1 & "'"
        arrfor(3) = "campofin = '" & Codigo2 & "'"
        arrfor(4) = "detaini = '" & Va1 & "'"
        arrfor(5) = "detafin = '" & Va2 & "'"
        If bMon Then arrfor(7) = "xMoneda = '" & cboMon.text & "'" Else arrfor(7) = ""
        arrparam(0) = cConexCom.DefaultDatabase
        arrparam(1) = TempoTAB

        If ExisteElem(0, cConexCom, TempoTAB) Then cConexCom.Execute "DROP TABLE " & TempoTAB
        cConexCom.Execute G_SQL + Cadfiltro
    
        Call ImpresionRptCad(Me.CrystalReport1, NombreRpt, arrfor, arrparam, , TituloRpt)
    End If
End Sub
Private Sub OpTodos_Click()
Text1.Enabled = False
Text2.Enabled = False
limpiar_t1_t2
If Option3.Value Then
  Label4.Visible = True
  Text3.Visible = True
ElseIf Option4.Value Then
  Label4.Visible = True: Label5.Visible = True
  Text3.Visible = True: Text4.Visible = True
End If
End Sub

Private Sub OpTodos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub

Private Sub Text1_DblClick()
Dim Adodc2 As ADODB.Recordset

Set Adodc2 = New ADODB.Recordset
If OpArt.Value Or Option1 Then
         VGForm1 = 3
         FormAyuArt1.Show 1
         If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
              MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
              Exit Sub
         End If
         If Text1 <> "" Then
              Text2.Enabled = True
              Text2.SetFocus
         End If
ElseIf Option2.Value Then
        Adodc2.Open "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", cConexCom, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
        frmReferencia.Label1.Caption = "Familias de Artículos"
        frmReferencia.Show vbModal
        Rem MVV Adodc2.Close
        If vGUtil(1) <> "" Then
                Text1 = (vGUtil(1))
        End If
ElseIf Option3.Value Then
        Adodc2.Open "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'", cConexCom, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'"
        frmReferencia.Label1.Caption = "Líneas de Artículos"
        frmReferencia.Show vbModal
        Rem MVV Adodc2.Close
        If vGUtil(1) <> "" Then
                Text1 = (vGUtil(1))
        End If
        If Text1 <> "" Then
                 Text2.Enabled = True
                 Text2.SetFocus
        End If
ElseIf Option4.Value Then
        Adodc2.Open "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'", cConexCom, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'"
        frmReferencia.Label1.Caption = "Grupos de Artículos"
        frmReferencia.Show vbModal
        Rem MVV Adodc2.Close
        If vGUtil(1) <> "" Then
                Text1 = (vGUtil(1))
        End If
        If Text1 <> "" Then
                 Text2.Enabled = True
                 Text2.SetFocus
        End If
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text1_DblClick
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not OpRango.Value Then
   OpRango = True
End If
If KeyAscii = 13 And Text1 <> "" Then
    If OpArt.Value Then
       If Existe_cod_art(Text1) <> "" Then
               Text2.Enabled = True
               Text2.SetFocus
       End If
   ElseIf Option2.Value Then
        If Existe(1, Text1, "FAMILIA", "FAM_CODIGO", False) = False Then
                MsgBox "El código de Familia no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                Text2.SetFocus
         End If
   ElseIf Option3.Value Then
        If Existe(1, Text1, "LINEAS", "LIN_CODIGO", False) = False Then
                MsgBox "El código de Línea no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                Text2.SetFocus
         End If
   ElseIf Option4.Value Then
        If Existe(1, Text1, "GRUPO", "GRU_CODIGO", False) = False Then
                MsgBox "El código de Grupo no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                Text2.SetFocus
         End If
     End If
 End If
End Sub

Private Sub Text2_DblClick()
Dim Adodc2 As ADODB.Recordset

Set Adodc2 = New ADODB.Recordset
If OpArt.Value Or Option1.Value Then
   VGForm1 = 3
   FormAyuArt1.Show 1
   If Text2 <> "" Then
        Command1.SetFocus
   End If
ElseIf Option2.Value Then
    Adodc2.Open "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", cConexCom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc2, "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
    frmReferencia.Label1.Caption = "Familias de Artículos"
    frmReferencia.Show vbModal
    Rem MVV Adodc2.Close
    If vGUtil(1) <> "" Then
      Text2 = (vGUtil(1))
    End If
ElseIf Option3.Value Then
        Adodc2.Open "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'", cConexCom, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT LIN_CODIGO,LIN_NOMBRE FROM LINEAS Where Fam_Codigo='" & Trim(Text3) & "'"
        frmReferencia.Label1.Caption = "Líneas de Artículos"
        frmReferencia.Show vbModal
        Rem MVV Adodc2.Close
        If vGUtil(1) <> "" Then
            Text2 = (vGUtil(1))
        End If
ElseIf Option4.Value Then
        Adodc2.Open "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'", cConexCom, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT GRU_CODIGO,GRU_NOMBRE FROM GRUPO Where Fam_Codigo='" & Trim(Text3) & "' and Lin_Codigo='" & Trim(Text4) & "'"
        frmReferencia.Label1.Caption = "Grupos de Artículos"
        frmReferencia.Show vbModal
        Rem MVV Adodc2.Close
        If vGUtil(1) <> "" Then
                Text2 = (vGUtil(1))
        End If
        If Text1 <> "" Then
                 Text2.Enabled = True
                 Text2.SetFocus
        End If
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text2_DblClick
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Text2 <> "" Then
     If OpArt.Value Then
        If Existe_cod_art(Text2) <> "" Then
           If Text1 > Text2 Then
                  MsgBox "El codigo fin debe ser mayor que el inicio", vbInformation, mensaje1
                  Exit Sub
           End If
           Command1.SetFocus
        End If
    ElseIf Option2.Value Then
         If Existe(1, Text2, "FAMILIA", "FAM_CODIGO", False) = False Then
             MsgBox "El código de Familia no existe", vbInformation, mensaje1
             Text2.SetFocus: Exit Sub
          Else
            Command1.SetFocus
          End If
    ElseIf Option3.Value Then
        If Existe(1, Text1, "LINEAS", "LIN_CODIGO", False) = False Then
                MsgBox "El código de Línea no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Command1.SetFocus
         End If
    ElseIf Option4.Value Then
        If Existe(1, Text1, "GRUPO", "GRU_CODIGO", False) = False Then
                MsgBox "El código de Grupo no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Command1.SetFocus
         End If
     End If
  End If
  If KeyAscii = 13 And Text2 = "" Then
      Command1.SetFocus
  End If
End Sub

Function Existe_cod_art(text As TextBox) As String
Dim rs As ADODB.Recordset
Dim RSQL As String

RSQL = "select  ACODIGO FROM maeart where ACODIGO = '" & text & "'" '
Set rs = New ADODB.Recordset
rs.Open RSQL, cConexCom, adOpenStatic, adLockOptimistic

If Not rs.EOF Then
    Existe_cod_art = rs(0)
Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
    Existe_cod_art = ""
End If
rs.Close

End Function

Private Sub limpiar_t1_t2()
Text1 = ""
Text2 = ""
Label4.Visible = False
Text3.Visible = False
Label5.Visible = False
Text4.Visible = False
End Sub

Private Sub Imprimir2()
On Error GoTo Mensaje
Dim cadena      As String
Dim Codigo1     As String
Dim Codigo2     As String
Dim RSQL        As String
Dim tex1        As String, tex2       As String
Dim Va1         As String, Va2        As String
ReDim arrfor(7), arrparam(2)
    Me.MousePointer = vbHourglass
    
    Set Adodc3 = New ADODB.Recordset
    RSQL = "Select AFamilia,Fam_Nombre from ((MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo) " & _
            " Left Join FAMILIA C on A.AFamilia=C.Fam_Codigo) " & _
            " Where Stalma='" & almacen & "' AND not(A.AFSERIE='N' AND A.AFLOTE='N' AND A.AFSTOCK='N') " & _
            " AND LTRIM(RTRIM(B.STCODIGO))<>'TEXTO' Order by AFamilia"
    
    Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount > 0 Then
        If Text1 = "" And Text2 = "" Then
            Adodc3.MoveFirst
            tex1 = IIf(IsNull(Adodc3("AFamilia")), "", Adodc3("AFamilia"))
            Va1 = IIf(IsNull(Adodc3("Fam_Nombre")), "", Adodc3("Fam_Nombre"))
            Adodc3.MoveLast
            tex2 = IIf(IsNull(Adodc3("AFamilia")), "", Adodc3("AFamilia"))
            Va2 = IIf(IsNull(Adodc3("Fam_Nombre")), "", Adodc3("Fam_Nombre"))
        End If
    End If
    Adodc3.Close
    
    Cadfiltro = ""
    G_SQL = "SELECT MAEART.AFAMILIA, TABALM.TADESCRI, MAEART.ADESCRI, STKART.STCODIGO, " & _
            " STKART.STSKDIS, MAEART.AUNIDAD, STKART.STALMA, familia.FAM_NOMBRE , STKART.STKPREULT, STKART.STKPREPRO," & _
            " STKART.STKPREPROUS into " & TempoTAB & " FROM MAEART INNER JOIN STKART ON MAEART.ACODIGO = STKART.STCODIGO " & _
            " INNER JOIN TABALM ON STKART.STALMA = TABALM.TAALMA INNER JOIN FAMILIA " & _
            " ON MAEART.AFAMILIA = FAMILIA.FAM_CODIGO "
        If Me.Check1.Value = 0 Then G_SQL = G_SQL + " WHERE STKART.STALMA='" & almacen & "'"
    
    If Chkstokcero.Value = 1 Then Cadfiltro = Cadfiltro & " and STKART.STskdis<>0 "
    
    If OpTodos.Value Then
        If Check1.Value = 0 Then 'SI NO ES consolidado
            If Chkprecio = 0 Then 'Si sin costo ...
                TituloRpt = "Inv080 -- Control de Inventarios"
                NombreRpt = "\inv080.rpt"
            Else ' Sino, con costo ...
                TituloRpt = "Inv070 -- Control de Inventarios"
                NombreRpt = "\inv070.rpt"
                bMon = True
            End If
            
        Else 'SI ES  consolidado ...
            If Chkprecio = 0 Then 'Si es sin costo ...
                TituloRpt = "Inv089 -- Control de Inventarios"
                NombreRpt = "\inv089.rpt"
            Else ' Sino, con costo ...
                TituloRpt = "Inv067 -- Control de Inventarios"
                NombreRpt = "\inv067.rpt"
                bMon = True
            End If
        End If
     
        If Combo2.ListIndex = 1 Then
            Rem mvv CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
        Else
            Rem mvv CrystalReport1.SortFields(0) = "+{STKART.STCODIGO}"
        End If
        
        arrfor(0) = "emp = '" & VGNemp & "'"
        arrfor(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
        arrfor(2) = "campoini = '" & tex1 & "'"
        arrfor(3) = "campofin = '" & tex2 & "'"
        arrfor(4) = "detaini = '" & Va1 & "'"
        arrfor(5) = "detafin = '" & Va2 & "'"
        If bMon Then arrfor(6) = "xMoneda = '" & cboMon.text & "'" Else arrfor(6) = ""
        arrparam(0) = cConexCom.DefaultDatabase
        arrparam(1) = TempoTAB
        If ExisteElem(0, cConexCom, TempoTAB) Then cConexCom.Execute "DROP TABLE " & TempoTAB
        cConexCom.Execute G_SQL + Cadfiltro
        Call ImpresionRpt(Me.CrystalReport1, NombreRpt, arrfor, arrparam, , TituloRpt)
        Exit Sub
    End If
'***** Fin todos los articulos .......

    If Trim(Text1) = "" Then
          MsgBox "Ingrese el codigo de la familia", vbExclamation, "Error"
          OpRango = True
          Text1.SetFocus
          Exit Sub
    End If
    Codigo2 = UCase(Trim(Text2))
    Codigo1 = UCase(Trim(Text1))

    RSQL = "Select Fam_Nombre from FAMILIA Where Fam_Codigo='" & Codigo1 & "'"
    Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount = 1 Then
        Va1 = Adodc3("Fam_Nombre")
    End If
    Adodc3.Close
  
    RSQL = "Select Fam_Nombre from FAMILIA Where Fam_Codigo='" & Codigo2 & "'"
    Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount = 1 Then
        Va2 = Adodc3("Fam_Nombre")
    End If
    Adodc3.Close

    If Check1.Value = 1 Then 'Si es consolidado sin marca "con costo".....
        If Chkprecio = 0 Then 'Si es sin costo ...
            TituloRpt = "Inv089 -- Control de Inventarios"
            NombreRpt = "\inv089.rpt"
        Else
            TituloRpt = "Inv067 -- Control de Inventarios"
            NombreRpt = "\inv067.rpt"
            bMon = True
        End If
        
        If Text2 <> "" Then
            Cadfiltro = Cadfiltro + " AND (MAEART.AFAMILIA >= '" & Codigo1 & "' AND MAEART.AFAMILIA <= '" & Codigo2 & "')"
        Else
            Codigo2 = Codigo1: Va2 = Va1
            Cadfiltro = Cadfiltro + " AND (MAEART.AFAMILIA >= '" & Codigo1 & "' AND MAEART.AFAMILIA <= '" & Codigo2 & "')  "
        End If
    Else 'Si es con costo sin marcar "consolidado".....
        If Chkprecio = 0 Then
            TituloRpt = "Inv080 -- Control de Inventarios"
            NombreRpt = "\inv080.rpt"
        Else
            TituloRpt = "Inv070 -- Control de Inventarios"
            NombreRpt = "\inv070.rpt"
            bMon = True
        End If
        
        If Text2 <> "" Then
            Cadfiltro = Cadfiltro + " and (MAEART.AFAMILIA >= '" & Codigo1 & "' AND MAEART.AFAMILIA <= '" & Codigo2 & "')  "
        Else
            Codigo2 = Codigo1: Va2 = Va1
            Cadfiltro = Cadfiltro + " and (MAEART.AFAMILIA >= '" & Codigo1 & "' AND  MAEART.AFAMILIA <= '" & Codigo2 & "') "
        End If
    End If
    
    If Combo2.ListIndex = 1 Then
        Rem MVV CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
    Else
        Rem MVV ordenrerporte = "+{STKART.STCODIGO}"
    End If
    
    arrfor(0) = "emp = '" & VGNemp & "'"
    arrfor(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    arrfor(2) = "campoini = '" & Codigo1 & "'"
    arrfor(3) = "campofin = '" & Codigo2 & "'"
    arrfor(4) = "detaini = '" & Va1 & "'"
    arrfor(5) = "detafin = '" & Va2 & "'"
    If bMon Then arrfor(6) = "xMoneda = '" & cboMon.text & "'" Else arrfor(6) = ""
    arrparam(0) = cConexCom.DefaultDatabase
    arrparam(1) = TempoTAB
    
    If ExisteElem(0, cConexCom, TempoTAB) Then cConexCom.Execute "DROP TABLE " & TempoTAB
    cConexCom.Execute G_SQL + Cadfiltro
    Call ImpresionRpt(Me.CrystalReport1, NombreRpt, arrfor, arrparam, , TituloRpt)
Exit Sub
Mensaje:
    Me.MousePointer = 1
    Captura_error
End Sub

Private Sub Imprimir3()
On Error GoTo Mensaje
Dim cadena        As String
Dim Codigo1       As String
Dim Codigo2       As String
Dim RSQL          As String
Dim tex1          As String, tex2         As String
Dim Va1           As String, Va2          As String

If Trim(Text3) = "" Then
      MsgBox "Ingrese el código de la familia", vbExclamation, "Error"
      Text3.SetFocus
      Exit Sub
End If

Set Adodc3 = New ADODB.Recordset
    CrystalReport1.Reset
If OpTodos.Value Then
    RSQL = "Select AModelo,Lin_Nombre from "
    RSQL = RSQL & "((MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo) "
    RSQL = RSQL & "Left Join LINEAS C on A.AFamilia=C.Fam_Codigo and A.Amodelo=C.Lin_Codigo) "
    RSQL = RSQL & "Where AFamilia='" & Text3.text & "' and Stalma='" & almacen & "' Order By Amodelo"
    
    Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount > 0 Then
      If Text1 = "" And Text2 = "" Then
        Adodc3.MoveFirst
        tex1 = IIf(IsNull(Adodc3("AModelo")), "", Adodc3("AModelo"))
        Va1 = IIf(IsNull(Adodc3("Lin_Nombre")), "", Adodc3("Lin_Nombre"))
        Adodc3.MoveLast
        tex2 = IIf(IsNull(Adodc3("AModelo")), "", Adodc3("AModelo"))
        Va2 = IIf(IsNull(Adodc3("Lin_Nombre")), "", Adodc3("Lin_Nombre"))
      End If
    End If
    Adodc3.Close
    
    If Check1.Value = 0 Then
        If Chkprecio = 0 Then
                CrystalReport1.WindowTitle = "Inv085 -- Control de Inventarios"
                CrystalReport1.ReportFileName = cRutP & "\inv085.rpt"
        Else
                CrystalReport1.WindowTitle = "Inv072 -- Control de Inventarios"
                CrystalReport1.ReportFileName = cRutP & "\inv072.rpt"
                bMon = True
        End If
        cadena = "{STKART.STALMA}='" & almacen & "' and {MAEART.AFAMILIA}='" & Text3.text & "'"
        If Chkstokcero.Value = 1 Then
           cadena = cadena & " and {STKART.STskdis}<>0 "
        End If
        
    Else
        If Chkprecio = 0 Then
                CrystalReport1.WindowTitle = "Inv084 -- Control de Inventarios"
                CrystalReport1.ReportFileName = cRutP & "\inv084.rpt"
        Else
                CrystalReport1.WindowTitle = "Inv069 -- Control de Inventarios"
                CrystalReport1.ReportFileName = cRutP & "\inv069.rpt"
                bMon = True
        End If
       
        cadena = "{MAEART.AFAMILIA}='" & Text3.text & "'"
        If Chkstokcero.Value = 1 Then
           cadena = "{STKART.STskdis}<>0 "
        End If

     End If
 
     Rem mvv Ubi_Tab CrystalReport1
     CrystalReport1.DiscardSavedData = True
     CrystalReport1.Destination = crptToWindow
     CrystalReport1.SelectionFormula = cadena
     CrystalReport1.WindowShowPrintBtn = True
     CrystalReport1.WindowShowRefreshBtn = True
     CrystalReport1.WindowShowSearchBtn = True
     CrystalReport1.WindowShowPrintSetupBtn = True
     CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
     CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
     CrystalReport1.Formulas(2) = "campoini = '" & tex1 & "'"
     CrystalReport1.Formulas(3) = "campofin = '" & tex2 & "'"
     CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
     CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
     If bMon Then CrystalReport1.Formulas(6) = "xMoneda = '" & cboMon.text & "'" Else CrystalReport1.Formulas(6) = ""
     If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
     Exit Sub
End If
If Trim(Text1) = "" Then
      MsgBox "Ingrese el codigo de la Línea", vbExclamation, "Error"
      OpRango = True
      Text1.SetFocus
      Exit Sub
End If
Codigo2 = UCase(Trim(Text2))
Codigo1 = UCase(Trim(Text1))
  
  RSQL = "Select Lin_Nombre from LINEAS Where Lin_Codigo='" & Codigo1 & "'"
  Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va1 = Adodc3("Lin_Nombre")
  End If
  Adodc3.Close
  
  RSQL = "Select Lin_Nombre from LINEAS Where Lin_Codigo='" & Codigo2 & "'"
  Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va2 = Adodc3("Lin_Nombre")
  End If
  Adodc3.Close

If Check1.Value = 1 Then
        If Chkprecio = 0 Then
              CrystalReport1.WindowTitle = "Inv084 -- Control de Inventarios"
              CrystalReport1.ReportFileName = cRutP & "\inv084.rpt"
        Else
              CrystalReport1.WindowTitle = "Inv069 -- Control de Inventarios"
              CrystalReport1.ReportFileName = cRutP & "\inv069.rpt"
              bMon = True
        End If
        Rem MVV Ubi_Tab CrystalReport1
        If Text2 <> "" Then
                Rem MVV cadena = "({MAEART.AMODELO} in '" & Codigo1 & "' to '" & Codigo2 & "')  and {STKART.STCODIGO}={MAEART.ACODIGO} and {MAEART.AFAMILIA} = '" & Text3.text & "'"
                cadena = "{MAEART.AMODELO} in '" & Codigo1 & "' to '" & Codigo2 & "'  and {MAEART.AFAMILIA} = '" & Text3.text & "'"
        Else
                Codigo2 = Codigo1: Va2 = Va1
                Rem MVV cadena = "({MAEART.AMODELO} in '" & Codigo1 & "' to '" & Codigo2 & "')  and {STKART.STCODIGO}={MAEART.ACODIGO} and {MAEART.AFAMILIA} = '" & Text3.text & "'"
                cadena = "{MAEART.AMODELO} in '" & Codigo1 & "' to '" & Codigo2 & "'  and {MAEART.AFAMILIA} = '" & Text3.text & "'"
        End If
Else
        If Chkprecio = 0 Then
                CrystalReport1.WindowTitle = "Inv085 -- Control de Inventarios"
                CrystalReport1.ReportFileName = cRutP & "\inv085.rpt"
        Else
                CrystalReport1.WindowTitle = "Inv072 -- Control de Inventarios"
                CrystalReport1.ReportFileName = cRutP & "\inv072.rpt"
                bMon = True
        End If
        
        If Text2 <> "" Then
                Rem MVV cadena = "{STKART.STALMA}='" & almacen & "'  and {MAEART.AFAMILIA} = '" & Text3.text & "' and ({MAEART.AMODELO} >= '" & Codigo1 & "' and {MAEART.AMODELO}<= '" & Codigo2 & "')  "
                cadena = " and {MAEART.AFAMILIA} = '" & Text3.text & "' and {MAEART.AMODELO} >= '" & Codigo1 & "' and {MAEART.AMODELO}<= '" & Codigo2 & "'"
                If Me.Check1.Value = 0 Then cadena = cadena + " and {STKART.STALMA}='" & almacen & "'"
        Else
                Codigo2 = Codigo1: Va2 = Va1
                cadena = "and {MAEART.AFAMILIA} = '" & Text3.text & "' and {MAEART.AMODELO} >= '" & Codigo1 & "' and {MAEART.AMODELO}<= '" & Codigo2 & "'"
                If Me.Check1.Value = 0 Then cadena = cadena + " and {STKART.STALMA}='" & almacen & "'"
        End If
End If
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    If Combo2.ListIndex = 1 Then CrystalReport1.SortFields(0) = "+{LINEAS.LIN_NOMBRE}"
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & Codigo1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & Codigo2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
    If bMon Then CrystalReport1.Formulas(6) = "xMoneda = '" & cboMon.text & "'" Else CrystalReport1.Formulas(6) = ""
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1

Exit Sub
Mensaje:
    Captura_error
End Sub
Private Sub Imprimir4()
Dim cadena As String
Dim Codigo1 As String
Dim Codigo2 As String
Dim RSQL As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String
Dim arrfor(7), arrparam(2)
    If Trim(Text3) = "" Then
        MsgBox "Ingrese el código de la familia", vbExclamation, "Error"
        Text3.SetFocus
        Exit Sub
    ElseIf Trim(Text4) = "" Then
        MsgBox "Ingrese el código de la Línea", vbExclamation, "Error"
        Text4.SetFocus
        Exit Sub
    End If

    Set Adodc3 = New ADODB.Recordset

    If OpTodos.Value Then
        RSQL = "Select AGrupo,Acodigo,Gru_Nombre from"
        RSQL = RSQL & "((MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo)"
        RSQL = RSQL & "Left Join GRUPO C on A.AFamilia=C.Fam_Codigo and A.Amodelo=C.Lin_Codigo and A.Agrupo=C.Gru_Codigo) "
        RSQL = RSQL & "Where AFamilia='" & Text3.text & "' and Amodelo='" & Text4.text & "'"
        If Me.Check1.Value = 0 Then RSQL = RSQL + " and Stalma='" & almacen & "'"
        RSQL = RSQL + " Order by Agrupo"
        
        Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
        If Adodc3.RecordCount > 0 Then
            If Text1 = "" And Text2 = "" Then
                Adodc3.MoveFirst
                tex1 = IIf(IsNull(Adodc3("AGrupo")), "", Adodc3("AGrupo"))
                Va1 = IIf(IsNull(Adodc3("Gru_Nombre")), "", Adodc3("Gru_Nombre"))
                Adodc3.MoveLast
                tex2 = IIf(IsNull(Adodc3("AGrupo")), "", Adodc3("AGrupo"))
                Va2 = IIf(IsNull(Adodc3("Gru_Nombre")), "", Adodc3("Gru_Nombre"))
            End If
        End If
        Adodc3.Close
        
        If Check1.Value = 0 Then
            If Chkprecio = 0 Then
                CrystalReport1.WindowTitle = "Inv083 -- Control de Inventarios"
                CrystalReport1.ReportFileName = cRutP & "\inv083.rpt"
            Else
                CrystalReport1.WindowTitle = "Inv071 -- Control de Inventarios"
                CrystalReport1.ReportFileName = cRutP & "\inv071.rpt"
                bMon = True
            End If
            cadena = "{STKART.STALMA}='" & almacen & "' and {MAEART.AFAMILIA}='" & Text3.text & "' and {MAEART.AMODELO}='" & Text4.text & "'"
        Else
            If Chkprecio = 0 Then
                CrystalReport1.WindowTitle = "Inv082 -- Control de Inventarios"
                CrystalReport1.ReportFileName = cRutP & "\inv082.rpt"
            Else
                CrystalReport1.WindowTitle = "Inv068 -- Control de Inventarios"
                CrystalReport1.ReportFileName = cRutP & "\inv068.rpt"
                bMon = True
            End If
            cadena = "{MAEART.AFAMILIA}='" & Text3.text & "' and {MAEART.AMODELO}='" & Text4.text & "'"
        End If
        
        If Chkstokcero.Value = 1 Then
            cadena = cadena & " and {STKART.STskdis}<>0 "
        End If
        CrystalReport1.DiscardSavedData = True
        CrystalReport1.Destination = crptToWindow
        CrystalReport1.SelectionFormula = cadena
        CrystalReport1.WindowShowPrintBtn = True
        CrystalReport1.WindowShowRefreshBtn = True
        CrystalReport1.WindowShowSearchBtn = True
        CrystalReport1.WindowShowPrintSetupBtn = True
        CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
        CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
        CrystalReport1.Formulas(2) = "campoini = '" & tex1 & "'"
        CrystalReport1.Formulas(3) = "campofin = '" & tex2 & "'"
        CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
        CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
        If bMon Then CrystalReport1.Formulas(6) = "xMoneda = '" & cboMon.text & "'" Else CrystalReport1.Formulas(6) = ""
        If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
        Exit Sub
    End If

    If Trim(Text1) = "" Then
        MsgBox "Ingrese el código del artículo", vbExclamation, "Error"
        OpRango = True
        Text1.SetFocus
        Exit Sub
    End If
    
    Codigo2 = UCase(Trim(Text2))
    Codigo1 = UCase(Trim(Text1))

    RSQL = "Select Gru_Nombre from GRUPO Where Gru_Codigo='" & Codigo1 & "'"
    Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount = 1 Then
        Va1 = Adodc3("Gru_Nombre")
    End If
    Adodc3.Close
  
    RSQL = "Select Gru_Nombre from Grupo Where Gru_Codigo='" & Codigo2 & "'"
    Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount = 1 Then
        Va2 = Adodc3("Gru_Nombre")
    End If
    Adodc3.Close

    If Check1.Value = 1 Then
        If Chkprecio = 0 Then
            CrystalReport1.WindowTitle = "Inv082 -- Control de Inventarios"
            CrystalReport1.ReportFileName = cRutP & "\inv082.rpt"
        Else
            CrystalReport1.WindowTitle = "Inv068 -- Control de Inventarios"
            CrystalReport1.ReportFileName = cRutP & "\inv068.rpt"
            bMon = True
        End If
        
        If Text2 <> "" Then
            cadena = " ({MAEART.AGRUPO} in '" & Codigo1 & "' to '" & Codigo2 & "')  and {STKART.STCODIGO}={MAEART.ACODIGO}"
            cadena = cadena & " and {MAEART.AFAMILIA}='" & Text3.text & "' and {MAEART.AMODELO}='" & Text4.text & "'"
        Else
            Codigo2 = Codigo1: Va2 = Va1
            cadena = "({MAEART.AGRUPO} in '" & Codigo1 & "' to '" & Codigo2 & "')  and {STKART.STCODIGO}={MAEART.ACODIGO}"
            cadena = cadena & " and {MAEART.AFAMILIA}='" & Text3.text & "' and {MAEART.AMODELO}='" & Text4.text & "'"
        End If
    Else
        If Chkprecio = 0 Then
            CrystalReport1.ReportFileName = cRutP & "\inv083.rpt"
        Else
            CrystalReport1.ReportFileName = cRutP & "\inv071.rpt"
            bMon = True
        End If
        
        If Text2 <> "" Then
            cadena = " ({MAEART.AGRUPO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
            cadena = cadena & " and {MAEART.AFAMILIA}='" & Text3.text & "' and {MAEART.AMODELO}='" & Text4.text & "'"
            If Me.Check1.Value = 0 Then cadena = cadena + " and {STKART.STALMA}='" & almacen & "'"
            
'            cadena = "{al_stkartgrup.STALMA}='" & almacen & "'  and ({al_stkartgrup.AGRUPO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
'            cadena = cadena & " and {al_stkartgrup.AFAMILIA}='" & Text3.text & "' and {al_stkartgrup.AMODELO}='" & Text4.text & "'"
        Else
            Codigo2 = Codigo1: Va2 = Va1
            cadena = "({MAEART.AGRUPO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
            cadena = cadena & " and {MAEART.AFAMILIA}='" & Text3.text & "' and {MAEART.AMODELO}='" & Text4.text & "'"
            If Me.Check1.Value = 0 Then cadena = cadena + " and {STKART.STALMA}='" & almacen & "'"
        End If
    End If
    If Chkstokcero.Value = 1 Then
        cadena = cadena & " and {STKART.STskdis}<>0 "
    End If
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & tex1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & tex2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
    If bMon Then CrystalReport1.Formulas(6) = "xMoneda = '" & cboMon.text & "'" Else CrystalReport1.Formulas(6) = ""
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End Sub

Private Sub Text3_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
If Option3.Value Or Option4.Value Then
         Adodc2.Open "Select FAM_CODIGO,FAM_NOMBRE,FAM_CTA from FAMILIA ", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select FAM_CODIGO,FAM_NOMBRE,FAM_CTA from FAMILIA "
         frmReferencia.Label1.Caption = "Familias"
         frmReferencia.Show vbModal
         Rem MVV Adodc2.Close
         If vGUtil(1) <> "" Then
                 Text3 = (vGUtil(1))
         End If
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text3_DblClick
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Text3 <> "" Then
     If Text4.Visible = True Then Text4.SetFocus Else OpTodos.SetFocus
  End If
End Sub

Private Sub Text4_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
If Option4.Value Then
         Adodc2.Open "Select LIN_CODIGO,LIN_NOMBRE from LINEAS WHERE FAM_CODIGO ='" & Text3 & "' ", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select LIN_CODIGO,LIN_NOMBRE from LINEAS WHERE FAM_CODIGO ='" & Text3 & "' "
         frmReferencia.Label1.Caption = "Líneas"
         frmReferencia.Show vbModal
         Rem MVV Adodc2.Close
         If vGUtil(1) <> "" Then
            Text4 = (vGUtil(1))
         End If
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text4_DblClick
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Text4 <> "" Then
     OpTodos.SetFocus
  End If
End Sub

Private Sub Imprimir5()
Dim Codigo1           As String
Dim Codigo2           As String
Dim cadena            As String
Dim RSQL              As String
Dim tex1              As String, tex2     As String
Dim Va1               As String, Va2      As String

Codigo1 = UCase(Trim(Text1))
Set Adodc3 = New ADODB.Recordset

If OpTodos.Value Then
    RSQL = "Select ACodigo,Adescri from "
    RSQL = RSQL & "MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo "
    RSQL = RSQL & "Where Stalma='" & almacen & "' Order by Acodigo"
    
    Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
    If Adodc3.RecordCount > 0 Then
      If Text1 = "" And Text2 = "" Then
        Adodc3.MoveFirst
        tex1 = IIf(IsNull(Adodc3("ACodigo")), "", Adodc3("ACodigo"))
        Va1 = IIf(IsNull(Adodc3("ADescri")), "", Adodc3("Adescri"))
        Adodc3.MoveLast
        tex2 = IIf(IsNull(Adodc3("ACodigo")), "", Adodc3("ACodigo"))
        Va2 = IIf(IsNull(Adodc3("ADescri")), "", Adodc3("ADescri"))
      End If
    End If
    Adodc3.Close

    If Check1.Value = 0 Then
            If ChkSerie = 1 Then
                    CrystalReport1.WindowTitle = "Inv136 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = cRutP & "\inv136.rpt"
            ElseIf Chkprecio = 0 Then
                    CrystalReport1.WindowTitle = "Inv138 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = cRutP & "\inv138.rpt"
            Else
                    CrystalReport1.WindowTitle = "Inv140-- Control de Inventarios"
                    CrystalReport1.ReportFileName = cRutP & "\inv140.rpt"
            End If
            cadena = "{STKART.STALMA}='" & almacen & "' "
            If Chkstokcero.Value = 1 Then
               cadena = cadena & " and {STKART.STskdis}<>0 "
            End If
            
    Else
            If ChkSerie = 1 Then
                    CrystalReport1.WindowTitle = "Inv137 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = cRutP & "\inv137.rpt"
            ElseIf Chkprecio = 0 Then
                    CrystalReport1.WindowTitle = "Inv139 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = cRutP & "\inv139.rpt"
            Else
                    CrystalReport1.WindowTitle = "Inv140 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = cRutP & "\inv141.rpt"
            End If
            If Chkstokcero.Value = 1 Then
               cadena = "{STKART.STskdis}<>0 "
            End If
            
    End If
    Ubi_Tab CrystalReport1
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = cadena
    If Combo2.ListIndex = 1 Then
      CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
    Else
      CrystalReport1.SortFields(0) = "+{STKART.STCODIGO}"
    End If
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & tex1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & tex2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
    Exit Sub
End If

If Text1 = "" Then
        MsgBox "Ingrese el codigo", vbExclamation, "Error"
        Text1.SetFocus
        Exit Sub
End If
  Codigo2 = UCase(Trim(Text2))
  RSQL = "Select ADescri from MAEART Where ACodigo='" & Codigo1 & "'"
  Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va1 = Adodc3("Adescri")
  End If
  Adodc3.Close
  
  RSQL = "Select ADescri from MAEART Where ACodigo='" & Codigo2 & "'"
  Adodc3.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
  If Adodc3.RecordCount = 1 Then
    Va2 = Adodc3("Adescri")
  End If
  Adodc3.Close

If Option1.Value Then
    If Check1.Value = 1 Then
            If ChkSerie = 1 Then
                    CrystalReport1.WindowTitle = "Inv137 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = cRutP & "\inv137.rpt"
            ElseIf Chkprecio = 0 Then
                    CrystalReport1.WindowTitle = "Inv139 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = cRutP & "\inv139.rpt"
            Else
                    CrystalReport1.WindowTitle = "Inv141 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = cRutP & "\inv141.rpt"
            End If
            If Text2 <> "" Then
                    Codigo2 = Text2
                    cadena = "({STKART.STCODIGO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
            Else
                    Codigo2 = Codigo1: Va2 = Va1
                    cadena = "{STKART.STCODIGO} = '" & Codigo1 & "' "
            End If
    Else
            If ChkSerie = 1 Then
                    CrystalReport1.WindowTitle = "Inv136 -- Control de Inventarios"
                    CrystalReport1.ReportFileName = cRutP & "\inv136.rpt"
            ElseIf Chkprecio = 0 Then
                    CrystalReport1.ReportFileName = cRutP & "\inv138.rpt"
                    CrystalReport1.WindowTitle = "Inv138 -- Control de Inventarios"
            Else
                    CrystalReport1.ReportFileName = cRutP & "\inv040.rpt"
                    CrystalReport1.WindowTitle = "Inv140 -- Control de Inventarios"
            End If
            If Text2 <> "" Then
                    Codigo2 = Text2
                    cadena = " {STKART.STALMA}='" & almacen & "' and ({STKART.STCODIGO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
            Else
                    Codigo2 = Codigo1: Va2 = Va1
                    cadena = "{STKART.STALMA}='" & almacen & "' and {STKART.STCODIGO} = '" & Codigo1 & "' "
            End If
    End If
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    If Combo2.ListIndex = 1 Then
      CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
    Else
      CrystalReport1.SortFields(0) = "+{STKART.STCODIGO}"
    End If
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "campoini = '" & Codigo1 & "'"
    CrystalReport1.Formulas(3) = "campofin = '" & Codigo2 & "'"
    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
End If
End Sub


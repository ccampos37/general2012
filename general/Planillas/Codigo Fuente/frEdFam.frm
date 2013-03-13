VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frEdFam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Derechohabiente"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "frEdFam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6870
   Begin VB.CommandButton CmdImp 
      Caption         =   "&Imprimir"
      Height          =   315
      Left            =   4515
      TabIndex        =   46
      Top             =   5850
      Width           =   1260
   End
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   2790
      TabIndex        =   42
      Top             =   5835
      Width           =   1260
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   1035
      TabIndex        =   41
      Top             =   5835
      Width           =   1260
   End
   Begin VB.Frame Frame2 
      Caption         =   "Domicilio Propio"
      Enabled         =   0   'False
      Height          =   2025
      Left            =   120
      TabIndex        =   26
      Top             =   3735
      Width           =   6630
      Begin Crystal.CrystalReport Reporte 
         Left            =   405
         Top             =   645
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin AplisetControlText.Aplitext xUbigeo 
         Height          =   315
         Left            =   1470
         TabIndex        =   40
         Top             =   1620
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   556
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xReferencia 
         Height          =   315
         Left            =   1470
         TabIndex        =   38
         Top             =   1275
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   556
         MaxLength       =   40
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xZona 
         Height          =   315
         Left            =   3585
         TabIndex        =   36
         Top             =   930
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         MaxLength       =   20
         Text            =   ""
      End
      Begin VB.ComboBox xTipoZona 
         Height          =   315
         ItemData        =   "frEdFam.frx":030A
         Left            =   1470
         List            =   "frEdFam.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   930
         Width           =   2100
      End
      Begin AplisetControlText.Aplitext xInterior 
         Height          =   315
         Left            =   4215
         TabIndex        =   33
         Top             =   585
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         MaxLength       =   4
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xNumero 
         Height          =   315
         Left            =   2205
         TabIndex        =   31
         Top             =   585
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         MaxLength       =   4
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xNombreVia 
         Height          =   315
         Left            =   3600
         TabIndex        =   29
         Top             =   225
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         MaxLength       =   20
         Text            =   ""
      End
      Begin VB.ComboBox xTipoVia 
         Height          =   315
         ItemData        =   "frEdFam.frx":0403
         Left            =   1470
         List            =   "frEdFam.frx":042B
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   240
         Width           =   2100
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Ubigeo (I.N.E.I.)"
         Height          =   195
         Index           =   15
         Left            =   105
         TabIndex        =   39
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   37
         Top             =   1335
         Width           =   780
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de Zona"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   34
         Top             =   990
         Width           =   1200
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Interior"
         Height          =   195
         Index           =   17
         Left            =   3615
         TabIndex        =   32
         Top             =   645
         Width           =   480
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Index           =   16
         Left            =   1470
         TabIndex        =   30
         Top             =   630
         Width           =   555
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de Vía"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Principal"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   6630
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   3720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3375
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.PictureBox Ole1 
         DataSource      =   "Data1"
         Height          =   1335
         Left            =   4020
         ScaleHeight     =   1275
         ScaleWidth      =   825
         TabIndex        =   47
         Top             =   555
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmDescargaFoto 
         Caption         =   "Descargar Foto"
         Height          =   270
         Left            =   5265
         TabIndex        =   45
         Top             =   1770
         Width           =   1245
      End
      Begin VB.CommandButton cmCargaFoto 
         Caption         =   "Cargar &Foto"
         Height          =   270
         Left            =   5265
         TabIndex        =   44
         Top             =   1425
         Width           =   1245
      End
      Begin VB.CheckBox xIDP 
         Caption         =   "Domicilio Propio"
         Height          =   210
         Left            =   4440
         TabIndex        =   25
         Top             =   3195
         Width           =   1440
      End
      Begin AplisetControlText.Aplitext xDocIncap 
         Height          =   315
         Left            =   1470
         TabIndex        =   24
         Top             =   3135
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         MaxLength       =   20
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin VB.ComboBox xMotivoBaja 
         Height          =   315
         ItemData        =   "frEdFam.frx":04BB
         Left            =   4425
         List            =   "frEdFam.frx":04C5
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2805
         Width           =   2085
      End
      Begin VB.ComboBox xSituacion 
         Height          =   315
         ItemData        =   "frEdFam.frx":04E5
         Left            =   1470
         List            =   "frEdFam.frx":04EF
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2805
         Width           =   2100
      End
      Begin AplisetControlText.Aplitext xCarta 
         Height          =   315
         Left            =   4425
         TabIndex        =   17
         Top             =   2475
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         MaxLength       =   20
         Locked          =   -1  'True
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin VB.ComboBox xVinculo 
         Height          =   315
         ItemData        =   "frEdFam.frx":0509
         Left            =   1470
         List            =   "frEdFam.frx":0519
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2475
         Width           =   2085
      End
      Begin VB.ComboBox xSexo 
         Height          =   315
         ItemData        =   "frEdFam.frx":054D
         Left            =   4425
         List            =   "frEdFam.frx":0557
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2145
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker xFechaNac 
         Height          =   315
         Left            =   1470
         TabIndex        =   12
         Top             =   2145
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24838145
         CurrentDate     =   36656
      End
      Begin AplisetControlText.Aplitext xNumDoc 
         Height          =   285
         Left            =   1470
         TabIndex        =   10
         Top             =   1845
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   503
         MaxLength       =   15
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin VB.ComboBox xTipoDoc 
         Height          =   315
         ItemData        =   "frEdFam.frx":0576
         Left            =   1470
         List            =   "frEdFam.frx":0592
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1515
         Width           =   2085
      End
      Begin AplisetControlText.Aplitext xNombre 
         Height          =   285
         Left            =   1470
         TabIndex        =   6
         Top             =   1215
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xApeMat 
         Height          =   285
         Left            =   1470
         TabIndex        =   5
         Top             =   915
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xApePat 
         Height          =   285
         Left            =   1470
         TabIndex        =   2
         Top             =   615
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   5910
         Picture         =   "frEdFam.frx":066E
         Top             =   330
         Width           =   480
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   135
         TabIndex        =   43
         Top             =   240
         Width           =   3405
      End
      Begin VB.Image xFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   1755
         Left            =   3795
         Stretch         =   -1  'True
         Top             =   285
         Width           =   1395
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Incapacidad"
         Height          =   195
         Index           =   8
         Left            =   165
         TabIndex        =   23
         Top             =   3210
         Width           =   1275
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Motivo"
         Height          =   195
         Index           =   11
         Left            =   3780
         TabIndex        =   21
         Top             =   2865
         Width           =   480
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Situación"
         Height          =   195
         Index           =   7
         Left            =   165
         TabIndex        =   19
         Top             =   2910
         Width           =   660
      End
      Begin VB.Label l1 
         Caption         =   "Carta"
         Height          =   195
         Index           =   10
         Left            =   3780
         TabIndex        =   18
         Top             =   2535
         Width           =   435
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Vínculo"
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   15
         Top             =   2565
         Width           =   555
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Sexo"
         Height          =   195
         Index           =   9
         Left            =   3780
         TabIndex        =   13
         Top             =   2205
         Width           =   360
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Nac."
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   11
         Top             =   2205
         Width           =   840
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Numero Doc."
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   8
         Top             =   1890
         Width           =   945
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   7
         Top             =   1575
         Width           =   1185
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   4
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Materno"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   3
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Paterno"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   660
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frEdFam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim RSDEREC As New ADODB.Recordset

Private Sub CMACEPTAR_CLICK()
    Dim MAXCODDER As Integer
    If Not VALIDAR Then Exit Sub
    With RSDEREC
        If VPTAREA = "NUEVO" Then
            .AddNew
            !CODTRAB = frFamily.xTrab.Tag
'            MAXCODDER = ESNULO(GetValor("SELECT MAX(CODDER) FROM FAMILIAR", DBSYSTEM), 0) + 1
'            !CODDER = MAXCODDER
        End If
        !TIPODOC = xTipoDoc.ListIndex
        !NUMDOC = xNumDoc.Text
        !ApePat = xApePat.Text
        !ApeMat = xApeMat.Text
        !NOMBRE = xNombre.Text
        !FechaNac = xFechaNac.Value
        !Sexo = xSexo.ListIndex
        !VINCULO = xVinculo.ListIndex
        !CARTA = xCarta.Text
        !SITUACION = xSituacion.ListIndex
        !MOTIVOBAJA = xMotivoBaja.ListIndex
        !DOCINCAP = xDocIncap.Text
        !IDP = xIDP.Value
        !NOMBREVIA = IIf(xIDP.Value = 0, "", xNombreVia.Text)
        !NUMERO = IIf(xIDP.Value = 0, "", xNumero.Text)
        !INTERIOR = IIf(xIDP.Value = 0, "", xInterior.Text)
        !ZONA = IIf(xIDP.Value = 0, "", xZona.Text)
        !REFERENCIA = IIf(xIDP.Value = 0, "", xReferencia.Text)
        !TIPOVIA = IIf(xIDP.Value = 0, 0, xTipoVia.ListIndex)
        !TIPOZONA = IIf(xIDP.Value = 0, 0, xTipoZona.ListIndex)
        !UBIGEO = IIf(xIDP.Value = 0, "", xUbigeo.Tag)
        .Update
            On Error GoTo ERRFOTO
            If xFoto.Tag <> "" Then
                FileCopy xFoto.Tag, (REGSISTEMA.PATHFOTOS & "\" & frFamily.xTrab.Tag & !CODDER & ".FTD")
            Else
                Kill (REGSISTEMA.PATHFOTOS & "\" & frFamily.xTrab.Tag & !CODDER & ".FTD")
            End If
    End With
    Unload Me
    Exit Sub
ERRFOTO:
    Resume Next
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Public Function VALIDAR() As Boolean
    VALIDAR = False
    If xApePat.Text = "" Then
        MsgBox "DEBE INGRESAR UN APELLIDO PATERNO", vbCritical
        xApePat.SetFocus
        Exit Function
    End If
    If xApeMat.Text = "" Then
        MsgBox "DEBE INGRESAR UN APELLIDO MATERNO", vbCritical
        xApeMat.SetFocus
        Exit Function
    End If
    If xNombre.Text = "" Then
        MsgBox "DEBE INGRESAR UN NOMBRE", vbCritical
        xNombre.SetFocus
        Exit Function
    End If
    If xNumDoc.Text = "" Then
        MsgBox "DEBE INGRESAR UN NÚMERO DE DOCUMENTO VÁLIDO", vbCritical
        xNumDoc.SetFocus
        Exit Function
    End If
    If xCarta.Text = "" And xVinculo.ListIndex = 3 Then
        MsgBox "SI EL VINCULO ES GESTANTE, DEBE TENER UN NÚMERO DE CARTA DE ATENCIÓN MÉDICA VÁLIDO", vbCritical
        xCarta.SetFocus
        Exit Function
    End If
    If xFechaNac.Value >= Date Then
        MsgBox "FECHA DE NACIMIENTO INVÁLIDA. DEBE SER MENOR O IGUAL AL DIA DE HOY", vbCritical
        xFechaNac.SetFocus
        Exit Function
    End If
    If xFechaNac.Value < DateAdd("YYYY", -18, Date) And xVinculo.ListIndex = 0 Then
        MsgBox "EL SISTEMA LE RECUERDA QUE SOLO ESTÁN AFECTOS AL PDT SUNAT LOS HIJOS MENORES DE 18 AÑOS. SI ES UN NIÑO ESPECIAL, DEBERÁ ASIGNARLE EL DOCUMENTO DE INCAPACIDAD ", vbCritical
    End If
    If Trim(xSituacion.Text) = "" Then
        MsgBox "Seleccione la sistuación del DerechoHabiente", vbCritical
        xSituacion.SetFocus
        Exit Function
    End If
    If xIDP.Value = 1 Then
        If xNombreVia.Text = "" Then
            MsgBox "FALTA EL NOMBRE DE LA VIA. EJEMPLO:CALLE - MIGUEL GRAU", vbCritical
            xNombreVia.SetFocus
            Exit Function
        End If
        If xZona.Text = "" Then
            MsgBox "DEBE INGRESAR EL NOMBRE DE LA ZONA. EJEMPLO: URBANIZACIÓN - JORGE CHAVEZ", vbCritical
            xZona.SetFocus
            Exit Function
        End If
        If xReferencia.Text = "" Then
            MsgBox "FALTA REFERNCIA DEL DOMICILIO", vbCritical
            xReferencia.SetFocus
        End If
        If xUbigeo.Tag = "" Then
            MsgBox "DEBERÁ SELECCIONAR EL CÓDIGO DE UBICACIÓN GEOGRÁFICA DEL DOMICILIO DEL DERECHOHABIENTE", vbCritical
            Exit Function
        End If
    End If
    VALIDAR = True
End Function

Private Sub CMCARGAFOTO_Click()
    frOpenGr.Show 1
    If VGUTIL(0) <> "" Then
        xFoto.Picture = LoadPicture(VGUTIL(0))
        xFoto.Tag = VGUTIL(0)
        VGUTIL(0) = ""
    Else
        MsgBox "ACCIÓN CANCELADA", vbInformation
    End If
End Sub

Private Sub CMDESCARGAFOTO_Click()
    Set xFoto.Picture = Nothing
    xFoto.Tag = ""
End Sub

Private Sub CMDIMP_Click()
Dim XCUEN As Long
    Screen.MousePointer = 11
    DBSYSTEM.Execute "DELETE FROM FTMPFOTO"
    DBSYSTEM.Close: DBSYSTEM.Open: DBSYSTEM.Close: DBSYSTEM.Open
    Data1.DatabaseName = App.PATH & "\BDAUXCOM.MDB"
    Data1.RecordSource = "FTMPFOTO"
    Do While True
        Data1.Refresh
        Data1.Recordset.AddNew
        Data1.Recordset.Fields("CODIGO") = "A"
        Ole1.DataField = "FOTO"
        If xFoto.Tag <> "" Then
            Ole1.Picture = LoadPicture(xFoto.Tag)
        Else
            Ole1.Picture = LoadPicture(REGSISTEMA.PATH & "OBJBLANK.BMP")
        End If
        Data1.Recordset.Fields("CODIGO") = " "
        Data1.Recordset.Update
        If Data1.Recordset.RecordCount > 0 Then Exit Do
    Loop
    
    With Reporte
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0038.RPT"
        .DataFiles(0) = App.PATH & "\BDAUXCOM.MDB"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowTitle = "PLAN0038 - FICHA DEL DERECHO HABIENTE"
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XPRIN='" & Label1.Caption & "'"
        .Formulas(2) = "XAPEPAT='" & xApePat.Text & "'"
        .Formulas(3) = "XAPEMAT='" & xApeMat.Text & "'"
        .Formulas(4) = "XNOMB='" & xNombre.Text & "'"
        .Formulas(5) = "XTIPDOC='" & xTipoDoc.Text & "'"
        .Formulas(6) = "XNUMDOC='" & xNumDoc.Text & "'"
        .Formulas(7) = "XFECHNAC='" & Format(xFechaNac, "DD/MM/YYYY") & "'"
        .Formulas(8) = "XVINC='" & UCase(xVinculo.Text) & "'"
        .Formulas(9) = "XSITU='" & xSituacion.Text & "'"
        .Formulas(10) = "XDOCINC='" & xDocIncap.Text & "'"
        .Formulas(11) = "XSEX='" & xSexo.Text & "'"
        .Formulas(12) = "XCARTA='" & xCarta.Text & "'"
        .Formulas(13) = "XMOTIV='" & xMotivoBaja.Text & "'"
        If xIDP.Value Then
            .Formulas(14) = "XDOMIC='SI'"
          Else: .Formulas(14) = "XDOMIC='NO'"
        End If
        .Formulas(15) = "XCODVIA='" & xTipoVia.Text & "'"
        .Formulas(16) = "XNOMVIA='" & xNombreVia.Text & "'"
        .Formulas(17) = "XNUM='" & xNumero.Text & "'"
        .Formulas(18) = "XINT='" & xInterior.Text & "'"
        .Formulas(19) = "XCODZON='" & xTipoZona.Text & "'"
        .Formulas(20) = "XNOMZON='" & xZona.Text & "'"
        .Formulas(21) = "XREF='" & xReferencia.Text & "'"
        .Formulas(22) = "XUBIGEO='" & xUbigeo.Text & "'"
        .Formulas(23) = "XHORA='" & Format(Time, "HH:MM") & "'"
        .Formulas(24) = "XRUC='" & REGSISTEMA.RUC & "'"
        If .Status <> 2 Then .Action = 1
        .WindowTitle = "FICHA DEL TRABAJADOR"
    End With
    Screen.MousePointer = 1
End Sub

Private Sub Form_Load()
    Label1.Caption = frFamily.xTrab.Text
    Me.Caption = "DERECHOHABIENTE DE " & frFamily.xTrab.Text
    RSDEREC.Open "FAMILIAR", DBSYSTEM, adOpenKeyset, adLockOptimistic
    xTipoDoc.ListIndex = 0
    xSexo.ListIndex = 0
    xVinculo.ListIndex = 0
    xTipoVia.ListIndex = 0
    xTipoZona.ListIndex = 0
    xMotivoBaja.ListIndex = 0
    xFechaNac.Value = Date
    If VPTAREA <> "NUEVO" Then
        RSDEREC.FIND "CODDER=" & VPTAREA
        CARGADATOS
    End If
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSDEREC = Nothing
End Sub

Private Sub XFECHANAC_LOSTFOCUS()
    If DateAdd("YYYY", 18, xFechaNac.Value) < Date Then
        xDocIncap.Locked = False
    Else
        xDocIncap.Locked = True
    End If
End Sub

Private Sub XIDP_Click()
    If xIDP.Value = 1 Then
        Frame2.Enabled = True
    Else
        Frame2.Enabled = False
    End If
End Sub

Private Sub XSITUACION_Click()
    If xSituacion.ListIndex = 1 Then xMotivoBaja.Locked = False Else xMotivoBaja.Locked = True
End Sub

Private Sub XUBIGEO_DblClick()
    frUbigeo.Show 1
    If VPCODTMP <> "" Then
        xUbigeo.Tag = VPCODTMP
        xUbigeo.Text = VPTRASPRM
    End If
End Sub

Private Sub XVINCULO_Click()
    If xVinculo.ListIndex = 3 Then
        xCarta.Locked = False
        If xSexo.ListIndex = 0 Then
            MsgBox "EL SEXO ES INVÁLIDO, DEBE SER DE TIPO FEMENINO PARA QUE SE ENCUENTRE EN CONDICIÓN DE GESTANTE", vbCritical
            xSexo.SetFocus
        End If
    Else
        xCarta.Text = ""
        xCarta.Locked = True
    End If
End Sub

Public Sub CARGADATOS()
    On Error GoTo ERRCARGA
    With RSDEREC
        xApePat.Text = "" & !ApePat
        xApeMat.Text = "" & !ApeMat
        xNombre.Text = "" & !NOMBRE
        xTipoDoc.ListIndex = !TIPODOC
        xNumDoc.Text = "" & !NUMDOC
        xFechaNac.Value = !FechaNac
        xSexo.ListIndex = !Sexo
        xVinculo.ListIndex = !VINCULO
        xCarta.Text = "" & !CARTA
        xSituacion.ListIndex = !SITUACION
        xMotivoBaja.ListIndex = !MOTIVOBAJA
        xDocIncap.Text = "" & !DOCINCAP
        xIDP.Value = !IDP
        xNombreVia.Text = "" & !NOMBREVIA
        xNumero.Text = "" & !NUMERO
        xInterior.Text = "" & !INTERIOR
        xZona.Text = "" & !ZONA
        xReferencia.Text = "" & !REFERENCIA
        xTipoVia.ListIndex = !TIPOVIA
        xTipoZona.ListIndex = !TIPOZONA
        xUbigeo.Text = "" & !UBIGEO
        xUbigeo.Tag = "" & !UBIGEO
        If UCase(Dir$(REGSISTEMA.PATHFOTOS & "\" & frFamily.xTrab.Tag & !CODDER & ".FTD")) = UCase(frFamily.xTrab.Tag & !CODDER & ".FTD") Then
            xFoto.Picture = LoadPicture(REGSISTEMA.PATHFOTOS & "\" & frFamily.xTrab.Tag & !CODDER & ".FTD")
            xFoto.Tag = REGSISTEMA.PATHFOTOS & "\" & frFamily.xTrab.Tag & !CODDER & ".FTD"
        End If
    End With
    Exit Sub
ERRCARGA:
    Resume Next
End Sub


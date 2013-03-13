VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Begin VB.Form BcoCTSCred 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Depositos de CTS - Medio Magnetico"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "BcoCTSCred.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCerrar 
      Caption         =   "&Cerrar"
      Height          =   315
      Left            =   4590
      TabIndex        =   32
      Top             =   4800
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Banco de Crédito"
      TabPicture(0)   =   "BcoCTSCred.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "-"
      TabPicture(1)   =   "BcoCTSCred.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "-"
      TabPicture(2)   =   "BcoCTSCred.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame1 
         Caption         =   "Configuración"
         Height          =   4035
         Left            =   105
         TabIndex        =   1
         Top             =   510
         Width           =   5610
         Begin VB.CommandButton cmExporta01 
            Caption         =   "&Exporta CTS"
            Height          =   330
            Left            =   4080
            TabIndex        =   30
            Top             =   3555
            Width           =   1410
         End
         Begin AplisetControlText.Aplitext T8 
            Height          =   315
            Left            =   1725
            TabIndex        =   29
            Top             =   3390
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext T12 
            Height          =   315
            Left            =   4065
            TabIndex        =   27
            Top             =   3060
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext T7 
            Height          =   315
            Left            =   1725
            TabIndex        =   25
            Top             =   3060
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext T11 
            Height          =   315
            Left            =   4065
            TabIndex        =   23
            Top             =   2730
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext T6 
            Height          =   315
            Left            =   1725
            TabIndex        =   21
            Top             =   2730
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext T10 
            Height          =   315
            Left            =   4065
            TabIndex        =   19
            Top             =   2400
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext T5 
            Height          =   315
            Left            =   1725
            TabIndex        =   17
            Top             =   2400
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext T9 
            Height          =   315
            Left            =   4065
            TabIndex        =   15
            Top             =   1740
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext T4 
            Height          =   315
            Left            =   1725
            TabIndex        =   13
            Top             =   2070
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext T3 
            Height          =   315
            Left            =   1725
            TabIndex        =   11
            Top             =   1740
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext T2 
            Height          =   315
            Left            =   1725
            TabIndex        =   9
            Top             =   1410
            Width           =   3780
            _ExtentX        =   6668
            _ExtentY        =   556
            Text            =   ""
         End
         Begin AplisetControlText.Aplitext T1 
            Height          =   315
            Left            =   1725
            TabIndex        =   7
            Top             =   1080
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            Text            =   ""
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "BcoCTSCred.frx":0060
            Left            =   1725
            List            =   "BcoCTSCred.frx":006A
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   750
            Width           =   1440
         End
         Begin AplisetControlText.Aplitext xCodigo 
            Height          =   315
            Left            =   1725
            TabIndex        =   3
            Top             =   390
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            MaxLength       =   14
            Text            =   ""
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Banco de Crédito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3840
            TabIndex        =   31
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label Label14 
            Caption         =   "Localidad"
            Height          =   255
            Left            =   255
            TabIndex        =   28
            Top             =   3465
            Width           =   960
         End
         Begin VB.Label Label13 
            Caption         =   "Sector"
            Height          =   210
            Left            =   3270
            TabIndex        =   26
            Top             =   3120
            Width           =   645
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Sector"
            Height          =   195
            Left            =   270
            TabIndex        =   24
            Top             =   3120
            Width           =   825
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Nombre U."
            Height          =   195
            Left            =   3255
            TabIndex        =   22
            Top             =   2775
            Width           =   765
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Urbanizacion"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   2790
            Width           =   1290
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Num. Dep."
            Height          =   195
            Left            =   3255
            TabIndex        =   18
            Top             =   2475
            Width           =   765
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Departamento"
            Height          =   195
            Left            =   210
            TabIndex        =   16
            Top             =   2430
            Width           =   1365
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Lote"
            Height          =   195
            Left            =   3435
            TabIndex        =   14
            Top             =   1785
            Width           =   315
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Manzana"
            Height          =   195
            Left            =   225
            TabIndex        =   12
            Top             =   2145
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Número de Calle"
            Height          =   195
            Left            =   210
            TabIndex        =   10
            Top             =   1815
            Width           =   1170
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nombre de Calle"
            Height          =   195
            Left            =   210
            TabIndex        =   8
            Top             =   1470
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Calle"
            Height          =   195
            Left            =   195
            TabIndex        =   6
            Top             =   1140
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   210
            TabIndex        =   4
            Top             =   795
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código de Empresa"
            Height          =   195
            Left            =   195
            TabIndex        =   2
            Top             =   435
            Width           =   1380
         End
      End
   End
End
Attribute VB_Name = "BcoCTSCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub cmExporta01_Click()
    On Error GoTo Err1
    Dim xFile As String, CadBan As String, xCad As String
    frSelDir.Show 1
    If VPTAREA = "" Then Exit Sub
    If Right(VPTAREA, 1) <> "\" Then VPTAREA = VPTAREA & "\"
    xFile = VPTAREA & "CTSAll.txt"
    If Dir$(xFile) <> "" Then
        If MsgBox("Ya existe en esta ruta un archivo correspondiente al pago de CTS del Banco de Credito, Desea Ud. reemplazar el archivo por el nuevo que está procesando", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        Kill xFile
    End If
    Dim RSAUX As New ADODB.Recordset, SumTodo As Long, SumaCab As Single
    Dim xCant As Long, ValorDolar As Single
    DBSYSTEM.Execute "UPDATE  [##ULTPAGOS" & VGL_COMPUTER & "]  SET  [##ULTPAGOS" & VGL_COMPUTER & "] .FECHANACIMIENTO = TRABAJADORES.FECHANAC FROM  [##ULTPAGOS" & VGL_COMPUTER & "] , TRABAJADORES WHERE  [##ULTPAGOS" & VGL_COMPUTER & "] .CODTRAB = TRABAJADORES.CODTRAB"
    RSAUX.Open "SELECT * FROM  [##ULTPAGOS" & VGL_COMPUTER & "]  WHERE BANCO='CRED'", DBSYSTEM, adOpenStatic, adLockReadOnly
    SumaCab = 0
    xCant = 0
    ValorDolar = Val(MDIPrincipal.BarraEstado.Panels(3).Text)
    Do While Not RSAUX.EOF
        If InStr(RSAUX!CTABANCO, "$") > 0 Then
            If Combo1.ListIndex = 1 Then SumaCab = SumaCab + Round(RSAUX!Neto / ValorDolar, 2)
            xCant = xCant + 1
        Else
            If Combo1.ListIndex = 0 Then SumaCab = SumaCab + RSAUX!Neto
            xCant = xCant + 1
        End If
        RSAUX.MoveNext
    Loop
    Open xFile For Append As #1
    Dim DirEmpresa As String
    DirEmpresa = Left(DevuelveValor("Select DIRECCIÓN FROM EMPRESA", DBSYSTEM) & String(40, " "), 40)
    xCad = SumaCab
    xCad = Left(xCad, InStr(xCad, ".") - 1) & Right(Trim(xCad), 2)
    CadBan = "2" & Left(Combo1.Text, 2) & Mid(REGSISTEMA.RUC, 3, 8) & "00000000000" & Format(Val(xCad), String(15, "0")) & Format(xCant, "00000") & "R1" & Left(REGSISTEMA.EMPRESA & String(40, " "), 40) & DirEmpresa & Left(xCodigo.Text & String(14, " "), 14)
    Print #1, CadBan
    RSAUX.MoveFirst
    Do While Not RSAUX.EOF
        If InStr(RSAUX!CTABANCO, "$") > 0 Then
            xCad = Trim(Format(Round(RSAUX!Neto / ValorDolar, 2), "0.00"))
        Else
            xCad = Trim(Format(RSAUX!Neto, "0.00"))
        End If
        xCad = Left(xCad, InStr(xCad, ".") - 1) & Right(xCad, 2)
        CadBan = "3" & Left(Combo1.Text, 2) & Mid(REGSISTEMA.RUC, 3, 8) & Left(SoloNumeros(RSAUX!CTABANCO) & "           ", 11) & Format(Val(xCad), String(15, "0")) & "R1 " & Left(RSAUX!NOMBRES & String(40, " "), 40) & " " & DirEmpresa & Format(RSAUX!FECHANACIMIENTO, "yyyymmdd") & Right("         " & RSAUX!DOCIDEN, 9) & "1" & Left(xCodigo.Text & String(14, " "), 14)
        Print #1, CadBan
        RSAUX.MoveNext
    Loop
    Close #1
    Set RSAUX = Nothing
    'CARGA DE LAS CUENTAS QUE ESTAN SIN CUENTAS
    xFile = VPTAREA & "CTSNew.txt"
    If Dir$(xFile) <> "" Then
        If MsgBox("Ya existe en esta ruta un archivo correspondiente al Creacion de nuevas cuentas y Pago de CTS del Banco de Credito, Desea Ud. reemplazar el archivo por el nuevo que está procesando", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        Kill xFile
    End If
    
    RSAUX.Open "SELECT * FROM  [##ULTPAGOS" & VGL_COMPUTER & "]  WHERE BANCO='NONE'", DBSYSTEM, adOpenStatic, adLockReadOnly
    CadBan = "2" & Mid(REGSISTEMA.RUC, 3, 8) & Left(REGSISTEMA.EMPRESA & String(75, " "), 75) & Left(T1.Text + "      ", 5) & Left(T2.Text & String(40, " "), 40) & Left(T3.Text & "    ", 4) & Left(T4.Text & String(10, " "), 10) & Left(T9.Text & "   ", 3) & Left(T5.Text & String(4, " "), 4) & Left(T6.Text & "     ", 4) & Left(T11.Text & String(20, " "), 20) & Left(T7.Text & "    ", 4) & Left(T12.Text & String(11, " "), 11) & Left(T8.Text & String(4, " "), 4) & String(12, " ") & "9309" & Left(xCodigo.Text & String(14, " "), 14)
    Open xFile For Append As #1
    Print #1, CadBan
    If RSAUX.RecordCount > 0 Then RSAUX.MoveFirst
    Do While Not RSAUX.EOF
        xCad = Trim(Format(RSAUX!Neto, "0.00"))
        xCad = Left(xCad, InStr(xCad, ".") - 1) & Right(xCad, 2)
        CadBan = "3N11" & Right("         " & RSAUX!DOCIDEN, 9) & "1" & Left(DevuelveValor("SELECT ApePat FROM Trabajadores WHERE CodTrab='" & RSAUX!CODTRAB & "'", DBSYSTEM) & String(25, " "), 25) & Left(DevuelveValor("SELECT ApeMat FROM Trabajadores WHERE CodTrab='" & RSAUX!CODTRAB & "'", DBSYSTEM) & String(25, " "), 25) & Left(DevuelveValor("SELECT Nombre FROM Trabajadores WHERE CodTrab='" & RSAUX!CODTRAB & "'", DBSYSTEM) & String(25, " "), 25) & Left(T1.Text + "      ", 5) & Left(T2.Text & String(40, " "), 40) & Left(T3.Text & "    ", 4) & Left(T4.Text & String(10, " "), 10) & Left(T9.Text & "   ", 3) & Left(T5.Text & String(4, " "), 4) & Left(T6.Text & "     ", 4) & Left(T11.Text & String(20, " "), 20) & Left(T7.Text & "    ", 4) & Left(T12.Text & String(11, " "), 11) & Left(T8.Text & String(4, " "), 4) & String(12, " ") & "0091" & Format(RSAUX!FECHANACIMIENTO, "yyyymmdd") & "1110" & Left(xCodigo.Text & String(14, " "), 14) & "2N"
        Print #1, CadBan
        RSAUX.MoveNext
    Loop
    Close #1
    Set RSAUX = Nothing
    'CREACION DEL ARCHIVO DE EMPRESAS
    xFile = VPTAREA & "VerNew.txt"
    If Dir$(xFile) <> "" Then
        If MsgBox("Ya existe en esta ruta un archivo correspondiente de Empresas de CTS del Banco de Credito, Desea Ud. reemplazar el archivo por el nuevo que está procesando", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        Kill xFile
    End If
    CadBan = Left(xCodigo.Text & String(12, " "), 12) & Mid(REGSISTEMA.RUC, 3, 8) & Left(REGSISTEMA.EMPRESA & String(50, " "), 50) & Left(Combo1.Text, 2)
    Open xFile For Append As #1
    Print #1, CadBan
    Close #1
    MsgBox "Proceso completado. Se han generado los archivos de pago de CTS para el Banco de Credito ", vbInformation
    Exit Sub
Err1:
    MsgBox ERR.Description
    Resume Next
    Exit Sub
    Resume
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 0
End Sub

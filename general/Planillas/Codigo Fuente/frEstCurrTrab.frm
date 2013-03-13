VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frEstCurrTrab 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formacion Laboral del Trabajador"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   Icon            =   "frEstCurrTrab.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   5880
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   360
      Width           =   900
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1335
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   345
      Width           =   4410
   End
   Begin VB.Frame Frame 
      Height          =   645
      Left            =   615
      TabIndex        =   34
      Top             =   5445
      Width           =   7905
      Begin VB.CommandButton Command1 
         Caption         =   "&Imprimir"
         Height          =   360
         Index           =   4
         Left            =   4455
         TabIndex        =   16
         Top             =   195
         Width           =   960
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Grabar"
         Height          =   360
         Index           =   3
         Left            =   3465
         TabIndex        =   14
         Top             =   195
         Width           =   960
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Eliminar"
         Height          =   360
         Index           =   2
         Left            =   2145
         TabIndex        =   3
         Top             =   195
         Width           =   960
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Modificar"
         Height          =   360
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   195
         Width           =   960
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Nuevo"
         Height          =   360
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   195
         Width           =   960
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancelar"
         Height          =   360
         Index           =   5
         Left            =   5595
         TabIndex        =   15
         Top             =   195
         Width           =   960
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   6
         Left            =   6570
         TabIndex        =   17
         Top             =   180
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   4080
      TabIndex        =   67
      Top             =   1920
      Width           =   1830
      Begin VB.CommandButton Command2 
         Caption         =   "Estudios"
         Height          =   315
         Left            =   100
         TabIndex        =   68
         Top             =   225
         Width           =   1620
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Laboral"
         Height          =   315
         Left            =   100
         TabIndex        =   69
         Top             =   615
         Width           =   1620
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Idiomas"
         Height          =   315
         Left            =   100
         TabIndex        =   70
         Top             =   975
         Width           =   1620
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   100
         TabIndex        =   71
         Top             =   1455
         Width           =   1620
      End
   End
   Begin VB.Frame FrMain 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   60
      TabIndex        =   36
      Top             =   765
      Width           =   8955
      Begin TabDlg.SSTab SSTab1 
         Height          =   4350
         Left            =   540
         TabIndex        =   37
         Top             =   180
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   7673
         _Version        =   393216
         Style           =   1
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "Estudios"
         TabPicture(0)   =   "frEstCurrTrab.frx":0ABA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Laboral"
         TabPicture(1)   =   "frEstCurrTrab.frx":0AD6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Idiomas"
         TabPicture(2)   =   "frEstCurrTrab.frx":0AF2
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame5"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame2 
            Height          =   3705
            Left            =   -74835
            TabIndex        =   54
            Top             =   465
            Width           =   7515
            Begin VB.ComboBox Combo7 
               Height          =   315
               Left            =   2145
               TabIndex        =   8
               Top             =   1710
               Width           =   3180
            End
            Begin VB.ComboBox Combo6 
               Height          =   315
               Left            =   2145
               TabIndex        =   6
               Top             =   1005
               Width           =   3180
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Pagado"
               Height          =   285
               Left            =   6240
               TabIndex        =   13
               Top             =   1710
               Width           =   960
            End
            Begin VB.TextBox Txtad 
               Height          =   690
               Left            =   2145
               MultiLine       =   -1  'True
               TabIndex        =   12
               Top             =   2820
               Width           =   5100
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   2145
               TabIndex        =   4
               Top             =   345
               Width           =   3180
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   2145
               TabIndex        =   5
               Top             =   675
               Width           =   3180
            End
            Begin AplisetControlText.Aplitext Aplitext5 
               Height          =   285
               Left            =   2145
               TabIndex        =   11
               Top             =   2415
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   503
               Text            =   ""
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   300
               Left            =   5580
               TabIndex        =   10
               Top             =   2040
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   529
               _Version        =   393216
               Format          =   61734913
               CurrentDate     =   36864
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   300
               Left            =   2145
               TabIndex        =   9
               Top             =   2040
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   529
               _Version        =   393216
               Format          =   61734913
               CurrentDate     =   36864
               MinDate         =   2
            End
            Begin AplisetControlText.Aplitext Aplitext3 
               Height          =   315
               Left            =   2145
               TabIndex        =   7
               Top             =   1350
               Width           =   3180
               _ExtentX        =   5609
               _ExtentY        =   556
               Text            =   ""
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Estudios"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   240
               Left            =   5730
               TabIndex        =   74
               Top             =   555
               Width           =   915
            End
            Begin VB.Image Image2 
               Height          =   480
               Left            =   6720
               Picture         =   "frEstCurrTrab.frx":0B0E
               Top             =   300
               Width           =   480
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Estudio"
               Height          =   195
               Left            =   360
               TabIndex        =   64
               Top             =   405
               Width           =   885
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Carrera"
               Height          =   195
               Left            =   360
               TabIndex        =   63
               Top             =   735
               Width           =   510
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Centro Estudios"
               Height          =   195
               Left            =   360
               TabIndex        =   62
               Top             =   1065
               Width           =   1470
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Centro Estudios"
               Height          =   195
               Left            =   360
               TabIndex        =   61
               Top             =   1410
               Width           =   1710
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Grado Obtenido"
               Height          =   195
               Left            =   360
               TabIndex        =   60
               Top             =   1755
               Width           =   1125
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Inicio"
               Height          =   195
               Left            =   360
               TabIndex        =   59
               Top             =   2100
               Width           =   870
            End
            Begin VB.Label Label8 
               Caption         =   "Fecha Fin"
               Height          =   195
               Left            =   4470
               TabIndex        =   58
               Top             =   2085
               Width           =   1005
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Nivel"
               Height          =   195
               Left            =   360
               TabIndex        =   57
               Top             =   2460
               Width           =   360
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Adicional"
               Height          =   195
               Left            =   360
               TabIndex        =   56
               Top             =   2850
               Width           =   645
            End
            Begin VB.Label LblCodEst 
               Height          =   345
               Left            =   6135
               TabIndex        =   55
               Top             =   390
               Visible         =   0   'False
               Width           =   930
            End
         End
         Begin VB.Frame Frame5 
            Height          =   3270
            Left            =   300
            TabIndex        =   51
            Top             =   690
            Width           =   7125
            Begin VB.Frame FrameOP 
               Caption         =   "Nivel"
               Height          =   1785
               Left            =   1725
               TabIndex        =   52
               Top             =   1155
               Width           =   2100
               Begin VB.OptionButton Option1 
                  Caption         =   "Nativo"
                  Height          =   240
                  Left            =   210
                  TabIndex        =   30
                  Top             =   270
                  Width           =   1680
               End
               Begin VB.OptionButton Option2 
                  Caption         =   "Basico"
                  Height          =   225
                  Left            =   210
                  TabIndex        =   31
                  TabStop         =   0   'False
                  Top             =   615
                  Width           =   1470
               End
               Begin VB.OptionButton Option3 
                  Caption         =   "Intermedio"
                  Height          =   195
                  Left            =   210
                  TabIndex        =   32
                  TabStop         =   0   'False
                  Top             =   975
                  Width           =   1455
               End
               Begin VB.OptionButton Option4 
                  Caption         =   "Avanzado"
                  Height          =   210
                  Left            =   210
                  TabIndex        =   33
                  TabStop         =   0   'False
                  Top             =   1335
                  Width           =   1485
               End
            End
            Begin VB.ComboBox Combo4 
               Height          =   315
               Left            =   1725
               TabIndex        =   29
               Top             =   570
               Width           =   4035
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Idioma"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   240
               Left            =   5280
               TabIndex        =   75
               Top             =   2700
               Width           =   720
            End
            Begin VB.Image Image3 
               Height          =   600
               Left            =   6150
               Picture         =   "frEstCurrTrab.frx":0E18
               Stretch         =   -1  'True
               Top             =   2355
               Width           =   645
            End
            Begin VB.Label Label21 
               Caption         =   "Idioma"
               Height          =   330
               Left            =   570
               TabIndex        =   53
               Top             =   585
               Width           =   810
            End
         End
         Begin VB.Frame Frame3 
            Height          =   3870
            Left            =   -74880
            TabIndex        =   38
            Top             =   375
            Width           =   7620
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   1980
               TabIndex        =   19
               Top             =   810
               Width           =   3435
            End
            Begin VB.ComboBox Combo5 
               Height          =   315
               Left            =   1980
               TabIndex        =   20
               Top             =   1170
               Width           =   4800
            End
            Begin VB.TextBox TxtCom 
               Height          =   555
               Left            =   1980
               TabIndex        =   28
               Top             =   3195
               Width           =   4800
            End
            Begin AplisetControlText.Aplitext Aplitext13 
               Height          =   285
               Left            =   5190
               TabIndex        =   26
               Top             =   2535
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   503
               Text            =   ""
            End
            Begin AplisetControlText.Aplitext Aplitext12 
               Height          =   285
               Left            =   1980
               TabIndex        =   25
               Top             =   2535
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   503
               Text            =   ""
            End
            Begin AplisetControlText.Aplitext Aplitext11 
               Height          =   285
               Left            =   1980
               TabIndex        =   27
               Top             =   2865
               Width           =   4800
               _ExtentX        =   8467
               _ExtentY        =   503
               Text            =   ""
            End
            Begin AplisetControlText.Aplitext Aplitext6 
               Height          =   285
               Left            =   1980
               TabIndex        =   18
               Top             =   465
               Width           =   4800
               _ExtentX        =   8467
               _ExtentY        =   503
               Text            =   ""
            End
            Begin MSComCtl2.DTPicker DTPicker4 
               Height          =   300
               Left            =   4995
               TabIndex        =   22
               Top             =   1530
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   529
               _Version        =   393216
               Format          =   61734913
               CurrentDate     =   36864
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   300
               Left            =   1980
               TabIndex        =   21
               Top             =   1530
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   529
               _Version        =   393216
               Format          =   61734913
               CurrentDate     =   36864
            End
            Begin AplisetControlText.Aplitext xBasico 
               Height          =   300
               Left            =   1980
               TabIndex        =   23
               Top             =   1875
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   529
               MaxLength       =   20
               Text            =   ""
               Redondear       =   -1  'True
               TipoDato        =   "N"
            End
            Begin AplisetControlText.Aplitext Aplitext1 
               Height          =   285
               Left            =   1980
               TabIndex        =   24
               Top             =   2205
               Width           =   4800
               _ExtentX        =   8467
               _ExtentY        =   503
               Text            =   ""
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Centro Laboral"
               Height          =   195
               Left            =   195
               TabIndex        =   50
               Top             =   525
               Width           =   1035
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Cargo"
               Height          =   195
               Left            =   195
               TabIndex        =   49
               Top             =   855
               Width           =   420
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Giro de Negocio"
               Height          =   195
               Left            =   195
               TabIndex        =   48
               Top             =   1230
               Width           =   1155
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Inicio"
               Height          =   195
               Left            =   195
               TabIndex        =   47
               Top             =   1590
               Width           =   870
            End
            Begin VB.Label Label14 
               Caption         =   "Fecha Fin"
               Height          =   285
               Left            =   3975
               TabIndex        =   46
               Top             =   1545
               Width           =   825
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Sueldo Estimado Anual"
               Height          =   195
               Left            =   180
               TabIndex        =   45
               Top             =   1920
               Width           =   1635
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Funciones"
               Height          =   195
               Left            =   165
               TabIndex        =   44
               Top             =   2250
               Width           =   735
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Categoria"
               Height          =   195
               Left            =   135
               TabIndex        =   43
               Top             =   2580
               Width           =   675
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Condicion"
               Height          =   195
               Left            =   4395
               TabIndex        =   42
               Top             =   2580
               Width           =   705
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Motivo Salida"
               Height          =   195
               Left            =   150
               TabIndex        =   41
               Top             =   2910
               Width           =   960
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Comentario"
               Height          =   195
               Left            =   135
               TabIndex        =   40
               Top             =   3195
               Width           =   795
            End
            Begin VB.Label LblCodLab 
               Height          =   315
               Left            =   6750
               TabIndex        =   39
               Top             =   885
               Visible         =   0   'False
               Width           =   765
            End
         End
         Begin VB.Label LblCodId 
            Height          =   345
            Left            =   -69750
            TabIndex        =   65
            Top             =   435
            Visible         =   0   'False
            Width           =   1515
         End
      End
   End
   Begin VB.Frame FrGrid 
      Height          =   4545
      Left            =   75
      TabIndex        =   66
      Top             =   735
      Width           =   8955
      Begin Crystal.CrystalReport CR1 
         Left            =   510
         Top             =   3765
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4290
         Left            =   105
         TabIndex        =   0
         Top             =   180
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   7567
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   15924991
         BackColorFixed  =   8421504
         FocusRect       =   2
         GridLines       =   0
         SelectionMode   =   1
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   8505
      Picture         =   "frEstCurrTrab.frx":2462
      Top             =   195
      Width           =   720
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "Trabajador :"
      Height          =   195
      Left            =   300
      TabIndex        =   35
      Top             =   375
      Width           =   855
   End
End
Attribute VB_Name = "frEstCurrTrab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_MAYOR, RS1, RS5, RS2, RS3, RS4 As ADODB.Recordset
Dim GOP, NUEVO As Boolean
Dim VAR_IDI As String
Dim Index_OP As Integer
Dim VAR_GRIDTEXTO, VAR_GRIDDETALLE As String

Private Sub Combo1_Click()
Combo1.SelStart = 0
Combo1.SelLength = Len(Combo1.Text)
End Sub
Private Sub COMBO1_GOTFOCUS()
Combo1.SelStart = 0
Combo1.SelLength = Len(Combo1.Text)
End Sub
Private Sub COMBO1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
Private Sub COMBO2_Click()
Combo2.SelStart = 0
Combo2.SelLength = Len(Combo2.Text)
End Sub
Private Sub COMBO2_GOTFOCUS()
Combo2.SelStart = 0
Combo2.SelLength = Len(Combo2.Text)
End Sub
Private Sub COMBO2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
Private Sub COMBO3_Click()
Combo3.SelStart = 0
Combo3.SelLength = Len(Combo3.Text)
End Sub
Private Sub COMBO3_GOTFOCUS()
Combo3.SelStart = 0
Combo3.SelLength = Len(Combo3.Text)
End Sub
Private Sub COMBO3_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then
        SendKeys "{TAB}"
        KEYCODE = 0
    End If
End Sub
Private Sub COMBO4_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then
        SendKeys "{TAB}"
        KEYCODE = 0
    End If
End Sub

Private Sub COMBO5_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then
        SendKeys "{TAB}"
        KEYCODE = 0
    End If
End Sub
Private Sub Command1_Click(INDEX As Integer)
Select Case INDEX
Case 0
    NUEVO = True
    Command1(0).Enabled = False
    Command1(1).Enabled = False
    Command1(2).Enabled = False
    Command1(3).Enabled = False
    Command1(4).Enabled = False
    Command1(5).Enabled = False
    Command1(6).Enabled = False
    MSFlexGrid1.Enabled = False
    VAR_MODO_EDIT = True
    Frame1.Visible = True
    Command2.SetFocus
Case 1
    VERIFICAR
    If Len(VAR_GRIDTEXTO) > 0 Then
        If Len(VAR_GRIDDETALLE) > 0 Then
            NUEVO = False
            LLENAR
        End If
    End If
    MSFlexGrid1.Row = POSICIONINICIAL
Case 2
    VERIFICAR
    If Len(VAR_GRIDTEXTO) > 0 Then
        If Len(VAR_GRIDDETALLE) > 0 Then
            ELIMITEM
            CARGAR
        End If
    End If
Case 3
    GRABA
    If GOP Then
        CARGAR
        DISENABLED_FR
        VAR_MODO_EDIT = False
        MSFlexGrid1.Enabled = True
    End If
Case 4
    With CR1
        .Reset
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .WindowTitle = "PLRH0004.RPT - " & Me.Caption
        .ReportFileName = REGSISTEMA.REPORTES & "PLRH0004.RPT"
        .SelectionFormula = "{TRABAJADORES.CODTRAB}='" & Text3.Text & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        
        .SubreportToChange = "PlRH0001.rpt"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""

        .SubreportToChange = "PlRH0003.rpt"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""

        .SubreportToChange = "PlRH0002.rpt"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""

        If .Status <> 2 Then .Action = 1
    End With
Case 5
    DISENABLED_FR
    VAR_MODO_EDIT = False
    MSFlexGrid1.Enabled = True
Case 6
    Unload Me
End Select
End Sub
Private Sub Command2_Click()
Index_OP = 0
ENABLED_FR
Frame1.Visible = False
VAR_MODO_EDIT = True
End Sub
Private Sub Command3_Click()
    Index_OP = 1
    ENABLED_FR
    VAR_MODO_EDIT = True
    Frame1.Visible = False
End Sub
Private Sub Command4_Click()
Index_OP = 2
ENABLED_FR
VAR_MODO_EDIT = True
Frame1.Visible = False
End Sub
Private Sub Command5_Click()
Command1(0).Enabled = True
    Command1(1).Enabled = True
    Command1(2).Enabled = True
    'Command1(3).Enabled = True
    Command1(4).Enabled = True
    'Command1(5).Enabled = True
    Command1(6).Enabled = True
    MSFlexGrid1.Enabled = True
    VAR_MODO_EDIT = False
Frame1.Visible = False
End Sub
Private Sub DTPICKER1_KeyDown(KEYCODE As Integer, Shift As Integer)
If KEYCODE = 13 Then
    SendKeys "{TAB}"
    KEYCODE = 0
End If
End Sub
Private Sub DTPICKER2_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then
        SendKeys "{TAB}"
        KEYCODE = 0
    End If
End Sub
Private Sub DTPICKER3_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then
        SendKeys "{TAB}"
        KEYCODE = 0
    End If
End Sub
Private Sub DTPICKER4_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then
        SendKeys "{TAB}"
        KEYCODE = 0
    End If
End Sub
Private Sub Form_Activate()
CARGAR
MSFlexGrid1.Enabled = True
End Sub
Private Sub FORM_KEYDOWN(KEYCODE As Integer, Shift As Integer)
If KEYCODE = vbKeyEscape Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
VAR_MODO_EDIT = False
Set RS1 = New ADODB.Recordset
Set RS2 = New ADODB.Recordset
Set RS3 = New ADODB.Recordset
Set RS4 = New ADODB.Recordset
Set RS5 = New ADODB.Recordset
Set RS_MAYOR = New ADODB.Recordset
Frame1.Visible = False
RS4.Open "SELECT * FROM TIPO_CENTROE ORDER BY TIPOCE", DBSYSTEM, adOpenStatic, adLockOptimistic
If RS4.RecordCount > 0 Then
    While Not RS4.EOF
        Combo6.AddItem RS4.Fields(0)
        RS4.MoveNext
    Wend
End If
RS4.Close
RS4.Open "SELECT * FROM TIPO_GRADO ORDER BY TIPOGRADO", DBSYSTEM, adOpenStatic, adLockOptimistic
If RS4.RecordCount > 0 Then
    While Not RS4.EOF
        Combo7.AddItem RS4.Fields(0)
        RS4.MoveNext
    Wend
End If
RS4.Close
RS4.Open "SELECT * FROM TIPO_ESTUDIOS ORDER BY TIPOEST_DESCRIP", DBSYSTEM, adOpenStatic, adLockOptimistic
If RS4.RecordCount > 0 Then
    While Not RS4.EOF
        Combo1.AddItem RS4.Fields(1)
        RS4.MoveNext
    Wend
End If
RS4.Close
RS5.Open "SELECT * FROM DESCESTUDIOS ORDER BY DESCESTUDIO", DBSYSTEM, adOpenStatic, adLockOptimistic
If RS5.RecordCount > 0 Then
    While Not RS5.EOF
        Combo2.AddItem RS5.Fields(1)
        RS5.MoveNext
    Wend
End If
RS5.Close
RS4.Open "SELECT * FROM CARGOS ORDER BY CARGO", DBSYSTEM, adOpenStatic, adLockOptimistic
If RS4.RecordCount > 0 Then
    While Not RS4.EOF
        Combo3.AddItem RS4.Fields(0)
        RS4.MoveNext
    Wend
End If
RS4.Close
RS5.Open "SELECT * FROM TIPO_IDIOMAS ORDER BY IDIO_DESCRIP", DBSYSTEM, adOpenStatic, adLockOptimistic
If RS5.RecordCount > 0 Then
    While Not RS5.EOF
        Combo4.AddItem RS5.Fields(0)
        RS5.MoveNext
    Wend
End If
RS5.Close
RS4.Open "SELECT * FROM GIROEMPRESA ORDER BY GIRO", DBSYSTEM, adOpenStatic, adLockOptimistic
If RS4.RecordCount > 0 Then
    While Not RS4.EOF
        Combo5.AddItem RS4.Fields(0)
        RS4.MoveNext
    Wend
End If
RS4.Close
DISENABLED_FR
SSTab1.Tab = 0
With MSFlexGrid1
    .Cols = 8
    .ColWidth(0) = 0
    .ColWidth(1) = 3500
    .ColWidth(2) = 3500
    .ColWidth(3) = 3500
    .ColWidth(4) = 2500
    .ColWidth(5) = 2500
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
End With
End Sub
Private Sub FORM_UNLOAD(CANCEL As Integer)
    VAR_SHOW = 0
End Sub

Private Sub MSFLEXGRID1_Click()
MSFlexGrid1.COL = 1
End Sub
Private Sub MSFLEXGRID1_DblClick()
Dim X, POSICIONINICIAL As Integer
        POSICIONINICIAL = MSFlexGrid1.Row
        MSFlexGrid1.Row = POSICIONINICIAL
        MSFlexGrid1.COL = 1
            If MSFlexGrid1.CellBackColor = &H808080 Then
                If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = "-" Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = "+"
                    For X = POSICIONINICIAL + 1 To MSFlexGrid1.Rows - 1
                        MSFlexGrid1.Row = X
                        MSFlexGrid1.COL = 1
                        If MSFlexGrid1.CellBackColor <> &H808080 Then
                            MSFlexGrid1.RowHeight(MSFlexGrid1.Row) = 0
                        Else
                            MSFlexGrid1.Row = POSICIONINICIAL
                            Exit For
                        End If
                    Next
                Else
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = "-"
                    For X = POSICIONINICIAL + 1 To MSFlexGrid1.Rows - 1
                        MSFlexGrid1.Row = X
                        MSFlexGrid1.COL = 1
                        If MSFlexGrid1.CellBackColor <> &H808080 Then
                            MSFlexGrid1.RowHeight(MSFlexGrid1.Row) = 220
                        Else
                            MSFlexGrid1.Row = POSICIONINICIAL
                            Exit For
                        End If
                    Next
                End If
            Else
                VERIFICAR
                If Len(VAR_GRIDTEXTO) > 0 Then
                    If Len(VAR_GRIDDETALLE) > 0 Then
                        NUEVO = False
                        LLENAR
                        VAR_MODO_EDIT = True
                    End If
                End If
                MSFlexGrid1.Row = POSICIONINICIAL
            End If
End Sub
Private Sub MSFLEXGRID1_GOTFOCUS()
MSFlexGrid1.COL = 1
End Sub
Private Sub MSFLEXGRID1_KeyDown(KEYCODE As Integer, Shift As Integer)
If KEYCODE = vbKeyDelete Then
    VERIFICAR
    If Len(VAR_GRIDTEXTO) > 0 Then
        If Len(VAR_GRIDDETALLE) > 0 Then
            ELIMITEM
            CARGAR
        End If
    End If
End If
End Sub
Private Sub OPTION1_CLICK()
VAR_IDI = Option1.Caption
End Sub
Private Sub OPTION1_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then
        Command1(3).SetFocus
        KEYCODE = 0
    End If
End Sub
Private Sub OPTION2_Click()
VAR_IDI = Option2.Caption
End Sub
Private Sub OPTION2_KeyDown(KEYCODE As Integer, Shift As Integer)
If KEYCODE = 13 Then
        Command1(3).SetFocus
        KEYCODE = 0
    End If
End Sub
Private Sub OPTION3_Click()
VAR_IDI = Option3.Caption
End Sub
Private Sub OPTION3_KeyDown(KEYCODE As Integer, Shift As Integer)
If KEYCODE = 13 Then
        Command1(3).SetFocus
        KEYCODE = 0
    End If
End Sub
Private Sub OPTION4_Click()
VAR_IDI = Option4.Caption
End Sub
Public Sub ENABLED_FR()
        FrGrid.Visible = False
        FrMain.Visible = True
Select Case Index_OP
Case 0
                SSTab1.TabVisible(1) = False
                SSTab1.TabVisible(2) = False
        If NUEVO Then
            Combo1.Text = ""
            Combo2.Text = ""
            Combo6.Text = ""
            Aplitext3.Text = ""
            Combo7.Text = ""
            Aplitext5.Text = ""
            Txtad.Text = ""
        End If
Case 1
                SSTab1.TabVisible(0) = False
                SSTab1.TabVisible(2) = False
        If NUEVO Then
                Aplitext6.Text = ""
                Aplitext11.Text = ""
                Aplitext12.Text = ""
                Aplitext13.Text = ""
                Combo3.Text = ""
                Combo5.Text = ""
        End If
Case 2
                SSTab1.TabVisible(0) = False
                SSTab1.TabVisible(1) = False
        If NUEVO Then
            Combo1.Text = ""
            Option1.Value = True
        End If
End Select
        Command1(0).Enabled = False
        Command1(1).Enabled = False
        Command1(2).Enabled = False
        Command1(3).Enabled = True
        Command1(4).Enabled = False
        Command1(5).Enabled = True
        Command1(6).Enabled = False
End Sub
Public Sub DISENABLED_FR()
        FrGrid.Visible = True
        FrMain.Visible = False
        Command1(0).Enabled = True
        Command1(1).Enabled = True
        Command1(2).Enabled = True
        Command1(3).Enabled = False
        Command1(4).Enabled = True
        Command1(5).Enabled = False
        Command1(6).Enabled = True
                SSTab1.TabVisible(0) = True
                SSTab1.TabVisible(1) = True
                SSTab1.TabVisible(2) = True
                MSFlexGrid1.Enabled = False
End Sub
Public Sub GRABA()
Dim SQLEXEC As String
Dim Codigo As Integer
Codigo = GENERA
Dim XVAR As Double
GOP = False
Select Case SSTab1.Tab
Case 0
    If Len(Trim(Combo1.Text)) = 0 Then
        MsgBox "DATOS INCOMPLETOS INGRESE EL TIPO DE ESTUDIO", vbInformation, "INFORMACION"
        Combo1.SetFocus
        Exit Sub
    End If
    If Len(Trim(Combo2.Text)) = 0 Then
        MsgBox "DATOS INCOMPLETOS INGRESE LA CARRERA", vbInformation, "INFORMACION"
        Combo2.SetFocus
        Exit Sub
    End If
    If Check1.Value = 1 Then
        XVAR = -1
    Else
        XVAR = 0
    End If
    If NUEVO Then
        SQLEXEC = "INSERT INTO ESTUDIOS VALUES (" & Codigo & ",'" & Text3.Text & "', '" & Combo1.Text & "','" & _
            Combo2.Text & "','" & Aplitext3.Text & "','" & Combo6.Text & "','" & Combo7.Text & "'," & _
            DateSQL(DTPicker1.Value) & ", " & DateSQL(DTPicker2.Value) & ",'" & Aplitext5.Text & "','" & Txtad.Text & "', " & XVAR & ")"
    Else
        SQLEXEC = "UPDATE ESTUDIOS SET TIPOEST_DESCRIP='" & Combo1.Text & "', EST_CARRERA='" & _
            Combo2.Text & "', DESCESTUDIOS='" & Aplitext3.Text & "', TIPO_CESTUDIO='" & Combo6.Text & "', EST_GRADO_OBT='" & Combo7.Text & "', EST_FINI=" & _
            DateSQL(DTPicker1.Value) & ", EST_FFIN=" & DateSQL(DTPicker2.Value) & ", EST_NIVEL='" & Aplitext5.Text & "', EST_ADIC='" & Txtad.Text & "', PAGEMP=" & XVAR & "  WHERE CODIGO =" & LblCodEst.Caption
    End If
Case 1
    If Len(Trim(Combo3.Text)) = 0 Then
        MsgBox "DATOS INCOMPLETOS INGRESE EL CARGO", vbInformation, "INFORMACION"
        Combo3.SetFocus
        Exit Sub
    End If
    If NUEVO Then
        SQLEXEC = "INSERT INTO LABORAL VALUES (" & Codigo & ",'" & Text3.Text & "', '" & Aplitext6.Text & "','" & _
            Combo3.Text & "','" & Combo5.Text & "'," & DateSQL(DTPicker3.Value) & "," & DateSQL(DTPicker4.Value) & ",'" & xBasico.Text & "','" & _
            Me.Aplitext1.Text & "','" & Aplitext12.Text & "','" & Aplitext13.Text & "','" & Aplitext11.Text & "','" & TxtCom.Text & "')"
    Else
        SQLEXEC = "UPDATE LABORAL SET LAB_CEN_LABORAL='" & Aplitext6.Text & "', CARGO='" & _
            Combo3.Text & "', GIRO='" & Combo5.Text & "', LAB_FINI=" & DateSQL(DTPicker3.Value) & ", LAB_FFIN=" & DateSQL(DTPicker4.Value) & ", LAB_SUELDO_EST_ANUAL='" & _
            xBasico.Text & "', FUNCION='" & Aplitext1.Text & "', CATEGORIA='" & Aplitext12.Text & "', LAB_CONDICION='" & Aplitext13.Text & "', LAB_MOT_SALIDA='" & Aplitext11.Text & "', LAB_COMENTARIO='" & TxtCom.Text & _
            "' WHERE CODIGO =" & LblCodLab.Caption
    End If
Case 2
    If Len(Trim(Combo4.Text)) = 0 Then
        MsgBox "DATOS INCOMPLETOS INGRESE EL IDIOMA", vbInformation, "INFORMACION"
        Combo4.SetFocus
        Exit Sub
    End If
    If NUEVO Then
        SQLEXEC = "INSERT INTO IDIOMAS VALUES (" & Codigo & ",'" & Text3.Text & "', '" & Combo4.Text & "','" & VAR_IDI & "')"
    Else
        SQLEXEC = "UPDATE IDIOMAS SET IDIO_DESCRIP='" & Combo4.Text & "', IDI_NIVEL='" & VAR_IDI & "' WHERE CODIGO=" & LblCodId.Caption
    End If
End Select
DBSYSTEM.Execute SQLEXEC
GOP = True
End Sub
Public Sub CARGAR()
RS_MAYOR.Open "SELECT DISTINCT TIPOEST_DESCRIP FROM ESTUDIOS WHERE COD_TRAB='" & Text3.Text & "'", DBSYSTEM, adOpenStatic, adLockOptimistic
RS2.Open "SELECT * FROM LABORAL WHERE COD_TRAB='" & Text3.Text & "'", DBSYSTEM, adOpenStatic, adLockOptimistic
RS3.Open "SELECT * FROM IDIOMAS WHERE COD_TRAB='" & Text3.Text & "'", DBSYSTEM, adOpenStatic, adLockOptimistic
With MSFlexGrid1
Dim X As Integer
    .Cols = 9
    .Rows = 4
    For X = 0 To .Cols - 1
        .TextMatrix(1, X) = "CURRICULUM VITAE"
        .Row = 1
        .COL = X
        .CellAlignment = 4
        .CellFontBold = True
        .CellFontSize = 10
    Next
    .RowHeight(1) = 250
    .MergeCells = flexMergeFree
    .MergeRow(1) = True
    For X = 0 To .Cols - 1
        .TextMatrix(3, X) = "FORMACION PROFESIONAL"
        .Row = 3
        .COL = X
        .CellAlignment = 4
        .CellBackColor = &H808080
        .CellFontBold = True
        .CellForeColor = vbWhite
    Next
    .TextMatrix(3, 1) = "-"
    .Row = 3
    .COL = 1
    .CellFontBold = True
    .MergeRow(3) = True
    If RS_MAYOR.RecordCount Then
        While Not RS_MAYOR.EOF
            .Rows = MSFlexGrid1.Rows + 1
            For X = 0 To .Cols - 1
                .TextMatrix(MSFlexGrid1.Rows - 1, X) = UCase(RS_MAYOR.Fields(0))
                .Row = .Rows - 1
                .COL = X
                .CellBackColor = &HC0FFFF
                .CellFontBold = True
                .CellAlignment = 2
            Next
            .MergeRow(.Rows - 1) = True
                    RS1.Open "SELECT * FROM ESTUDIOS WHERE COD_TRAB='" & Text3.Text & "' AND TIPOEST_DESCRIP='" & RS_MAYOR.Fields(0) & "'", DBSYSTEM, adOpenStatic, adLockOptimistic
                    If RS1.RecordCount Then
                        While Not RS1.EOF
                            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                            If RS1.Fields(11) = True Then
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 8) = "**"
                            End If
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = UCase(RS1.Fields(0))
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = UCase(RS1.Fields(3))
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) = UCase(RS1.Fields(4))
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = UCase(RS1.Fields(5))
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 4) = UCase(RS1.Fields(6))
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = UCase(RS1.Fields(7))
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 6) = UCase(RS1.Fields(8))
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 7) = UCase(RS1.Fields(9))
                            RS1.MoveNext
                        Wend
                    End If
                    RS1.Close
                RS_MAYOR.MoveNext
        Wend
    End If
    .Rows = .Rows + 2
    For X = 0 To .Cols - 1
        .TextMatrix(.Rows - 1, X) = "LABORALES"
        .Row = .Rows - 1
        .COL = X
        .CellAlignment = 4
        .CellBackColor = &H808080
        .CellFontBold = True
        .CellForeColor = vbWhite
    Next
    .TextMatrix(.Rows - 1, 1) = "-"
    .Row = .Rows - 1
    .COL = 1
    .CellFontBold = True
   .MergeRow(.Rows - 1) = True
    If RS2.RecordCount Then
        While Not RS2.EOF
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = UCase(RS2.Fields(0))
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = UCase(RS2.Fields(2))
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) = UCase(RS2.Fields(3))
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = UCase(RS2.Fields(4))
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 4) = UCase(RS2.Fields(7))
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = UCase(RS2.Fields(8))
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 6) = UCase(RS2.Fields(5))
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 7) = UCase(RS2.Fields(6))
            RS2.MoveNext
        Wend
    End If
    RS2.Close
    .Rows = .Rows + 2
    For X = 0 To .Cols - 1
        .TextMatrix(.Rows - 1, X) = "IDIOMAS"
        .Row = .Rows - 1
        .COL = X
        .CellAlignment = 4
        .CellBackColor = &H808080
        .CellFontBold = True
        .CellForeColor = vbWhite
    Next
    .TextMatrix(.Rows - 1, 1) = "-"
    .Row = .Rows - 1
    .COL = 1
    .CellFontBold = True
    .MergeRow(.Rows - 1) = True
    If RS3.RecordCount Then
        While Not RS3.EOF
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = UCase(RS3.Fields(0))
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = UCase(RS3.Fields(2))
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) = UCase(RS3.Fields(3))
            RS3.MoveNext
        Wend
    End If
    RS3.Close
End With
RS_MAYOR.Close
End Sub
Public Sub LLENAR()
        FrGrid.Visible = False
        FrMain.Visible = True
Select Case VAR_GRIDTEXTO
Case "FORMACION PROFESIONAL"
    RS1.Open "SELECT * FROM ESTUDIOS WHERE COD_TRAB='" & Text3.Text & "' AND CODIGO=" & VAR_GRIDDETALLE, DBSYSTEM, adOpenStatic, adLockOptimistic
    If RS1.RecordCount Then
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = False
        SSTab1.TabVisible(0) = True
        LblCodEst.Caption = RS1.Fields(0)
        Combo1.Text = RS1.Fields(2)
        Combo2.Text = RS1.Fields(3)
        Combo6.Text = RS1.Fields(5)
        Aplitext3.Text = RS1.Fields(4)
        Combo7.Text = RS1.Fields(6)
        DTPicker1.Value = RS1.Fields(7)
        DTPicker2.Value = RS1.Fields(8)
        Aplitext5.Text = RS1.Fields(9)
        Txtad.Text = RS1.Fields(10)
        If RS1.Fields(11) = True Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
    End If
    RS1.Close
Case "LABORALES"
    RS1.Open "SELECT * FROM LABORAL WHERE COD_TRAB='" & Text3.Text & "' AND CODIGO=" & VAR_GRIDDETALLE, DBSYSTEM, adOpenStatic, adLockOptimistic
    If RS1.RecordCount Then
        SSTab1.TabVisible(0) = False
        SSTab1.TabVisible(2) = False
        SSTab1.TabVisible(1) = True
        LblCodLab.Caption = RS1.Fields(0)
        Aplitext6.Text = RS1.Fields(2)
        Combo3.Text = RS1.Fields(3)
        Combo5.Text = RS1.Fields(4)
        DTPicker3.Value = RS1.Fields(5)
        DTPicker4.Value = RS1.Fields(6)
        xBasico.Text = RS1.Fields(7)
        Aplitext1.Text = RS1.Fields(8)
        Aplitext12.Text = RS1.Fields(9)
        Aplitext13.Text = RS1.Fields(10)
        Aplitext11.Text = RS1.Fields(11)
        TxtCom.Text = RS1.Fields(12)
    End If
    RS1.Close
Case "IDIOMAS"
    RS1.Open "SELECT * FROM IDIOMAS WHERE COD_TRAB='" & Text3.Text & "' AND CODIGO=" & VAR_GRIDDETALLE, DBSYSTEM, adOpenStatic, adLockOptimistic
    If RS1.RecordCount Then
        SSTab1.TabVisible(0) = False
        SSTab1.TabVisible(2) = True
        SSTab1.TabVisible(1) = False
        Combo4.Text = RS1.Fields(2)
        Select Case UCase(RS1.Fields(3))
        Case "NATIVO"
            Option1.Value = True
        Case "BASICO"
            Option2.Value = True
        Case "INTERMEDIO"
            Option3.Value = True
        Case "AVANZADO"
            Option4.Value = True
        Case Else
            Option1.Value = True
        End Select
        LblCodId.Caption = RS1.Fields(0)
    End If
    RS1.Close
Case Else
    'SENTENCIA
End Select
Command1(0).Enabled = False
Command1(1).Enabled = False
Command1(2).Enabled = False
Command1(3).Enabled = True
Command1(4).Enabled = False
Command1(5).Enabled = True
Command1(6).Enabled = False
End Sub
Public Function GENERA() As Integer
Dim RS_AUX As ADODB.Recordset
Dim SSQL As String
Set RS_AUX = New ADODB.Recordset
Select Case SSTab1.Tab
Case 0
    SSQL = "SELECT * FROM ESTUDIOS ORDER BY CODIGO"
Case 1
    SSQL = "SELECT * FROM LABORAL ORDER BY CODIGO"
Case 2
    SSQL = "SELECT * FROM IDIOMAS ORDER BY CODIGO"
End Select
RS_AUX.Open SSQL, DBSYSTEM, adOpenStatic, adLockOptimistic
If RS_AUX.RecordCount Then
    RS_AUX.MoveLast
    GENERA = RS_AUX.Fields(0) + 1
Else
    GENERA = 1
End If
RS_AUX.Close
End Function
Public Sub ELIMITEM()
Dim TABLA, SQLDELETE As String
If MsgBox("DESEA ELIMINAR EL REGISTRO DE '" & VAR_GRIDTEXTO & "'", vbYesNo, "CONFIRMACION") = vbYes Then
    Select Case VAR_GRIDTEXTO
    Case "FORMACION PROFESIONAL"
        TABLA = "ESTUDIOS"
    Case "LABORALES"
        TABLA = "LABORAL"
    Case "IDIOMAS"
        TABLA = "IDIOMAS"
    End Select
    SQLDELETE = "DELETE FROM " & TABLA & " WHERE CODIGO=" & CInt(VAR_GRIDDETALLE)
    DBSYSTEM.Execute SQLDELETE
End If
End Sub
Public Sub VERIFICAR()
Dim X, POSICIONINICIAL, UBICA As Integer
VAR_GRIDTEXTO = ""
VAR_GRIDDETALLE = ""
If Len(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) > 0 And MSFlexGrid1.CellBackColor <> &H808000 Then
        POSICIONINICIAL = MSFlexGrid1.Row
        UBICA = 3
        For X = 0 To MSFlexGrid1.Rows - 1
            MSFlexGrid1.Row = X
            MSFlexGrid1.COL = 1
            If MSFlexGrid1.CellBackColor <> &HC0FFFF Then
                If POSICIONINICIAL = MSFlexGrid1.Row Then
                    VAR_GRIDDETALLE = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
                    Exit For
                End If
                 MSFlexGrid1.Row = X
                 MSFlexGrid1.COL = 2
                 If Len(Me.MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)) > 0 And MSFlexGrid1.CellBackColor = &H808080 Then
                    VAR_GRIDTEXTO = MSFlexGrid1.TextMatrix(X, 2)
                 End If

            End If
        Next
End If
End Sub
Private Sub OPTION4_KeyDown(KEYCODE As Integer, Shift As Integer)
If KEYCODE = 13 Then
        Command1(3).SetFocus
        KEYCODE = 0
    End If
End Sub
Private Sub TXTAD_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then
        Command1(3).SetFocus
        KEYCODE = 0
    End If
End Sub
Private Sub TXTCOM_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then
        Command1(3).SetFocus
        KEYCODE = 0
    End If
End Sub



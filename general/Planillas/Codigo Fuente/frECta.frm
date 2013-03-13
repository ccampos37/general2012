VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frECta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar Cuenta"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "frECta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Otras Configuraciones"
      Height          =   3105
      Left            =   105
      TabIndex        =   9
      Top             =   2085
      Width           =   5130
      Begin VB.CheckBox Check12 
         Caption         =   "Variable Auxiliar Total03"
         Height          =   195
         Left            =   2805
         TabIndex        =   20
         Top             =   2640
         Width           =   2085
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Provisiones de C.T.S."
         Height          =   195
         Left            =   255
         TabIndex        =   19
         Top             =   2100
         Width           =   1920
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Provisiones de Gratificación"
         Height          =   195
         Left            =   255
         TabIndex        =   18
         Top             =   2362
         Width           =   2325
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Provisiones de Vacaciones"
         Height          =   195
         Left            =   255
         TabIndex        =   17
         Top             =   2625
         Width           =   2445
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Variable Auxiliar Total01"
         Height          =   195
         Left            =   2805
         TabIndex        =   16
         Top             =   2115
         Width           =   2085
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Variable Auxiliar Total02"
         Height          =   195
         Left            =   2805
         TabIndex        =   15
         Top             =   2385
         Width           =   2085
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Remuneración afecta al S.C.T.R."
         Height          =   195
         Left            =   255
         TabIndex        =   14
         Top             =   1185
         Width           =   3240
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Remuneración afecta a Quinta Categoria"
         Height          =   195
         Left            =   255
         TabIndex        =   13
         Top             =   1455
         Width           =   3645
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Remuneración afecta al I.E.S."
         Height          =   195
         Left            =   255
         TabIndex        =   12
         Top             =   915
         Width           =   3240
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Remuneración de Aportaciones a EsSalud"
         Height          =   195
         Left            =   255
         TabIndex        =   11
         Top             =   645
         Width           =   3585
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Remuneración Asegurable de AFP"
         Height          =   195
         Left            =   255
         TabIndex        =   10
         Top             =   375
         Width           =   3240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   15
         X2              =   5070
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         Index           =   0
         X1              =   15
         X2              =   5085
         Y1              =   1860
         Y2              =   1860
      End
   End
   Begin VB.CommandButton cmCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   2918
      TabIndex        =   8
      Top             =   5370
      Width           =   1365
   End
   Begin VB.CommandButton cmAcepta 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   1058
      TabIndex        =   7
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   1830
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   5145
      Begin VB.CheckBox Check1 
         Caption         =   "Agrupar como Otros"
         Height          =   195
         Left            =   1995
         TabIndex        =   21
         Top             =   1500
         Width           =   2880
      End
      Begin AplisetControlText.Aplitext xCodigo 
         Height          =   285
         Left            =   1980
         TabIndex        =   4
         Top             =   375
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         MaxLength       =   6
         Text            =   ""
         TipoCodigo      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xNombre 
         Height          =   315
         Left            =   1980
         TabIndex        =   5
         Top             =   705
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         MaxLength       =   30
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xPlanilla 
         Height          =   315
         Left            =   1980
         TabIndex        =   6
         Top             =   1065
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label a1 
         AutoSize        =   -1  'True
         Caption         =   "Columna de Planilla"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   1125
         Width           =   1380
      End
      Begin VB.Label a1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   765
         Width           =   555
      End
      Begin VB.Label a1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   420
         Width           =   495
      End
   End
End
Attribute VB_Name = "frECta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSCOLPL2 As New ADODB.Recordset

Private Sub cmAcepta_Click()
    If xNombre.Text = "" Then
        MsgBox "DEBE INGRESAR UN NOMBRE DE GRUPO VÁLIDO", vbCritical
        xNombre.SetFocus
        Exit Sub
    End If
    If xPlanilla.Tag = "" Then
        MsgBox "NO HA SELECCIONADO UN CÓDIGO DE PLANILLA. SELECCIONELO HACIENDO DOBLE Click SOBRE LA CASILLA CORRESPONDIENTE", vbCritical
        xPlanilla.SetFocus
        Exit Sub
    End If
    If vpTarea = "NUEVO" Then
        If frCuentas.EXISTE(xCodigo.Text) Then
            MsgBox "EL CÓDIGO INGRESADO YA EXISTE, ESCRIBA OTRO", vbCritical
            xCodigo.SetFocus
            Exit Sub
        End If
        Dim RSAUX As New ADODB.Recordset
        RSAUX.Open "SELECT CODIGO FROM CONCEPTOS WHERE CODIGO='" & xCodigo.Text & "'", DbSystem, adOpenStatic, adLockReadOnly
        If RSAUX.RecordCount <> 0 Then
            Set RSAUX = Nothing
            MsgBox "EL CÓDIGO INGRESADO YA EXISTE COMO PARTE DEL MODULO DE CONCEPTOS DE REMUNERACIONES.", vbInformation
            Exit Sub
        End If
        Set RSAUX = Nothing
        DbSystem.Execute "INSERT INTO CTAGRUPO (CODGRUPO, NOMBRE, TIPO, PLANILLA) SELECT '" & xCodigo.Text & "', '" & xNombre.Text & "'," & (frCuentas.xTipo.ListIndex + 1) & ",'" & xPlanilla.Tag & "'"
    Else
        DbSystem.Execute "UPDATE CTAGRUPO SET NOMBRE='" & xNombre.Text & "', PLANILLA='" & xPlanilla.Tag & "' WHERE CODGRUPO='" & xCodigo.Text & "'"
    End If
    Unload Me
End Sub

Private Sub CMCANCELA_Click()
    Unload Me
End Sub

Private Sub FORM_ACTIVATE()
    If vpTarea = "NUEVO" Then Exit Sub
    RSCOLPL2.FIND "CODIGO='" & xPlanilla.Tag & "'"
    If RSCOLPL2.EOF Then
        Beep
        MsgBox "LA COLUMNA DE PLANILLA A LA QUE SE REFIERE LA CUENTA YA NO EXISTE, DEBERÁ SELECCIONAR OTRO", vbCritical
        Exit Sub
    Else
        xPlanilla.Text = RSCOLPL2!CODIGO & " : " & RSCOLPL2!NOMBRE
    End If
    xCodigo.Locked = True
    xNombre.SetFocus
End Sub

Private Sub FORM_Load()
    If frCuentas.xTipo.ListIndex = 0 Then
        RSCOLPL2.Open "SELECT CODIGO,NOMBRE FROM COLUMPL WHERE TIPO=2 ORDER BY NOMBRE", DbSystem, adOpenKeyset, adLockOptimistic
    Else
        RSCOLPL2.Open "SELECT CODIGO,NOMBRE FROM COLUMPL WHERE TIPO=3 ORDER BY NOMBRE", DbSystem, adOpenKeyset, adLockOptimistic
    End If
End Sub

Private Sub FORM_UnLoad(CANCEL As Integer)
    RSCOLPL2.Close
    Set RSCOLPL2 = Nothing
End Sub

Private Sub XPLANILLA_DblClick()
    frmComun.CONECTAR RSCOLPL2
    frmComun.Show 1
    If vgUtil(1) <> "" Then
        xPlanilla.Tag = vgUtil(1)
        xPlanilla.Text = vgUtil(1) & " : " & vgUtil(2)
    End If
End Sub


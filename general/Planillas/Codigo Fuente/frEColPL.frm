VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frEColPL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Columnas de Planilla"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frEColPL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   2460
      Width           =   1395
   End
   Begin VB.CommandButton cmAcepta 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1380
      TabIndex        =   9
      Top             =   2460
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Caption         =   "Columnas de Planilla"
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5685
      Begin VB.ComboBox xTipo 
         Height          =   315
         ItemData        =   "frEColPL.frx":030A
         Left            =   1200
         List            =   "frEColPL.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1590
         Width           =   2505
      End
      Begin AplisetControlText.Aplitext xValor 
         Height          =   525
         Left            =   1200
         TabIndex        =   7
         Top             =   1020
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   926
         MaxLength       =   200
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xNombre 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   690
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xCodigo 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         MaxLength       =   8
         Text            =   ""
         TipoCodigo      =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   270
         TabIndex        =   4
         Top             =   1710
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   270
         TabIndex        =   3
         Top             =   1110
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   210
         TabIndex        =   1
         Top             =   390
         Width           =   495
      End
   End
End
Attribute VB_Name = "frEColPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSCOLS As New ADODB.Recordset

Private Sub cmAcepta_Click()
    If xNombre.Text = "" Then
        MsgBox "TODA COLUMNA DE PLANILLA DEBE TENER UN NOMBRE VÁLIDO", vbCritical
        xNombre.SetFocus
        Exit Sub
    End If
    If xCodigo.Text = "" Then
        MsgBox "NO HA INGRESADO UN CÓDIGO PARA LA COLUMNA DE PLANILLA", vbCritical
        xCodigo.SetFocus
        Exit Sub
    End If
    If vpTarea = "NUEVO" Then
        If DevuelveValor("SELECT CODIGO FROM COLUMPL WHERE CODIGO='" & xCodigo.Text & "'", DbSystem) = xCodigo.Text Then
            MsgBox "EL CÓDIGO YA EXISTE, CAMBIE EL CÓDIGO", vbInformation
            xCodigo.SetFocus
            Exit Sub
        End If
        DbSystem.Execute "UPDATE COLUMPL SET INDICE=INDICE+1 WHERE INDICE>" & vpNumTmp
        RSCOLS.AddNew
        RSCOLS!CODIGO = UCase(xCodigo.Text)
        RSCOLS!INDICE = vpNumTmp + 1
    End If
    RSCOLS!NOMBRE = xNombre.Text
    RSCOLS!VALOR = "" & xValor.Text
    RSCOLS!TIPO = xTipo.ListIndex + 1
    RSCOLS.Update
    Unload Me
End Sub

Private Sub CMCANCELA_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    RSCOLS.Open "COLUMPL", DbSystem, adOpenDynamic, adLockPessimistic
    If vpTarea = "NUEVO" Then
        xTipo.ListIndex = 0
    Else
        RSCOLS.FIND "CODIGO='" & vpTarea & "'"
        If RSCOLS.EOF Then Unload Me
        With RSCOLS
            xCodigo.Text = !CODIGO
            xNombre.Text = "" & !NOMBRE
            xValor.Text = "" & !VALOR
                xTipo.ListIndex = 0 + !TIPO - 1
        End With
    End If
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    RSCOLS.Close
    Set RSCOLS = Nothing
End Sub


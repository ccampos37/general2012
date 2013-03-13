VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frENomBl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edición del Nombre de Boleta"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frENomBl.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4650
      TabIndex        =   21
      Top             =   5250
      Width           =   1155
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   3300
      TabIndex        =   20
      Top             =   5250
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Caption         =   "Declaraciones"
      Height          =   2085
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5685
      Begin VB.ComboBox xTipo 
         Height          =   315
         ItemData        =   "frENomBl.frx":030A
         Left            =   3480
         List            =   "frENomBl.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1290
         Width           =   2025
      End
      Begin AplisetControlText.Aplitext xMensaje 
         Height          =   285
         Left            =   1530
         TabIndex        =   19
         Top             =   1650
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   503
         MaxLength       =   100
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin MSComCtl2.DTPicker xMes 
         Height          =   315
         Left            =   1530
         TabIndex        =   17
         Top             =   1290
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36496
      End
      Begin AplisetControlText.Aplitext xCCosto 
         Height          =   285
         Left            =   1530
         TabIndex        =   15
         Top             =   960
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xNombre 
         Height          =   285
         Left            =   1530
         TabIndex        =   13
         Top             =   630
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   503
         MaxLength       =   35
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xCodigo 
         Height          =   285
         Left            =   1530
         TabIndex        =   11
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   3090
         TabIndex        =   22
         Top             =   1350
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje en Bolt."
         Height          =   195
         Left            =   240
         TabIndex        =   18
         ToolTipText     =   "Mensaje al final de la boleta"
         Top             =   1680
         Width           =   1185
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Mes de Proceso"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1350
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1020
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   690
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código Interno"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fijar Dias"
      Height          =   2865
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   5685
      Begin VB.CheckBox chShowNW 
         Caption         =   "Mostrar Número de Semana"
         Height          =   225
         Left            =   3330
         TabIndex        =   8
         Top             =   2460
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   315
         Left            =   3330
         TabIndex        =   7
         Top             =   1860
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36496
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   315
         Left            =   3330
         TabIndex        =   5
         Top             =   1215
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36496
      End
      Begin VB.ComboBox x1erDia 
         Height          =   315
         ItemData        =   "frENomBl.frx":033E
         Left            =   3330
         List            =   "frENomBl.frx":0357
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   570
         Width           =   2265
      End
      Begin MSComCtl2.MonthView mvAlma 
         Height          =   2370
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483646
         BackColor       =   -2147483644
         Appearance      =   1
         MaxSelCount     =   31
         MultiSelect     =   -1  'True
         ShowToday       =   0   'False
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   24707074
         TitleBackColor  =   -2147483647
         TitleForeColor  =   -2147483639
         CurrentDate     =   36496
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3330
         TabIndex        =   6
         Top             =   1620
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3330
         TabIndex        =   4
         Top             =   990
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Primer Día de la Semana"
         Height          =   195
         Left            =   3330
         TabIndex        =   2
         Top             =   330
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frENomBl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMACEPTAR_CLICK()
    If xNombre.Text = "" Then
        MsgBox "TODO NOMBRE DE BOLETA DEBE TENER UNA CADENA DESCRIPTIVA. INGRESE EL NOMBRE DE LA BOLETA", vbCritical
        xNombre.SetFocus
        Exit Sub
    End If
    If xCCosto.Text = "" Then
        If MsgBox("NO HA SELECCIONADO UN CENTRO DE COSTO. EL SISTEMA ASUMIRÁ QUE ES UN NOMBRE DE BOLETA DE CARACTER GENERAL (USO PARA TODOS LOS CENTROS DE COSTO). PRESIONE SI PARA CONTINUAR", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    If xFechaFin.Value < xFechaIni.Value Then
        MsgBox "LA FECHA INICIAL NO PUEDE SER MAYOR A LA FINAL. ESTABLESCA ADECUADAMENTE LAS FECHAS Y LUEGO VUELVA A INTENTAR GRABAR", vbCritical
        Exit Sub
    End If
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub CHSHOWNW_Click()
    mvAlma.ShowWeekNumbers = IIf(chShowNW = 0, False, True)
End Sub

Private Sub Form_Activate()
    xNombre.SetFocus
End Sub

Private Sub Form_Load()
    xTipo.ListIndex = 0
    x1erDia.ListIndex = 1
End Sub

Private Sub MVALMA_SELCHANGE(ByVal STARTDATE As Date, ByVal ENDDATE As Date, CANCEL As Boolean)
    xFechaIni.Value = STARTDATE
    xFechaFin.Value = ENDDATE
End Sub

Private Sub X1ERDIA_Click()
    mvAlma.StartOfWeek = x1erDia.ListIndex + 1
End Sub

Private Sub XMES_LOSTFOCUS()
    If Day(xMes.Value) <> 1 Then
        xMes.Day = 1
    End If
End Sub


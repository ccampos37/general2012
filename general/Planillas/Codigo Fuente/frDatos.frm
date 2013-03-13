VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Generales de la Empresa"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "frDatos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   330
      Left            =   3165
      TabIndex        =   17
      Top             =   5475
      Width           =   1290
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   1507
      TabIndex        =   16
      Top             =   5475
      Width           =   1290
   End
   Begin VB.Frame Frame2 
      Caption         =   "Representante Legal"
      Height          =   1950
      Left            =   135
      TabIndex        =   28
      Top             =   3345
      Width           =   4020
      Begin AplisetControlText.Aplitext xRL_Documento 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   1545
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xRL_TipoDoc 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   1245
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xRL_Nombre 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   945
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xRL_ApePat 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   345
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xRL_ApeMat 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   645
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Documento"
         Height          =   195
         Index           =   12
         Left            =   180
         TabIndex        =   33
         Top             =   1590
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Documento"
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   32
         Top             =   1290
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   31
         Top             =   990
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Materno"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   30
         Top             =   690
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Paterno"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   29
         Top             =   390
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información General"
      Height          =   3075
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   5655
      Begin AplisetControlText.Aplitext xCodigoPostal 
         Height          =   300
         Left            =   4215
         TabIndex        =   4
         Top             =   990
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         MaxLength       =   15
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xTelefono2 
         Height          =   300
         Left            =   4215
         TabIndex        =   10
         Top             =   2595
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         MaxLength       =   15
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xInterior 
         Height          =   300
         Left            =   1740
         TabIndex        =   3
         Top             =   990
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         MaxLength       =   5
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xDistrito 
         Height          =   285
         Left            =   1740
         TabIndex        =   6
         Top             =   1635
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         MaxLength       =   25
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xProvincia 
         Height          =   285
         Left            =   1740
         TabIndex        =   7
         Top             =   1950
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         MaxLength       =   25
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xDepartamento 
         Height          =   285
         Left            =   1740
         TabIndex        =   8
         Top             =   2265
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         MaxLength       =   25
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xTelefono1 
         Height          =   300
         Left            =   1740
         TabIndex        =   9
         Top             =   2580
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         MaxLength       =   15
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext XAlias 
         Height          =   300
         Left            =   1740
         TabIndex        =   1
         Top             =   360
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   529
         MaxLength       =   30
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xDireccion 
         Height          =   300
         Left            =   1740
         TabIndex        =   2
         Top             =   675
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   529
         MaxLength       =   50
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xAreaUrbana 
         Height          =   300
         Left            =   1740
         TabIndex        =   5
         Top             =   1305
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   529
         MaxLength       =   25
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono 2"
         Height          =   195
         Index           =   15
         Left            =   3150
         TabIndex        =   27
         Top             =   2655
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono 1"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   26
         Top             =   2633
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   25
         Top             =   2310
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   24
         Top             =   1995
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Distrito"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   23
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Area Urbana"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   22
         Top             =   1358
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Postal"
         Height          =   195
         Index           =   14
         Left            =   3150
         TabIndex        =   21
         Top             =   1050
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Interior"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   20
         Top             =   1035
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   19
         Top             =   728
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Alias para Reportes"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   413
         Width           =   1380
      End
   End
   Begin VB.Image xLogotipo 
      BorderStyle     =   1  'Fixed Single
      Height          =   1740
      Left            =   4275
      Top             =   3555
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Logotipo"
      Height          =   195
      Index           =   13
      Left            =   4275
      TabIndex        =   34
      Top             =   3330
      Width           =   615
   End
End
Attribute VB_Name = "frDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSEMP As New ADODB.Recordset

Private Sub CMACEPTAR_CLICK()
    GRABAR
    Unload Me
End Sub

Private Sub CMCERRAR_CLICK()
    Unload Me
End Sub

Private Sub Form_Load()
    RSEMP.Open "EMPRESA", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSEMP.RecordCount <> 1 Then
        cmAceptar.Enabled = False
    End If
    With RSEMP
        XAlias.Text = "" & !Alias
        xDireccion.Text = "" & !DIRECCIÓN
        xInterior.Text = "" & !INTERIOR
        xCodigoPostal.Text = "" & !CODIGOPOSTAL
        xAreaUrbana.Text = "" & !AREAURBANA
        xDistrito.Text = "" & !DISTRITO
        xProvincia.Text = "" & !PROVINCIA
        xDepartamento.Text = "" & !DEPARTAMENTO
        xTelefono1.Text = "" & !TELEFONO1
        xTelefono2.Text = "" & !TELEFONO2
        xRL_ApePat.Text = "" & !RL_APEPAT
        xRL_ApeMat.Text = "" & !RL_APEMAT
        xRL_Nombre.Text = "" & !RL_NOMBRE
        xRL_TipoDoc.Text = "" & !RL_TIPODOC
        xRL_Documento.Text = "" & !RL_DOCUMENTO
        'ASIGNAR AQUÍ EL LOGOTIPO
    End With
End Sub

Public Sub GRABAR()
    With RSEMP
        DBSYSTEM.Execute "UPDATE EMPRESA SET ALIAS ='" & Left(XAlias.Text, 20) & "',DIRECCIÓN='" & xDireccion.Text & "',INTERIOR='" & xInterior.Text & "',CODIGOPOSTAL='" & xCodigoPostal.Text & "',AREAURBANA='" & xAreaUrbana.Text & "',DISTRITO='" & xDistrito.Text & "',PROVINCIA='" & xProvincia.Text & "',DEPARTAMENTO='" & xDepartamento.Text & "',TELEFONO1='" & xTelefono1.Text & "',TELEFONO2='" & xTelefono2.Text & "',RL_APEPAT='" & xRL_ApePat.Text & "',RL_APEMAT='" & xRL_ApeMat.Text & "',RL_NOMBRE='" & xRL_Nombre.Text & "',RL_TIPODOC='" & xRL_TipoDoc.Text & "',RL_DOCUMENTO='" & xRL_Documento.Text & "'"
    End With
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSEMP = Nothing
End Sub


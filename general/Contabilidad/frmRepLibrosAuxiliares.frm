VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepLibrosAuxiliares 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libros Auxiliares"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   4620
   Begin VB.Frame fraCajaBancos 
      Caption         =   "Para Caja Bancos"
      Height          =   750
      Left            =   0
      TabIndex        =   15
      Top             =   2535
      Width           =   4620
      Begin VB.OptionButton optOpcionCajaBancos 
         Caption         =   "Registro Caja Resumido por Comprobante"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   17
         Top             =   450
         Width           =   3555
      End
      Begin VB.OptionButton optOpcionCajaBancos 
         Caption         =   "Registro Caja Resumido por Banco"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   16
         Top             =   225
         Width           =   3555
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Seleccionar Libro Auxiliar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   0
      TabIndex        =   13
      Top             =   1800
      Width           =   4620
      Begin VB.ComboBox cboLibroAuxiliar 
         Height          =   315
         Left            =   75
         TabIndex        =   14
         Top             =   270
         Width           =   3735
      End
   End
   Begin VB.Frame fraDetallado 
      Caption         =   "Diario General Detallado"
      Height          =   930
      Left            =   0
      TabIndex        =   7
      Top             =   3810
      Width           =   4635
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1980
         TabIndex        =   18
         Top             =   585
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   529
         _Version        =   393216
         Format          =   57999361
         CurrentDate     =   37474
      End
      Begin MSComCtl2.DTPicker DTPickerFecInicio 
         Height          =   300
         Left            =   1980
         TabIndex        =   19
         Top             =   225
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   529
         _Version        =   393216
         Format          =   57999361
         CurrentDate     =   37474
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Height          =   300
         Left            =   240
         TabIndex        =   9
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   615
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1830
      Left            =   0
      TabIndex        =   4
      Top             =   -75
      Width           =   4620
      Begin VB.OptionButton Option1 
         Caption         =   "Letras x Cobrar"
         Height          =   210
         Left            =   135
         TabIndex        =   12
         Top             =   1440
         Width           =   1800
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Letras x Pagar"
         Height          =   270
         Index           =   3
         Left            =   135
         TabIndex        =   11
         Top             =   1095
         Width           =   2145
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Registro de Ventas"
         Height          =   270
         Index           =   2
         Left            =   135
         TabIndex        =   10
         Top             =   810
         Width           =   2145
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Registro de Compras"
         Height          =   270
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   495
         Width           =   2145
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Caja y Bancos"
         Height          =   270
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   225
         Width           =   2250
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   3240
      Width           =   4620
      Begin VB.CheckBox chkAcumula 
         Caption         =   "Acumulado"
         Height          =   270
         Left            =   210
         TabIndex        =   3
         Top             =   180
         Width           =   1725
      End
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   345
      Index           =   1
      Left            =   2385
      TabIndex        =   1
      Top             =   5220
      Width           =   1050
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   345
      Index           =   0
      Left            =   1035
      TabIndex        =   0
      Top             =   5220
      Width           =   1050
   End
End
Attribute VB_Name = "frmRepLibrosAuxiliares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const N_OPCIONES_CAJABANCOS As Integer = 2
Const N_OPCIONES_LIBROSAUX As Integer = 5
Dim aOpcionLibroaux(N_OPCIONES_LIBROSAUX) As Integer

Private Sub Form_Load()
  Call LlenarcboLibroAuxiliar
  Call ConfiguraForm
End Sub

Sub LlenarcboLibroAuxiliar()
 Dim i As Integer
  cboLibroAuxiliar.AddItem "Caja y Bancos"
  cboLibroAuxiliar.AddItem "Registro de Compras"
  cboLibroAuxiliar.AddItem "Registro de Ventas"
  cboLibroAuxiliar.AddItem "Letras x Pagar"
  cboLibroAuxiliar.AddItem "Letras x Cobrar"
  For i = 0 To N_OPCIONES_LIBROSAUX
    aOpcionLibroaux(i) = i + 1
  Next
End Sub

Private Sub optOpcion_Click(Index As Integer)
  ConfiguraSeleccion (Index)
End Sub

Private Sub cboLibroAuxiliar_Click()
  Select Case cboLibroAuxiliar.ListIndex
    Case 0:
      fraCajaBancos.Enabled = True
      Call ConfiguraOpcionCajaBancos(True)
  
    Case 1, 2, 3, 4:
      fraCajaBancos.Enabled = False
      Call ConfiguraOpcionCajaBancos(False)
  End Select
 
End Sub

Sub ConfiguraForm()
  fraCajaBancos.Enabled = False
  Call ConfiguraOpcionCajaBancos(False)
  Width = 4740
  Height = 6090

End Sub

Sub ConfiguraOpcionCajaBancos(valor As Boolean)
  'Valor ==> T: Habilita  .F.=Deshabilita
  Dim i As Integer
  For i = 0 To N_OPCIONES_CAJABANCOS - 1
    optOpcionCajaBancos(i).Enabled = valor
  Next
End Sub

Sub ConfiguraSeleccion(valor As Integer)
  Select Case valor
    Case 0:
      fraCajaBancos.Enabled = True
      Call ConfiguraOpcionCajaBancos(True)
  
    Case 1, 2, 3, 4:
      fraCajaBancos.Enabled = False
      Call ConfiguraOpcionCajaBancos(False)
  End Select
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Select Case Index
    Case 0
    
    Case 1: Unload Me
  
  End Select

End Sub

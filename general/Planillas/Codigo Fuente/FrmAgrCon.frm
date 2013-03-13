VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form FrmAgrCon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Conceptos"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "FrmAgrCon.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin AplisetControlText.Aplitext xImporte 
      Height          =   300
      Left            =   2205
      TabIndex        =   1
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext xConcep 
      Height          =   315
      Left            =   2205
      TabIndex        =   0
      Top             =   255
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   556
      Text            =   ""
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3053
      TabIndex        =   4
      Top             =   1695
      Width           =   1425
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1223
      TabIndex        =   3
      Top             =   1695
      Width           =   1425
   End
   Begin VB.ComboBox xTipo 
      Height          =   315
      ItemData        =   "FrmAgrCon.frx":030A
      Left            =   2205
      List            =   "FrmAgrCon.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1155
      Width           =   3285
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Remuneración"
      Height          =   195
      Left            =   225
      TabIndex        =   7
      Top             =   1200
      Width           =   1620
   End
   Begin VB.Label Label2 
      Caption         =   "Importe   "
      Height          =   330
      Left            =   255
      TabIndex        =   6
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripción del Concepto"
      Height          =   195
      Left            =   255
      TabIndex        =   5
      Top             =   285
      Width           =   1830
   End
End
Attribute VB_Name = "FrmAgrCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VarGrabar As Boolean
Public Importe As Double
Public CONCEPTO As String
Public TIPO As Integer
Private Sub CMDACEPTAR_CLICK()
    If Trim(xConcep.Text) = "" Then
        MsgBox "Debe ingresar una Descripcion al concepto", vbExclamation
        Exit Sub
    End If
    If Val(xImporte.Text) = 0 Then
        MsgBox "El Monto del Importe debe ser mayor a 0", vbExclamation
        Exit Sub
    End If
    VarGrabar = True
    Importe = Val(xImporte.Text)
    CONCEPTO = xConcep.Text
    TIPO = xTipo.ListIndex
    Unload Me
End Sub

Private Sub CMDCANCELAR_CLICK()
    VarGrabar = False
    Unload Me
End Sub

Private Sub FORM_LOAD()
    Importe = 0
    CONCEPTO = ""
    TIPO = 0
    xTipo.ListIndex = 0
    xConcep.Text = ""
    xImporte.Text = "0"
    VarGrabar = False
End Sub

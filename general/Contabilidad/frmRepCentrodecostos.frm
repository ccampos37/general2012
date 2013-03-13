VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepCentrodeCostos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   0
      Width           =   5415
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1410
         TabIndex        =   9
         Top             =   240
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MM - MMMM"
         Format          =   61341699
         UpDown          =   -1  'True
         CurrentDate     =   37505
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   255
         Width           =   390
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de reporte"
      ForeColor       =   &H000000FF&
      Height          =   1155
      Left            =   3120
      TabIndex        =   5
      Top             =   720
      Width           =   2655
      Begin VB.OptionButton Option4 
         Caption         =   "Resumido"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Detallado"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordenado Por : "
      ForeColor       =   &H000000FF&
      Height          =   1155
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   2655
      Begin VB.OptionButton Option1 
         Caption         =   "Centro Costo / Cuenta"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cuenta / Ccentro Costo"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1665
      TabIndex        =   1
      Top             =   1980
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3015
      TabIndex        =   0
      Top             =   1980
      Width           =   1275
   End
End
Attribute VB_Name = "frmRepCentrodeCostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub cmdAceptar_Click()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
  Dim arrform(2) As Variant, arrparm(6) As Variant
  Set VGvardllgen = New dllgeneral.dll_general
  
  
  arrparm(0) = VGParamSistem.BDEmpresa
  arrparm(1) = VGParametros.empresacodigo
  arrparm(2) = VGParamSistem.Anoproceso
  arrparm(3) = Format(Month(DTPicker1.Value), "00")
  arrparm(4) = 1
  
  arrform(0) = "@mes='" & VGvardllgen.DesMes(VGParamSistem.Mesproceso) & "'"
  arrform(1) = "@Tituloreporte='Movimientos por Centro de Costos '"

  If Option3.Value = True Then
       arrparm(4) = 0
     If Option1.Value = True Then
        Call ImpresionRptProc("ct_listacentrocostodetallado.rpt", arrform, arrparm, , "Reporte detallado ")
       Else
        Call ImpresionRptProc("ct_cuentaxcentrocostodetallado.rpt", arrform, arrparm, , "Reporte detallado ")
     End If
   Else
     If Option1.Value = True Then
          Call ImpresionRptProc("ct_listacentrocostoresumido.rpt", arrform, arrparm, , "Reporte resumido ")
       Else
          Call ImpresionRptProc("ct_cuentaxcentrocostoresumido.rpt", arrform, arrparm, , "Reporte resumido ")
     End If
 End If
   End Sub

Private Sub cmdCancelar_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  Option1.Value = True
  Option3.Value = True
  
    DTPicker1.Value = DateSerial(CInt(VGParamSistem.Anoproceso), CInt(VGParamSistem.Mesproceso), 1)

End Sub


VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Begin VB.Form frmRepMayor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMayor XXXXXXXXX"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5955
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   15
      Width           =   5955
      Begin VB.CheckBox chkAcumula 
         Caption         =   "Acumulado"
         Height          =   270
         Left            =   210
         TabIndex        =   0
         Top             =   180
         Width           =   1725
      End
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   345
      Index           =   0
      Left            =   1695
      TabIndex        =   4
      Top             =   2730
      Width           =   1050
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   345
      Index           =   1
      Left            =   3075
      TabIndex        =   5
      Top             =   2715
      Width           =   1050
   End
   Begin VB.Frame fraDetallado 
      Height          =   1725
      Left            =   0
      TabIndex        =   6
      Top             =   585
      Width           =   5955
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   360
         Left            =   1035
         TabIndex        =   3
         Top             =   720
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   635
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
      Begin MSComCtl2.DTPicker DTPickerFecFinal 
         Height          =   300
         Left            =   4050
         TabIndex        =   2
         Top             =   270
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24444929
         CurrentDate     =   37474
      End
      Begin MSComCtl2.DTPicker DTPickerFecInicio 
         Height          =   300
         Left            =   1035
         TabIndex        =   1
         Top             =   270
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24444929
         CurrentDate     =   37474
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta"
         Height          =   255
         Left            =   75
         TabIndex        =   9
         Top             =   810
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Height          =   300
         Left            =   75
         TabIndex        =   8
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   285
         Left            =   3150
         TabIndex        =   7
         Top             =   315
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmRepMayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_tituloreporte As String
Public m_caso As String

Private Sub Form_Load()
  Call ConfiguraForm
  'Call SeleccionarMes(CInt(VGParamSistem.Mesproceso), CInt(VGParamSistem.Anoproceso))
  Call SeleccionarMes(CInt(VGParamSistem.Mesproceso), CInt("2002"))
End Sub

Sub ConfiguraForm()
  Me.Caption = m_tituloreporte
  Me.Width = 6075: Me.Height = 3660
  'Call CentrarForm(MDIPrincipal, Me)
  Ctr_Ayuda1.conexion VGcnx
  Ctr_Ayuda1.Filtro = "cuentacodigo<>'00'"
End Sub

Property Let tituloreporte(Valor As String)
  m_tituloreporte = Valor
End Property

Property Let caso(Valor As String)
  m_caso = Valor
End Property

Private Sub cmdBotones_Click(Index As Integer)
 Select Case Index
  Case 0:
     Select Case m_caso
        Case 1: 'Impresion de Libro Mayor Analítico
            
            
        
        Case 2: 'Impresión de Libro Mayor General
     
     
     
     End Select
  
  Case 1: Unload Me
 
 End Select

End Sub

Sub SeleccionarMes(nMes As Integer, nAnno As Integer)
  DTPickerFecInicio.Value = Format("01/" & nMes & "/" & nAnno, "dd/mm/yyyy")
  DTPickerFecFinal.Value = DateAdd("d", -1, DateAdd("m", 1, DTPickerFecInicio.Value))
End Sub

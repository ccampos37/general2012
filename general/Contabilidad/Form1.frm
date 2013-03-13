VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1965
      Width           =   3750
   End
   Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Analitico 
      Height          =   300
      Left            =   165
      TabIndex        =   0
      Top             =   1275
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   529
      XcodMaxLongitud =   11
      xcodwith        =   900
      NomTabla        =   "v_analiticoentidad"
      TituloAyuda     =   "Busqueda de Analitico"
      ListaCampos     =   "entidadruc(1),entidadrazonsocial(1),analiticocodigo(1)"
      XcodCampo       =   "entidadruc"
      XListCampo      =   "entidadrazonsocial"
      ListaCamposDescrip=   "Ruc,Descripcion,Codigo"
      ListaCamposText =   "entidadruc,entidadrazonsocial,analiticocodigo"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CtrAyu_Analitico_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
   Text1.Text = ColecCampos("ENTIDADRUC").Value
   'CtrAyu_Analitico.Ejecutar
End Sub

Private Sub Form_Load()
  CtrAyu_Analitico.conexion VGcnx
End Sub

VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7380
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mant 
      Caption         =   "&Mantenimiento"
      Begin VB.Menu mant_art 
         Caption         =   "&Articulos"
      End
      Begin VB.Menu mant_prov 
         Caption         =   "&Proveedores"
      End
      Begin VB.Menu mant_clie 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu mant_trans 
         Caption         =   "&Transacciones"
      End
      Begin VB.Menu mant_uni 
         Caption         =   "&Unidades"
      End
      Begin VB.Menu mnu_ayu 
         Caption         =   "&Ayudas"
         Begin VB.Menu mant_fam 
            Caption         =   "&Familia"
         End
         Begin VB.Menu mn_gru 
            Caption         =   "&Grupo"
         End
         Begin VB.Menu mant_linea 
            Caption         =   "&Linea"
         End
      End
   End
   Begin VB.Menu mn_tra 
      Caption         =   "&Transacciones"
      Begin VB.Menu mn_ent 
         Caption         =   "&Entradas"
      End
      Begin VB.Menu mn_sal 
         Caption         =   "&Salidas"
         Begin VB.Menu mn_salint 
            Caption         =   "&Internas"
         End
         Begin VB.Menu mn_salexr 
            Caption         =   "&mn_salext"
         End
      End
   End
   Begin VB.Menu mn_consul 
      Caption         =   "&Consultas"
      Begin VB.Menu mn_stkart 
         Caption         =   "&Stock Articulos"
      End
      Begin VB.Menu mn_doc 
         Caption         =   "&Documentos"
      End
   End
   Begin VB.Menu mn_rep 
      Caption         =   "&Reportes"
      Begin VB.Menu mn_kar 
         Caption         =   "&Kardex"
         Begin VB.Menu mnu_karart 
            Caption         =   "&Articulos"
         End
         Begin VB.Menu mnu_karval 
            Caption         =   "&Valorizados"
         End
      End
      Begin VB.Menu mn_stkrep 
         Caption         =   "&Stock"
      End
      Begin VB.Menu mn_gen 
         Caption         =   "&Generales"
      End
   End
   Begin VB.Menu mn_sist 
      Caption         =   "&Procesos"
      Begin VB.Menu mn_valor 
         Caption         =   "&Valorizacion"
      End
      Begin VB.Menu mn_guiarem 
         Caption         =   "&Guias de Remision"
         Begin VB.Menu mnu_anularGui 
            Caption         =   "&Anular"
         End
         Begin VB.Menu mnu_devGuia 
            Caption         =   "&Devolucion"
         End
      End
      Begin VB.Menu mn_prodoc 
         Caption         =   "&Documentos"
         Begin VB.Menu mnu_modDoc 
            Caption         =   "&Modificar"
         End
         Begin VB.Menu mnu_eliDoc 
            Caption         =   "&Eliminar"
         End
      End
      Begin VB.Menu mn_re 
         Caption         =   "-"
      End
      Begin VB.Menu mn_val 
         Caption         =   "&Valorizacion de Art"
      End
      Begin VB.Menu mnu_corraart 
         Caption         =   "&Correcion Art"
      End
      Begin VB.Menu mnu_blanco 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_esp 
         Caption         =   "&Especiales"
         Begin VB.Menu mnu_esttra 
            Caption         =   "&Estacion de Trabajo"
         End
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mant_art_Click()
FormArticulos.Show 1
End Sub

Private Sub mant_clie_Click()
   'Form2.Show 1
   FrmArClien.Show 1
End Sub

Private Sub mant_fam_Click()
'Form5.Show 1
FrmArFam.Show 1
End Sub

Private Sub mant_linea_Click()
  Form6.Show 1
End Sub

Private Sub mant_trans_Click()
  FormTransa.Show 1
End Sub

Private Sub mant_uni_Click()
  Form3.Show 1
End Sub

Private Sub MDIForm_Load()
  
  VGAlma = "01"
 '  VGRuta = cRuta2
 '  VGLongCodigo = 8
End Sub

Private Sub mn_doc_Click()
  FormConValArt.Show 1
End Sub

Private Sub mn_ent_Click()
   FormRegistro.Show 1
End Sub

Private Sub mn_gen_Click()
  'codigo
   formRep.Show 1
End Sub

Private Sub mn_gru_Click()
  Form7.Show 1
End Sub

Private Sub mn_salexr_Click()
    FormRegistro.Show 1
End Sub

Private Sub mn_salint_Click()
  FormRegistro.Show 1
End Sub

Private Sub mn_stk_Click()
   FormStkAlm.Show 1
End Sub

Private Sub mn_stkart_Click()
   FormConStk.Show
End Sub

Private Sub mn_stkrep_Click()
   FormStkAlm.Show 1
End Sub

Private Sub mnu_anularGui_Click()
   VGElimina = False
   FormEliminaDoc.Show 1
End Sub

Private Sub mnu_devGuia_Click()
   VGGuiaSal = False
   FormGuiaSal.Show 1
End Sub

Private Sub mnu_esttra_Click()
    FormCamAlm.Show 1
End Sub

Private Sub mnu_karart_Click()
  'form10
End Sub


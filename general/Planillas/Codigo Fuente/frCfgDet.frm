VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frCfgDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Vista "
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "frCfgDet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2955
      Top             =   3525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frCfgDet.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   105
      TabIndex        =   1
      Top             =   3660
      Width           =   1290
   End
   Begin MSComctlLib.ListView LColumnas 
      Height          =   3420
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   6033
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4095
      Picture         =   "frCfgDet.frx":0796
      Top             =   3600
      Width           =   480
   End
End
Attribute VB_Name = "frCfgDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Command2_Click()
    frCfgDet.Visible = False
    FrmDetalle.DgDet.ReBind
    Dim I As Integer
    For I = 1 To LColumnas.ListItems.Count
        If LColumnas.ListItems.ITEM(I).Checked = False Then
            FrmDetalle.DgDet.Columns(I - 1).Visible = False
         Else: FrmDetalle.DgDet.Columns(I - 1).Visible = True
        End If
    Next
    FrmDetalle.FORMHEAD
    FrmDetalle.DgDet.Refresh
End Sub

Private Sub Form_Activate()
    Form_Load
End Sub

Private Sub Form_Load()
    CARGAR
End Sub
Public Sub CARGAR()
    Dim I As Integer
Dim XLIST As ListItem
    LColumnas.ListItems.Clear
    Set XLIST = LColumnas.ListItems.Add(, , "CODTRAB", , 1)
        XLIST.SubItems(1) = "CODIGO"
    Set XLIST = LColumnas.ListItems.Add(, , "NOMBRES", , 1)
        XLIST.SubItems(1) = "APELLIDOS Y NOMBRES"
    For I = 1 To FrmDetalle.LvColumn.ListItems.Count
        Set XLIST = LColumnas.ListItems.Add(, , Trim(FrmDetalle.LvColumn.ListItems.ITEM(I)), , 1)
        XLIST.SubItems(1) = FrmDetalle.LvColumn.ListItems.ITEM(I).SubItems(2)
        
    Next
    
    For I = 0 To FrmDetalle.DgDet.Columns.Count - 1
         If FrmDetalle.DgDet.Columns(I).Visible Then
            LColumnas.ListItems.ITEM(I + 1).Checked = True
         End If
    Next
End Sub


Private Sub FORM_UNLOAD(CANCEL As Integer)
   CANCEL = 1
   Command2_Click
End Sub


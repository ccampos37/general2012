VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Descomprimir 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Recuperar Backups"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1125
      Top             =   3540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Recuperar Backups"
      Height          =   435
      Left            =   3045
      TabIndex        =   0
      Top             =   3540
      Width           =   1905
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   165
      Top             =   3570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Descomprimir.frx":0000
            Key             =   "P1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Descomprimir.frx":381C
            Key             =   "C2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Descomprimir.frx":3C70
            Key             =   "C1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Descomprimir.frx":40C4
            Key             =   "R1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Descomprimir.frx":4518
            Key             =   "R2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Descomprimir.frx":496C
            Key             =   "P2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvCCostos 
      Height          =   3465
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   6112
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Archivo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fecha Creacion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Comentario"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Descomprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESTINO As String
Dim RUTABKP As String, ARCHIVOBKP As String

Private Sub COMMAND1_CLICK()
On Error Resume Next
    If Trim(RUTABKP) <> "" Then
     DESTINO = Trim(RUTABKP)
     PROCESO2
    Else
     MsgBox "UD. NO HA SELECCIONADO NINGUN DIRECTORIO DONDE UBICAR EL BACKUP."
     Exit Sub
    End If
End Sub
Public Sub PROCESO2()  'DESCOMPRIMIR
On Error GoTo ERR
            D = Mid(Me.lvCCostos.ListItems(lvCCostos.SelectedItem.Index).Text, 1, 7)
            FileCopy RUTABKP & "\" & Me.lvCCostos.ListItems(lvCCostos.SelectedItem.Index).Text, DESTINO & "\" & D & ".MDB"
            MsgBox "LA OPERACIÓN SE REALIZÓ SATISFACTORIAMENTE, LA RUTA DEL BACKUP ES " & DESTINO, vbInformation, "INFORMACIÓN"
Exit Sub
ERR:
        MsgBox "ERROR AL DESCOMPRIMIR EL ARCHIVO", vbCritical, "ERROR"
        Exit Sub
End Sub
Private Sub FORM_LOAD()
Dim DIAR As String, MESR As String, ANNOR As String, FECHAR As String
Dim RS_BKP As ADODB.Recordset
Set RS_BKP = New ADODB.Recordset
Dim XLIST As ListItem
lvCCostos.ListItems.Clear
RS_BKP.Open "SELECT BKP_RUTA FROM EMPRESA", DbSystem, adOpenKeyset, adLockOptimistic
If RS_BKP.RecordCount Then
    RUTABKP = RS_BKP.Fields(0)
    ARCHIVOBKP = UCase(Dir((RUTABKP & "\A*.BKP")))
       While ARCHIVOBKP <> ""
            Set XLIST = lvCCostos.ListItems.Add(, "C" & ARCHIVOBKP, ARCHIVOBKP, 1, 1)
            DIAR = Mid(ARCHIVOBKP, 2, 2)
            MESR = Mid(ARCHIVOBKP, 4, 2)
            ANNOR = Mid(ARCHIVOBKP, 6, 2)
            FECHAR = DIAR & "/" & MESR & "/" & ANNOR
            XLIST.SubItems(1) = Format(FECHAR, "DD/MM/YYYY")
            XLIST.SubItems(2) = "BACKUP BASE AUXILIAR PLANILLAS"
            ARCHIVOBKP = Dir
       Wend
       ARCHIVOBKP = UCase(Dir((RUTABKP & "\P*.BKP")))
       While ARCHIVOBKP <> ""
       Set XLIST = lvCCostos.ListItems.Add(, "C" & ARCHIVOBKP, ARCHIVOBKP, 1, 1)
            DIAR = Mid(ARCHIVOBKP, 2, 2)
            MESR = Mid(ARCHIVOBKP, 4, 2)
            ANNOR = Mid(ARCHIVOBKP, 6, 2)
            FECHAR = DIAR & "/" & MESR & "/" & ANNOR
            XLIST.SubItems(1) = Format(FECHAR, "DD/MM/YYYY")
            XLIST.SubItems(2) = "BACKUP BASE PRINCIPAL PLANILLAS"
            ARCHIVOBKP = Dir
       Wend
End If
End Sub


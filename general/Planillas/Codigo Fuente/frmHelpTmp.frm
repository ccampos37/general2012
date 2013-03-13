VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHelpTmp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda del Sistema"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "frmHelpTmp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Pegar"
      Default         =   -1  'True
      Height          =   360
      Left            =   3870
      TabIndex        =   4
      Top             =   5895
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4950
      Top             =   1650
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpTmp.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpTmp.frx":0794
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpTmp.frx":0BE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   5400
      TabIndex        =   1
      Top             =   5895
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5760
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   10160
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Variables de Conceptos"
      TabPicture(0)   =   "frmHelpTmp.frx":0F3C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lista"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Pegar1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Pegar1(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Pegar1(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Pegar1(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Pegar1(4)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Pegar1(5)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Pegar1(6)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Pegar1(7)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Pegar1(8)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Pegar1(9)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Pegar1(10)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Pegar1(11)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Pegar1(12)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Pegar1(13)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Pegar1(14)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Pegar1(15)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Pegar1(16)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Pegar1(17)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Pegar1(18)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      Begin VB.TextBox Text1 
         Height          =   720
         HideSelection   =   0   'False
         Left            =   180
         TabIndex        =   24
         Top             =   4890
         Width           =   6150
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "Val"
         Height          =   360
         Index           =   18
         Left            =   5295
         TabIndex        =   23
         Top             =   4410
         Width           =   510
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "<="
         Height          =   360
         Index           =   17
         Left            =   3000
         TabIndex        =   22
         Top             =   4425
         Width           =   390
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   ">="
         Height          =   360
         Index           =   16
         Left            =   2517
         TabIndex        =   21
         Top             =   4425
         Width           =   390
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "<"
         Height          =   360
         Index           =   15
         Left            =   2034
         TabIndex        =   20
         Top             =   4425
         Width           =   390
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   ">"
         Height          =   360
         Index           =   14
         Left            =   1551
         TabIndex        =   19
         Top             =   4425
         Width           =   390
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "<>"
         Height          =   360
         Index           =   13
         Left            =   1068
         TabIndex        =   18
         Top             =   4425
         Width           =   390
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "Iif"
         Height          =   360
         Index           =   12
         Left            =   3480
         TabIndex        =   17
         Top             =   3930
         Width           =   510
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "Date"
         Height          =   360
         Index           =   11
         Left            =   4695
         TabIndex        =   16
         Top             =   4410
         Width           =   510
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "&&"
         Height          =   360
         Index           =   10
         Left            =   4080
         TabIndex        =   15
         Top             =   4410
         UseMaskColor    =   -1  'True
         Width           =   510
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "Not"
         Height          =   360
         Index           =   9
         Left            =   5295
         TabIndex        =   14
         Top             =   3930
         Width           =   510
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "Or"
         Height          =   360
         Index           =   8
         Left            =   4695
         TabIndex        =   13
         Top             =   3930
         Width           =   510
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "And"
         Height          =   360
         Index           =   7
         Left            =   4080
         TabIndex        =   12
         Top             =   3930
         Width           =   510
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "="
         Height          =   360
         Index           =   6
         Left            =   585
         TabIndex        =   11
         Top             =   4410
         Width           =   390
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   ")"
         Height          =   360
         Index           =   5
         Left            =   3000
         TabIndex        =   10
         Top             =   3930
         Width           =   390
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "("
         Height          =   360
         Index           =   4
         Left            =   2520
         TabIndex        =   9
         Top             =   3930
         Width           =   390
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "/"
         Height          =   360
         Index           =   3
         Left            =   2040
         TabIndex        =   8
         Top             =   3930
         Width           =   390
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "*"
         Height          =   360
         Index           =   2
         Left            =   1545
         TabIndex        =   7
         Top             =   3930
         Width           =   390
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "-"
         Height          =   360
         Index           =   1
         Left            =   1065
         TabIndex        =   6
         Top             =   3930
         Width           =   390
      End
      Begin VB.CommandButton Pegar1 
         Caption         =   "+"
         Height          =   360
         Index           =   0
         Left            =   585
         TabIndex        =   5
         Top             =   3930
         Width           =   390
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   2760
         Left            =   180
         TabIndex        =   3
         Top             =   1110
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   4868
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   345
         Picture         =   "frmHelpTmp.frx":0F58
         Top             =   450
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   270
         Picture         =   "frmHelpTmp.frx":139A
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   $"frmHelpTmp.frx":16DC
         Height          =   615
         Left            =   825
         TabIndex        =   2
         Top             =   450
         Width           =   5490
      End
   End
End
Attribute VB_Name = "frmHelpTmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LlamaFrm As Integer
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Command2_Click()
If LlamaFrm = 1 Then
    frECnpt.xFormula.Text = UCase(Text1.Text)
Else
    frAutoAd.xFormula.Text = UCase(Text1.Text)
End If
End Sub
Private Sub Form_Load()
    Dim xItem As ListItem
    Dim RsAux As New ADODB.Recordset
    RsAux.Open "SELECT * FROM VARIABLES2 ORDER BY CODIGO", DBSTARPLAN, adOpenStatic, adLockReadOnly
    Do While Not RsAux.EOF
        Set xItem = Lista.ListItems.Add(, , RsAux!CODIGO, , 1)
        xItem.SubItems(1) = RsAux!DESCRIPCION
        RsAux.MoveNext
    Loop
    Set RsAux = Nothing
    RsAux.Open "SELECT * FROM DATATRAB ORDER BY DESCDATA", DBSYSTEM, adOpenStatic, adLockReadOnly
    Do While Not RsAux.EOF
        Set xItem = Lista.ListItems.Add(, , RsAux!CODDATA, , 2)
        xItem.SubItems(1) = RsAux!DESCDATA
        RsAux.MoveNext
    Loop
    Set RsAux = Nothing
    RsAux.Open "SELECT * FROM CONCEPTOS ORDER BY NOMBRE", DBSYSTEM, adOpenStatic, adLockReadOnly
    Do While Not RsAux.EOF
        Set xItem = Lista.ListItems.Add(, , RsAux!CODIGO, , 3)
        xItem.SubItems(1) = RsAux!NOMBRE
        RsAux.MoveNext
    Loop
    Set RsAux = Nothing
End Sub
Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Lista.SortOrder = lvwAscending
    Lista.SortKey = ColumnHeader.Index - 1
End Sub
Private Sub Lista_DblClick()
    Text1.Text = Text1.Text & Lista.SelectedItem.Text
    Text1.SetFocus
    Text1.SelStart = Len(Text1.Text)
End Sub
Private Sub Pegar1_Click(Index As Integer)
    Text1.Text = Text1.Text & IIf(Pegar1(Index).Caption = "&&", "&", Pegar1(Index).Caption)
    Text1.SetFocus
    Text1.SelStart = Len(Text1.Text)
End Sub

VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmArticuloXCenCos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consumo de Articulos en Centro de Costos"
   ClientHeight    =   3810
   ClientLeft      =   300
   ClientTop       =   1410
   ClientWidth     =   6060
   Icon            =   "frmArticuloXCenCos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6060
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   708
      Left            =   2064
      Picture         =   "frmArticuloXCenCos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3084
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Salir"
      Height          =   708
      Left            =   3180
      Picture         =   "frmArticuloXCenCos.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3084
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar"
      Height          =   3000
      Left            =   36
      TabIndex        =   4
      Top             =   12
      Width           =   5970
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   10
         Top             =   432
         Width           =   1470
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   9
         Top             =   852
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   8
         Top             =   2472
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   312
         Left            =   1572
         TabIndex        =   0
         Top             =   1416
         Width           =   1404
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   47906817
         CurrentDate     =   36755
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   1572
         TabIndex        =   1
         Top             =   1932
         Width           =   1392
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         Format          =   47906817
         CurrentDate     =   36755
      End
      Begin VB.Label Label2 
         Caption         =   "Articulo Inicial"
         Height          =   252
         Left            =   312
         TabIndex        =   12
         Top             =   480
         Width           =   1128
      End
      Begin VB.Label Label1 
         Caption         =   "Articulo Final"
         Height          =   252
         Left            =   312
         TabIndex        =   11
         Top             =   900
         Width           =   1020
      End
      Begin VB.Label Label5 
         Caption         =   "Autorizado"
         Height          =   252
         Left            =   156
         TabIndex        =   7
         Top             =   2472
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta la Fecha"
         Height          =   360
         Left            =   144
         TabIndex        =   6
         Top             =   1908
         Width           =   1308
      End
      Begin VB.Label Label3 
         Caption         =   "Desde la Fecha "
         Height          =   336
         Left            =   144
         TabIndex        =   5
         Top             =   1440
         Width           =   1440
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   96
      Top             =   3144
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmArticuloXCenCos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Adodc3 As ADODB.Recordset

Private Sub Command1_Click()
Dim aform(4) As Variant
Dim aparam(5) As Variant
Dim Codigo1 As String, Codigo2 As String
aparam(0) = VGCNx.DefaultDatabase
aparam(1) = Text1.text
aparam(2) = Text2.text
aparam(3) = DTPicker1
aparam(4) = DTPicker2
aform(0) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
aform(1) = "fecini = '" & Format(DTPicker1, "dd/mm/yyyy") & "'"
aform(2) = "fecfin = '" & Format(DTPicker2, "dd/mm/yyyy") & "'"
aform(3) = "CCFIN = '" & Codigo2 & "'"
Call ImpresionRptProc("inv519.rpt", aform, aparam, , "inv519 Articulos x Centro de Costos")
Screen.MousePointer = 1
End Sub

Private Sub Command8_Click()
 Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub Form_Load()
If VGcc = 1 Then
   frmCenCos.Caption = "Consumo por centro de Costos"
   Text3.Visible = False
   Label5.Visible = False
Else
   frmCenCos.Caption = "Consumo por Persona Autorizado"
   Text3.Visible = True
   Label5.Visible = True
End If
DTPicker1 = DateAdd("m", -1, Date)
DTPicker2 = Date
LblCC1 = ""
LblCC2 = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   Text1_DblClick
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
 End If
End Sub

Private Sub Text2_DblClick()
    VGForm1 = 22
    FormAyuArt1.Show 1
    If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
         MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
         Exit Sub
    End If
    If Text1 <> "" Then
         Text2.Enabled = True
         Text2.SetFocus
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   Text2_DblClick
End If
End Sub

Private Sub Text3_DblClick()
FormAyuda.Show 1
End Sub
Private Sub Text1_DblClick()
    VGForm1 = 22
    FormAyuArt1.Show 1
    If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
         MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
         Exit Sub
    End If
    If Text1 <> "" Then
         Text2.Enabled = True
         Text2.SetFocus
    End If
End Sub




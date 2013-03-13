VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmKarValdetxEmpresa 
   Caption         =   "Kadex detallado por Empresa"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   5025
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmKarValdetxEmpresa.frx":0000
         Left            =   2070
         List            =   "FrmKarValdetxEmpresa.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   750
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FrmKarValdetxEmpresa.frx":0004
         Left            =   1710
         List            =   "FrmKarValdetxEmpresa.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   3150
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1710
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2220
         Width           =   1470
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1590
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Todos los Almacenes"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmKarValdetxEmpresa.frx":0022
         Left            =   2040
         List            =   "FrmKarValdetxEmpresa.frx":0024
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Todos los Establecimientos"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   690
         Left            =   3360
         Picture         =   "FrmKarValdetxEmpresa.frx":0026
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2400
         Width           =   705
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   708
         Left            =   4200
         Picture         =   "FrmKarValdetxEmpresa.frx":0468
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2400
         Width           =   705
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1710
         TabIndex        =   9
         Top             =   1275
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MMMM 'del' yyyy"
         Format          =   51511299
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.Label Label9 
         Caption         =   "Mes"
         Height          =   255
         Left            =   390
         TabIndex        =   14
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda   :"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   1755
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Articulo Inicial"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   2265
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "Articulo Final"
         Height          =   255
         Left            =   345
         TabIndex        =   11
         Top             =   2805
         Width           =   1020
      End
   End
End
Attribute VB_Name = "FrmKarValdetxEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim puntovta As ADODB.Recordset
Dim cRt As String
Dim almacen As String
Dim nTra, nConReg, nTotRec As Integer
Dim tipo As String
Private Sub Check1_Click()
If Check1.Value = 1 Then
   Check2.Enabled = False
   Combo1.Enabled = False
   Combo2.Enabled = False
   Check2.Value = 1
Else
   Check2.Enabled = True
   Combo1.Enabled = True
   Combo2.Enabled = True
      Check2.Value = 0
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   Combo1.Enabled = False
Else
   Combo1.Enabled = True
End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
rs.MoveFirst
rs.Move Combo1.ListIndex
almacen = Format(rs(0), "00")
VGAlma = almacen
End Sub

Private Sub cmdAceptar_Click()
Dim titulo As String
Dim aparam(7) As Variant
Dim aform(4) As Variant

'**************Valida el Ingreso**********
If Left(Combo3.text, 1) = "" Then
   MsgBox "Seleccione el Tipo de Moneda al la que desea el Informe de Valorización", vbInformation, "Faltan Datos"
   Exit Sub
End If
If Left(Combo3.text, 1) = "" Then
   MsgBox "Seleccione el Tipo de Moneda al la que desea el Informe de Valorización", vbInformation, "Faltan Datos"
   Exit Sub
End If
If Text1.text = "" And Text2.text = "" Then
    Text1.text = "000000"
    Text2.text = "999999"
 ElseIf Text1.text <> "" And Text2.text = "" Then
      Text2.text = Text1.text
End If
   aparam(0) = VGCNx.DefaultDatabase
   aparam(1) = VGparametros.empresacodigo
   If Check1.Value = 1 Then
       aparam(2) = "%%"
       aparam(3) = "%%"
    Else
       aparam(2) = Left(Combo2.text, 2)
       If Check2.Value = 1 Then
          aparam(3) = "%%"
        Else
          aparam(3) = Left(Combo1.text, 2)
       End If
   End If
   aparam(4) = UCase(Format(DTPicker1, "yyyymm"))
   aparam(5) = Text1.text
   aparam(6) = Text2.text
   If Check1.Value = 1 Then
      aform(0) = "PUNTOVTA = 'TODOS LOS ESTABLECIMIENTOS '"
     Else
       aform(0) = "PUNTOVTA = '" & UCase(Combo2.text) & "'"
   End If
   If Check2.Value = 1 Then
      aform(1) = "ALMACEN = 'TODOS LOS ALMACENES '"
     Else
       aform(1) = "ALMACEN = '" & UCase(Combo1.text) & "'"
   End If
   aform(2) = "Mes= '" & UCase(Format(DTPicker1, "MMMM - yyyy")) & "'"
  
  If Combo3.ListIndex <> 0 Then
     aform(3) = "MONEY= 'DOLAR'"
  Else
      aform(3) = "MONEY= 'SOLES'"
  End If
  Call ImpresionRptProc("al_kardexvaldetallado.RPT", aform, aparam, , "Kardex valorizado detallado por almacen ")
 End Sub
 Private Sub Carga_Almacen()
Dim RSQL As String
Dim I As Integer

RSQL = "Select TAALMA,TADESCRI FROM TabAlm where puntovtacodigo='" & Left(Combo2.text, 2) & "'"
RSQL = RSQL & " and almacenvalorizado=1 and empresacodigo='" & VGparametros.empresacodigo & "'"
Set rs = New ADODB.Recordset
rs.Open RSQL, VGCNx, adOpenStatic
Combo1.Clear
Do While Not rs.EOF
     Combo1.AddItem (Trim(rs(0)) & " - " & Trim(rs(1)))
     rs.MoveNext
     If rs.EOF Then Exit Do
Loop
If rs.RecordCount = 0 Then Exit Sub
rs.MoveFirst
For I = 0 To rs.RecordCount - 1
  If rs(0) = VGAlma Then
    Combo1.ListIndex = I
    Exit For
  Else
    rs.MoveNext
  End If
Next
End Sub
Private Sub Carga_puntovta()
Dim RSQL As String
Dim I As Integer
RSQL = "Select puntovtacodigo,puntovtadescripcion  FROM vt_puntoventa "
Set puntovta = New ADODB.Recordset
puntovta.Open RSQL, VGCNx, adOpenStatic
If puntovta.RecordCount = 0 Then Exit Sub
Do While Not puntovta.EOF
     Combo2.AddItem (Trim(puntovta(0)) & " - " & Trim(puntovta(1)))
     puntovta.MoveNext
     If puntovta.EOF Then Exit Do
Loop
puntovta.MoveFirst
For I = 0 To puntovta.RecordCount - 1
  If puntovta(0) = VGparametros.puntovta Then
    Combo2.ListIndex = I
    Exit For
  Else
    puntovta.MoveNext
  End If
Next
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Combo2_Click()
Carga_Almacen
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Dim RSQL As String
central Me
Carga_puntovta
Check2.Enabled = False
Combo1.Enabled = False
VGForm1 = 31
Check1.Value = 1
Check2.Value = 1
DTPicker1.Value = VGParamSistem.FechaTrabajo
End Sub

Private Sub Text1_DblClick()
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

Private Sub Text2_DblClick()
   FormAyuArt1.Show 1
End Sub


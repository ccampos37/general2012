VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmToma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Toma de Inventarios"
   ClientHeight    =   2790
   ClientLeft      =   2520
   ClientTop       =   1440
   ClientWidth     =   4560
   Icon            =   "frmtoma.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4560
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   795
      Left            =   855
      Picture         =   "frmtoma.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1770
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   795
      Left            =   2475
      Picture         =   "frmtoma.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame Frame15 
      Caption         =   "Seleccionar"
      Height          =   1020
      Left            =   330
      TabIndex        =   0
      Top             =   270
      Width           =   4050
      Begin VB.OptionButton Option16 
         Caption         =   "Reconteo"
         Height          =   255
         Left            =   2025
         TabIndex        =   2
         Top             =   390
         Width           =   1095
      End
      Begin VB.OptionButton Option15 
         Caption         =   "Conteo"
         Height          =   375
         Left            =   555
         TabIndex        =   1
         Top             =   345
         Width           =   1035
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3276
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   1560
      Left            =   450
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1185
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1095
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1185
         MaxLength       =   20
         TabIndex        =   8
         Top             =   780
         Width           =   1470
      End
      Begin VB.OptionButton OpTodos 
         Caption         =   "Todos los Artículos"
         Height          =   270
         Left            =   105
         TabIndex        =   7
         Top             =   225
         Width           =   1785
      End
      Begin VB.OptionButton OpRango 
         Caption         =   "Rango"
         Height          =   255
         Left            =   105
         TabIndex        =   6
         Top             =   495
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         Height          =   255
         Left            =   510
         TabIndex        =   11
         Top             =   1095
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio "
         Height          =   255
         Left            =   510
         TabIndex        =   10
         Top             =   795
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmToma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim almacen As String
Dim Conexion As String
Dim Adodc2 As ADODB.Recordset
Dim cSel1 As New ADODB.Recordset
Dim csql As String
'Existe un formulario para la toma de inventario ,y grabarlo en un archivo
'el formulario es FrmInvFis - asi que pueda marca algunos articulos y se quedan grabados


Private Sub Command1_Click()
 Dim CADENA As String
  On Error GoTo Err
     Screen.MousePointer = 11
           If Option15.Value Then
               CrystalReport1.WindowTitle = "Inv087 -- Control de Inventarios"
               CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv087.rpt"
               CrystalReport1.formulas(0) = "empresa ='" & VGparametros.RucEmpresa & "'"
            ElseIf Option16.Value Then
               CrystalReport1.WindowTitle = "Inv088 -- Control de Inventarios"
               CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv088.rpt"
               CADENA = "{STKART.STALMA}='" & VGAlma & "'"
               CrystalReport1.SelectionFormula = CADENA
               CrystalReport1.formulas(0) = "emp ='" & VGparametros.RucEmpresa & "'"
            End If
            Ubi_Tab CrystalReport1
            CrystalReport1.WindowShowPrintBtn = True
            CrystalReport1.WindowShowRefreshBtn = True
            CrystalReport1.WindowShowSearchBtn = True
            CrystalReport1.WindowShowPrintSetupBtn = True
            CrystalReport1.DiscardSavedData = True
            
            CrystalReport1.formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
            CrystalReport1.DiscardSavedData = True
            CrystalReport1.Destination = crptToWindow
            If CrystalReport1.Status <> 2 Then
               CrystalReport1.Action = 1
            End If
      Screen.MousePointer = 1
      Exit Sub
Err:
    MsgBox "No se encontro el reporte", vbInformation, "Aviso"
    Screen.MousePointer = 1
End Sub

Private Sub Command8_Click()
 Unload Me
End Sub



'Private Sub Text1_DblClick()
'         Set Adodc2 = New ADODB.Recordset
'         Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD,p.ACODIGO2 from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "'  ", Vgcnx, adOpenStatic, adLockOptimistic
'         frmReferencia.conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD,p.ACODIGO2 from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  "
'         frmReferencia.Label1.Caption = "Artículos"
'         frmReferencia.show vbmodal
'         Adodc2.Close
'         If vGUtil(1) <> "" Then
'                 Text1 = (vGUtil(1))
'         End If
'         If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
'                 MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
'                 Exit Sub
'        End If
'        If Text1 <> "" Then
'                 Text2.Enabled = True
'                 Text2.SetFocus
'        End If
'End Sub
'
'Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then Text1_DblClick
'End Sub
'
'Private Sub Text1_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'     Set Adodc2 = New ADODB.Recordset
'     Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD,p.ACODIGO2 from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  ", Vgcnx, adOpenStatic, adLockOptimistic
'
'
'End If
'End Sub
'
'
'Private Sub Text2_DblClick()
'    Set Adodc2 = New ADODB.Recordset
'    Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "'  ", Vgcnx, adOpenStatic, adLockOptimistic
'    frmReferencia.conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'  "
'    frmReferencia.Label1.Caption = "Artículos"
'    frmReferencia.show vbmodal
'    Adodc2.Close
'    If vGUtil(1) <> "" Then
'        Text2 = (vGUtil(1))
'    End If
'   If Text2 <> "" Then
'        Command1.SetFocus
'   End If
'End Sub
'
'Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then Text2_DblClick
'End Sub
Private Sub Option16_Click()

End Sub

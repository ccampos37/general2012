VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormKardexMov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimiento de Kardex"
   ClientHeight    =   3540
   ClientLeft      =   240
   ClientTop       =   990
   ClientWidth     =   4425
   Icon            =   "FormKardexMov.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4425
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   255
      Top             =   2145
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   1140
      Picture         =   "FormKardexMov.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2580
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   2700
      Picture         =   "FormKardexMov.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2580
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   2340
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   4185
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FormKardexMov.frx":114E
         Left            =   1290
         List            =   "FormKardexMov.frx":1176
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   735
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1290
         TabIndex        =   12
         Top             =   1485
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1290
         TabIndex        =   11
         Top             =   1845
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   645
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormKardexMov.frx":11DE
         Left            =   1290
         List            =   "FormKardexMov.frx":11E0
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   2760
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   4440
         TabIndex        =   16
         Top             =   1410
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   21364737
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   4470
         TabIndex        =   17
         Top             =   1050
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   21364737
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin VB.Label Label9 
         Caption         =   "Mes"
         Height          =   255
         Left            =   345
         TabIndex        =   18
         Top             =   735
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Artículos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Del"
         Height          =   255
         Left            =   690
         TabIndex        =   14
         Top             =   1470
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Al"
         Height          =   255
         Left            =   690
         TabIndex        =   13
         Top             =   1845
         Width           =   135
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Al"
         Height          =   255
         Left            =   4845
         TabIndex        =   2
         Top             =   675
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Del"
         Height          =   255
         Left            =   4830
         TabIndex        =   1
         Top             =   690
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Moneda"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   5880
      TabIndex        =   8
      Top             =   1410
      Width           =   1695
      Begin VB.OptionButton Option1 
         Caption         =   "Soles"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Dolares"
         Height          =   375
         Left            =   375
         TabIndex        =   9
         Top             =   645
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FormKardexMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As New ADODB.Connection
Dim almacen As String
Dim almacen1 As String

Private Sub Combo1_Click()
 almacen = Format(Combo1.ListIndex + 1, "00")
End Sub

Private Sub Combo2_Click()
almacen1 = Format(Combo1.ListIndex + 1, "00")
End Sub

Private Sub Command1_Click()
Dim Aux, CADENA As String
Dim Codigo2 As String
Dim Dia1 As Integer, Dia2 As Integer, Mes1 As Integer, Mes2 As Integer, Agno1 As Integer, Agno2 As Integer
If Text1 <> "" And Text2 <> "" Then
    Codigo2 = "CENTRAL"
    Dia1 = 1
    Mes1 = Combo3.ListIndex + 1
    Agno1 = Year(DTPicker2)
    Mes2 = Combo3.ListIndex + 1
    Agno2 = Year(DTPicker2)
    Dia2 = Last_Day(Mes2, Agno2)
    CrystalReport1.WindowTitle = "Inv023 -- Control de Inventarios"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv023.rpt"
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = "{MovAlmDet.DEALMA} = '" & almacen & "' and {MovAlmDet.DECODIGO} in '" & Text1 & "' to '" & Text2 & "'  and {MovAlmCab.CAFECDOC} in Date (" & Agno1 & ", " & Mes1 & "," & Dia1 & ") to Date (" & Agno2 & "," & Mes2 & "," & Dia2 & ") and ({MovAlmCab.CATD} <> 'GS'  Or  {MovAlmCab.CACODMOV} <> 'GV' AND {MovAlmCab.CASITGUI} <> 'F')"
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.formulas(0) = "TXTALMACEN = '" & UCase(Combo1.text) & "'"
    CrystalReport1.formulas(1) = "Mes= '" & UCase(Combo3.text) & "'"
    CrystalReport1.formulas(2) = "EMP= '" & UCase(VGparametros.RucEmpresa) & "'"
    CrystalReport1.formulas(3) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.formulas(4) = "Alma ='" & almacen & "'"
    CrystalReport1.formulas(5) = "MinArt ='" & Text1 & "'"
    CrystalReport1.formulas(6) = "MaxArt ='" & Text2 & "'" '
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End If
End Sub

Private Sub Carga_Almacen()
Dim rsql As String
Dim rs As New ADODB.Recordset

rsql = "select  TADESCRI FROM TabAlm "
'Set db = Workspaces(0).OpenDatabase(cRuta2)
'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)

Set rs = VGCNx.Execute(rsql)
While Not rs.EOF
    Combo1.AddItem (rs(0))
    Combo2.AddItem (rs(0))
    rs.MoveNext
Wend
rs.Close
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Load()
Carga_Almacen
If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
central Me
Option1.Value = True
DTPicker1.Value = Date
DTPicker2.Value = Date
If Month(Date) <> 1 Then
        Combo3.ListIndex = Month(DTPicker2) - 1
End If
End Sub


Private Sub Text1_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "'  ", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "'  "
frmReferencia.Label1.Caption = "Artículos"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
        Text1 = (vGUtil(1))
End If
If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
        MsgBox "Ingrese un Código menor al Fin ", vbOKOnly, "Error"
        Exit Sub
End If
If Text1 <> "" Then
        Text2.Enabled = True
        Text2.SetFocus
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text1_DblClick
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1 <> "" Then
      If Existe_cod_art(Text1) <> "" Then
              Text2.Enabled = True
              Text2.SetFocus
      End If
End If
End Sub

Private Sub Text2_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
Adodc2.Open "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "'  ", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "Select p.ACODIGO, p.ADESCRI,p.AUNIDAD from MaeArt p, StkArt n Where n.STSKDIS<> 0 and  p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "'  "
frmReferencia.Label1.Caption = "Artículos"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
        Text2 = (vGUtil(1))
End If
 If Trim(Text2) <> "" Then Command1.SetFocus
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text2_DblClick
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 And Trim(Text2) <> "" Then
      Text2 = Trim(Text2)
      If Existe_cod_art(Text2) <> "" Then
            If Text1 > Text2 Then
                   MsgBox "El código fin debe ser mayor que el inicio", vbExclamation, "Aviso"
                   Exit Sub
            End If
            Command1.SetFocus
      End If
End If
End Sub

Function Existe_cod_art(text As TextBox) As String
Dim rs As Recordset
Dim rsql As String
rsql = "select  ACODIGO FROM maeart where ACODIGO = '" & text & "'" '
'Set db = Workspaces(0).OpenDatabase(cRuta2)
'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)

Set rs = VGCNx.Execute(rsql)
If Not rs.EOF Then
    Existe_cod_art = rs(0)
Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
    Existe_cod_art = ""
End If
rs.Close
End Function

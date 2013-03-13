VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form formkardexValResumen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kardex Valorizado Resumido "
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "formkardexValResumen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3420
      Left            =   120
      TabIndex        =   1
      Top             =   30
      Width           =   4875
      Begin VB.CheckBox Check1 
         Caption         =   "Todos los Almacenes"
         Height          =   435
         Left            =   1260
         TabIndex        =   10
         Top             =   315
         Width           =   1695
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   585
         Left            =   1470
         Picture         =   "formkardexValResumen.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2475
         Width           =   735
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   600
         Left            =   2370
         Picture         =   "formkardexValResumen.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2475
         Width           =   735
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "formkardexValResumen.frx":0CC6
         Left            =   1440
         List            =   "formkardexValResumen.frx":0CD0
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1815
         Width           =   2115
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "formkardexValResumen.frx":0CE4
         Left            =   1425
         List            =   "formkardexValResumen.frx":0CE6
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   990
         Width           =   2760
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1425
         TabIndex        =   2
         Top             =   1395
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MMMM 'del' yyyy"
         Format          =   62652419
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda   :"
         Height          =   300
         Left            =   480
         TabIndex        =   7
         Top             =   1890
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen  :"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   495
         TabIndex        =   5
         Top             =   1035
         Width           =   795
      End
      Begin VB.Label Label9 
         Caption         =   "Mes         :"
         Height          =   255
         Left            =   495
         TabIndex        =   4
         Top             =   1470
         Width           =   855
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   195
      Left            =   3885
      TabIndex        =   0
      Top             =   4065
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   344
      _Version        =   393216
      Format          =   62652417
      CurrentDate     =   36710
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   150
      Top             =   2310
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "formkardexValResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim almacen As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Combo1.Enabled = False
 Else
   Combo1.Enabled = True
End If
End Sub

Private Sub CmdAceptar_Click()
Dim rsql As String
Dim Va1 As String, Va2 As String
Dim arrform(6) As Variant
Dim arrparam(3) As Variant
Dim Reporte As String, titrep As String


Set Adodc3 = New ADODB.Recordset  'Para sacar la descripcion del rango elegido

  If Left(Combo3.text, 1) <> "D" And Left(Combo3.text, 1) <> "S" Then
     MsgBox "Debe Seleccionar el Tipo de Moneda...! ", vbInformation, "Corregir"
     Exit Sub
  End If
   
  If Combo3.ListIndex <> 0 Then
     titrep = "Inv503 -- Control de Inventarios"
     Reporte = "inv503.rpt"
  Else
     titrep = "Inv502 -- Control de Inventarios"
     Reporte = "inv502.rpt"
  End If
  
  Dim ccadena As String
  If Reporte = "" Then
      MsgBox "No hay registros a imprimir", vbInformation, "Aviso"
      Screen.MousePointer = 1
      Exit Sub
  End If
  
 ' ccadena = "{MORESMES.SMALMA}='" & almacen & "' and {MORESMES.SMMESPRO}='" & CStr(Format(DTPicker1.Year, "0000") & Format(DTPicker1.Month, "00")) & "' AND {@SALDO}<>0"

  
 arrform(0) = "ALMACEN = '" & UCase(Combo1.text) & "'"
  arrform(1) = "Mes= '" & UCase(Format(DTPicker1, "MMMM - yyyy")) & "'"
  arrform(2) = "EMPRESA= '" & UCase(VGparametros.RucEmpresa) & "'"
  arrform(3) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
  If Combo3.ListIndex <> 0 Then
     arrform(4) = "MONEY= 'DOLAR'"
  Else
      arrform(4) = "MONEY= 'SOLES'"
  End If
  
If Check1.Value = 1 Then
   arrform(0) = "ALMACEN ='Todos'"
   almacen = "%%"
   titrep = "Inv504 -- Control de Inventarios"
   Reporte = "inv504.rpt"
End If
  arrparam(0) = VGCNx.DefaultDatabase
  arrparam(1) = almacen
  arrparam(2) = CStr(Format(DTPicker1.Year, "0000") + Format(DTPicker1.Month, "00"))
 
Call ImpresionRptProc(Reporte, arrform, arrparam, "", titrep)
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
        rs.MoveFirst
        rs.Move Combo1.ListIndex
        almacen = Format(rs(0), "00")
End Sub

Private Sub Form_Load()
central Me
Carga_Almacen
If Combo1.ListIndex = 0 Then VGForm1 = 6
DTPicker1.Value = Date
End Sub
Private Sub Carga_Almacen()
Dim rsql As String
Dim I As Integer
rsql = "Select TAALMA,TADESCRI FROM TabAlm "
Set rs = New ADODB.Recordset
rs.Open rsql, VGCNx, adOpenStatic
Do While Not rs.EOF
     Combo1.AddItem (rs(1))
     rs.MoveNext
     If rs.EOF Then Exit Do
Loop
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


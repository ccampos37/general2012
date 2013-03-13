VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmInvFis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Toma de Inventarios"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   3735
      Picture         =   "FrmInvFis.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5205
      Width           =   825
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   675
      Left            =   2415
      Picture         =   "FrmInvFis.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5205
      Width           =   825
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Articulo"
      Height          =   4935
      Left            =   216
      TabIndex        =   0
      Top             =   144
      Width           =   6510
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Marcar Todo"
         Height          =   375
         Index           =   0
         Left            =   3060
         TabIndex        =   2
         Top             =   225
         Width           =   1230
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Desmarcar Todo"
         Height          =   375
         Index           =   1
         Left            =   4755
         TabIndex        =   1
         Top             =   225
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid FG 
         Height          =   3960
         Left            =   150
         TabIndex        =   3
         Top             =   765
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   6985
         _Version        =   393216
         Cols            =   3
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   690
      Top             =   5565
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmInvFis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As Recordset

'Este formulario sirve para generar una tabla de los registro que van a tener un
'inventario inicial e imprime los articulos que voy a realizar la toma de inventario
Private Sub CmdAceptar_Click()
Dim I As Integer
Dim insertar1 As String
Dim corre As String
Dim adodc1 As New ADODB.Recordset

adodc1.Open "Select icorrela from INVFIS Order By ICorrela ", VGCNx, adOpenStatic, adLockOptimistic
If adodc1.RecordCount = 0 Then
  corre = "0001"
ElseIf adodc1.RecordCount > 0 Then
  adodc1.MoveLast
  corre = Format(CInt(adodc1("iCorrela")) + 1, "0000")
End If
adodc1.Close

  For I = 0 To FG.Rows - 1
   If FG.TextMatrix(I, 0) = "»" Then
      insertar1 = "insert into INVFIS (ICORRELA,ICODART) values ('" & corre & "','" & FG.TextMatrix(I, 1) & "')"
      VGCNx.Execute insertar1
      'varform.Salida.AddItem (FG.TextMatrix(I, 1) & vbTab & FG.TextMatrix(I, 2) & vbTab & FG.TextMatrix(I, 3) & vbTab & FG.TextMatrix(I, 5) & vbTab & FG.TextMatrix(I, 6))
   End If
  Next I
  insertar1 = "insert into HISTO_INV (HCORRELA,HFECHA,HALMA) values ('" & corre & "','" & Date & "'," & VGAlma & ")"
  VGCNx.Execute insertar1

'''''''''''''
'If Option1.Value Then
      cadena = "{INVFIS.ICORRELA}='" & corre & "'"
      CrystalReport1.SelectionFormula = cadena
      CrystalReport1.WindowShowPrintBtn = True
      CrystalReport1.WindowShowRefreshBtn = True
      CrystalReport1.WindowShowSearchBtn = True
      CrystalReport1.WindowShowPrintSetupBtn = True
      CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv097.rpt"
      Ubi_Tab CrystalReport1
      CrystalReport1.DiscardSavedData = True
      CrystalReport1.Destination = crptToWindow
      CrystalReport1.formulas(0) = "emp ='" & VGparametros.RucEmpresa & "'"
      CrystalReport1.formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
' Else
'    CrystalReport1.ReportFileName =  VGParamSistem.RutaReport & "tomainv2.rpt"
'    Ubi_Tab CrystalReport1
'    CrystalReport1.Formulas(0) = "emp ='" & VGparametros.RucEmpresa & "'"
'    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
'    CrystalReport1.DiscardSavedData = True
'End If
  CrystalReport1.WindowTitle = "Reporte de Inventario Físico - Conteo"
  If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click(Index As Integer)
Dim I As Integer
Select Case Index
 Case 0:
        For I = 1 To FG.Rows - 1
         FG.TextMatrix(I, 0) = "»"
        Next I
 Case 1:
        For I = 1 To FG.Rows - 1
         FG.TextMatrix(I, 0) = ""
        Next I
End Select
End Sub

Private Sub Flex1_Click()
If Flex1.TextMatrix(Flex1.Row, 0) = "»" Then
  Flex1.TextMatrix(Flex1.Row, 0) = ""
Else
  Flex1.TextMatrix(Flex1.Row, 0) = "»"
End If
End Sub

Private Sub FG_Click()
  If FG.Row = 0 Then Exit Sub
  If FG.TextMatrix(FG.Row, 0) = "»" Then
     FG.TextMatrix(FG.Row, 0) = " "
  Else
     FG.TextMatrix(FG.Row, 0) = "»"
  End If
End Sub

Private Sub Form_Activate()
Dim cCod As String
Dim nStock As Double
If Rs.RecordCount = 0 Then
    MsgBox "No hay articulos disponibles en el almacen", vbInformation, "Aviso"
'    Form_Unload (0)
    Exit Sub
  End If
  Rs.MoveFirst
  FG.Visible = False
  cCod = "": nStock = 0
  Do While Not Rs.EOF
     cCod = Rs(0)
     If Rs("Alma1") = VGAlma Or Rs("Alma1") = "xx" Then
            FG.AddItem (" " & vbTab & Rs(0) & vbTab & Rs(1) & vbTab & Rs(2) & vbTab & IIf(IsNull(Rs(3)), 0, Format(Rs(3), "##0.000")) & vbTab & Rs(4) & vbTab & Rs(5) & vbTab & Rs(6))
     End If
     Rs.MoveNext
     If Rs.EOF Then Exit Do
  Loop
  FG.Visible = True
End Sub

Private Sub Form_Load()
  Dim real As Double
 ''Dim db As Database
  Dim Cod As String
  Dim rsql As String
  Dim varform As Form
  central Me
  Set varform = Me
  Cod = ""
'  varform.Salida.Rows = 1
     
  rsql = "select  p.ACODIGO, p.ADESCRI,p.AUNIDAD, n.STSKDIS, p.AFSERIE, p.AFLOTE  ,p.AFAMILIA,n.STALMA as Alma1 from MaeArt p, StkArt n where p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'   and n.STSKDIS <>0  ORDER BY ACODIGO "

  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  Set Rs = VGCNx.Execute(rsql)
  
  FG.FormatString = "^Seleccion|  Codigo|   Descripcion|  Unidad | Stock   |Se | Lt | Familia "
  FG.Row = 0
  FG.ColWidth(0) = 910
  FG.ColWidth(1) = 1200
  FG.ColWidth(2) = 3250
  FG.ColWidth(3) = 800
  FG.ColWidth(4) = 1200
  FG.ColWidth(5) = 2
  FG.ColWidth(6) = 2
  FG.ColWidth(7) = 1000
 
  FG.ColAlignment(1) = 1
'  varform.Salida.FormatString = "Codigo|Descripcion|Unidad|sr|lt"
'  varform.Salida.Row = 0
'  varform.Salida.ColWidth(0) = 1500
'  varform.Salida.ColWidth(1) = 1500
'  varform.Salida.ColWidth(2) = 500
'  varform.Salida.ColWidth(3) = 200
'  varform.Salida.ColWidth(4) = 200
  FG.Rows = 1
  
End Sub

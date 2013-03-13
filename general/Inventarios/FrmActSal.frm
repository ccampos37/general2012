VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmActSal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualiza Saldos"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "FrmActSal.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1785
      Left            =   285
      TabIndex        =   2
      Top             =   75
      Width           =   4725
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1905
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   435
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1875
         TabIndex        =   6
         Top             =   1080
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MMMM 'del' yyyy"
         Format          =   53870595
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35431
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
         Height          =   270
         Left            =   570
         TabIndex        =   5
         Top             =   465
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Mes"
         Height          =   255
         Left            =   585
         TabIndex        =   4
         Top             =   1080
         Width           =   1410
      End
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   2880
      Picture         =   "FrmActSal.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2100
      Width           =   800
   End
   Begin VB.CommandButton CmdSaldos 
      Caption         =   " S&aldos"
      Height          =   735
      Left            =   1695
      Picture         =   "FrmActSal.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2100
      Width           =   800
   End
End
Attribute VB_Name = "FrmActSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim db As Database
Dim almacen As String
Dim Conexion As String
Dim Adodc3 As ADODB.Recordset

Private Sub CmdSaldos_Click()
Dim adodc1 As New ADODB.Recordset
Dim Ado2 As New ADODB.Recordset
Dim Sql1 As String, Sql2 As String
Dim Mess As String, C As Integer
Dim ING As Double, SAL As Double
Screen.MousePointer = 11

Mess = Format(Month(DTPicker1), "00")
Sql1 = "Select DECODIGO,DETD,SUM(DECANTID) as Cantidad from MOVALMDET A Inner Join MOVALMCAB B on "
Sql1 = Sql1 & "A.DETD=B.CATD and A.DEALMA=B.CAALMA and A.DENUMDOC=B.CANUMDOC "
Sql1 = Sql1 & "Where DeAlma='" & almacen & "' and Month(CAFECDOC)=" & Mess & " AND   YEAR(CAFECDOC)=" & Year(DTPicker1) & ""
Sql1 = Sql1 & "and CASITGUI<>'A'  AND CACODMOV <> 'GF'  AND DECODIGO<>'TEXTO' Group By DECODIGO,DETD Order By DECODIGO"

adodc1.Open Sql1, VGcnx, adOpenDynamic, adLockOptimistic
C = 0
If adodc1.RecordCount = 0 Then MsgBox "No existe Movimientos de almacén", vbInformation, "Aviso": adodc1.Close: Screen.MousePointer = 1: Exit Sub
adodc1.MoveFirst
'''Actualiza las entradas y salidas a cero
Sql2 = "Update MORESMES set SMCANENT=0,SMCANSAL=0 Where SMALMA='" & almacen & "' "
Sql2 = Sql2 & "and  SMMESPRO='" & Year(DTPicker1) & Mess & "'"
VGcnx.Execute Sql2

adodc1.MoveFirst
Do While Not adodc1.EOF
  'Verifica si existe el articulo en MORESMES
  
  Ado2.Open "Select SMCODIGO from MORESMES Where SMALMA='" & almacen & "' and SMMESPRO='" & Year(DTPicker1) & Mess & "' and SMCODIGO='" & adodc1("DECODIGO") & "'", VGcnx, adOpenDynamic, adLockOptimistic
  If Ado2.RecordCount = 0 Then
     VGcnx.Execute "Insert into MORESMES (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMSALDOINI) Values('" & almacen & "','" & adodc1("DECODIGO") & "','" & Year(DTPicker1) & Mess & "',0,0,0)"
  End If
  Ado2.Close
  If UCase(adodc1("DECODIGO")) <> "TEXTO" Then
    If adodc1("DETD") = "NI" Or adodc1("DETD") = "NC" Then
      Sql2 = "Update MORESMES set SMCANENT= SMCANENT + " & adodc1("CANTIDAD") & " Where SMALMA='" & almacen & "' "
      Sql2 = Sql2 & "and  SMMESPRO='" & Year(DTPicker1) & Mess & "' and SMCODIGO='" & adodc1("DECODIGO") & "'"
      VGcnx.Execute Sql2
      C = C + 1
    Else
     
      Sql2 = "Update MORESMES set SMCANSAL= SMCANSAL + " & adodc1("CANTIDAD") & " Where SMALMA='" & almacen & "' "
      Sql2 = Sql2 & "and  SMMESPRO='" & Year(DTPicker1) & Mess & "' and SMCODIGO='" & adodc1("DECODIGO") & "'"
      VGcnx.Execute Sql2
      C = C + 1
    End If
  End If
  adodc1.MoveNext
Loop
Stock
adodc1.Close
Screen.MousePointer = 1
MsgBox "Se hicieron " & C & " Actualizaciones", vbInformation, "Inventarios"
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Stock()
Dim adodc1 As New ADODB.Recordset
Dim Ado2 As New ADODB.Recordset
Dim Sql1 As String, Sql2 As String
Dim C As String, Stock As Double
Screen.MousePointer = 11
C = 0
Sql1 = "Select SMALMA,SMCODIGO,SUM(SMCANENT) AS ENTRADA,SUM(SMCANSAL) AS SALIDA from MORESMES   "
Sql1 = Sql1 & "Where SMALMA='" & almacen & "'  Group by SMCODIGO,SMALMA"
adodc1.Open Sql1, VGcnx, adOpenDynamic, adLockOptimistic

If adodc1.RecordCount = 0 Then MsgBox "No existe Stock de almacén", vbInformation, "Aviso": adodc1.Close: Screen.MousePointer = 1: Exit Sub
adodc1.MoveFirst
Do While Not adodc1.EOF
    Stock = adodc1("ENTRADA") - adodc1("SALIDA")
    
    'Verifica si existe el articulo en STKART
    Ado2.Open "Select STCODIGO from STKART Where STALMA='" & almacen & "' and STCODIGO='" & adodc1("SMCODIGO") & "'", VGcnx, adOpenDynamic, adLockOptimistic
    If Ado2.RecordCount = 0 Then
      VGcnx.Execute "Insert into STKART (STALMA,STCODIGO,STSKDIS) Values('" & almacen & "','" & adodc1("SMCODIGO") & "'," & Stock & ")"
    Else
      Sql2 = "Update STKART set STSKDIS=" & Stock & " Where STALMA='" & almacen & "' "
      Sql2 = Sql2 & "and  STCODIGO='" & adodc1("SMCODIGO") & "'"
      VGcnx.Execute Sql2
    End If
    Ado2.Close
    
    C = C + 1
  adodc1.MoveNext
Loop
adodc1.Close
MsgBox "Se hicieron " & C & " Actualizaciones", vbInformation, "Inventarios"
Screen.MousePointer = 1
End Sub



Private Sub Combo1_Click()
almacen = Trim(Mid(Combo1.text, 1, 2))
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub



Private Sub Form_Load()
Carga_Almacen
central Me
DTPicker1.Value = Date
End Sub

Private Sub Carga_Almacen()
Dim rsql As String
Dim rs As Recordset
 
rsql = "select TAALMA,TADESCRI FROM TabAlm "
'Set db = Workspaces(0).OpenDatabase(cRuta2)
'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = VGcnx.Execute(rsql)
While Not rs.EOF
    Combo1.AddItem (rs(0)) & " " & (rs(1))
    rs.MoveNext
Wend
Combo1.ListIndex = 0
rs.Close
End Sub



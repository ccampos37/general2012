VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmkardexDoc 
   Caption         =   "Kardex Valorizado por Dcto."
   ClientHeight    =   4470
   ClientLeft      =   1965
   ClientTop       =   3615
   ClientWidth     =   5805
   Icon            =   "frmkardexDoc.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5805
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3270
      Picture         =   "frmkardexDoc.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3585
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   1845
      Picture         =   "frmkardexDoc.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   3270
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmkardexDoc.frx":114E
         Left            =   1755
         List            =   "frmkardexDoc.frx":1158
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2685
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   270
         Left            =   1755
         TabIndex        =   1
         Top             =   825
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   476
         _Version        =   393216
         CustomFormat    =   "MMMM - yyyy"
         Format          =   47448067
         CurrentDate     =   36913
         MinDate         =   36404
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1755
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2250
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1755
         MaxLength       =   2
         TabIndex        =   4
         Top             =   2175
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1755
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1740
         Width           =   495
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmkardexDoc.frx":116C
         Left            =   1755
         List            =   "frmkardexDoc.frx":1179
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1260
         Width           =   2250
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo Moneda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2730
         Width           =   1425
      End
      Begin VB.Label lbltrans2 
         Caption         =   "lbltrans2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2370
         TabIndex        =   14
         Top             =   2175
         Width           =   2835
      End
      Begin VB.Label lbltrans1 
         Caption         =   "lbltrans1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2370
         TabIndex        =   13
         Top             =   1755
         Width           =   2865
      End
      Begin VB.Label Label5 
         Caption         =   "Al Movimiento"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2175
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Del Movimiento"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Movimiento"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1260
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Mes"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   870
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Por Almacen"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmkardexDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dato As String * 1
Dim almacen As String * 2
'Dim db As Database
Private Sub Combo1_Click()
    almacen = Format(Combo1.ListIndex + 1, "00")
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub



Private Sub Combo3_Click()
    Text1 = ""
    Text2 = ""
    lbltrans1 = ""
    lbltrans2 = ""
 If Combo3.ListIndex = 2 Then
        Command1.SetFocus
 Else
     If Not Text1.Enabled Then
       Text1.SetFocus
     End If
  End If
End Sub

Private Sub Combo3_Change()
If Combo1.ListIndex = 0 Then
          VGRegEnt = 1
Else
           VGRegEnt = 2
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Command1_Click()
Screen.MousePointer = 11
If Combo3.ListIndex = 2 Then
   Imprimir2
Else
    imprimir
End If
Screen.MousePointer = 1
End Sub

Private Sub Command7_Click()
 Unload Me
End Sub


Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
   Carga_Almacen
   If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
   DTPicker1 = Date
   
   If Combo3.ListCount > 0 Then Combo3.ListIndex = 0
   central Me
   lbltrans1 = ""
   lbltrans2 = ""
   Text1 = "AJ"
   Text2 = "DV"
   lbltrans1 = "AJUSTE"
   lbltrans2 = Mid("DEVOLUCION DE PRODUCCION", 1, 20)
   If Combo3.ListIndex = 0 Then
          dato = Trim("I")
        Else
          dato = Trim("S")
    End If
    Combo2.ListIndex = 0
End Sub

Private Sub Carga_Almacen()
   Dim rsql As String
   Dim rs As New ADODB.Recordset

   rsql = "select  TADESCRI FROM TabAlm "
   'Set db = Workspaces(0).OpenDatabase(cRuta2)
   'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(rsql)
   If rs.RecordCount > 0 Then
   While Not rs.EOF
      Combo1.AddItem (rs(0))
      rs.MoveNext
   Wend
   End If
   rs.Close
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
       Text1_DblClick
    ElseIf KeyCode = 46 Then
        lbltrans1 = ""
    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  '****************** TRANSACCION
  If KeyAscii = 13 And Len(Text1.text) = 2 Then
           buscar_trans Text1, lbltrans1
           lbltrans1 = Mid(lbltrans1, 1, 20)
           If lbltrans1 <> "" Then Text2.SetFocus
     Else
         If KeyAscii = 8 Then lbltrans1 = ""
    End If
End Sub
Private Sub Text1_DblClick()
VGForm = 20
If Combo3.ListIndex = 0 Then
  VGRegEnt = 1
Else
  VGRegEnt = 2
End If
Text1 = ""
 FormAyuTransa.Show 1
    If Text1 <> "" Then
      lbltrans1 = Mid(lbltrans1, 1, 20)
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then
     Text2_DblClick
 ElseIf KeyCode = 46 Then
    lbltrans2 = ""
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  '****************** TRANSACCIONES
     If KeyAscii = 13 And Len(Text2.text) = 2 Then
           buscar_trans Text2, lbltrans2
           If lbltrans2 <> "" Then Command1.SetFocus
           lbltrans2 = Mid(lbltrans2, 1, 20)
     Else       'habilitado (False)
         If KeyAscii = 8 Then lbltrans2 = ""
   
    End If
End Sub
Private Sub Text2_DblClick()
VGForm = 20
If Combo3.ListIndex = 0 Then
  VGRegEnt = 1
Else
  VGRegEnt = 2
End If
 FormAyuTransa.Show 1
    If Text1 <> "" Then
      'buscar_trans Text2, lbltrans2
       lbltrans1 = Mid(lbltrans1, 1, 20)
    End If
End Sub


Sub buscar_trans(texto As TextBox, lbl As Label)
  Dim rsql As String
  Dim rs As New ADODB.Recordset
        If Combo3.ListIndex = 0 Then
          dato = Trim("I")
        Else
          dato = Trim("S")
        End If
        rsql = "select  *  from TabTransa  where TT_CODMOV ='" & texto & "' and TT_TIPMOV ='" & dato & "'"
'        Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
         Set rs = VGCNx.Execute(rsql)
         If rs.EOF Then
            MsgBox "El tipo de transaccion no existe !", vbOKOnly, "Error"
            texto.SetFocus
            Exit Sub
         End If
         texto = UCase(texto)
         lbl = Mid(rs("TT_DESCRI"), 1, 30)
End Sub

Private Sub imprimir()
Dim CADENA As String
Text1 = Trim(Text1)
Text2 = Trim(Text2)
If Trim(Text1) = "" Then
    MsgBox "Ingrese el codigo", vbInformation, "Aviso"
    Text1.SetFocus
    Exit Sub
End If
If (Text1 = Text2) Or (Text2 = "") Then
      CADENA = "{MOVALMCAB.CAALMA}='" & Format(Combo1.ListIndex + 1, "00") & "'  and  {MOVALMCAB.CACODMOV}='" & Text1.text & "'  AND {MOVALMCAB.CATIPMOV}='" & dato & "' and MONTH({MOVALMCAB.CAFECDOC}) = " & Month(DTPicker1) & " and YEAR({MOVALMCAB.CAFECDOC}) = " & Year(DTPicker1) & ""      ''({STKART.STCODIGO} in '" & codigo1 & "' to '" & CODIGO2 & "')"
Else
      CADENA = "{MOVALMCAB.CAALMA}='" & Format(Combo1.ListIndex + 1, "00") & "' and ( {MOVALMCAB.CACODMOV} in '" & Text1.text & "'  to '" & Text2.text & "')  AND {MOVALMCAB.CATIPMOV}='" & dato & "' and MONTH({MOVALMCAB.CAFECDOC}) = " & Month(DTPicker1) & " and YEAR({MOVALMCAB.CAFECDOC}) = " & Year(DTPicker1) & ""
End If
   CrystalReport1.WindowTitle = "Inv031 - Sistema de Inventarios"
   CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv031.rpt"
   Ubi_Tab CrystalReport1
   CrystalReport1.DiscardSavedData = True
   CrystalReport1.Destination = crptToWindow
   CrystalReport1.SelectionFormula = CADENA
   CrystalReport1.WindowShowPrintBtn = True
   CrystalReport1.WindowShowRefreshBtn = True
   CrystalReport1.WindowShowSearchBtn = True
   CrystalReport1.WindowShowPrintSetupBtn = True
   CrystalReport1.formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
   CrystalReport1.formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
   CrystalReport1.formulas(2) = "almacen = '" & Combo1.text & "'"
   If CrystalReport1.Status <> 2 Then
      CrystalReport1.Action = 1
   End If
End Sub


Private Sub Imprimir2()
Dim CADENA As String
   CADENA = ""
   CrystalReport1.WindowTitle = "Inv031 - Sistema de Inventarios"
   CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv031.rpt"
   Ubi_Tab CrystalReport1
   CrystalReport1.DiscardSavedData = True
   CrystalReport1.Destination = crptToWindow
   CrystalReport1.SelectionFormula = CADENA
   CrystalReport1.WindowShowPrintBtn = True
   CrystalReport1.WindowShowRefreshBtn = True
   CrystalReport1.WindowShowSearchBtn = True
   CrystalReport1.WindowShowPrintSetupBtn = True
   CrystalReport1.formulas(0) = "EMP='" & VGparametros.RucEmpresa & "' "
   CrystalReport1.formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
   CrystalReport1.formulas(2) = "almacen = '" & Combo1.text & "'"
   If CrystalReport1.Status <> 2 Then
      CrystalReport1.Action = 1
   End If

End Sub

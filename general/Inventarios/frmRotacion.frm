VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRotacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rotación Mensual"
   ClientHeight    =   3975
   ClientLeft      =   810
   ClientTop       =   1530
   ClientWidth     =   6240
   Icon            =   "frmRotacion.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6240
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5895
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   3360
         TabIndex        =   4
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MMMM'del' yyyy"
         Format          =   24772611
         CurrentDate     =   36710
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmRotacion.frx":08CA
         Left            =   720
         List            =   "frmRotacion.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   3000
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1440
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   720
         TabIndex        =   3
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MMMM 'del' yyyy"
         Format          =   24772611
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.Label Label9 
         Caption         =   "Al"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Del"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Del"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Al"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Al"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "Del"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Articulos :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3240
      Picture         =   "frmRotacion.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3045
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   720
      Left            =   1950
      Picture         =   "frmRotacion.frx":0D10
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3045
      Width           =   735
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRotacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mesini, mesfin As String
Dim diferencia As Integer
'Dim db As Database
Dim almacen As String
Dim almacenAnt As String

Private Sub Combo1_Click()
'  almacen = Format(Combo1.ListIndex + 1, "00")
almacen = Mid(Combo1.text, 1, 2)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text1.SetFocus
End Sub

Private Sub Combo2_Click()
'almacen1 = Format(Combo1.ListIndex + 1, "00")
almacen = Mid(Combo2.text, 1, 2)
End Sub

Private Sub Command1_Click()
 Dim Aux, cadena As String
 Dim mespro As String
 Dim mesp1 As Integer
 Dim mesp2 As Integer
 Dim mesreporte As Integer
 Dim nf As Integer
 Dim i As Integer
          Screen.MousePointer = 11
          If Text1 = "" And Text2 = "" Then
              MsgBox "No ingresó rango de artículos", vbExclamation, "Error"
              Screen.MousePointer = 1
              Exit Sub
          End If
          ' verificar si los codigos existen
          mesp1 = Month(DTPicker1)
          mesp2 = Month(DTPicker2)
          mesini = Year(DTPicker1) & Format(Month(DTPicker1), "00")
          mesfin = Year(DTPicker2) & Format(Month(DTPicker2), "00")
          If Year(DTPicker1) <> Year(DTPicker2) Then
            diferencia = (mesfin + 12) - (mesini + 100)
          Else
            diferencia = Val(mesfin) - Val(mesini)
          End If
          If diferencia < 0 Then
                MsgBox "Error en el rango de meses", vbExclamation, "Error"
                Screen.MousePointer = 1
                Exit Sub
          End If
          If diferencia > 12 Then
                   MsgBox "El rango de meses No puede ser mayor que 12", vbExclamation, "Aviso"
                   Screen.MousePointer = 1
                   Exit Sub
           End If
           '****************************
           rotacion
           '****************************
           If diferencia > 6 Then
              CrystalReport1.WindowTitle = "Inv052 -- Control de Inventarios"
              CrystalReport1.ReportFileName = cRutP & "inv052.rpt"
           Else
              CrystalReport1.WindowTitle = "Inv051-- Control de Inventarios"
              CrystalReport1.ReportFileName = cRutP & "inv051.rpt"
           End If
           Ubi_Tab CrystalReport1
           CrystalReport1.DiscardSavedData = True
           CrystalReport1.Destination = crptToWindow
            'cadena = "{MORESMES.SMALMA}='" & almacen & "' and({MOREMES.SMCODIGO} in '" & Text1 & "' to '" & Text2 & "')"
            'CrystalReport1.SelectionFormula = cadena
            CrystalReport1.WindowShowPrintBtn = True
            CrystalReport1.WindowShowRefreshBtn = True
            CrystalReport1.WindowShowSearchBtn = True
            CrystalReport1.WindowShowPrintSetupBtn = True
           CrystalReport1.Formulas(0) = "emp ='" & VGNemp & "'"
           CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
           CrystalReport1.Formulas(2) = "alm ='" & Trim(Mid(Combo1.text, 5, 20)) & "'"
           CrystalReport1.Formulas(3) = "factor =" & (diferencia + 1) & ""
           nf = 4
           mesreporte = mesp1 - 1
           For i = 1 To 12
              If (diferencia + 1) >= i Then
                 CrystalReport1.Formulas(nf) = "Mes" & i & "  = '" & Format(mesreporte, "00") & "'"
              Else
                CrystalReport1.Formulas(nf) = "Mes" & i & "  = 'XX' "
              End If
              nf = nf + 1
              mesreporte = mesreporte + 1
              If mesreporte = 12 Then mesreporte = 0
           Next i
           
           If CrystalReport1.Status <> 2 Then
                CrystalReport1.Action = 1
           End If
           Screen.MousePointer = 1
End Sub

Private Sub Command7_Click()
 Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then DTPicker2.SetFocus
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1.SetFocus
End Sub

Private Sub Form_Load()
   Carga_Almacen
   DTPicker1 = DateAdd("m", -3, Date)
   DTPicker2 = Date
   'Combo1.ListIndex = 0
   'Combo1.ListIndex = 0
   central Me
   VGForm1 = 9
End Sub

Private Sub Text1_DblClick()
 VGForm1 = 9
 almacenAnt = VGAlma
 VGAlma = almacen
  FormAyuArt1.Show 1
   If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
      MsgBox "Ingrese un código menor al fin ", vbOKOnly, "Error"
      VGAlma = almacenAnt
      Exit Sub
   End If
   If Text1 <> "" Then
        Text2.Enabled = True
        Text2.SetFocus
   End If
   VGAlma = almacenAnt
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   Text1_DblClick
End If
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
   FormAyuArt1.Show 1
   almacenAnt = VGAlma
  VGAlma = almacen
   If Text2 <> "" Then
        DTPicker1.SetFocus
   End If
  VGAlma = almacenAnt
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
  Text2_DblClick
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  
 If KeyAscii = 13 And Text2 <> "" Then
      
      If Existe_cod_art(Text2) <> "" Then
         If Text1 > Text2 Then
         
         MsgBox "El codigo fin debe ser mayor que el inicio"
                               Exit Sub
         End If
         DTPicker1.SetFocus
      End If
      
   End If
End Sub

Function Existe_cod_art(text As TextBox) As String
 
 Dim rs As Recordset
 Dim RSQL As String
  RSQL = "select  ACODIGO FROM maeart where ACODIGO = '" & text & "'" '
  'Set db = Workspaces(0).OpenDatabase(cRuta2)
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
  Set rs = cConexCom.Execute(RSQL)
  If Not rs.EOF Then
    Existe_cod_art = rs(0)
  Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly + vbInformation, "Error"
    Existe_cod_art = ""
  End If
   rs.Close
End Function

Private Sub Carga_Almacen()
   Dim RSQL As String
   Dim rs As Recordset
   Dim i As Integer
   RSQL = "select TAALMA,TADESCRI FROM TabAlm "
   'Set db = Workspaces(0).OpenDatabase(cRuta2)
   'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = cConexCom.Execute(RSQL)
   If rs.RecordCount > 0 Then
   While Not rs.EOF
     Combo1.AddItem (rs(0)) & "   " & (rs(1))
     Combo2.AddItem (rs(0)) & "   " & (rs(1))
     rs.MoveNext
   Wend
   End If
   rs.MoveFirst
   For i = 0 To rs.RecordCount - 1
      If rs(0) = VGAlma Then
        Combo1.ListIndex = i
        Exit For
      Else
        rs.MoveNext
      End If
    Next
   Combo1_Click
   rs.Close
  
End Sub

Private Sub rotacion()
 Dim RSQL, Rsql1 As String
 Dim Adoreg1 As ADODB.Recordset
 Dim AdoReg2 As ADODB.Recordset
 Dim mes As Long
 Dim sumaent, SumaSal As Double
 Dim contador As Integer
 Dim codigoaux As String
    cConexCom.Execute "delete * from rotacion"
   RSQL = "select smcodigo,smcanent,smcansal,smmespro from moresmes where " & _
   "smalma ='" & almacen & "'   and smcodigo  between '" & Text1 & "'  and '" & Text2 & "' and smmespro between '" & mesini & "' and '" & mesfin & "' order  by smcodigo,smmespro"
    Set Adoreg1 = New ADODB.Recordset
    Set AdoReg2 = New ADODB.Recordset
    Rsql1 = "select * from rotacion"
    Adoreg1.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
    AdoReg2.Open Rsql1, cConexCom, adOpenDynamic, adLockOptimistic
    If Adoreg1.EOF Then
      MsgBox "No se ha registrado  articulo para ese rango", vbInformation, "Aviso"
      Exit Sub
    End If
    AdoReg2.AddNew
    codigoaux = Adoreg1("smcodigo")
    AdoReg2("tcod") = Adoreg1("smcodigo")
    mes = Val(mesini)
    While Not Adoreg1.EOF
           contador = 1
            If Not Adoreg1.EOF Then
                If Adoreg1("smcodigo") <> codigoaux Then
                    contador = 1
                    mes = Val(mesini)
                    AdoReg2.AddNew
                    AdoReg2("tcod") = Adoreg1("smcodigo")
                    codigoaux = Adoreg1("smcodigo")
                End If
                If Val(Adoreg1("smmespro")) = mes Then
                       contador = IIf(Adoreg1("smmespro") = mesini, 1, Val(Adoreg1("smmespro")) - Val(mesini) + 1)
                       AdoReg2("te" & Format(contador, "00")) = IIf(IsNull(Adoreg1("smcanent")), 0, Adoreg1("smcanent"))
                       AdoReg2("ts" & Format(contador, "00")) = IIf(IsNull(Adoreg1("smcansal")), 0, Adoreg1("smcansal"))
                       Adoreg1.MoveNext
                End If
                
            End If
            contador = contador + 1
            mes = mes + 1
           'If Not Adoreg1.EOF Then Adoreg1.MoveNext
   Wend
    AdoReg2.UpdateBatch
   Set AdoReg2 = Nothing
   Set Adoreg1 = Nothing
End Sub


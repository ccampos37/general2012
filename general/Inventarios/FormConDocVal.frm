VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormConDocVal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentos Valorizados"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   Icon            =   "FormConDocVal.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCon 
      Caption         =   "&Dcmto."
      Height          =   675
      Left            =   2055
      MouseIcon       =   "FormConDocVal.frx":08CA
      Picture         =   "FormConDocVal.frx":1B3C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5040
      Width           =   775
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   195
      TabIndex        =   0
      Top             =   75
      Width           =   8175
      Begin VB.CommandButton CmdConsultar 
         Caption         =   "C&onsultar"
         Height          =   315
         Left            =   6285
         TabIndex        =   5
         Top             =   705
         Width           =   1275
      End
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FormConDocVal.frx":1F7E
         Left            =   4275
         List            =   "FormConDocVal.frx":1F88
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   990
         MaxLength       =   10
         TabIndex        =   1
         Top             =   285
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   315
         Left            =   4275
         TabIndex        =   4
         Top             =   675
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         _Version        =   393216
         Format          =   47120385
         CurrentDate     =   36704
         MinDate         =   35431
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   990
         TabIndex        =   3
         Top             =   675
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         _Version        =   393216
         Format          =   47120385
         CurrentDate     =   36704
         MinDate         =   35431
      End
      Begin VB.Label Label13 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3615
         TabIndex        =   12
         Top             =   675
         Width           =   600
      End
      Begin VB.Label Label12 
         Caption         =   "Desde"
         Height          =   255
         Left            =   195
         TabIndex        =   11
         Top             =   690
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Criterio"
         Height          =   300
         Left            =   3585
         TabIndex        =   10
         Top             =   285
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   210
         TabIndex        =   9
         Top             =   285
         Width           =   645
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   360
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Rango"
      Height          =   675
      Left            =   3750
      Picture         =   "FormConDocVal.frx":1FA4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5040
      Width           =   775
   End
   Begin VB.CommandButton Command7 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   5460
      Picture         =   "FormConDocVal.frx":23E6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   775
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3585
      Left            =   210
      TabIndex        =   6
      Top             =   1275
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6324
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FormConDocVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim criterio(1 To 3) As String
Dim tipo As String
Dim NumDoc As String
Dim rsql As String
Dim adodc1 As ADODB.Recordset

Private Sub CarObj()        ' Carga Objetos
  'FG.FormatString = "Tipo Doc.|Numero de Doc| Tr| Fecha | Proveedor|Cliente|Td REF|Num.Doc Ref."
  
  DataGrid1.Columns(0).Locked = True
  DataGrid1.Columns(0).WrapText = True
  DataGrid1.Columns(0).Caption = "Tipo Doc."
  DataGrid1.Columns(1).Caption = "Numero Doc"
  DataGrid1.Columns(2).Caption = "Tr"
  DataGrid1.Columns(3).Caption = "Fecha"
  DataGrid1.Columns(4).Caption = "Provedor"
  DataGrid1.Columns(5).Caption = "Cliente"
  DataGrid1.Columns(6).Caption = "Ref"
  DataGrid1.Columns(7).Caption = "Numero"
  DataGrid1.Columns(0).Width = 800
  DataGrid1.Columns(1).Width = 1500
  DataGrid1.Columns(2).Width = 800
  DataGrid1.Columns(3).Width = 1000
  DataGrid1.Columns(4).Width = 1000
  DataGrid1.Columns(5).Width = 1000
  DataGrid1.Columns(6).Width = 800
  DataGrid1.Columns(7).Width = 1500

 'DataGrid1.Columns(0).WrapText = False
End Sub

Private Sub CmbOrden_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then DTPicker2.SetFocus
End Sub

Private Sub CmdCon_Click()
 If adodc1.RecordCount > 0 Then
     Screen.MousePointer = 11
     NumDoc = adodc1("CANUMDOC")
     imprimir
     Screen.MousePointer = 1
 End If
End Sub

Private Sub CmdConsultar_Click()
If DTPicker2 > DTPicker3 Then
  MsgBox "La fecha inicial no puede ser mayor", vbOKOnly, "Aviso"
  DTPicker2.SetFocus
  Exit Sub
End If
adodc1.Close
rsql = "select  m.CATD, m.CANUMDOC, m.CACODMOV ,m.CAFECDOC, m.CACODPRO,m.CACODCLI, m.CARFTDOC ,m.CARFNDOC from MovAlmCab m where  m.CAALMA ='" & VGAlma & "' and m.CATD='" & tipo & "'  and   casitgui <>'A' and   m.cafecdoc  between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "'  ORDER BY m.CANUMDOC" '
                  '0         1           2             3          4           5          6         7
adodc1.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
CarObj1
End Sub

Private Sub CmdImprimir_Click()
 If adodc1.RecordCount > 0 Then
     Screen.MousePointer = 11
     NumDoc = adodc1("CANUMDOC")
    ' Imprimir
     Imprimir2
     Screen.MousePointer = 1
 End If
End Sub

Private Sub Command7_Click()
adodc1.Close
Unload Me
End Sub

Private Sub DataGrid1_DblClick()
cmdImprimir.SetFocus: CmdImprimir_Click
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdImprimir.SetFocus: CmdImprimir_Click
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then DTPicker3.SetFocus
End Sub

Private Sub DTPicker3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then DataGrid1.SetFocus
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
central Me                                 ' Centra el Formulario
DTPicker2 = DateAdd("m", -2, Date)
DTPicker3 = Date
tipo = "NI"
Set adodc1 = New ADODB.Recordset
DataGrid1.ClearFields                       ' Limpia las Columnas
rsql = "select  m.CATD, m.CANUMDOC, m.CACODMOV ,m.CAFECDOC, m.CACODPRO,m.CACODCLI, m.CARFTDOC ,m.CARFNDOC from MovAlmCab m where  m.CAALMA ='" & VGAlma & "' and m.CATD='" & tipo & "'  and   casitgui <>'A' and   m.cafecdoc  between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "'  ORDER BY m.CANUMDOC" '
                  '0         1           2             3          4           5          6         7
adodc1.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
CarObj1
CmbOrden.ListIndex = 0
End Sub

Private Sub CarObj1()        ' Carga Objetos
  'FG.FormatString = "Tipo Doc.|Numero de Doc| Tr| Fecha | Proveedor|Cliente|Td REF|Num.Doc Ref."
  Init_ControlDataGrid DataGrid1
  Set DataGrid1.DataSource = adodc1
  DataGrid1.Columns(0).Locked = True
  DataGrid1.Columns(0).Caption = "Tipo Doc."
  DataGrid1.Columns(0).Alignment = dbgCenter
  DataGrid1.Columns(1).Caption = "          Nro. Doc."
  DataGrid1.Columns(1).Alignment = dbgRight
  DataGrid1.Columns(2).Caption = "   Transa."
  DataGrid1.Columns(2).Alignment = dbgCenter
  DataGrid1.Columns(3).Caption = "     Fecha"
  DataGrid1.Columns(3).Alignment = dbgCenter
  DataGrid1.Columns(4).Caption = "   Proveedor"
  DataGrid1.Columns(4).Alignment = dbgRight
  DataGrid1.Columns(5).Caption = "   Cliente"
  DataGrid1.Columns(5).Alignment = dbgRight
  DataGrid1.Columns(6).Caption = " Doc. Ref."
  DataGrid1.Columns(6).Alignment = dbgCenter
  DataGrid1.Columns(7).Caption = "      Numero Ref."
  DataGrid1.Columns(7).Alignment = dbgRight
  DataGrid1.Columns(0).Width = 800
  DataGrid1.Columns(1).Width = 1500
  DataGrid1.Columns(2).Width = 800
  DataGrid1.Columns(3).Width = 1500
  DataGrid1.Columns(4).Width = 1500
  DataGrid1.Columns(5).Width = 1500
  DataGrid1.Columns(6).Width = 800
  DataGrid1.Columns(7).Width = 1500
 
End Sub

Private Sub Text1_Change()
 Dim C As String
 Dim Ant As Integer
 If adodc1.RecordCount <> 0 Then
  Ant = adodc1.Bookmark
  adodc1.AbsolutePosition = 1
  Text1 = Trim(Text1)
   If Text1 <> "" Then
    Select Case CmbOrden.ListIndex
        Case 0
           C = "CANUMDOC LIKE '" & Trim(UCase(Text1)) & "*'"
        Case 1
            C = "CACODPRO LIKE '" & Trim(UCase(Text1)) & "*' "
    End Select
       adodc1.Find C
       
   If adodc1.EOF Then
    adodc1.AbsolutePosition = Ant
   End If
  End If
 End If
End Sub

Private Sub imprimir()
  Dim CADENA As String
  Dim Codigo2 As String
    Codigo2 = "NOTA DE INGRESO"
    CrystalReport1.WindowTitle = "Inv020  -- Sistema de Inventarios "
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv020.rpt"       'notaingsal
    CrystalReport1.WindowState = crptMaximized
    Ubi_Tab CrystalReport1
    CADENA = "{MOVALMCAB.CAALMA} = '" & VGAlma & "'  and {MOVALMCAB.CATD} = '" & tipo & "' and {MOVALMCAB.CANUMDOC} = '" & NumDoc & "'"
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = CADENA
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.formulas(0) = "empresa ='" & VGparametros.RucEmpresa & "'"
    CrystalReport1.formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.formulas(2) = "almacen ='" & VGNomAlm & "'"
    CrystalReport1.formulas(3) = "nota ='" & Codigo2 & "'"
    CrystalReport1.formulas(4) = ""
    CrystalReport1.formulas(5) = ""
    If CrystalReport1.Status <> 2 Then
       CrystalReport1.Action = 1
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Text1 <> "" Then
    Text1 = Trim(Text1)
    Text1 = Format(Text1, String(10, "0"))
  Else
     CmbOrden.SetFocus
  End If
End If

End Sub

Private Sub Imprimir2()         'Prueba 05/10/2000

Dim Codigo1 As String
Dim Codigo2 As String
Dim CADENA As String
Dim cTip As String, tipo As String
Codigo1 = Trim(Text1)
Screen.MousePointer = 11
CrystalReport1.WindowTitle = "Inv098 -- Control de Inventarios"
CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv098.rpt"

CADENA = "({MOVALMCAB.CAFECDOC} IN DATE (" & Format(DTPicker2, "yyyy") & "," & Format(DTPicker2, "mm") & "," & Format(DTPicker2, "dd") & ") "
CADENA = CADENA & "to DATE (" & Format(DTPicker3, "yyyy") & "," & Format(DTPicker3, "mm") & "," & Format(DTPicker3, "dd") & ")) "
CADENA = CADENA & " and {MOVALMCAB.CATD}='NI'  AND {MOVALMCAB.CAALMA}='" & VGAlma & "'"

    CADENA = CADENA & " And {MOVALMCAB.CASITGUI} <>'A' "
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
    CrystalReport1.formulas(2) = "Tipo ='" & tipo & "'"
    CrystalReport1.formulas(3) = "Almacen ='" & VGNomAlm & "'"
    CrystalReport1.formulas(4) = "FecIni ='" & DTPicker2 & "'"
    CrystalReport1.formulas(5) = "FecFin='" & DTPicker3 & "'"
   ' CrystalReport1.WindowTitle = "Reporte de Guías de Remisión " & Tipo

    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
   

End Sub

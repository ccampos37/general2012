VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmArTipoCambio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Cambio"
   ClientHeight    =   4365
   ClientLeft      =   1350
   ClientTop       =   2580
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6465
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   105
      Top             =   2985
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   180
      TabIndex        =   10
      Top             =   60
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmTipoCambio.frx":0000
         Left            =   1080
         List            =   "FrmTipoCambio.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3990
         MaxLength       =   4
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Año   :"
         Height          =   255
         Left            =   3390
         TabIndex        =   12
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Mes   :"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   6255
      Begin VB.CommandButton CmdIng 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   120
         Picture         =   "FrmTipoCambio.frx":0090
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdModi 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   1170
         Picture         =   "FrmTipoCambio.frx":04D2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdEli 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2220
         Picture         =   "FrmTipoCambio.frx":0914
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdRep 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   3255
         Picture         =   "FrmTipoCambio.frx":0D56
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   4320
         Picture         =   "FrmTipoCambio.frx":1198
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5340
         Picture         =   "FrmTipoCambio.frx":15DA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2235
      Left            =   180
      TabIndex        =   13
      Top             =   840
      Width           =   6015
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1830
         Left            =   135
         TabIndex        =   19
         Top             =   210
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   3228
         _Version        =   393216
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
   Begin VB.Frame Frame4 
      Height          =   2175
      Left            =   180
      TabIndex        =   15
      Top             =   900
      Width           =   6015
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   2820
         TabIndex        =   20
         Top             =   375
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   503
         _Version        =   393216
         CalendarForeColor=   12582912
         CalendarTitleBackColor=   -2147483638
         Format          =   47382529
         CurrentDate     =   36714
      End
      Begin VB.TextBox TxVenta 
         Height          =   285
         Left            =   2820
         TabIndex        =   7
         Text            =   "Text4"
         Top             =   1380
         Width           =   1335
      End
      Begin VB.TextBox TxCompra 
         Height          =   285
         Left            =   2820
         TabIndex        =   6
         Text            =   "Text3"
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Venta    :"
         Height          =   255
         Left            =   1380
         TabIndex        =   18
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo Compra   :"
         Height          =   315
         Left            =   1380
         TabIndex        =   17
         Top             =   900
         Width           =   1635
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha             : "
         Height          =   315
         Left            =   1380
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmArTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As ADODB.Recordset
Dim cOpen  As Boolean
Dim nOperador As Integer
Dim nTra As Integer
Dim nTra2 As Integer
Dim cEstado As Boolean
Dim cBase As String

Private Sub CmdEli_Click()
Dim nPosi As Integer
Dim cSql1 As String

On Error GoTo EliErr
If adodc1.RecordCount > 0 Then
    cSql1 = "Delete from Tipo_Cambio Where TIPOCAMB_FECHA= '" & adodc1("TIPOCAMB_FECHA") & "'"

    If MsgBox("Seguro de Eliminar", vbQuestion + vbOKCancel, "Inventarios") = vbOK Then
            nPosi = Pos_Dato(adodc1)
            cBase = cRuta4
            If UCase(Dir$(cBase)) = UCase(VGNameCont & ".MDB") Then
                
                nTra = 1
                VGcnxCT.BeginTrans
                VGcnxCT.Execute cSql1
                VGcnxCT.CommitTrans
                nTra = 0: adodc1.Requery
            Else
                nTra2 = 1
                VGCNx.BeginTrans
                VGCNx.Execute cSql1
                VGCNx.CommitTrans
                nTra2 = 0: adodc1.Requery
            End If
            If nPosi <> 0 Then adodc1.AbsolutePosition = nPosi
    End If
    If DataGrid1.Visible Then DataGrid1.SetFocus
    Set_Data
Else
    MsgBox "No existe ningún registro para Eilminar", vbInformation, "Inventarios"
End If
Exit Sub
EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGcnxCT.RollbackTrans
    If nTra2 = 1 Then VGCNx.RollbackTrans
    nTra = 0: nTra2 = 0
End Sub

Private Sub CmdGrabar_Click()
Dim CSQL2 As String
Dim cValor As String
Dim nMes As String
Dim nAno As String

On Error GoTo GrabErr

If nOperador = 1 Then                  ' Si es Ingreso
'        cValor = ValidFecha2(Format(DTPicker1, "DD/MM/YYYY"))
'        If cValor = "" Then
'           MsgBox "Ingrese la Fecha Correctamente", vbExclamation + vbOKOnly, "Advertencia"
'           DTPicker1.SetFocus
'           Exit Sub
'        End If
        
    nMes = Combo1.ListIndex + 1
    If Month(CDate(DTPicker1)) <> nMes Then
       MsgBox "El Mes ingresado en la Fecha no coincide con lo señalado previamente", vbExclamation + vbOKOnly, "Advertencia"
       DTPicker1.SetFocus
       Exit Sub
    End If
    nAno = Val(Text2)
    If Year(CDate(DTPicker1)) <> nAno Then
       MsgBox "El Año ingresado en la Fecha no coincide con lo señalado previamente", vbExclamation + vbOKOnly, "Advertencia"
       DTPicker1.SetFocus
       Exit Sub
    End If
    
    cBase = cRuta4
    If UCase(Dir$(cBase)) = UCase(VGNameCont & ".MDB") Then
        If Existe(3, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True) Then
            MsgBox "El Tipo de Cambio ya existe", vbInformation, "Inventarios"
            DTPicker1.SetFocus: Exit Sub
        End If
    Else
        If Existe(1, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True) Then
            MsgBox "El Tipo de Cambio ya existe", vbInformation, "Inventarios"
            DTPicker1.SetFocus: Exit Sub
        End If
    End If
    
    CSQL2 = "Insert Into Tipo_Cambio (TIPOMON_CODIGO,TIPOCAMB_FECHA,TIPOCAMB_COMPRA,TIPOCAMB_VENTA)"
    CSQL2 = CSQL2 & "  Values ('ME','" & Format(DTPicker1, "dd/mm/yyyy") & "'," & TxCompra & "," & TxVenta & ")"
ElseIf nOperador = 2 Then               'Si es Modificación
    CSQL2 = "Update Tipo_Cambio set  TIPOCAMB_COMPRA =" & TxCompra & ",TIPOCAMB_VENTA = " & TxVenta & ""
    CSQL2 = CSQL2 & "  Where TIPOCAMB_FECHA = '" & Format(DTPicker1, "dd/mm/yyyy") & "' and TIPOMON_CODIGO = 'ME'"
End If

'Si  existe Contabilidad
cBase = cRuta4
If UCase(Dir$(cBase)) = UCase(cNomBd4) Then
    nTra = 1
    VGcnxCT.BeginTrans
    VGcnxCT.Execute CSQL2
    VGcnxCT.CommitTrans
    nTra = 0
Else
    nTra2 = 1
    VGCNx.BeginTrans
    VGCNx.Execute CSQL2
    VGCNx.CommitTrans
    nTra2 = 0
End If

    adodc1.Requery
    adodc1.Find "TIPOCAMB_FECHA = '" & Format(DTPicker1, "dd/mm/yyyy") & "'"

If nOperador = 1 Then
    Limpiar
    OculObj (True)
    DTPicker1.SetFocus
ElseIf nOperador = 2 Then
    OculObj (False)
    nOperador = 0
    Frame1.Enabled = True
    Set_Data
    DataGrid1.SetFocus
End If

Exit Sub
GrabErr:
    MsgBox Err.Description
    'Si no existe Contabilidad
    If nTra = 1 Then VGcnxCT.RollbackTrans
    If nTra2 = 1 Then VGCNx.RollbackTrans
    nTra = 0: nTra2 = 0
End Sub

Private Sub CmdIng_Click()
If Val(Text2) < 1900 Or Val(Text2) > 2100 Then
      MsgBox "El Año asignado es incorrecto", vbInformation, "Inventarios"
      Text2.SetFocus
      Exit Sub
End If
OculObj (True): nOperador = 1
Limpiar
Frame4.Caption = "Ingreso de Tipo de Cambio"
DTPicker1.Enabled = True: DTPicker1.SetFocus
Frame1.Enabled = False
DTPicker1.SetFocus
End Sub

Private Sub CmdModi_Click()
If adodc1.RecordCount > 0 Then
    If Val(Text2) < 1900 Or Val(Text2) > 2100 Then
      MsgBox "El Año asignado es incorrecto", vbInformation, "Inventarios"
      Text2.SetFocus
      Exit Sub
    End If
    nOperador = 2: Limpiar
    Frame4.Caption = "Modificación de Tipo de Cambio"
    Edit_Ven
    OculObj (True)
    DTPicker1.Enabled = False
Else
    MsgBox "No existe ningún registro para modificar", vbInformation, "Inventarios"
End If
Frame1.Enabled = False
TxCompra.SetFocus
End Sub

Private Sub CmdRep_Click()
Dim nMes As String
Dim CTIME As String
If Text2 <> "" Then
    If Val(Text2) > 1900 Or Val(Text2) < 2100 Then
        nMes = Combo1.ListIndex + 1
        CTIME = Format(Time, "hh:mm:ss")
        CrystalReport1.WindowTitle = "Sistema  de Inventarios"
        CrystalReport1.formulas(0) = "Hora = '" & CTIME & "'"
        CrystalReport1.formulas(1) = "Empresa = '" & Mid(VGparametros.RucEmpresa, 1, 20) & "'"
        'Para el reporte establecer la nueva ubicacion de la Base de Datos si existiera Contabilidad (ahora BDComun)
        cBase = cRuta4
        If UCase(Dir$(cBase)) = UCase(cNomBd4) Then
            CrystalReport1.WindowTitle = "Inv041 -- Control de Inventarios"
            CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv041.Rpt"
        Else
            CrystalReport1.WindowTitle = "Inv042 -- Control de Inventarios"
            CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv042.Rpt"
        End If
        Call Ubi_Tab(CrystalReport1)
        CrystalReport1.SelectionFormula = " Month ({TIPO_CAMBIO.TIPOCAMB_FECHA})=" & Format(Trim(nMes), "00") & "   and  Year ({TIPO_CAMBIO.TIPOCAMB_FECHA}) =" & Text2 & ""
        CrystalReport1.WindowTop = 100
        CrystalReport1.WindowLeft = 150
        CrystalReport1.DiscardSavedData = True
        CrystalReport1.WindowShowPrintBtn = True
        CrystalReport1.WindowShowRefreshBtn = True
        CrystalReport1.WindowShowSearchBtn = True
        CrystalReport1.WindowShowPrintSetupBtn = True
        CrystalReport1.Destination = crptToWindow
         If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
     End If
 End If
End Sub

Private Sub CmdSalir_Click()
If nOperador = 1 Or nOperador = 2 Then
        OculObj (False): nOperador = 0
        Frame1.Enabled = True
        DataGrid1.SetFocus
Else
    Unload Me
End If
End Sub

Private Sub Combo1_Click()
If Len(Text2) = 4 Then
   If Val(Text2) < 1900 Or Val(Text2) > 2100 Then
      MsgBox "El Año asignado es incorrecto", vbInformation, "Inventarios"
      Text2.SetFocus
      Exit Sub
   End If
   Frame3.Enabled = True
   CarObj
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Form_Activate()
Combo1.SetFocus
End Sub

Private Sub Form_Load()
Dim cMes As String
central Me
Init_ControlDataGrid DataGrid1
cEstado = False
cMes = Mid(Format(Date, "dd/mm/yyyy"), 4, 2)
Combo1.ListIndex = Val(cMes) - 1
Text2 = Year(Date)
cEstado = True
cOpen = False
CarObj
DTPicker1 = Date
cOpen = True
End Sub

Private Sub Limpiar()
TxCompra = 0: TxVenta = 0
End Sub

Private Sub Text2_GotFocus()
Enfoque Text2
End Sub

Private Sub Text2_Change()
If cEstado = True Then
   Frame3.Enabled = True
   CarObj
  'End If
End If
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
  cEstado = False
   Text2 = ""
   CarObj
  cEstado = True
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    CmdIng.SetFocus
Else
  If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub OculObj(bTip As Boolean)
'Si es Ingreso o modificación en True
Frame4.Visible = bTip
CmdIng.Enabled = Not bTip
CmdModi.Enabled = Not bTip
CmdEli.Enabled = Not bTip
CmdRep.Enabled = Not bTip
CmdIng.Enabled = Not bTip
CmdGrabar.Enabled = bTip
Frame2.Visible = Not bTip
End Sub

Private Sub TxCompra_GotFocus()
Enfoque TxCompra
End Sub

Private Sub TxCompra_KeyPress(KeyAscii As Integer)
Dim I As Integer

If KeyAscii = 13 Then
   SendKeys "{tab}"
Else
  If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And Chr$(KeyAscii) <> "." And KeyAscii <> 8 Then
     KeyAscii = 0
  Else
     If Chr$(KeyAscii) = "." Then
        For I = 1 To Len(TxCompra)
            If Mid(TxCompra, I, 1) = "." Then KeyAscii = 0: Exit Sub
        Next
        
     End If
  End If
End If
End Sub

Private Sub TxVenta_GotFocus()
Enfoque TxVenta
End Sub

Private Sub TxVenta_KeyPress(KeyAscii As Integer)
Dim I As Integer

If KeyAscii = 13 Then
   SendKeys "{tab}"
Else
  If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And Chr$(KeyAscii) <> "." And KeyAscii <> 8 Then
     KeyAscii = 0
  Else
     If Chr$(KeyAscii) = "." Then
        For I = 1 To Len(TxVenta)
            If Mid(TxVenta, I, 1) = "." Then KeyAscii = 0: Exit Sub
        Next
        
     End If
  End If
End If
End Sub

Private Sub Edit_Ven()
Dim cSel1 As ADODB.Recordset
Dim cSql1 As String
cSql1 = "Select TIPOCAMB_FECHA,TIPOCAMB_COMPRA,TIPOCAMB_VENTA from Tipo_Cambio where TIPOCAMB_FECHA = '" & adodc1(0) & "'"
Limpiar
    
Set cSel1 = New ADODB.Recordset
'Si existe Contabilidad
cBase = cRuta4
If UCase(Dir$(cBase)) = UCase(cNomBd4) Then
    
    cSel1.Open cSql1, VGcnxCT, adOpenStatic
Else
    cSel1.Open cSql1, VGCNx, adOpenStatic
End If

If cSel1.RecordCount > 0 Then
    DTPicker1 = adodc1(0)
    TxCompra = adodc1(1)
    TxVenta = adodc1(2)
Else
    MsgBox "El registro ha sido Eliminado", vbInformation, "Inventarios"
End If
cSel1.Close
End Sub

Private Sub CarObj()
Dim nMes As String
Dim cAno As String

If cOpen = True Then adodc1.Close

Set adodc1 = New ADODB.Recordset
nMes = Combo1.ListIndex + 1
cEstado = False

If Text2 <> "" Then
    cAno = Text2
Else
    cAno = 0
End If

cEstado = True
cBase = cRuta4
If UCase(Dir$(cBase)) = UCase(cNomBd4) Then
    adodc1.Open "Select TIPOCAMB_FECHA,TIPOCAMB_COMPRA,TIPOCAMB_VENTA From TIPO_CAMBIO WHERE MONTH(TIPOCAMB_FECHA) = " & Format(Trim(nMes), "00") & "  AND YEAR(TIPOCAMB_FECHA) = " & Format(Trim(cAno), "0000") & " order by TIPOCAMB_FECHA", VGcnxCT, adOpenStatic
Else
    adodc1.Open "Select TIPOCAMB_FECHA,TIPOCAMB_COMPRA,TIPOCAMB_VENTA From TIPO_CAMBIO WHERE MONTH(TIPOCAMB_FECHA) = " & Format(Trim(nMes), "00") & "  AND YEAR(TIPOCAMB_FECHA) = " & Format(Trim(cAno), "0000") & " order by TIPOCAMB_FECHA", VGCNx, adOpenStatic
End If
Set_Data
End Sub

Private Sub Set_Data()
Set DataGrid1.DataSource = adodc1
DataGrid1.Columns(0).Caption = "Fecha"
DataGrid1.Columns(0).Alignment = dbgCenter
DataGrid1.Columns(0).Width = 1500
DataGrid1.Columns(1).NumberFormat = "##.#00 "
DataGrid1.Columns(1).Alignment = dbgRight
DataGrid1.Columns(1).Caption = "            Compra"
DataGrid1.Columns(1).Width = 1800
DataGrid1.Columns(2).NumberFormat = "##.#00 "
DataGrid1.Columns(2).Caption = "            Venta"
DataGrid1.Columns(2).Alignment = dbgRight
DataGrid1.Columns(2).Width = 1800
End Sub

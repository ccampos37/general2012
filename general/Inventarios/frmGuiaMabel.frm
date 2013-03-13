VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGuiaMabel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guia de Ingreso de IASA a Mabel"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "frmGuiaMabel.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   5370
      Begin VB.CommandButton Command1 
         Caption         =   "&Consultar"
         Height          =   375
         Left            =   3585
         TabIndex        =   10
         Top             =   525
         Width           =   1455
      End
      Begin VB.TextBox TxtProveedor 
         Height          =   285
         Left            =   3660
         MaxLength       =   8
         TabIndex        =   4
         Top             =   300
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24838145
         CurrentDate     =   36699
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24838145
         CurrentDate     =   36699
      End
      Begin VB.Label Label1 
         Caption         =   "Desde la Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   4020
         TabIndex        =   8
         Top             =   315
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta la Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   1785
      Picture         =   "frmGuiaMabel.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4065
      Width           =   735
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3315
      Picture         =   "frmGuiaMabel.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4065
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   1155
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4895
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
      ColumnCount     =   3
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
      BeginProperty Column02 
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
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmGuiaMabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Adoreg1 As ADODB.Recordset
Dim RSQL As String
Dim IASA As String
Dim NumDoc As String
Private Sub CmdAceptar_Click()
 Screen.MousePointer = 11

 If Adoreg1.RecordCount > 0 Then
      NumDoc = Adoreg1("CANUMDOC")
      imprimir
 End If

 Screen.MousePointer = 1
End Sub

Private Sub CmdCancelar_Click()
 Unload Me
End Sub

Private Sub Command1_Click()
    RSQL = "select  CATD, CANUMDOC, CAFECDOC from MovAlmCab " & _
           "where  CAALMA ='" & VGAlma & "' and CATD='NI' and CACODPRO='" & VGIASA & _
           "'and CASITGUI <> 'A'  AND  cafecdoc  between " & DateSQL(DTPicker1.Value) & " and " & DateSQL(DTPicker2.Value) & _
           " ORDER BY CANUMDOC" '
    Set Adoreg1 = New ADODB.Recordset
    Adoreg1.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
    If Adoreg1.RecordCount = 0 Then
         MsgBox "No hay documentos de IASA", vbExclamation, "Aviso"
         Set DataGrid1.DataSource = Adoreg1
          CarObj
         CmdAceptar.Enabled = False
         Exit Sub
    End If
    Set DataGrid1.DataSource = Adoreg1
    DataGrid1.Refresh
    CarObj
    CmdAceptar.Enabled = True
End Sub

Private Sub Form_Load()
  central Me
  DTPicker1 = DateAdd("m", -1, Date)
  DTPicker2 = Date
  IASA = "1235"
  CarObj
  CmdAceptar.Enabled = False
End Sub

Private Sub imprimir()
Dim cadena As String
            CrystalReport1.WindowTitle = "Inv044 -- Control de Inventarios"
            CrystalReport1.ReportFileName = cRutP & "inv044.rpt"       'notamabel
            CrystalReport1.WindowState = crptMaximized
            Ubi_Tab CrystalReport1
            cadena = "{MOVALMCAB.CAALMA} = '" & VGAlma & "'  and {MOVALMCAB.CATD} = 'NI' and {MOVALMCAB.CANUMDOC} = '" & NumDoc & "' "
            CrystalReport1.DiscardSavedData = True
            CrystalReport1.Destination = crptToWindow
            CrystalReport1.SelectionFormula = cadena
            CrystalReport1.WindowShowPrintBtn = True
            CrystalReport1.WindowShowRefreshBtn = True
            CrystalReport1.WindowShowSearchBtn = True
            CrystalReport1.WindowShowPrintSetupBtn = True
            CrystalReport1.Formulas(0) = "empresa ='" & VGNemp & "'"
            CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
            If CrystalReport1.Status <> 2 Then
               CrystalReport1.Action = 1
            End If
End Sub


Private Sub CarObj()        ' Carga Objetos

 
 DataGrid1.Columns(0).Locked = True
 DataGrid1.Columns(0).WrapText = True
 DataGrid1.Columns(0).Caption = "   TIPO"
 DataGrid1.Columns(1).Caption = "   DOCUMENTO"
 DataGrid1.Columns(2).Caption = "   FECHA"
 DataGrid1.Columns(0).Width = 1200
 DataGrid1.Columns(1).Width = 2750
 DataGrid1.Columns(2).Width = 1000

End Sub

Private Sub TxtProveedor_DblClick()
  VGForm1 = 21
  TxTProveedor = ""
  FormAyuProv.Show 1
  If Trim(TxTProveedor) <> "" Then
      Command1.SetFocus
  End If
End Sub

Private Sub TxtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
    TxtProveedor_DblClick
   End If
   End Sub

Private Sub TxtProveedor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And TxTProveedor <> "" Then
           TxTProveedor = Trim(TxTProveedor)
           If prove(TxTProveedor) <> "" Then
              Command1.SetFocus
           End If
   Else
           TxtArticulo = ""
   End If
End Sub

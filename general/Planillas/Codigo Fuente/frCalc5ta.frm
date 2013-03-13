VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#8.0#0"; "ApliCTxt.ocx"
Begin VB.Form frCalc5ta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de Renta de 5ta. Categoria"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   Icon            =   "frCalc5ta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "Por Areas de Trabajo"
      Height          =   210
      Left            =   195
      TabIndex        =   5
      Top             =   720
      Value           =   -1  'True
      Width           =   1830
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Por Centros de Costo"
      Height          =   210
      Left            =   195
      TabIndex        =   4
      Top             =   1020
      Width           =   1830
   End
   Begin VB.CommandButton cmSelecTrab 
      Caption         =   "Seleccion (F5)"
      Height          =   990
      Left            =   240
      Picture         =   "frCalc5ta.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2325
      Width           =   870
   End
   Begin VB.CommandButton cmContinuar 
      Caption         =   "Continuar >>"
      Height          =   375
      Left            =   6150
      TabIndex        =   1
      Top             =   285
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6150
      TabIndex        =   0
      Top             =   735
      Width           =   1275
   End
   Begin MSDataGridLib.DataGrid DGLista 
      Height          =   2325
      Left            =   1200
      TabIndex        =   2
      Top             =   2325
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   4101
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
      Caption         =   "Trabajadores Seleccionados"
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
            LCID            =   10250
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
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin AplisetControlText.Aplitext xMes 
      Height          =   285
      Left            =   225
      TabIndex        =   6
      Top             =   285
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker xFechaFin 
      Height          =   285
      Left            =   1110
      TabIndex        =   7
      Top             =   1665
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   24444929
      CurrentDate     =   36699
   End
   Begin MSComCtl2.DTPicker xFechaIni 
      Height          =   285
      Left            =   1110
      TabIndex        =   8
      Top             =   1335
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   24444929
      CurrentDate     =   36699
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   1965
      Left            =   2535
      TabIndex        =   9
      Top             =   270
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   3466
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Periodos en Cronogramas"
         Object.Width           =   5733
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "FechaIni"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "FechaFin"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mes de Trabajo"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   45
      Width           =   1110
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicial"
      Height          =   195
      Left            =   150
      TabIndex        =   11
      Top             =   1380
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label l2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Final"
      Height          =   195
      Left            =   150
      TabIndex        =   10
      Top             =   1725
      Visible         =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "frCalc5ta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTrab As New ADODB.Recordset
Dim xItem As ListItem

Private Sub cmSelecTrab_Click()
    frSelect.Show 1
    If Not xFechaIni.Visible Then
        MsgBox "Deberá seleccionar un periodo de pago", vbCritical
        Exit Sub
    End If
    RsTrab.Requery
    Set DGLista.DataSource = RsTrab
End Sub

Private Sub Form_Load()
    RsTrab.Open "_tmpSelect", DBAuxCom, adOpenStatic
    Set DGLista.DataSource = RsTrab
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    xFechaIni.Visible = True
    xFechaFin.Visible = True
    l1.Visible = True
    l2.Visible = True
    xFechaIni.Value = CDate(Item.SubItems(1))
    xFechaFin.Value = CDate(Item.SubItems(2))
End Sub

Private Sub xMes_DblClick()
    Lista.ListItems.Clear
    Dim RsMeses As New ADODB.Recordset
    RsMeses.Open "SELECT MesActivo, Nombre FROM MesesAct ORDER BY MesActivo", DbSystem, adOpenStatic
    If RsMeses.RecordCount = 0 Then
        MsgBox "No se han encontrado meses en actividad", vbCritical
        Set RsMeses = Nothing
        Exit Sub
    End If
    frmComun.Conectar RsMeses
    frmComun.Show 1
    If vgUtil(1) <> "" Then
        xMes.Text = RsMeses!Nombre
        xMes.Tag = RsMeses!mesactivo
    Else
        Set RsMeses = Nothing
        Exit Sub
    End If
    Set RsMeses = Nothing
    'Reciclaje de RsMeses
    RsMeses.Open "SELECT Codigo, Nombre, FechaIni, FechaFin FROM NomBol WHERE Cerrado<>1 AND Mes=" & DateSQL(CDate(xMes.Tag)) & " ORDER BY FechaIni", DbSystem, adOpenStatic
    Do While Not RsMeses.EOF
        Set xItem = Lista.ListItems.Add(, , RsMeses!Nombre, , 1)
        xItem.SubItems(1) = RsMeses!FechaIni
        xItem.SubItems(2) = RsMeses!FechaFin
        xItem.Tag = RsMeses!Codigo
        RsMeses.MoveNext
    Loop
    l1.Visible = False
    l2.Visible = False
    xFechaIni.Visible = False
    xFechaFin.Visible = False
    Set RsMeses = Nothing
End Sub

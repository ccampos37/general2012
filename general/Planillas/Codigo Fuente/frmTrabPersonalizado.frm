VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmTrabPersonalizado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utilitario"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   Icon            =   "frmTrabPersonalizado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   9855
      TabIndex        =   38
      Top             =   420
      Width           =   1410
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Quitar"
         Height          =   330
         Left            =   105
         TabIndex        =   40
         Top             =   180
         Width           =   1185
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   41
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label lblcampo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   150
         TabIndex        =   44
         Top             =   705
         Width           =   1185
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1155
      Left            =   4410
      TabIndex        =   23
      Top             =   390
      Width           =   4890
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Promedio  Valor"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   2730
         TabIndex        =   43
         Top             =   885
         Width           =   1425
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Promedio"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   2715
         MaskColor       =   &H00404040&
         TabIndex        =   42
         Top             =   570
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   315
         Left            =   3465
         TabIndex        =   34
         Top             =   120
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   1095
         TabIndex        =   29
         Top             =   780
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   37121
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1095
         TabIndex        =   28
         Top             =   510
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   37121
         MaxDate         =   2958435
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1080
         TabIndex        =   24
         Top             =   150
         Width           =   2310
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   315
         Index           =   7
         Left            =   3555
         TabIndex        =   35
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final"
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   6
         Left            =   150
         TabIndex        =   32
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   4
         Left            =   150
         TabIndex        =   30
         Top             =   555
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Promedio"
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   2
         Left            =   135
         TabIndex        =   25
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Columna "
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   3
         Left            =   165
         TabIndex        =   26
         Top             =   285
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   5
         Left            =   180
         TabIndex        =   31
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final"
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   7
         Left            =   180
         TabIndex        =   33
         Top             =   885
         Width           =   990
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   4185
      Left            =   0
      ScaleHeight     =   4185
      ScaleWidth      =   11505
      TabIndex        =   11
      Top             =   1770
      Width           =   11505
      Begin MSDataGridLib.DataGrid DgDet 
         Height          =   4005
         Left            =   75
         TabIndex        =   12
         Top             =   90
         Width           =   11340
         _ExtentX        =   20003
         _ExtentY        =   7064
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
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
               Type            =   1
               Format          =   "0"
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
            MarqueeStyle    =   4
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir Text"
      Height          =   360
      Left            =   9630
      TabIndex        =   10
      Top             =   6495
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Configuración de la Impresión"
      Height          =   1305
      Left            =   60
      TabIndex        =   5
      Top             =   6330
      Width           =   9270
      Begin AplisetControlText.Aplitext xTitulo 
         Height          =   300
         Index           =   0
         Left            =   1515
         TabIndex        =   6
         Top             =   300
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   529
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xTitulo 
         Height          =   300
         Index           =   1
         Left            =   1515
         TabIndex        =   7
         Top             =   765
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   529
         Text            =   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encabezado 01"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   9
         Top             =   825
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo del Informe"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdImpPla 
      Caption         =   "&Imprimir Reporte"
      Height          =   360
      Left            =   9645
      TabIndex        =   4
      Top             =   7110
      Width           =   1845
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   90
      TabIndex        =   0
      Top             =   420
      Width           =   3750
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         Height          =   315
         Left            =   2055
         TabIndex        =   36
         Top             =   615
         Width           =   1200
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   885
         TabIndex        =   16
         Top             =   150
         Width           =   2400
      End
      Begin VB.TextBox xCod 
         Height          =   285
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   315
         Index           =   0
         Left            =   2145
         TabIndex        =   37
         Top             =   675
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Columna "
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Código "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   870
         TabIndex        =   2
         Top             =   255
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Columna "
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   1
         Left            =   165
         TabIndex        =   22
         Top             =   285
         Width           =   675
      End
   End
   Begin MSComctlLib.ListView LvColumn 
      Height          =   2280
      Left            =   1935
      TabIndex        =   13
      Top             =   3120
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   4022
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descrip"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1065
      Index           =   2
      Left            =   165
      TabIndex        =   17
      Top             =   510
      Width           =   3720
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1290
      Index           =   3
      Left            =   255
      TabIndex        =   18
      Top             =   6435
      Width           =   9135
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   330
      Index           =   4
      Left            =   9705
      TabIndex        =   19
      Top             =   6570
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   330
      Index           =   5
      Left            =   9765
      TabIndex        =   20
      Top             =   7185
      Width           =   1770
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4095
      Index           =   6
      Left            =   225
      TabIndex        =   21
      Top             =   1920
      Width           =   11355
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1170
      Index           =   9
      Left            =   4485
      TabIndex        =   27
      Top             =   465
      Width           =   4860
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1140
      Index           =   8
      Left            =   9960
      TabIndex        =   39
      Top             =   495
      Width           =   1365
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   9105
      Top             =   5550
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1335
      Top             =   5775
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Trabajadores"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   8715
      TabIndex        =   46
      Top             =   6060
      Width           =   1395
   End
   Begin VB.Label lblTotalTrab 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   10215
      TabIndex        =   45
      Top             =   6060
      Width           =   1380
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reporte de Trabajadores"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   -105
      TabIndex        =   14
      Top             =   75
      Width           =   3675
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   11250
      Picture         =   "frmTrabPersonalizado.frx":08CA
      Top             =   30
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reporte de Trabajadores"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   -60
      TabIndex        =   15
      Top             =   60
      Width           =   3540
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      DrawMode        =   1  'Blackness
      FillColor       =   &H00808080&
      Height          =   210
      Left            =   11355
      Shape           =   2  'Oval
      Top             =   360
      Width           =   450
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   300
      Left            =   11685
      Shape           =   2  'Oval
      Top             =   225
      Width           =   120
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Trabajadores"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   8700
      TabIndex        =   47
      Top             =   6075
      Width           =   1395
   End
End
Attribute VB_Name = "frmTrabPersonalizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstrabajadores As ADODB.Recordset
Dim rsTRABTEMP As ADODB.Recordset
Dim Head_GridDet As String
Public Property Let GET_TABLE(SourceTable As String)

If ExisteTablaAux("[##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]"
DBSYSTEM.Execute "SELECT CODTRAB,NOMBRES INTO [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]  FROM " & SourceTable

End Property

Private Sub CMDAGREGAR_CLICK()
On Error GoTo handler
Dim CAD As String
If Combo1.Text = "" Then Exit Sub
CAD = "select " & Combo1.Text & " from [##_TMPTRABPERSONALIZADOTOTAL" & VGL_COMPUTER & "] "

'/////////////////////////LIMA 15/08/2001/////////////////////////////////////////////////////
If DevuelveValor(CAD, DBSYSTEM) <> "" Then
    '*************************************************************CREADO POR BASILIO
    CAD = ""
    For K = 0 To DgDet.Columns.Count - 1
            CAD = CAD & "TMP." & DgDet.Columns(K).Caption & ","
    Next K
    CAD = CAD & "TMP2." & Combo1.Text
    If DevuelveValor("select " & Combo1.Text & " from [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]", DBSYSTEM) <> "" Then
        MsgBox "Este campo ya existe en la tabla"
        Exit Sub
    End If
    'ALMACENAR EL TMP GRID EN UN TMPGLOBAL
    If ExisteTablaAux("[##_TMPGRIDDET" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##_TMPGRIDDET" & VGL_COMPUTER & "]"
    DBSYSTEM.Execute "SELECT * INTO [##_TMPGRIDDET" & VGL_COMPUTER & "] FROM [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]"
    'EMININAR EL TMP
    If ExisteTablaAux("[##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]"
    CAD = "Select " & CAD & "  into  [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "] FROM [##_TMPGRIDDET" & VGL_COMPUTER & "] TMP,[##_TMPTRABPERSONALIZADOTOTAL" & VGL_COMPUTER & "] TMP2 WHERE TMP.CODTRAB=TMP2.CODTRAB"
    DBSYSTEM.Execute CAD
    Set rsTRABTEMP = New ADODB.Recordset
        rsTRABTEMP.Open "[##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]", DBSYSTEM, adOpenDynamic, adLockOptimistic
    Set DgDet.DataSource = rsTRABTEMP
    '******************************************************************
Else
    CAD = "SELECT CODDATA FROM DATATRAB WHERE  CODDATA='" & Combo1.Text & "'"
    If DevuelveValor(CAD, DBSYSTEM) <> "" Then
        '*************************************
        CAD = ""
        For K = 0 To DgDet.Columns.Count - 1
            CAD = CAD & "TMP." & DgDet.Columns(K).Caption & ","
        Next K
        CAD = CAD & "TRA." & Combo1.Text
        
        If DevuelveValor("select " & Combo1.Text & " from [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]", DBSYSTEM) <> "" Then
            MsgBox "Este campo ya existe en la tabla"
            Exit Sub
        End If
        'ALMACENAR EL TMP GRID EN UN TMPGLOBAL
        If ExisteTablaAux("[##_TMPGRIDDET" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##_TMPGRIDDET" & VGL_COMPUTER & "]"
            DBSYSTEM.Execute "SELECT * INTO [##_TMPGRIDDET" & VGL_COMPUTER & "] FROM [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]"
        'EMININAR EL TMP DEL GRID
        If ExisteTablaAux("[##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]"
        
        CAD = "Select " & CAD & "  into  [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "] FROM [##_TMPGRIDDET" & VGL_COMPUTER & "] TMP,TRABAJADORES TRA WHERE TMP.CODTRAB=TRA.CODTRAB"
        '*************************************
        'CREAR LA TABLA
        DBSYSTEM.Execute CAD
        Set rsTRABTEMP = New ADODB.Recordset
        rsTRABTEMP.Open "[##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]", DBSYSTEM, adOpenDynamic, adLockOptimistic
        Set DgDet.DataSource = rsTRABTEMP
        
        
    End If
End If
        
        
Call WHIT_DATAGRID(DgDet)
Exit Sub
handler:
   If ERR.Number = -2147217900 Then
        
   End If
   
   Exit Sub
   Resume
End Sub

Private Sub CMDELIMINAR_CLICK()
If Head_GridDet = "" Then
    Exit Sub
    MsgBox "Seleccione la columna a eliminar"
End If
If MsgBox("Desea quitar la columna " & Head_GridDet, vbYesNo + vbQuestion) = vbNo Then Exit Sub
If Head_GridDet = "CODTRAB" Then Exit Sub
If Head_GridDet = "NOMBRES" Then Exit Sub
  
  If ExisteCampo(Head_GridDet, "[##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]", DBSYSTEM) = False Then
        MsgBox "El campo " & Head_GridDet & "ya No existe en la tabla", vbInformation
        Head_GridDet = ""
        lblcampo.Caption = ""
        Exit Sub
  End If
  
  DBSYSTEM.Execute "ALTER TABLE [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]  DROP COLUMN  " & Head_GridDet
  Set rsTRABTEMP = New ADODB.Recordset
  rsTRABTEMP.Open "[##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]", DBSYSTEM, adOpenDynamic, adLockOptimistic
  Set DgDet.DataSource = rsTRABTEMP
  
  Head_GridDet = ""
  lblcampo.Caption = ""
Call WHIT_DATAGRID(DgDet)
End Sub

Private Sub CMDIMPPLA_Click()
''**********COMINEZA LO BUENO---------
If rsTRABTEMP.EOF Then Exit Sub
Dim K As Integer

Screen.MousePointer = 11
If ExisteTablaAux("[##_TRABFINAL" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##_TRABFINAL" & VGL_COMPUTER & "]"
    'TODO ES DECLARADO VARCHAR POR SER CONCEPTOS MIXTOS
    DBSYSTEM.Execute "CREATE TABLE [##_TRABFINAL" & VGL_COMPUTER & "] (CODTRAB VARCHAR(8),CONCEPTO VARCHAR(50),VALOR VARCHAR(100),ORDEN INT)"
    
    
rsTRABTEMP.MoveFirst
 Do While Not rsTRABTEMP.EOF
        For K = 2 To DgDet.Columns.Count - 1
             DBSYSTEM.Execute "INSERT INTO [##_TRABFINAL" & VGL_COMPUTER & "] VALUES('" & rsTRABTEMP.Fields("CODTRAB") & "  " & rsTRABTEMP.Fields("NOMBRES") & "','" & DgDet.Columns(K).Caption & "','" & rsTRABTEMP.Fields(K) & "'," & K & ")"
        Next K
  rsTRABTEMP.MoveNext
 Loop

rsTRABTEMP.MoveFirst
    With CrystalReport1
        .Reset
        .WindowTitle = "PLAN0093 - REPORTE PERSONALIZADO DE CONCEPTOS FIJOS Y PROMEDIOS"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0093.rpt"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .StoredProcParam(0) = "[##_TRABFINAL" & VGL_COMPUTER & "]"
        .Destination = crptToWindow
        .WindowState = crptMaximized
'        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
'        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XCABEZA='" & Trim(xTitulo(0).Text) & "'"
        .Formulas(1) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(2) = "XRUC='" & REGSISTEMA.RUC & "'"
        .Formulas(3) = "XENCA1='" & Trim(xTitulo(1).Text) & "'"
        If .Status <> 2 Then .Action = 1
    End With
Screen.MousePointer = 1
End Sub

Private Sub Command1_Click()
   If DevuelveValor("select " & Combo2.Text & " from [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]", DBSYSTEM) <> "" Then
        MsgBox "Este campo ya existe en la Vista"
        Exit Sub
    End If
    If Combo2.Text = "" Then MsgBox "Debe seleccionar un concepto": Exit Sub
 '*-------------------------------------------
 Dim nFECHAINI As Long, nFECHAFIN As Long
 Dim nContarfecha As Integer
 Dim sMes As String, sAno As String, sAux As String
 nFECHAINI = FechS(DTPicker1.Value, Sqlf)
 nFECHAFIN = FechS(DTPicker2.Value, Sqlf)
 If nFECHAINI > nFECHAFIN Then
    MsgBox "La Fecha Inicial debe ser menor a la Fecha Final"
    Exit Sub
 End If
'*------------------------------------------------------
sAux = "01/" & Format(Month(DTPicker2.Value), "00") & "/" & Year(DTPicker2.Value)
nFECHAFIN = FechS(sAux, Sqlf)
sAux = "01/" & Format(Month(DTPicker1.Value), "00") & "/" & Year(DTPicker1.Value)
nFECHAINI = FechS(sAux, Sqlf)
Dim RSx1 As ADODB.Recordset
Dim RSPROMEDIO As ADODB.Recordset
Dim CAD As String

CambiaPanelBD True
Set RSx1 = New ADODB.Recordset
    RSx1.Open "[##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]", DBSYSTEM, adOpenDynamic, adLockOptimistic

If ExisteTablaAux("[##_TMPPROMEDIO" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##_TMPPROMEDIO" & VGL_COMPUTER & "]"
DBSYSTEM.Execute "CREATE TABLE [##_TMPPROMEDIO" & VGL_COMPUTER & "](CODTRAB VARCHAR(8),CONCEPTO VARCHAR(10),MONTO  Numeric(20,2)  DEFAULT 0,PRO INT DEFAULT 0,PROVAL INT DEFAULT 0)"

Do While nFECHAINI <= nFECHAFIN
   RSx1.MoveFirst
   Do While Not RSx1.EOF
        Set RSPROMEDIO = New ADODB.Recordset
        CAD = "SELECT * FROM [##_TMPPROMEDIO" & VGL_COMPUTER & "] WHERE CODTRAB='" & RSx1.Fields("CODTRAB") & "'"
        RSPROMEDIO.Open CAD, DBSYSTEM, adOpenDynamic, adLockOptimistic
            If RSPROMEDIO.EOF Then 'NO EXISTE -> CREAR
                RSPROMEDIO.AddNew
                        RSPROMEDIO!CODTRAB = RSx1!CODTRAB
                        RSPROMEDIO!CONCEPTO = Combo2.Text
                        RSPROMEDIO!MONTO = sumarConceptoMensual(RSx1!CODTRAB, Combo2.Text, Trim(Format(Month(sAux), "00") & Year(sAux)))
                        If Option1.Value = True Then
                            RSPROMEDIO!PRO = 1
                        Else
                            If sumarConceptoMensual(RSx1!CODTRAB, Combo2.Text, Trim(Format(Month(sAux), "00") & Year(sAux))) <> 0 Then
                                RSPROMEDIO!PROVAL = 1
                            End If
                        End If
                RSPROMEDIO.Update
            Else 'YA EXISTE -> ACUMULAR
                        RSPROMEDIO!MONTO = RSPROMEDIO!MONTO + sumarConceptoMensual(RSx1!CODTRAB, Combo2.Text, Trim(Format(Month(sAux), "00") & Year(sAux)))
                        If Option1.Value = True Then
                            RSPROMEDIO!PRO = RSPROMEDIO!PRO + 1
                        Else
                            If sumarConceptoMensual(RSx1!CODTRAB, Combo2.Text, Trim(Format(Month(sAux), "00") & Year(sAux))) <> 0 Then
                                RSPROMEDIO!PROVAL = RSPROMEDIO!PROVAL + 1
                            End If
                        End If
                    RSPROMEDIO.Update
            End If
        RSx1.MoveNext
   Loop
sAux = "01/" & Format(CInt(Month(sAux)) + 1, "00") & "/" & Year(DTPicker1.Value)
nFECHAINI = FechS(sAux, Sqlf)
nContarfecha = nContarfecha + 1
Loop
'******************************************************************
'ALMACENAR EL TMP GRID EN UN TMPGLOBAL
Dim sPromedio As String
If ExisteTablaAux("[##_TMPGRIDDET" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##_TMPGRIDDET" & VGL_COMPUTER & "]"
DBSYSTEM.Execute "SELECT * INTO [##_TMPGRIDDET" & VGL_COMPUTER & "] FROM [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]"
'EMININAR EL TMP
If ExisteTablaAux("[##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]"
        If Option1.Value = True Then
          sPromedio = "TMP2.PRO"
        Else
          sPromedio = "TMP2.PROVAL"
        End If
CAD = "Select  TMP.*," & Combo2.Text & "=CASE WHEN " & sPromedio & "=0 THEN 0 ELSE round((TMP2.MONTO/" & sPromedio & "),2) END  into  [##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "] FROM" & _
" [##_TMPGRIDDET" & VGL_COMPUTER & "] TMP,[##_TMPPROMEDIO" & VGL_COMPUTER & "] TMP2 WHERE TMP.CODTRAB=TMP2.CODTRAB"
 
DBSYSTEM.Execute CAD
Set rsTRABTEMP = New ADODB.Recordset
   rsTRABTEMP.Open "[##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]", DBSYSTEM, adOpenDynamic, adLockOptimistic
Set DgDet.DataSource = rsTRABTEMP
'******************************************************************
CambiaPanelBD False
Call WHIT_DATAGRID(DgDet)
End Sub

Private Sub DGDET_HEADCLICK(ByVal COLINDEX As Integer)
Static BIT_ORDEN As Boolean
        Head_GridDet = DgDet.Columns(COLINDEX).Caption
        lblcampo.Caption = "" & Head_GridDet
        
If Head_GridDet = "CODTRAB" Or Head_GridDet = "NOMBRES" Then
 If BIT_ORDEN = False Then
    rsTRABTEMP.Sort = Head_GridDet & " ASC"
    BIT_ORDEN = True
 Else
    rsTRABTEMP.Sort = Head_GridDet & " DESC"
    BIT_ORDEN = False
 End If
End If

End Sub

Private Sub Form_Load()
Dim RSAUX As ADODB.Recordset
Dim K As Integer

Set rsTRABTEMP = New ADODB.Recordset
rsTRABTEMP.Open "[##_TMPTRABPERSONALIZADO" & VGL_COMPUTER & "]", DBSYSTEM, adOpenDynamic, adLockOptimistic
Set DgDet.DataSource = rsTRABTEMP
'para ver el total trabajador
lblTotalTrab.Caption = rsTRABTEMP.RecordCount

Set rstrabajadores = New ADODB.Recordset
rstrabajadores.Open "[##_TMPTRABPERSONALIZADOTOTAL" & VGL_COMPUTER & "]", DBSYSTEM, adOpenDynamic, adLockOptimistic


Set RSAUX = New ADODB.Recordset
RSAUX.Open "SELECT * FROM DATATRAB", DBSYSTEM, adOpenDynamic, adLockReadOnly

Do While Not RSAUX.EOF
        If RSAUX!TIPODATA = "N" Then
            Combo1.AddItem RSAUX!CODDATA
        End If
    RSAUX.MoveNext
Loop

Dim RSXXXX As ADODB.Recordset
Dim TIPOXX As String
Set RSXXXX = New ADODB.Recordset
RSXXXX.Open "sp_existecampo '" & REGSISTEMA.BASESQL & "','VWTRABAJ'", DBSTARPLAN, adOpenDynamic, adLockReadOnly

Do While Not RSXXXX.EOF
        TIPOXX = UCase(RSXXXX.Fields(1))
        If TIPOXX = "INT" Or TIPOXX = "SMALLINT" _
            Or Mid(TIPOXX, 1, 3) = "TIN" Or Mid(TIPOXX, 1, 2) = "DE" _
            Or Mid(TIPOXX, 1, 2) = "MO" Or Mid(TIPOXX, 1, 3) = "SMA" _
            Or Mid(TIPOXX, 1, 3) = "FLO" Or TIPOXX = "REAL" Or Mid(TIPOXX, 1, 4) = "NUME" Then
            
            Combo1.AddItem RSXXXX.Fields(0)
        End If
    RSXXXX.MoveNext
Loop
'*****************************************
Set RSXXXX = New ADODB.Recordset
RSXXXX.Open "SELECT * FROM CONCEPTOS WHERE TIPO=1 OR TIPO=2", DBSYSTEM, adOpenDynamic, adLockReadOnly
Do While Not RSXXXX.EOF
            Combo2.AddItem RSXXXX.Fields(0)
    RSXXXX.MoveNext
Loop
'*******************************************
RSXXXX.Close
Call WHIT_DATAGRID(DgDet)
End Sub

Private Sub IMAGE1_CLICK()
frmAcerca_util.Show 1
End Sub

Private Function sumarConceptoMensual(ByVal CODTRAB As String, ByVal CONCEP As String, ByVal MESANO As String) As Double
On Error GoTo handler
    Dim CAD As String
    CAD = " SELECT BOL.CODTRAB,MOV.CONCEPTO,SUM(MOV.MONTO) AS MONTO  " & _
          " FROM BOL" & MESANO & " BOL,MOV" & MESANO & " MOV WHERE BOL.CODTRAB='" & CODTRAB & "' AND  " & _
          " MOV.CONCEPTO='" & CONCEP & "'  AND BOL.INUMBOL=MOV.INUMBOL  " & _
          " GROUP BY BOL.CODTRAB,MOV.CONCEPTO "
    Dim RSX As ADODB.Recordset
    Set RSX = New ADODB.Recordset
    RSX.Open CAD, DBSYSTEM, adOpenDynamic, adLockReadOnly
        sumarConceptoMensual = RSX!MONTO
 Exit Function
handler:
sumarConceptoMensual = 0
End Function
Private Sub WHIT_DATAGRID(MEDATA As DataGrid)
 MEDATA.Columns(0).Width = 850
 MEDATA.Columns(1).Width = 3000
 
 For K% = 2 To MEDATA.Columns.Count - 1
     MEDATA.Columns(K).Width = 1000
     MEDATA.Columns(K).Alignment = dbgRight
     MEDATA.Columns(K).NumberFormat = "###0.00"
 Next K
 
End Sub


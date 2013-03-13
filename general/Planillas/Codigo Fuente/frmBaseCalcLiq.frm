VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBaseCalcLiq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Base de Cálculo de Liquidación"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "frmBaseCalcLiq.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCalcular 
      Caption         =   "&Recalcular"
      Height          =   360
      Left            =   195
      TabIndex        =   32
      Top             =   5520
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   4830
      TabIndex        =   9
      Top             =   5520
      Width           =   1185
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   4095
      Left            =   195
      TabIndex        =   3
      Top             =   1335
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "CTS"
      TabPicture(0)   =   "frmBaseCalcLiq.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Suma1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "xData1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "xFecCTS"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "xFechaCese"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Vacaciones"
      TabPicture(1)   =   "frmBaseCalcLiq.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shape1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Suma2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label3(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "xFecha2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "xFecVac"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "xData2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdEliminar"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdAgregar"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Gratificaciones"
      TabPicture(2)   =   "frmBaseCalcLiq.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Suma3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label3(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label1(2)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "xFec3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "xFecGrat"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "xData3"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Command4"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Command5"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      Begin MSComCtl2.DTPicker xFechaCese 
         Height          =   300
         Left            =   -71655
         TabIndex        =   23
         Top             =   645
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24510465
         CurrentDate     =   36850
      End
      Begin MSComCtl2.DTPicker xFecCTS 
         Height          =   300
         Left            =   -74205
         TabIndex        =   21
         Top             =   645
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36850
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Agregar"
         Height          =   315
         Left            =   -74715
         TabIndex        =   16
         Top             =   3420
         Width           =   960
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Quitar"
         Height          =   315
         Left            =   -73620
         TabIndex        =   15
         Top             =   3420
         Width           =   960
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Agregar"
         Height          =   315
         Left            =   -74715
         TabIndex        =   11
         Top             =   3420
         Width           =   960
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Quitar"
         Height          =   315
         Left            =   -73620
         TabIndex        =   10
         Top             =   3420
         Width           =   960
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   315
         Left            =   285
         TabIndex        =   5
         Top             =   3420
         Width           =   960
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Quitar"
         Height          =   315
         Left            =   1380
         TabIndex        =   4
         Top             =   3420
         Width           =   960
      End
      Begin MSDataGridLib.DataGrid xData2 
         Height          =   2310
         Left            =   285
         TabIndex        =   6
         Top             =   1020
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   4075
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   15395562
         HeadLines       =   2
         RowHeight       =   17
         FormatLocked    =   -1  'True
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Detalle del Cálculo"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Concepto"
            Caption         =   "Conceptos Computables"
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
            DataField       =   "Importe"
            Caption         =   "Importe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00 "
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            BeginProperty Column00 
               ColumnWidth     =   3030.236
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               DividerStyle    =   1
               ColumnWidth     =   1214.929
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid xData3 
         Height          =   2310
         Left            =   -74715
         TabIndex        =   12
         Top             =   1020
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   4075
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   15395562
         HeadLines       =   2
         RowHeight       =   17
         FormatLocked    =   -1  'True
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Detalle del Cálculo"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Concepto"
            Caption         =   "Conceptos Computables"
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
            DataField       =   "Importe"
            Caption         =   "Importe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00 "
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            BeginProperty Column00 
               ColumnWidth     =   3030.236
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               DividerStyle    =   1
               ColumnWidth     =   1214.929
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid xData1 
         Height          =   2310
         Left            =   -74715
         TabIndex        =   17
         Top             =   1020
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   4075
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   15395562
         HeadLines       =   2
         RowHeight       =   17
         FormatLocked    =   -1  'True
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Detalle del Cálculo"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Concepto"
            Caption         =   "Conceptos Computables"
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
            DataField       =   "Importe"
            Caption         =   "Importe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00 "
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            BeginProperty Column00 
               ColumnWidth     =   3030.236
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               DividerStyle    =   1
               ColumnWidth     =   1214.929
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker xFecVac 
         Height          =   300
         Left            =   795
         TabIndex        =   28
         Top             =   645
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36850
      End
      Begin MSComCtl2.DTPicker xFecha2 
         Height          =   300
         Left            =   3345
         TabIndex        =   29
         Top             =   645
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24510465
         CurrentDate     =   36850
      End
      Begin MSComCtl2.DTPicker xFecGrat 
         Height          =   300
         Left            =   -74205
         TabIndex        =   30
         Top             =   645
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36850
      End
      Begin MSComCtl2.DTPicker xFec3 
         Height          =   300
         Left            =   -71655
         TabIndex        =   31
         Top             =   645
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24510465
         CurrentDate     =   36850
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   -74685
         TabIndex        =   27
         Top             =   705
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Término"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   -72360
         TabIndex        =   26
         Top             =   690
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Término"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   25
         Top             =   690
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   24
         Top             =   705
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Término"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   -72360
         TabIndex        =   22
         Top             =   690
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   -74685
         TabIndex        =   20
         Top             =   705
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Cálculo"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   -71910
         TabIndex        =   19
         Top             =   3465
         Width           =   930
      End
      Begin VB.Label Suma1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   270
         Left            =   -70890
         TabIndex        =   18
         Top             =   3420
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Cálculo"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   -71910
         TabIndex        =   14
         Top             =   3465
         Width           =   930
      End
      Begin VB.Label Suma3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   270
         Left            =   -70890
         TabIndex        =   13
         Top             =   3420
         Width           =   1335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Cálculo"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3090
         TabIndex        =   8
         Top             =   3465
         Width           =   930
      End
      Begin VB.Label Suma2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   270
         Left            =   4110
         TabIndex        =   7
         Top             =   3420
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   3360
         Left            =   180
         Top             =   525
         Width           =   5460
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   3360
         Left            =   -74820
         Top             =   525
         Width           =   5460
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   3360
         Left            =   -74820
         Top             =   525
         Width           =   5460
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Trabajador"
      Height          =   930
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   5820
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   270
         Width           =   630
      End
      Begin VB.Label xTrab 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Daniel Yafac Baquedano"
         Height          =   285
         Left            =   195
         TabIndex        =   1
         Top             =   510
         Width           =   5310
      End
   End
End
Attribute VB_Name = "frmBaseCalcLiq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSCTS As New ADODB.Recordset
Dim RSVAC As New ADODB.Recordset
Dim RSGRAT As New ADODB.Recordset
Private Sub CMCALCULAR_CLICK()
    DBAUXCOM.Execute "DELETE FROM  [##TMPLIQUIDA" & VGL_COMPUTER & "] "
    REALIZARCALCULOCTS
End Sub

Private Sub CMDAGREGAR_CLICK()
    Dim XSTR As String
    XSTR = InputBox("INGRESE LA DESCRIPCIÓN DEL CONCEPTO PARA VACACIONES", "AGREGAR CONCEPTO")
    Set RSVAC = Nothing
    RSVAC.Open " [##TMPLIQUIDA" & VGL_COMPUTER & "] ", DBAUXCOM, adOpenKeyset, adLockOptimistic
    If XSTR <> "" Then
        frValor.Show 1
        If Val(VPTAREA) <> 0 Then
            RSVAC.AddNew
            RSVAC!TIPO = 2
            RSVAC!Importe = Val(VPTAREA)
            RSVAC!CONCEPTO = XSTR
            RSVAC.Update
            XDATA2_AFTERCOLUPDATE 0
        End If
    End If
    Call ACTRS
End Sub

Private Sub CMDELIMINAR_CLICK()
    If RSVAC.EOF Then Exit Sub
    If MsgBox("REALMENTE DESEA ELIMINAR EL REGISTRO SELECCIONADO", vbYesNo + vbQuestion) = vbYes Then RSVAC.Delete
    XDATA2_AFTERCOLUPDATE 0
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If RSCTS.EOF Then Exit Sub
    If MsgBox("REALMENTE DESEA ELIMINAR EL REGISTRO SELECCIONADO", vbYesNo + vbQuestion) = vbYes Then RSCTS.Delete
    XDATA1_AFTERCOLUPDATE 0
End Sub

Private Sub COMMAND3_Click()
    Dim XSTR As String
    XSTR = InputBox("INGRESE LA DESCRIPCIÓN DEL CONCEPTO PARA CTS", "AGREGAR CONCEPTO")
    If XSTR <> "" Then
        Set RSCTS = Nothing
        RSCTS.Open " [##TMPLIQUIDA" & VGL_COMPUTER & "] ", DBAUXCOM, adOpenKeyset, adLockOptimistic
        frValor.Show 1
        If Val(VPTAREA) <> 0 Then
            RSCTS.AddNew
            RSCTS!TIPO = 1
            RSCTS!Importe = Val(VPTAREA)
            RSCTS!CONCEPTO = XSTR
            RSCTS.Update
            XDATA1_AFTERCOLUPDATE 0
        End If
    End If
    Call ACTRS
End Sub

Private Sub COMMAND4_Click()
    If RSGRAT.EOF Then Exit Sub
    If MsgBox("REALMENTE DESEA ELIMINAR EL REGISTRO SELECCIONADO", vbYesNo + vbQuestion) = vbYes Then RSGRAT.Delete
    XDATA3_AFTERCOLUPDATE 0
End Sub

Private Sub Command5_Click()
    Dim XSTR As String
    XSTR = InputBox("INGRESE LA DESCRIPCIÓN DEL CONCEPTO PARA GRATIFICACIONES", "AGREGAR CONCEPTO")
    Set RSGRAT = Nothing
    RSGRAT.Open " [##TMPLIQUIDA" & VGL_COMPUTER & "] ", DBAUXCOM, adOpenKeyset, adLockOptimistic
    If XSTR <> "" Then
        frValor.Show 1
        If Val(VPTAREA) <> 0 Then
            RSGRAT.AddNew
            RSGRAT!TIPO = 3
            RSGRAT!Importe = Val(VPTAREA)
            RSGRAT!CONCEPTO = XSTR
            RSGRAT.Update
            XDATA3_AFTERCOLUPDATE 0
        End If
    End If
    Call ACTRS
End Sub

Private Sub Form_Load()
    If ExisteTablaAux(" [##TMPLIQUIDA" & VGL_COMPUTER & "] ") Then DBAUXCOM.Execute "DROP TABLE  [##TMPLIQUIDA" & VGL_COMPUTER & "] "
    DBAUXCOM.Execute "CREATE TABLE  [##TMPLIQUIDA" & VGL_COMPUTER & "]  (TIPO INTEGER, CONCEPTO VARCHAR(35), IMPORTE  Numeric(20,2) )"
End Sub

Public Sub REALIZARCALCULOCTS()
Dim GENERAL  As Boolean
    Dim VALOR As Single
    Dim RSCNPT As New ADODB.Recordset
    'CALCULO DE LA CTS
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    RSCNPT.Open "SELECT * FROM FORMULASCTS WHERE AFECTOPRO<>0", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSCNPT.EOF Or RSCNPT.RecordCount = 0 Then
        MsgBox "MENSAJE DEL SISTEMA: EL SISTEMA NO HA ENCONTRADO FÓRMULAS DE CTS", vbInformation
        Set RSCNPT = Nothing
        Exit Sub
    End If
    Do While Not RSCNPT.EOF
        GENERAL = RSCNPT!GENE
        If InStr(RSCNPT!FORMULA, "@") = 0 Then
        VALOR = DevuelveValor("SELECT " & RSCNPT!FORMULA & " AS VALOR_DEV FROM VWTRABAJ WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
            If IsNull(VALOR) Then VALOR = 0
        Else
            VALOR = DevuelveValor("SELECT " & CAMBIACADENA(RSCNPT!FORMULA, xTrab.Tag, GENERAL, xFecCTS.Value, xFechaCese.Value) & " AS VALOR_DEV FROM VWTRABAJ WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
        End If
        If VALOR <> 0 Then
            VALOR = Round(VALOR, 2)
            DBAUXCOM.Execute "INSERT INTO  [##TMPLIQUIDA" & VGL_COMPUTER & "]  VALUES (1,'" & RSCNPT!NOMBRE & "'," & VALOR & ")"
        End If
        RSCNPT.MoveNext
    Loop
    Set RSCNPT = Nothing
    'CALCULO DE LAS VACACIONES
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    If Tab1.TabEnabled(1) Then
        RSCNPT.Open "SELECT * FROM FORMULASVAC WHERE TIPO=0 AND AFECTOPRO<>0", DBSYSTEM, adOpenStatic, adLockReadOnly
        If RSCNPT.EOF Or RSCNPT.RecordCount = 0 Then
            MsgBox "MENSAJE DEL SISTEMA: EL SISTEMA NO HA ENCONTRADO FÓRMULAS DE VACACIONES", vbInformation
            Set RSCNPT = Nothing
            Exit Sub
        End If
        Do While Not RSCNPT.EOF
            GENERAL = RSCNPT!GENE
            If InStr(RSCNPT!FORMULA, "@") = 0 Then
            VALOR = DevuelveValor("SELECT " & RSCNPT!FORMULA & " AS VALOR_DEV FROM VWTRABAJ WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
                If IsNull(VALOR) Then VALOR = 0
            Else
                VALOR = DevuelveValor("SELECT " & CAMBIACADENA(RSCNPT!FORMULA, xTrab.Tag, GENERAL, xFecVac.Value, xFechaCese.Value) & " AS VALOR_DEV FROM VWTRABAJ WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
            End If
            If VALOR <> 0 Then
                VALOR = Round(VALOR, 2)
                DBAUXCOM.Execute "INSERT INTO  [##TMPLIQUIDA" & VGL_COMPUTER & "]  VALUES (2,'" & RSCNPT!NOMBRE & "'," & VALOR & ")"
            End If
            RSCNPT.MoveNext
        Loop
    End If
    'CALCULO DE LAS GRATIFICACIONES
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    If Tab1.TabEnabled(2) Then
        Set RSCNPT = Nothing
        RSCNPT.Open "SELECT * FROM FORMULASGRATI WHERE TIPO=0 AND AFECTOPRO<>0", DBSYSTEM, adOpenStatic, adLockReadOnly
        If RSCNPT.EOF Or RSCNPT.RecordCount = 0 Then
            MsgBox "MENSAJE DEL SISTEMA: EL SISTEMA NO HA ENCONTRADO FÓRMULAS DE GRATIFICACIONES", vbInformation
            Set RSCNPT = Nothing
            Exit Sub
        End If
        Do While Not RSCNPT.EOF
            If InStr(RSCNPT!FORMULA, "@") = 0 Then
            VALOR = DevuelveValor("SELECT " & RSCNPT!FORMULA & " AS VALOR_DEV FROM VWTRABAJ WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
                If IsNull(VALOR) Then VALOR = 0
            Else
                VALOR = DevuelveValor("SELECT " & CAMBIACADENA(RSCNPT!FORMULA, xTrab.Tag, GENERAL, xFecGrat.Value, xFechaCese.Value) & " AS VALOR_DEV FROM VWTRABAJ WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
            End If
            If VALOR <> 0 Then
                VALOR = Round(VALOR, 2)
                DBAUXCOM.Execute "INSERT INTO  [##TMPLIQUIDA" & VGL_COMPUTER & "]  VALUES (3,'" & RSCNPT!NOMBRE & "'," & VALOR & ")"
            End If
            RSCNPT.MoveNext
        Loop
    End If
    Set RSCNPT = Nothing
    Call ACTRS
End Sub
Private Sub ACTRS()
    Set RSCTS = Nothing
    Set RSVAC = Nothing
    Set RSGRAT = Nothing
    RSCTS.Open "SELECT * FROM  [##TMPLIQUIDA" & VGL_COMPUTER & "]  WHERE TIPO=1", DBAUXCOM, adOpenDynamic, adLockOptimistic
    RSVAC.Open "SELECT * FROM  [##TMPLIQUIDA" & VGL_COMPUTER & "]  WHERE TIPO=2", DBAUXCOM, adOpenDynamic, adLockOptimistic
    RSGRAT.Open "SELECT * FROM  [##TMPLIQUIDA" & VGL_COMPUTER & "]  WHERE TIPO=3", DBAUXCOM, adOpenDynamic, adLockOptimistic
    Set xData1.DataSource = RSCTS
    Set xData2.DataSource = RSVAC
    Set xData3.DataSource = RSGRAT
    XDATA1_AFTERCOLUPDATE 0
    XDATA2_AFTERCOLUPDATE 0
    XDATA3_AFTERCOLUPDATE 0
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    With frLiquidacion
        .xBase1.Caption = Suma1.Caption
        .xBase2.Caption = Suma2.Caption
        .xBase3.Caption = Suma3.Caption
    End With
    Set RSCTS = Nothing
    Set RSVAC = Nothing
    Set RSGRAT = Nothing
End Sub

Public Function CAMBIACADENA(ByVal CADENA As String, ByVal CODTRAB As String, GENERAL2 As Boolean, Optional F0 As Date, Optional F1 As Date) As String
    On Error Resume Next
    Dim POSARROBA As Integer, POS1 As Integer, PROCESO As String, CAMPO As String, POS2 As Integer
    Dim VALOR As Double
    POSARROBA = 1
    POSARROBA = InStr(POSARROBA, CADENA, "@")
    Do While POSARROBA <> 0
        POS1 = InStr(POSARROBA, CADENA, "(")
        PROCESO = Mid(CADENA, POSARROBA + 1, POS1 - (POSARROBA + 1))
        POS2 = InStr(POSARROBA, CADENA, ")")
        CAMPO = Mid(CADENA, POS1 + 1, POS2 - (POS1 + 1))
        Select Case UCase(PROCESO)
            Case "PROMEDIO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, F0, F1, PROMEDIO, CAMPO, GENERAL2)
            Case "ULTIMOVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, F0, F1, ULTIMOVALOR, CAMPO, GENERAL2)
            Case "PRIMERVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, F0, F1, PRIMERVALOR, CAMPO, GENERAL2)
            Case "SUMA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, F0, F1, SUMA, CAMPO, GENERAL2)
            Case "MEDIA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, F0, F1, MEDIA, CAMPO, GENERAL2)
            Case "PROMEDIOVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, F0, F1, PROMEDIOVALOR, CAMPO, GENERAL2)
            Case "PRIMERO", GENERAL2
                VALOR = CALCULOCONCEPTOS(CODTRAB, F0, F1, PRIMERO, CAMPO, GENERAL2)
            Case "ULTIMO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, F0, F1, ULTIMO, CAMPO, GENERAL2)
            Case "MAYORVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, F0, F1, MAYORVALOR, CAMPO, GENERAL2)
            Case "MENORVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, F0, F1, MENORVALOR, CAMPO, GENERAL2)
            Case "NUMERO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, F0, F1, NUMERO, CAMPO, GENERAL2)
        End Select
        If IsNull(VALOR) Then VALOR = 0
        CADENA = Replace(CADENA, Mid(CADENA, POSARROBA, (POS2 - POSARROBA) + 1), "" & VALOR)
        POSARROBA = InStr(POSARROBA, CADENA, "@")
    Loop
    CAMBIACADENA = CADENA
End Function

Private Sub XDATA1_AFTERCOLUPDATE(ByVal COLINDEX As Integer)
    If RSCTS.EOF Then Exit Sub
    RSCTS.MOVE 0
    Suma1.Caption = Format(DevuelveValor("SELECT SUM(IMPORTE) AS T1 FROM  [##TMPLIQUIDA" & VGL_COMPUTER & "]  WHERE TIPO=1", DBAUXCOM), "0.00 ")
End Sub

Private Sub XDATA2_AFTERCOLUPDATE(ByVal COLINDEX As Integer)
    If RSVAC.EOF Then Exit Sub
    RSVAC.MOVE 0
    Suma2.Caption = Format(DevuelveValor("SELECT SUM(IMPORTE) AS T1 FROM  [##TMPLIQUIDA" & VGL_COMPUTER & "]  WHERE TIPO=2", DBAUXCOM), "0.00 ")
End Sub

Private Sub XDATA3_AFTERCOLUPDATE(ByVal COLINDEX As Integer)
    If RSGRAT.EOF Then Exit Sub
    RSGRAT.MOVE 0
    Suma3.Caption = Format(DevuelveValor("SELECT SUM(IMPORTE) AS T1 FROM  [##TMPLIQUIDA" & VGL_COMPUTER & "]  WHERE TIPO=3", DBAUXCOM), "0.00 ")
End Sub



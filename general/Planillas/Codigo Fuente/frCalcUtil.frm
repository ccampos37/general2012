VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frCalcUtil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de UTILIDADES"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   Icon            =   "frCalcUtil.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Calculo"
      Enabled         =   0   'False
      Height          =   960
      Left            =   5325
      TabIndex        =   28
      Top             =   960
      Width           =   4500
      Begin AplisetControlText.Aplitext XUtil 
         Height          =   300
         Left            =   2910
         TabIndex        =   36
         Top             =   225
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         Text            =   "0"
         SinBlancos      =   -1  'True
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xPorc 
         Height          =   285
         Left            =   1365
         TabIndex        =   35
         Top             =   225
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         MaxLength       =   15
         Text            =   "0"
         SinBlancos      =   -1  'True
         TipoDato        =   "N"
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   270
         Left            =   2145
         Top             =   585
         Width           =   2220
      End
      Begin VB.Label Label12 
         Caption         =   "Participacion a distribuir"
         Height          =   270
         Left            =   150
         TabIndex        =   34
         Top             =   615
         Width           =   1755
      End
      Begin VB.Label Label11 
         Caption         =   "Utilidad"
         Height          =   225
         Left            =   2235
         TabIndex        =   33
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "%  de Particip."
         Height          =   285
         Left            =   150
         TabIndex        =   32
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label xPartDist 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0.00 "
         Height          =   255
         Left            =   2160
         TabIndex        =   37
         Top             =   600
         Width           =   2190
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir Constancias"
      Height          =   360
      Left            =   8115
      TabIndex        =   27
      Top             =   6390
      Visible         =   0   'False
      Width           =   1650
   End
   Begin AplisetControlText.Aplitext xDias 
      Height          =   270
      Left            =   9150
      TabIndex        =   23
      Top             =   3000
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   476
      MaxLength       =   5
      Text            =   "0"
      Entero          =   -1  'True
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext xMeses 
      Height          =   270
      Left            =   7860
      TabIndex        =   21
      Top             =   3000
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   476
      MaxLength       =   5
      Text            =   "0"
      Entero          =   -1  'True
      TipoDato        =   "N"
   End
   Begin MSComctlLib.ProgressBar Prog 
      Height          =   135
      Left            =   120
      TabIndex        =   17
      Top             =   6930
      Visible         =   0   'False
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   8505
      TabIndex        =   12
      Top             =   180
      Width           =   1305
   End
   Begin VB.CommandButton cmGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   8505
      TabIndex        =   11
      Top             =   600
      Width           =   1305
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   6135
      Top             =   6180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmCalcular 
      Caption         =   "&Calcular"
      Height          =   765
      Left            =   6345
      Picture         =   "frCalcUtil.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   165
      Width           =   870
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   3540
      Left            =   135
      TabIndex        =   8
      Top             =   2790
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   6244
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "Trabajadores Seleccionados"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "CodTrab"
         Caption         =   "Codigo"
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
         DataField       =   "Nombres"
         Caption         =   "Apellidos y Nombres"
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
      BeginProperty Column02 
         DataField       =   "PartPer"
         Caption         =   "Part Segun Periodo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,###,##0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "PartRem"
         Caption         =   "Part Segur Rem. Perc."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,###,##0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "TotPart"
         Caption         =   "Total Participacion"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###,###,##0.00 "
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
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2700.284
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1214.929
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmSelecTrab 
      Caption         =   "(F5)"
      Height          =   765
      Left            =   5415
      Picture         =   "frCalcUtil.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   165
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "Planilla de Utilidades"
      Height          =   1845
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   5205
      Begin VB.OptionButton OpPer 
         Caption         =   "Calculo por Horas"
         Height          =   255
         Index           =   1
         Left            =   3465
         TabIndex        =   52
         Top             =   1485
         Width           =   1605
      End
      Begin VB.OptionButton OpPer 
         Caption         =   "Calculo por Dias"
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   31
         Top             =   1485
         Width           =   1605
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Dias efectivos"
         Height          =   195
         Left            =   3555
         TabIndex        =   10
         Top             =   750
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   330
         Left            =   1065
         TabIndex        =   6
         Top             =   1080
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMMM'del ' yyyy"
         Format          =   62062595
         CurrentDate     =   36816
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   330
         Left            =   1065
         TabIndex        =   5
         Top             =   675
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMMM'del ' yyyy"
         Format          =   62062595
         CurrentDate     =   36816
      End
      Begin AplisetControlText.Aplitext xPeriodo 
         Height          =   300
         Left            =   1065
         TabIndex        =   2
         Top             =   285
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   529
         Text            =   ""
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   300
         TabIndex        =   4
         Top             =   1155
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   300
         TabIndex        =   3
         Top             =   750
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         Height          =   195
         Left            =   300
         TabIndex        =   1
         Top             =   345
         Width           =   540
      End
   End
   Begin MSDataGridLib.DataGrid xDetalle 
      Height          =   2775
      Left            =   5265
      TabIndex        =   26
      Top             =   3555
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            DividerStyle    =   1
            ColumnWidth     =   1530.142
         EndProperty
      EndProperty
   End
   Begin VB.Label xTotRemu 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   255
      Left            =   3975
      TabIndex        =   53
      Top             =   6585
      Width           =   1065
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total x Periodo"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2850
      TabIndex        =   56
      Top             =   6390
      Width           =   1065
   End
   Begin VB.Label xTotPer 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   255
      Left            =   3975
      TabIndex        =   55
      Top             =   6360
      Width           =   1065
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total x Rem."
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2865
      TabIndex        =   54
      Top             =   6645
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Planilla"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5145
      TabIndex        =   13
      Top             =   6375
      Width           =   900
   End
   Begin VB.Label xImpRem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8250
      TabIndex        =   51
      Top             =   2445
      Width           =   1560
   End
   Begin VB.Label xTRem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6795
      TabIndex        =   50
      Top             =   2445
      Width           =   1470
   End
   Begin VB.Label XpartRem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5340
      TabIndex        =   49
      Top             =   2445
      Width           =   1470
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Imp. x cada N S/."
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8280
      TabIndex        =   48
      Top             =   2220
      Width           =   1500
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Total de Rem"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   6825
      TabIndex        =   47
      Top             =   2220
      Width           =   1350
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Partic. a Distrib."
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   5355
      TabIndex        =   46
      Top             =   2220
      Width           =   1380
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Importe por cada Nuevo Sol de Remuneraciones Percibidas"
      Height          =   210
      Left            =   5340
      TabIndex        =   45
      Top             =   1980
      Width           =   4470
   End
   Begin VB.Label xImpPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3555
      TabIndex        =   44
      Top             =   2445
      Width           =   1725
   End
   Begin VB.Label XPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1725
      TabIndex        =   43
      Top             =   2445
      Width           =   1845
   End
   Begin VB.Label xPartPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      TabIndex        =   42
      Top             =   2445
      Width           =   1635
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Imp. x cada D o H"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3705
      TabIndex        =   41
      Top             =   2220
      Width           =   1500
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Total  de H. Labora."
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   1755
      TabIndex        =   40
      Top             =   2220
      Width           =   1605
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Partic. a Distrib."
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   105
      TabIndex        =   39
      Top             =   2220
      Width           =   1515
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Importe por cada Hora 50 %"
      Height          =   210
      Left            =   105
      TabIndex        =   38
      Top             =   1980
      Width           =   5190
   End
   Begin VB.Label xTotRem 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   270
      Left            =   7875
      TabIndex        =   30
      Top             =   3270
      Width           =   1890
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Remuneracion"
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5265
      TabIndex        =   29
      Top             =   3270
      Width           =   4515
   End
   Begin VB.Label xFecha 
      Alignment       =   2  'Center
      Height          =   225
      Left            =   6105
      TabIndex        =   25
      Top             =   3030
      Width           =   1140
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " F. Ingreso"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5265
      TabIndex        =   24
      Top             =   3000
      Width           =   2040
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Horas"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   8460
      TabIndex        =   22
      Top             =   3000
      Width           =   1320
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dias"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   7275
      TabIndex        =   20
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tiempo Computable"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   5265
      TabIndex        =   19
      Top             =   2790
      Width           =   4515
   End
   Begin VB.Label xProg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "**** Texto *****"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   135
      TabIndex        =   18
      Top             =   6705
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label xNumTrabs 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   1995
      TabIndex        =   16
      Top             =   6405
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Trabajadores"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   135
      TabIndex        =   15
      Top             =   6435
      Width           =   1755
   End
   Begin VB.Label xTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   255
      Left            =   6210
      TabIndex        =   14
      Top             =   6345
      Width           =   1065
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   4395
      Left            =   75
      Top             =   2745
      Width           =   9780
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   780
      Left            =   75
      Top             =   1950
      Width           =   5235
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   780
      Left            =   5310
      Top             =   1950
      Width           =   4530
   End
End
Attribute VB_Name = "frCalcUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents RSTRABS As ADODB.Recordset
Attribute RSTRABS.VB_VarHelpID = -1
Dim RSCALC As New ADODB.Recordset
Dim ENPROCESO As Boolean
Dim PARTDIST As Double
Dim TOTPER As Double, IMPPER As Double
Dim TOTREM As Double, IMPREM As Double
Dim FLAG As Boolean
Private Sub DIASHORAS()
    On Error Resume Next
    Dim GENERAL As Boolean
    Screen.MousePointer = 11
    If RSTRABS.EOF Or RSTRABS.RecordCount = 0 Then
        MsgBox "MENSAJE DEL SISTEMA: NO SE PUEDE PROCESAR LA TAREA REQUERIDA, SI NO HA SELECCIONADO UNO O MAS TRABAJADORES. PRESIONE F5 PARA SELECCIONAR TRABAJADORES", vbInformation
        Exit Sub
    End If
    DBSYSTEM.Execute "DELETE FROM  [##TMPCTS2" & VGL_COMPUTER & "] "
    Dim XFEC As Date, NUMHORAS As Integer, NUMDIAS As Integer, XFEC2 As Date
    'PONER A1 EL DIA DEL MES DE INICIO
    XPer.Caption = "0.00 ": xImpPer.Caption = "0.00 "
    xFechaIni.Day = 1
    'PONER EL ULTIMO DIA DEL MES PARA LA FECHA FINAL
    xFechaFin.Day = 1
    xFechaFin.Value = DateAdd("M", 1, xFechaFin.Value)
    xFechaFin.Value = DateAdd("D", -1, xFechaFin.Value)
    Prog.Min = 0
    Prog.Max = Val(xNumTrabs.Caption)
    Prog.Visible = True
    Prog.Value = 0
    xProg.Visible = True
    xProg.Caption = "ASIGNANDO TIEMPO COMPUTABLE"
    ENPROCESO = True
    'CALCULO DEL TIEMPO COMPUTABLE
    Dim RSCNPT As New ADODB.Recordset
    If Check1.Value = 1 Then
        RSCNPT.Open "SELECT * FROM CONCEPTOS WHERE TIPO=0 AND TIPOINFO<>5", DBSYSTEM, adOpenStatic, adLockReadOnly
        If RSCNPT.EOF Or RSCNPT.RecordCount = 0 Then
            MsgBox "NO SE HAN ENCONTRADO CONCEPTOS INFORMATIVOS DE DIAS U HORAS TRABAJADAS COMPUTABLES A BENEFICIOS SOCIALES", vbInformation
            Set RSCNPT = Nothing
            Screen.MousePointer = 1
            Exit Sub
        End If
        RSTRABS.MoveFirst
        Do While Not RSTRABS.EOF
            RSCNPT.MoveFirst
            Prog.Value = Prog.Value + 1
            NUMDIAS = 0
            Do While Not RSCNPT.EOF
                Select Case RSCNPT!TIPOINFO
                    Case 0
                        NUMDIAS = NUMDIAS + CALCULOCONCEPTOS(RSTRABS!CODTRAB, xFechaIni.Value, xFechaFin.Value, SUMA, RSCNPT!Codigo, False)
                    Case 1
                        NUMHORAS = NUMHORAS + CALCULOCONCEPTOS(RSTRABS!CODTRAB, xFechaIni.Value, xFechaFin.Value, SUMA, RSCNPT!Codigo, False)
                    Case 3
                        NUMDIAS = NUMDIAS - CALCULOCONCEPTOS(RSTRABS!CODTRAB, xFechaIni.Value, xFechaFin.Value, SUMA, RSCNPT!Codigo, False)
                    Case 4
                        NUMHORAS = NUMHORAS - CALCULOCONCEPTOS(RSTRABS!CODTRAB, xFechaIni.Value, xFechaFin.Value, SUMA, RSCNPT!Codigo, False)
                End Select
                Me.Refresh
                RSCNPT.MoveNext
            Loop
            RSTRABS!HORAS = NUMHORAS
            RSTRABS!Dias = NUMDIAS
            RSTRABS.Update
            If OpPer(0).Value Then
               TOTPER = TOTPER + NUMDIAS
             Else:
               TOTPER = TOTPER + IIf(NUMHORAS = 0, NUMDIAS * 8, NUMHORAS)
            End If
            Me.Refresh
            RSTRABS.MoveNext
        Loop
        Set RSCNPT = Nothing
    Else
        RSTRABS.MoveFirst
        Dim NUMMESES As Long
        Do While Not RSTRABS.EOF
            Prog.Value = Prog.Value + 1
            XFEC = DevuelveValor("SELECT FECHAING FROM TRABAJADORES WHERE CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM)
            If IsNull(XFEC) Then
                MsgBox "EL TRABAJADOR " & RSTRABS!NOMBRES & " NO PRESENTA FECHA DE INGRESO, EL SISTEMA ABORTARÁ EL PROCESO", vbCritical
                Screen.MousePointer = 1
                Exit Sub
            End If
            If XFEC > xFechaIni.Value Then
                If Day(XFEC) <> 1 Then
                    NUMMESES = DateDiff("M", XFEC, xFechaFin.Value)
                    XFEC2 = DateAdd("D", -1, DateAdd("M", 1, CDate("01/" & Month(XFEC) & "/" & Year(XFEC))))
                    NUMDIAS = XFEC2 - XFEC
                Else
                    NUMMESES = DateDiff("M", XFEC, xFechaFin.Value) + 1
                    NUMDIAS = 0
                End If
            Else
                NUMMESES = DateDiff("M", xFechaIni.Value, xFechaFin.Value) + 1
                NUMDIAS = 0
            End If
            RSTRABS!HORAS = (NUMMESES * 30 * 8) + (NUMDIAS * 8)
            RSTRABS!Dias = (NUMMESES * 30) + NUMDIAS
            RSTRABS!FECHAING = XFEC
            RSTRABS.Update
            If OpPer(0).Value Then
               TOTPER = TOTPER + ((NUMMESES * 30) + NUMDIAS)
             Else:
               TOTPER = TOTPER + ((NUMMESES * 30 * 8) + (NUMDIAS * 8))
            End If
            RSTRABS.MoveNext
        Loop
    End If
    XPer.Caption = Format(TOTPER, "###,###,##0.00 ")
    xImpPer.Caption = Format((PARTDIST * 0.5) / TOTPER, "#0.0000000 ")
    xImpPer.Tag = (PARTDIST * 0.5) / TOTPER
    IMPPER = (PARTDIST * 0.5) / TOTPER
End Sub
Private Sub CALCULA()
    Dim VALOR As Single
    FLAG = False
    Dim RSCNPT As New ADODB.Recordset
    RSCNPT.Open "SELECT * FROM FORMULASUTIL ", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSCNPT.EOF Or RSCNPT.RecordCount = 0 Then
        MsgBox "MENSAJE DEL SISTEMA: EL SISTEMA NO HA ENCONTRADO FÓRMULAS DE UTILIDAD", vbInformation
        Set RSCNPT = Nothing
        Screen.MousePointer = 1
        Exit Sub
    End If
    xTRem.Caption = "0.00 ": xImpRem.Caption = "0.00 "
    Prog.Min = 0
    Prog.Max = Val(RSCNPT.RecordCount)
    Prog.Value = 0
    Prog.Visible = True
    xProg.Visible = True
    Do While Not RSCNPT.EOF
        Prog.Value = Prog.Value + 1
        xProg.Caption = "CALCULANDO " & RSCNPT!NOMBRE
        RSTRABS.MoveFirst
        Do While Not RSTRABS.EOF
            If InStr(RSCNPT!FORMULA, "@") = 0 Then
                VALOR = DevuelveValor("SELECT " & RSCNPT!FORMULA & " AS VALOR_DEV FROM VWTRABAJ WHERE CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM)
                If IsNull(VALOR) Then
                    VALOR = 0
                 Else
                    TOTREM = TOTREM + VALOR
                End If
            Else
                VALOR = DevuelveValor("SELECT " & CAMBIACADENA(RSCNPT!FORMULA, RSTRABS!CODTRAB) & " AS VALOR_DEV FROM VWTRABAJ WHERE CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM)
                TOTREM = TOTREM + VALOR
            End If
             If VALOR <> 0 Then
                VALOR = Round(VALOR, 2)
                DBSYSTEM.Execute "INSERT INTO  [##TMPCTS2" & VGL_COMPUTER & "]  VALUES ('" & RSTRABS!CODTRAB & "','" & RSCNPT!NOMBRE & "'," & VALOR & ")"
            End If
            RSTRABS!TOTREM = IIf(IsNull(RSTRABS!TOTREM), 0, RSTRABS!TOTREM) + VALOR
            RSTRABS.Update
            Me.Refresh
            RSTRABS.MoveNext
        Loop
        Me.Refresh
        RSCNPT.MoveNext
    Loop
    xProg.Visible = False
    Prog.Visible = False
    ENPROCESO = False
    RSTRABS.MoveFirst
    Screen.MousePointer = 1
    'CALCULANDO TOTALREM
    xTRem.Caption = Format(TOTREM, "###,###,##0.00 ")
    If TOTREM > 0 Then
        xImpRem.Caption = Format((PARTDIST * 0.5) / TOTREM, "#0.0000000 ")
        xImpRem.Tag = (PARTDIST * 0.5) / TOTREM
        IMPREM = (PARTDIST * 0.5) / TOTREM
    End If
    FLAG = True
End Sub
Private Sub CMCALCULAR_CLICK()
'On Error GoTo ERRCALC
    'If Not IsNumeric(IMPPER) Then Exit Sub
    xTotPer.Caption = "0.00 "
    xTotRemu.Caption = "0.00 "
    xTotal.Caption = "0.00 "
    ENPROCESO = True
    RSTRABS.MoveFirst
    Prog.Min = 0
    Prog.Max = Val(RSTRABS.RecordCount)
    Prog.Value = 0
    Prog.Visible = True
    xProg.Visible = True
    Do While Not RSTRABS.EOF
        Prog.Value = Prog.Value + 1
        xProg.Caption = "CALCULANDO UTILIDADES " & RSTRABS!NOMBRES
        If OpPer(0) Then
            RSTRABS!PARTPER = RSTRABS!Dias * IMPPER
         Else:
            RSTRABS!PARTPER = IIf(RSTRABS!HORAS = 0, RSTRABS!Dias * 8, RSTRABS!HORAS) * IMPPER
        End If
        RSTRABS!PARTREM = RSTRABS!TOTREM * IMPREM
        RSTRABS!TOTPART = RSTRABS!PARTPER + RSTRABS!PARTREM
        RSTRABS.Update
        RSTRABS.MoveNext
    Loop
    xProg.Visible = False
    Prog.Visible = False
    ENPROCESO = False
    RSTRABS.Requery
    'CALCULANDO TOTALES
    Dim RSTOT As New ADODB.Recordset
    RSTOT.Open "SELECT SUM(PARTPER) AS SUMPER,SUM(PARTREM) AS SUMREM,SUM(TOTPART) AS TOT FROM  [##TMPCTS1" & VGL_COMPUTER & "] ", DBSYSTEM
    If RSTOT.RecordCount <> 0 Then
        xTotPer.Caption = Format(RSTOT!SUMPER, "###,###,##0.00 ")
        xTotRemu.Caption = Format(RSTOT!SUMREM, "###,###,##0.00 ")
        xTotal.Caption = Format(RSTOT!TOT, "###,###,##0.00 ")
      Else:
        xTotPer.Caption = "0.00 "
        xTotRemu.Caption = "0.00 "
        xTotal.Caption = "0.00 "
    End If
    cmGrabar.Enabled = True
'    Exit Sub
'ERRCALC:
'    Exit Sub
End Sub

Private Sub CmdImprimir_Click()
    Dim SQL As String, TODOS As String
    Dim XDISTRIT As String
    If MsgBox("IMPRIMIR REGISTRO SELECCIONADO", vbYesNo + vbQuestion) = vbYes Then
        TODOS = "AND {PLANUTIL.CODTRAB}='" & RSTRABS!CODTRAB & "'"
      Else:
        TODOS = ""
    End If
    Screen.MousePointer = 11
    XDISTRIT = IIf(IsNull(DevuelveValor("SELECT DISTRITO FROM EMPRESA", DBSYSTEM)), " ", DevuelveValor("SELECT DISTRITO FROM EMPRESA", DBSYSTEM))
    With Reporte
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "\PLAN0059.RPT"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .Destination = crptToWindow
        .SelectionFormula = "{PLANUTIL.CODIGO}=" & VPTRASPRM & TODOS
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .WindowTitle = .ReportFileName
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XDISTRITO='" & XDISTRIT & "'"
        .Formulas(2) = "XRUC='" & REGSISTEMA.RUC & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub
Private Sub CREATRABSEL(OPC As Integer)
   'CREAR TABLA TEMPORAL DE LOS TRABAJADORES QUE SE HAN SELECCIONADO
    Dim RSTEMP As New ADODB.Recordset
    If ExisteTablaAux("SELCODTRAB") Then DBSYSTEM.Execute "DROP TABLE SELCODTRAB"
    DBSYSTEM.Execute "CREATE TABLE SELCODTRAB (CODTRAB VARCHAR(8) CONSTRAINT CLAVE PRIMARY KEY )"
    RSTEMP.Open "SELCODTRAB", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Select Case OPC
        Case 0:
            Dim NUM As Variant
            For Each NUM In xData.SelBookmarks
                xData.Bookmark = NUM
                xData.COL = 0
                RSTEMP.AddNew
                RSTEMP!CODTRAB = Trim(xData.Text)
                RSTEMP.Update
            Next
        Case 1:
            DBSYSTEM.Execute "INSERT INTO SELCODTRAB SELECT CODTRAB TMPCTS1 "
    End Select
End Sub

Private Sub CMGRABAR_CLICK()
    Dim xCodigo As Long, xSoles As Double
    If MsgBox("SEGURO DE GRABAR LOS CAMBIOS EN LA PLANILLA DE UTILIDAD", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    If RSTRABS.RecordCount = 0 Then
        MsgBox "NO EXISTE NADA POR GRABAR", vbInformation
        Exit Sub
    End If
    If VPTAREA = "NUEVO" Then
        If Trim(xPeriodo.Text) = "" Then
            MsgBox "FALTA ESPECIFICAR UN NOMBRE DESCRIPTIVO DE CÁLCULO PARA LA C.T.S.", vbInformation
            xPeriodo.SetFocus
            Exit Sub
        End If
    Else
        'SI ES MODIFICAR
        xCodigo = Val(VPTRASPRM)
        DBSYSTEM.Execute "DELETE FROM UTIL WHERE CODIGO=" & xCodigo
        DBSYSTEM.Execute "DELETE FROM PLANUTIL WHERE CODIGO=" & xCodigo
        DBSYSTEM.Execute "DELETE FROM DETALLEUTIL WHERE CODIGO=" & xCodigo
    End If
    
    DBSYSTEM.Execute _
    "INSERT INTO UTIL (NOMBRE, CERRADO, FECHAINI, FECHAFIN,CALPER,DIAEFECT,PORPART,UTILIDAD,PARTDIST,TOTPER,IMPXPER,TOTREM,IMPXREM) " & _
    "VALUES ('" & xPeriodo.Text & "',0," & DateSQL(xFechaIni.Value) & "," & DateSQL(xFechaFin.Value) & "," & Str(IIf(OpPer(0), -1, 0)) & "," & Str(IIf(Check1.Value = 1, -1, 0)) & "," & _
    Str(Val(xPorc.Text)) & "," & XUtil.Text & "," & Str(PARTDIST) & "," & Str(TOTPER) & "," & Str(IMPPER) & "," & Str(TOTREM) & "," & Str(IMPREM) & ")"
    
    xCodigo = DevuelveValor("SELECT MAX(CODIGO) AS COD1 FROM UTIL", DBSYSTEM)
    DBSYSTEM.Execute "INSERT INTO PLANUTIL (CODIGO, CODTRAB, NOMBRES, PARTPER,PARTREM,TOTPART,DIAS,HORAS,TOTREM,FECHAING) SELECT " & xCodigo & " AS CODIGO, CODTRAB,NOMBRES,PARTPER,PARTREM,TOTPART,DIAS,HORAS,TOTREM,FECHAING FROM TMPCTS1 IN '" & App.PATH & "\BDAUXCOM.MDB" & "'"
    DBSYSTEM.Execute "INSERT INTO DETALLEUTIL SELECT " & xCodigo & " AS CODIGO, CODTRAB, CONCEPTO, IMPORTE FROM TMPCTS2  IN '" & App.PATH & "\BDAUXCOM.MDB" & "' WHERE IMPORTE<>0"
    MsgBox "INFORMACIÓN GRABADA SATISFACTORIAMENTE", vbInformation
    Unload Me
End Sub

Private Sub CMSELECTRAB_CLICK()
 On Error GoTo ERRCALC
    If xFechaIni.Value = xFechaFin.Value Then
        MsgBox "FECHAS NO SON VÁLIDAS", vbInformation
        Exit Sub
    End If
    If xFechaFin.Value < xFechaIni.Value Then
        MsgBox "LA FECHA DE INICIO NO PUEDE SER MENOR QUE LA FECHA DE INICIO", vbExclamation
        Exit Sub
    End If
    If PARTDIST = 0 Then
        MsgBox "TIENE QUE CALCULAR LA PARTICIPACION DE UTILIDADES", vbExclamation
        xPorc.SetFocus
        Exit Sub
    End If
    frSelect.Show 1
    Dim RSTMSEL As New ADODB.Recordset
    RSTMSEL.Open " [##TMPSELECT" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset
    If RSTMSEL.RecordCount = 0 Then Exit Sub
    If MsgBox("DESEA ASIGNAR LOS TRABAJADORES AL CÁLCULO DE UTILIDADES", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    DBSYSTEM.Execute "DELETE FROM  [##TMPCTS1" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "DELETE FROM  [##TMPCTS2" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "INSERT INTO  [##TMPCTS1" & VGL_COMPUTER & "]  (CODTRAB,NOMBRES) SELECT CODTRAB, LTRIM(RTRIM(NOMBRES)) FROM  [##TMPSELECT" & VGL_COMPUTER & "] "
    xTotPer.Caption = "0.00 "
    xTotRemu.Caption = "0.00 "
    xTotal.Caption = "0.00 "
    Set RSTRABS = Nothing
    Set RSTRABS = New ADODB.Recordset
    RSTRABS.Open " [##TMPCTS1" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RSTRABS
    xNumTrabs.Caption = RSTRABS.RecordCount
    TOTPER = 0: IMPPER = 0
    TOTREM = 0: IMPREM = 0
    DIASHORAS
    CALCULA
ERRCALC:
    Exit Sub
    Screen.MousePointer = 1
End Sub

Private Sub Command1_Click()
    With Reporte
        .Reset
        .WindowTitle = "LISTADO DE CONCEPTOS AFECTOS A CTS"
        .ReportFileName = REGSISTEMA.REPORTES & "\PLAN0047.RPT"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        If .Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    PARTDIST = 0
    TOTPER = 0: IMPPER = 0
    TOTREM = 0: IMPREM = 0
    xFechaFin.Value = Date
    xFechaFin.Day = 1
    xFechaIni.Value = xFechaFin.Value
    xFechaIni.Value = DateAdd("M", -5, xFechaFin.Value)
    xFechaIni.Day = 1
    xPeriodo.Text = AMESES(xFechaIni.Month) & "." & Year(xFechaIni.Value) & " A " & AMESES(xFechaFin.Month) & "." & Year(xFechaFin.Value)
    If ExisteTablaAux(" [##TMPCTS1" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCTS1" & VGL_COMPUTER & "] "
    'CABECERAS
    DBSYSTEM.Execute "CREATE TABLE  [##TMPCTS1" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(100), PARTPER  Numeric(20,2) ,PARTREM  Numeric(20,2) ,TOTPART  Numeric(20,2) , DIAS INT,HORAS INT, TOTREM  Numeric(20,2) , FECHAING DATETIME)"
    'DETALLES
    If ExisteTablaAux(" [##TMPCTS2" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCTS2" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPCTS2" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), CONCEPTO VARCHAR(80), IMPORTE  Numeric(20,2) )"
    Select Case UCase(VPTAREA)
        Case "NUEVO"
            Me.Caption = "NUEVO CÁLCULO DE UTILIDADES"
            Frame2.Enabled = True
            CMDIMPRIMIR.Visible = False
            OpPer(0) = 1
            Check1.Value = 1
        Case "MODIFICAR"
            Me.Caption = "MODIFICACIÓN DEL CÁLCULO DE UTILIDADES"
            Frame1.Enabled = False
            Frame2.Enabled = True
            CARGADATOS
            cmGrabar.Enabled = False
            CMDIMPRIMIR.Visible = False
        Case "VISTA"
            Frame1.Enabled = False
            Frame2.Enabled = False
            xDetalle.AllowUpdate = False
            xMeses.Locked = True
            xDias.Locked = True
            cmSelecTrab.Enabled = False
            cmCalcular.Visible = False
            CMDIMPRIMIR.Visible = True
            Me.Caption = "CONSULTA DEL CÁLCULO DE UTILIDADES"
            CARGADATOS
        Case "PRUEBA"
            CMDIMPRIMIR.Visible = False
            cmGrabar.Visible = False
    End Select
    Set RSTRABS = New ADODB.Recordset
    RSTRABS.Open " [##TMPCTS1" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RSTRABS
    XDATA_ROWCOLCHANGE 0, 0
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSTRABS = Nothing
    Set RSCALC = Nothing
End Sub

Private Sub OPPER_Click(INDEX As Integer)
    Select Case INDEX
        Case 0:
            Label16.Caption = "TOTAL DE DIAS TRAB"
            Label17.Caption = "IMP. X CADA DIA"
        Case 1
            Label16.Caption = "TOTAL DE HORAS TRAB"
            Label17.Caption = "IMP. X CADA HORA"
    End Select
End Sub
Public Function CAMBIACADENA(ByVal CADENA As String, ByVal CODTRAB As String) As String
    Dim POSARROBA As Integer, POS1 As Integer, PROCESO As String, CAMPO As String, POS2 As Integer
    Dim VALOR As Double
    POSARROBA = 1
    POSARROBA = InStr(POSARROBA, CADENA, "@")
    Do While POSARROBA <> 0
        POS1 = InStr(POSARROBA, CADENA, "(")
        PROCESO = Mid(CADENA, POSARROBA + 1, POS1 - (POSARROBA + 1))
        POS2 = InStr(POSARROBA, CADENA, ")")
        CAMPO = Mid(CADENA, POS1 + 1, POS2 - (POS1 + 1))
'        IF (DEVUELVEVALOR("SELECT CODIGO FROM CONCEPTOS WHERE CODIGO='" & CAMPO & "'", DBSYSTEM)) = "" THEN
'            MSGBOX "EL CONCEPTO DE REMUNERACIÓN: " & CAMPO & " DE LA FÓRMULA DE CTS: " & CADENA & " NO EXISTE", VBINFORMATION, "ERROR DE CONFIGURACIÓN"
'            CAMBIACADENA = "0"
'            EXIT FUNCTION
'        END IF
        Select Case UCase(PROCESO)
            Case "PROMEDIO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, PROMEDIO, CAMPO, False)
            Case "ULTIMOVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, ULTIMOVALOR, CAMPO, False)
            Case "PRIMERVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, PRIMERVALOR, CAMPO, False)
            Case "SUMA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, SUMA, CAMPO, False)
            Case "MEDIA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, MEDIA, CAMPO, False)
            Case "PROMEDIOVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, PROMEDIOVALOR, CAMPO, False)
            Case "PRIMERO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, PRIMERO, CAMPO, False)
            Case "ULTIMO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, ULTIMO, CAMPO, False)
            Case "MAYORVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, MAYORVALOR, CAMPO, False)
            Case "MENORVALOR"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, MENORVALOR, CAMPO, False)
            Case "NUMERO"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, Numero, CAMPO, False)
            Case "NSECUENCIA"
                VALOR = CALCULOCONCEPTOS(CODTRAB, xFechaIni.Value, xFechaFin.Value, NSECUENCIA, CAMPO, False)
        End Select
        If IsNull(VALOR) Then VALOR = 0
        CADENA = Replace(CADENA, Mid(CADENA, POSARROBA, (POS2 - POSARROBA) + 1), "" & VALOR)
        POSARROBA = InStr(POSARROBA, CADENA, "@")
    Loop
    CAMBIACADENA = CADENA
End Function

Public Sub ACTUALIZAUTILIDADES()
    
End Sub



Private Sub XDATA_HEADCLICK(ByVal COLINDEX As Integer)
    RSTRABS.Sort = xData.Columns(COLINDEX).DataField
End Sub

Private Sub XDATA_ROWCOLCHANGE(LASTROW As Variant, ByVal LASTCOL As Integer)
    If Not FLAG Then Exit Sub
    If RSTRABS.EOF Or RSTRABS.RecordCount = 0 Then
        cmCalcular.Enabled = False
        xMeses.Text = 0
        xDias.Text = 0
        Exit Sub
       Else
        cmCalcular.Enabled = True
    End If
    If ENPROCESO Then Exit Sub
    xMeses.Text = IIf(IsNull(RSTRABS!Dias), 0, RSTRABS!Dias)
    If OpPer(1) Then
        xDias.Text = IIf(IsNull(RSTRABS!HORAS), 0, RSTRABS!HORAS)
      Else:
        xDias.Text = IIf(IsNull(RSTRABS!HORAS), 0, RSTRABS!Dias * 8)
    End If
    xFecha.Caption = IIf(IsNull(RSTRABS!FECHAING), "", RSTRABS!FECHAING)
    xTotRem.Caption = Format(IIf(IsNull(RSTRABS!TOTREM), "", RSTRABS!TOTREM), "###,###,##0.00 ")
    Set RSCALC = Nothing
    RSCALC.Open "SELECT * FROM  [##TMPCTS2" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xDetalle.DataSource = RSCALC
End Sub

Private Sub XDETALLE_AFTERCOLUPDATE(ByVal COLINDEX As Integer)
    RSCALC.MOVE 0
End Sub

Public Sub CARGADATOS()
    Dim DIAEF As Integer, OP As Integer
    Dim VAL1 As Double, VAL2 As Double, VAL3 As Double
    Dim VAL4 As Double, VAL5 As Double, VAL6 As Double
    Dim VAL7 As Double
    xPeriodo.Text = DevuelveValor("SELECT NOMBRE FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    xFechaIni.Value = DevuelveValor("SELECT FECHAINI FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    xFechaFin.Value = DevuelveValor("SELECT FECHAFIN FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    DIAEF = DevuelveValor("SELECT DIAEFECT FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    Check1.Value = IIf(DIAEF = 0, 0, 1)
    OP = DevuelveValor("SELECT CALPER FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    If OP = 0 Then
        OpPer(1) = True
      Else:
        OpPer(0) = True
    End If
    VAL1 = DevuelveValor("SELECT UTILIDAD FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    VAL2 = DevuelveValor("SELECT PORPART FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    
    XUtil.Text = VAL1
    xPorc.Text = Format(VAL2, "##0")
    VAL3 = VAL1 * (VAL2 / 100)
    xPartDist.Caption = Format(VAL3, "###,###,##0.00 ")
    xPartPer.Caption = Format(VAL3 * 0.5, "###,###,##0.00 ")
    XpartRem.Caption = Format(VAL3 * 0.5, "###,###,##0.00 ")
    VAL4 = DevuelveValor("SELECT TOTPER FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    XPer.Caption = Format(VAL4, "###,###,##0.00 ")
    VAL5 = DevuelveValor("SELECT IMPXPER FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    xImpPer.Caption = Format(VAL5, "##0.0000000 ")
    VAL6 = DevuelveValor("SELECT TOTREM FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    xTRem.Caption = Format(VAL6, "###,###,##0.00 ")
    VAL7 = DevuelveValor("SELECT IMPXREM FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    xImpRem.Caption = Format(VAL7, "##0.0000000 ")
    xTotPer.Caption = xPartPer.Caption
    xTotRemu.Caption = xPartPer.Caption
    xTotal.Caption = xPartDist.Caption
    PARTDIST = DevuelveValor("SELECT PARTDIST FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    TOTPER = DevuelveValor("SELECT TOTPER FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    IMPPER = DevuelveValor("SELECT IMPXPER FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    TOTREM = DevuelveValor("SELECT TOTREM FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    IMPREM = DevuelveValor("SELECT IMPXREM FROM UTIL WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    Dim X As Integer
    DBSYSTEM.Execute "INSERT INTO TMPCTS1 SELECT CODTRAB, NOMBRES, PARTPER,PARTREM,TOTPART,DIAS,HORAS,TOTREM,FECHAING FROM PLANUTIL IN '" & REGSISTEMA.PATHEMPRESA & "\PLANILLA.MDB' WHERE CODIGO=" & VPTRASPRM, X
    DBSYSTEM.Execute "INSERT INTO TMPCTS2 SELECT CODTRAB, CONCEPTO, IMPORTE FROM DETALLEUTIL IN '" & REGSISTEMA.PATHEMPRESA & "\PLANILLA.MDB' WHERE CODIGO=" & VPTRASPRM
    xNumTrabs.Caption = Format(X, "##0 ")
End Sub

Private Sub XPARTDIST_CHANGE()
    xPartPer.Caption = Format(PARTDIST * 0.5, "###,###,##0.00 ")
    XpartRem.Caption = Format(PARTDIST * 0.5, "###,###,##0.00 ")
End Sub

Private Sub XPORC_CHANGE()
    PARTDIST = Val(XUtil.Text) * (Val(xPorc.Text) / 100)
    xPartDist.Caption = Format(PARTDIST, "###,###,##0.00 ")
End Sub

Private Sub XUTIL_CHANGE()
    PARTDIST = Val(XUtil.Text) * (Val(xPorc.Text) / 100)
    xPartDist.Caption = Format(PARTDIST, "###,###,##0.00 ")
End Sub



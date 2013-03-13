VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frPagoCtaCteLiq 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cancelación de Cta. Cte. Egresos por Liquidaciones"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   4770
      TabIndex        =   4
      Top             =   4140
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   330
      Left            =   3420
      TabIndex        =   3
      Top             =   4140
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3435
      Left            =   120
      TabIndex        =   0
      Top             =   645
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6059
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   14477281
      HeadLines       =   2
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Listado de Deudas del Trabajador"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "CODMOV"
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
         DataField       =   "Descripcion"
         Caption         =   "Descripción"
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
         DataField       =   "SALDO"
         Caption         =   "Saldo Actual"
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
      BeginProperty Column03 
         DataField       =   "Monto"
         Caption         =   "Monto a Descontar"
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
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2670.236
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   285
      Left            =   1050
      TabIndex        =   2
      Top             =   225
      Width           =   4965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Trabajador"
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   255
      Width           =   765
   End
End
Attribute VB_Name = "frPagoCtaCteLiq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsDesc As New ADODB.Recordset
Private Sub COMMAND1_CLICK()
    VPTAREA = "Si"
    Unload Me
End Sub
Private Sub Command2_Click()
    VPTAREA = "No"
    Unload Me
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    RsDesc.Sort = DataGrid1.Columns(ColIndex).DataField
End Sub
Private Sub Form_Load()
    Label2.Caption = frLiquidacion.xTrab.Text
    RsDesc.Open " [##TMPLIQCTA" & VGL_COMPUTER & "] ", DBAUXCOM, adOpenKeyset, adLockOptimistic
    Set DataGrid1.DataSource = RsDesc
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RsDesc = Nothing
End Sub

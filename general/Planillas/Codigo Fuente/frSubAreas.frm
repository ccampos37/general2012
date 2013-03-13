VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frSubAreas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centros de Tareo"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frSubAreas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   3720
      TabIndex        =   1
      Top             =   4065
      Width           =   1425
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   3780
      Left            =   180
      TabIndex        =   0
      Top             =   165
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   6668
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      Caption         =   "Centros de Tareo"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Código"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3105.071
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frSubAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsAreas As New ADODB.Recordset

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub FORM_LOAD()
    RsAreas.Open "SUBAREAS", DBSYSTEM, adOpenKeyset, adLockOptimistic
    REFRESCAR
End Sub

Public Sub REFRESCAR()
    RsAreas.Requery
    With xData
        Set .DataSource = RsAreas
    End With
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RsAreas = Nothing
End Sub

Private Sub XDATA_HEADCLICK(ByVal ColIndex As Integer)
    RsAreas.Sort = xData.Columns(ColIndex).DataField
End Sub


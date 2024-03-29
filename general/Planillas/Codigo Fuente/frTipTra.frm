VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frTipTra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Trabajadores"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
   Icon            =   "frTipTra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   3810
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frTipTra.frx":08CA
      Height          =   4290
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7567
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
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
      Caption         =   "Tipos de Trabajadores"
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frTipTra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsTipos As New ADODB.Recordset
Dim REGACT As RegWin

Private Sub FORM_ACTIVATE()
    ActivarTools REGACT
End Sub

Private Sub FORM_LOAD()
    With REGACT
        .Buscar = False
        .Editar = False
        .ELIMINAR = False
        .Filtrar = False
        .IMPRIMIR = False
        .NUEVO = False
        .Preliminar = False
    End With
    RsTipos.Open "TIPOSTRAB", DbSystem, adOpenKeyset, adLockOptimistic
    Set DataGrid1.DataSource = RsTipos
End Sub

Private Sub FORM_UNLOAD(Cancel As Integer)
    Set RsTipos = Nothing
End Sub

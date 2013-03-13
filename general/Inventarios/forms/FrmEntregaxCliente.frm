VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmEntregaxCliente 
   Caption         =   "Clientes de destino"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5415
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   10455
      Begin VB.Frame Frame2 
         Height          =   2055
         Left            =   6840
         TabIndex        =   3
         Top             =   1920
         Width           =   3495
         Begin VB.Frame frmbotones 
            Height          =   1170
            Left            =   480
            TabIndex        =   5
            Top             =   720
            Width           =   2610
            Begin VB.CommandButton CmdSalir 
               Caption         =   "&Salir"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   1320
               Picture         =   "FrmEntregaxCliente.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton cmdAceptar 
               Caption         =   "&Aceptar"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               Picture         =   "FrmEntregaxCliente.frx":0442
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Permute Cliente de destino"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   4
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   1
         Top             =   360
         Width           =   4455
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3495
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6165
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
      Begin VB.Label Label2 
         Caption         =   "Dar Click en la Grilla para activar o desactivar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo  - Razon social"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmEntregaxCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim RSQL As New ADODB.Recordset
Dim nLongicampo(3) As Integer


Private Sub cmdBotones_Click(Index As Integer)

End Sub

Private Sub cmdAceptar_Click()
 If MsgBox(" Es correcto Si/No ", vbYesNo, "Confirmacion") = vbYes Then
    VGCNx.Execute (" update vt_cliente set clienteguiasterceros=" & Check1.Value & "  where clientecodigo='" & RSQL!clientecodigo & "'")
  Else
    frmbotones.Visible = False
End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click()
 Check1.Value = RSQL!clienteguiasterceros
 frmbotones.Visible = True
End Sub

Private Sub Form_Load()
  frmbotones.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

SQL = " select clientecodigo, clienterazonsocial,clienteguiasterceros  from vt_cliente"
Text1.text = UCase$(RTrim(Text1.text))
If Len(Text1) > 0 Then
Text1.text = UCase$(Text1.text)
If IsNumeric(Text1.text) Then
   SQL = SQL & " where clientecodigo='" & RTrim(Text1.text) & "'"
 Else
   SQL = SQL & " where clienterazonsocial like '%" & RTrim(Text1.text) & "%'"

End If
Set RSQL = VGCNx.Execute(SQL)
Set DataGrid1.DataSource = RSQL
With DataGrid1
      .Columns(0).Caption = "Codigo"
      .Columns(0).Width = 700
      .Columns(1).Caption = "Razon Social"
      .Columns(1).Width = 3800
      .MarqueeStyle = dbgHighlightRow
      .Refresh
  End With
DataGrid1.Refresh
Text1.SetFocus
End If
End Sub



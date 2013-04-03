VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FrmSunat682 
   Caption         =   "Formato 682"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   11655
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2415
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4260
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
               LCID            =   3082
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
               LCID            =   3082
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
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   6615
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1920
         TabIndex        =   11
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Carpeta y Nombre del Archivo a Exportar"
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   6840
      TabIndex        =   5
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton Transf 
         Caption         =   "Exportar a SUNAT"
         Height          =   975
         Left            =   1680
         Picture         =   "FrmSunat682.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Genera Archivo"
         Height          =   975
         Left            =   240
         Picture         =   "FrmSunat682.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Niagara Engraved"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3240
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmSunat682.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Height          =   720
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   6570
      Begin VB.ComboBox cmbNivel 
         Height          =   315
         Left            =   1035
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   225
         Width           =   1920
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   4290
         TabIndex        =   2
         Top             =   255
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MM - MMMM"
         Format          =   109510659
         UpDown          =   -1  'True
         CurrentDate     =   37505
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   3360
         TabIndex        =   4
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nivel :"
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   270
         Width           =   450
      End
   End
End
Attribute VB_Name = "FrmSunat682"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSAUX As New ADODB.Recordset
Dim lforma As Integer

Private Sub cmdCancelar_Click()
    Unload Me
End Sub



Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Call generar
End Sub

Private Sub Form_Load()
    Call CargaNivel
    DTPicker1.Value = DateSerial(CInt(VGParamSistem.Anoproceso), CInt(VGParamSistem.Mesproceso), 1)
End Sub
Private Sub CargaNivel()
    Dim i As Integer
    For i = 1 To VGnumnivelescuenta
        cmbNivel.AddItem Format(i, "0")
    Next
    cmbNivel.ListIndex = 0
End Sub
Private Sub Exportar(rs As Recordset)
Dim li_Arc As Integer, TotReg As Integer
Dim sumactas As Double
Dim xcuenta, xmoneda, xmonto, xfecha, xtotreg, xblanco15, xblanco20, xblanco40 As String
Dim xregistro1 As String
Dim sumas As String
sumas = "0"
xmonto = 0
xsumactas = 0
Me.MousePointer = 11
Me.MousePointer = vbHourglass

li_Arc = FreeFile

Open "C:\temp\" & RTrim(Text1.Text) & ".txt" For Output As #li_Arc
rs.MoveFirst
Do While Not rs.EOF
   With rs
        xregistro1 = rs!cuentacodigo & "|" & rs!saldoinidebe & "|" & rs!saldoinihaber & "|"
        xregistro1 = xregistro1 & "" & rs!Movacumdebe & "|" & rs!movacumhaber & "|"
        xregistro1 = xregistro1 & "" & sumas & "|" & sumas & "|"
   End With
   Print #li_Arc, xregistro1
    rs.MoveNext
Loop
rs.Close
Close #li_Arc
Me.MousePointer = vbDefault
Set rs = Nothing
Me.MousePointer = 0
MsgBox "Se ha generado el archivo c:\temp\" & Text1.Text & ".txt  satisfactoriamente ", vbInformation, "Mensaje"
Exit Sub
Error_PDT:
End Sub

Private Sub Transf_Click()
If Text1.Text = "" Then
   MsgBox ("Ingrese nombre del Archivo a transferir ")
   Text1.SetFocus
   Exit Sub
End If
Call Exportar(RSAUX)
End Sub
Private Sub generar()
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    'Elimar los Detalle antes de grabar
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_pro_GeneraSunat682"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@anno") = CInt(VGParamSistem.Anoproceso)
        .Parameters("@nivel") = cmbNivel.ListIndex + 1
        Set RSAUX = .Execute
    End With
If RSAUX.RecordCount > 0 Then
   Set DataGrid1.DataSource = RSAUX
   DataGrid1.Refresh
End If
End Sub

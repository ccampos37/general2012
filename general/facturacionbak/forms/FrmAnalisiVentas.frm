VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAnalisisVentas 
   Caption         =   "Analisis de ventas"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form2"
   ScaleHeight     =   9060
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1455
      Left            =   6240
      TabIndex        =   11
      Top             =   120
      Width           =   2535
      Begin MSComCtl2.DTPicker DTPdesde 
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   41177
      End
      Begin MSComCtl2.DTPicker DTPhasta 
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   41177
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   7800
      Width           =   8535
      Begin VB.CommandButton Command1 
         Caption         =   "Exportar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5280
         Picture         =   "FrmAnalisiVentas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6960
         Picture         =   "FrmAnalisiVentas.frx":047A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Cmdprocesa 
         Caption         =   "Procesar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3480
         Picture         =   "FrmAnalisiVentas.frx":0AC1
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1425
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   5805
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuPuntovta 
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "vt_puntoventa"
         TituloAyuda     =   "Punto de ventas"
         ListaCampos     =   "puntovtacodigo(1),puntovtadescripcion(1)"
         XcodCampo       =   "puntovtacodigo"
         XListCampo      =   "puntovtadescripcion"
         ListaCamposDescrip=   "codigo, descripcion"
         ListaCamposText =   "puntovtacodigo,puntovtadescripcion"
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "Punto de venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   2400
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   8535
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4815
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8493
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
      Begin VB.Label Label22 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   7320
         TabIndex        =   3
         Top             =   5400
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Cantidad de registros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   2
         Top             =   5400
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FrmAnalisisVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsmes As New ADODB.Recordset

Private Sub Cmdprocesa_Click()
On Error GoTo error1
If Ctr_AyuPuntovta.xclave = "" Then
   MsgBox (" Ingrese el punto de venta ")
   Exit Sub
End If
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "vt_analisisVentas_rpt"
    VGCommandoSP.Parameters.Refresh
With VGCommandoSP
     .Parameters("@base") = VGParamSistem.BDEmpresa
     .Parameters("@empresa") = VGParametros.empresacodigo
     .Parameters("@puntovta") = Ctr_AyuPuntovta.xclave
     .Parameters("@desde") = DTPdesde
     .Parameters("@hasta") = DTPhasta
     .Parameters("@tipo") = IIf(DTPhasta.Value = Date, 0, 1)
    Set rsmes = .Execute
End With
' Set rsmes = VGCNx.Execute("select * from " & RTrim(VGcomputer) & "_ana1")
If rsmes.RecordCount = 0 Then Exit Sub
Frame2.Visible = True
Set DataGrid1.DataSource = rsmes
DataGrid1.Refresh
Label22.Caption = rsmes.RecordCount
Exit Sub
error1:

Exit Sub
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim TITULO As String
Set DataGrid1.DataSource = Nothing
TITULO = "ANALISIS DE PUNTO DE VENTA " & RTrim(Ctr_AyuPuntovta.xnombre) & " DESDE " & DTPdesde.Value & " AL " & DTPhasta.Value & ""
Call exportarExcel(rsmes, TITULO)
Set DataGrid1.DataSource = rsmes
DataGrid1.Refresh
End Sub

Private Sub Form_Load()
Ctr_AyuPuntovta.conexion VGCNx
DTPdesde = VGParamSistem.FechaTrabajo
DTPhasta = VGParamSistem.FechaTrabajo

End Sub



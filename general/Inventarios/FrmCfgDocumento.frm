VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCfgDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Documentos"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "FrmCfgDocumento.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7920
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Reporte"
      Height          =   675
      Left            =   5400
      Picture         =   "FrmCfgDocumento.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   5400
      Width           =   775
   End
   Begin VB.CommandButton CmdCon 
      Caption         =   "&Consulta"
      Height          =   675
      Left            =   4380
      Picture         =   "FrmCfgDocumento.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5385
      Width           =   775
   End
   Begin VB.CommandButton CmdEli2 
      Caption         =   "&Eliminar"
      Height          =   675
      Left            =   3180
      Picture         =   "FrmCfgDocumento.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5385
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   6420
      Picture         =   "FrmCfgDocumento.frx":1590
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5385
      Width           =   775
   End
   Begin VB.CommandButton CmdIng 
      Caption         =   "&Ingreso"
      Height          =   675
      Left            =   1035
      Picture         =   "FrmCfgDocumento.frx":19D2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5385
      Width           =   775
   End
   Begin VB.CommandButton CmdModi 
      Caption         =   "&Modificar"
      Height          =   675
      Left            =   2100
      Picture         =   "FrmCfgDocumento.frx":1E14
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5385
      Width           =   775
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   675
      Left            =   2865
      Picture         =   "FrmCfgDocumento.frx":2256
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5385
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir2 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   4170
      Picture         =   "FrmCfgDocumento.frx":2698
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5385
      Visible         =   0   'False
      Width           =   775
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmCfgDocumento.frx":2ADA
      Height          =   4800
      Left            =   240
      TabIndex        =   16
      Top             =   360
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   8467
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "CtnCodigo"
         Caption         =   "TD"
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
         DataField       =   "TDO_CODSUN"
         Caption         =   "Sunat"
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
         DataField       =   "TDO_DESCRI"
         Caption         =   "                Descripción"
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
      BeginProperty Column03 
         DataField       =   "CtnNumser"
         Caption         =   "Serie"
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
      BeginProperty Column04 
         DataField       =   "CtnNumero"
         Caption         =   "Nro. Actual"
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
         ScrollBars      =   2
         BeginProperty Column00 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   3644.788
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1335.118
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5145
      Left            =   120
      TabIndex        =   18
      Top             =   90
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1935
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1935
         MaxLength       =   2
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   690
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1935
         MaxLength       =   30
         TabIndex        =   20
         Text            =   "Text3"
         Top             =   1020
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   5745
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "Text4"
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1935
         MaxLength       =   7
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   1350
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   5745
         MaxLength       =   7
         TabIndex        =   3
         Text            =   "Text6"
         Top             =   1350
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1935
         MaxLength       =   7
         TabIndex        =   4
         Text            =   "Text7"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5745
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "Text8"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "Text9"
         Top             =   2010
         Width           =   1110
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   5745
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "Text10"
         Top             =   2010
         Width           =   630
      End
      Begin VB.TextBox TxImpresora 
         Height          =   285
         Left            =   1935
         MaxLength       =   30
         TabIndex        =   7
         Text            =   "Text11"
         Top             =   2340
         Width           =   1440
      End
      Begin VB.TextBox TxControlador 
         Height          =   285
         Left            =   5745
         MaxLength       =   30
         TabIndex        =   8
         Text            =   "Text11"
         Top             =   2340
         Width           =   1395
      End
      Begin VB.TextBox TxPuerto 
         Height          =   285
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "Text11"
         Top             =   2685
         Width           =   1425
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   1935
         Left            =   285
         TabIndex        =   22
         Top             =   3060
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3413
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Cabecera"
         TabPicture(0)   =   "FrmCfgDocumento.frx":2AEF
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label14"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label13"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label4"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label21"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label22"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label23"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label24"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label25"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label26"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label27"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "ChCambio"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "ChSerie"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "ChPto"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "ChFecha"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "ChRazon"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "ChOrden"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "ChPedido"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "ChRuc"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "ChDireccion"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "ChCotizacion"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).ControlCount=   20
         TabCaption(1)   =   "Detalle"
         TabPicture(1)   =   "FrmCfgDocumento.frx":2B0B
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label32"
         Tab(1).Control(1)=   "Label31"
         Tab(1).Control(2)=   "Label12"
         Tab(1).Control(3)=   "Label15"
         Tab(1).Control(4)=   "ChModo"
         Tab(1).Control(5)=   "Chalmacen"
         Tab(1).Control(6)=   "ChStock"
         Tab(1).Control(7)=   "ChControl"
         Tab(1).ControlCount=   8
         Begin VB.CheckBox ChModo 
            Height          =   255
            Left            =   -69375
            TabIndex        =   36
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox Chalmacen 
            Height          =   255
            Left            =   -72900
            TabIndex        =   35
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox ChCotizacion 
            Height          =   255
            Left            =   5085
            TabIndex        =   34
            Top             =   1260
            Width           =   855
         End
         Begin VB.CheckBox ChDireccion 
            Height          =   255
            Left            =   5085
            TabIndex        =   33
            Top             =   750
            Width           =   855
         End
         Begin VB.CheckBox ChRuc 
            Height          =   255
            Left            =   2040
            TabIndex        =   32
            Top             =   1005
            Width           =   855
         End
         Begin VB.CheckBox ChPedido 
            Height          =   255
            Left            =   5085
            TabIndex        =   31
            Top             =   1530
            Width           =   855
         End
         Begin VB.CheckBox ChOrden 
            Height          =   255
            Left            =   2040
            TabIndex        =   30
            Top             =   1530
            Width           =   855
         End
         Begin VB.CheckBox ChRazon 
            Height          =   255
            Left            =   2040
            TabIndex        =   29
            Top             =   750
            Width           =   855
         End
         Begin VB.CheckBox ChFecha 
            Height          =   255
            Left            =   5085
            TabIndex        =   28
            Top             =   1005
            Width           =   855
         End
         Begin VB.CheckBox ChStock 
            Height          =   255
            Left            =   -72900
            TabIndex        =   27
            Top             =   810
            Width           =   855
         End
         Begin VB.CheckBox ChPto 
            Height          =   195
            Left            =   2040
            TabIndex        =   26
            Top             =   510
            Width           =   1005
         End
         Begin VB.CheckBox ChSerie 
            Height          =   225
            Left            =   5085
            TabIndex        =   25
            Top             =   480
            Width           =   1665
         End
         Begin VB.CheckBox ChCambio 
            Height          =   255
            Left            =   2040
            TabIndex        =   24
            Top             =   1260
            Width           =   1215
         End
         Begin VB.CheckBox ChControl 
            Height          =   255
            Left            =   -69375
            TabIndex        =   23
            Top             =   810
            Width           =   1020
         End
         Begin VB.Label Label32 
            Caption         =   "Almacen                       :"
            Height          =   255
            Left            =   -74700
            TabIndex        =   50
            Top             =   480
            Width           =   1770
         End
         Begin VB.Label Label31 
            Caption         =   "Mod. Desc. Articulo    :"
            Height          =   255
            Left            =   -71370
            TabIndex        =   49
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label27 
            Caption         =   "Fecha Documento :"
            Height          =   255
            Left            =   3405
            TabIndex        =   48
            Top             =   1005
            Width           =   1575
         End
         Begin VB.Label Label26 
            Caption         =   "R.U.C.                   :"
            Height          =   255
            Left            =   360
            TabIndex        =   47
            Top             =   1005
            Width           =   1530
         End
         Begin VB.Label Label25 
            Caption         =   "Razón Social         :"
            Height          =   255
            Left            =   360
            TabIndex        =   46
            Top             =   750
            Width           =   1455
         End
         Begin VB.Label Label24 
            Caption         =   "Dirección                :"
            Height          =   255
            Left            =   3405
            TabIndex        =   45
            Top             =   750
            Width           =   1665
         End
         Begin VB.Label Label23 
            Caption         =   "Orden Compra       :"
            Height          =   255
            Left            =   360
            TabIndex        =   44
            Top             =   1530
            Width           =   1575
         End
         Begin VB.Label Label22 
            Caption         =   "Cotización              :"
            Height          =   255
            Left            =   3405
            TabIndex        =   43
            Top             =   1260
            Width           =   1545
         End
         Begin VB.Label Label21 
            Caption         =   "Pedido                   :"
            Height          =   255
            Left            =   3405
            TabIndex        =   42
            Top             =   1530
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "Mueve Stock (NC/ND) :"
            Height          =   255
            Left            =   -74715
            TabIndex        =   41
            Top             =   810
            Width           =   1755
         End
         Begin VB.Label Label4 
            Caption         =   "Punto de Vta.        :"
            Height          =   240
            Left            =   360
            TabIndex        =   40
            Top             =   465
            Width           =   1560
         End
         Begin VB.Label Label13 
            Caption         =   "Nro. Serie               :  "
            Height          =   210
            Left            =   3405
            TabIndex        =   39
            Top             =   495
            Width           =   1755
         End
         Begin VB.Label Label14 
            Caption         =   "Tipo Cambio          :"
            Height          =   255
            Left            =   360
            TabIndex        =   38
            Top             =   1260
            Width           =   1800
         End
         Begin VB.Label Label15 
            Caption         =   "Control Stock Positvo  :"
            Height          =   255
            Left            =   -71370
            TabIndex        =   37
            Top             =   810
            Width           =   1830
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Doc.               :"
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label Label2 
         Caption         =   "Doc. Contable        :"
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   690
         Width           =   1665
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción            :"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   1020
         Width           =   1665
      End
      Begin VB.Label Label5 
         Caption         =   "Serie Documento      :"
         Height          =   255
         Left            =   3945
         TabIndex        =   60
         Top             =   1020
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Nro. Inicio               :"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   1350
         Width           =   1635
      End
      Begin VB.Label Label7 
         Caption         =   "Nro. Final                  :"
         Height          =   255
         Left            =   3945
         TabIndex        =   58
         Top             =   1350
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Ult. Nro. Correl.       :"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   1710
         Width           =   1560
      End
      Begin VB.Label Label9 
         Caption         =   "Cod. Sunat               :"
         Height          =   255
         Left            =   3945
         TabIndex        =   56
         Top             =   1710
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Formato de Impr.     :"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   1995
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Lineas Imp.              :"
         Height          =   255
         Left            =   3945
         TabIndex        =   54
         Top             =   1995
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Impresora               :"
         Height          =   285
         Left            =   240
         TabIndex        =   53
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label Label17 
         Caption         =   "Controlador              :"
         Height          =   285
         Left            =   3945
         TabIndex        =   52
         Top             =   2340
         Width           =   1650
      End
      Begin VB.Label Label18 
         Caption         =   "Puerto de Impr.      :"
         Height          =   285
         Left            =   240
         TabIndex        =   51
         Top             =   2685
         Width           =   1620
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmCfgDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As ADODB.Recordset
Dim nTipOper As Integer
Dim csql As String
Dim nTra As Integer
Dim cTip As String
Dim cSER As String
Dim nPos As Integer
 
Private Sub CmdCon_Click()
If adodc1.RecordCount > 0 Then
    cTip = adodc1("CTNCODIGO")
    cSER = adodc1("CtnNumSer")
    OculObj (False)
    Limpiar
    Mostrar
    InhObj (False)
    Frame1.Caption = "Consulta de Documento"
    Frame1.Visible = True
    Cmdgrabar.Visible = False
    CmdSalir2.Visible = True
End If
End Sub

Private Sub CmdEli2_Click()
Dim nNd As Integer
On Error GoTo ElErr

If adodc1.RecordCount > 0 Then
    If MsgBox("Desea Eliminar el Dato", vbQuestion + vbOKCancel, "Mensaje") = vbOK Then
        cTip = adodc1("CtnCodigo")
        cSER = adodc1("CtnNumSer")
        If adodc1("CtnCodigo") = "WK" Then
            MsgBox "Este tipo de Documento no se puede Eliminar, porque está siendo utilizado por el Sistema", vbInformation, "Sistema de Ventas"
            DataGrid1.SetFocus
            Exit Sub
        ElseIf adodc1("CtnCodigo") = "LE" Then
            MsgBox "Este tipo de Documento no se puede Eliminar, porque está siendo utilizado por el Sistema", vbInformation, "Sistema de Ventas"
            DataGrid1.SetFocus
            Exit Sub
        End If
        
        If Existe(1, cTip, "FacCab", "CFTD", False, cSER, "CFNUMSER") = False Then
            csql = "Delete From Num_DocumentoS where CtnCodigo = '" & cTip & "' and CtnNumSer = '" & cSER & "'"
            nNd = Pos_Dato(adodc1)
            nTra = 1
            VGCNx.BeginTrans
            VGCNx.Execute csql
            VGCNx.CommitTrans
            nTra = 0
        Else
            MsgBox "No se puede Eliminar el Documento, porque esta siendo utilizado", vbInformation, "Información"
        End If
    
        adodc1.Requery
        If nNd <> 0 Then adodc1.AbsolutePosition = nNd
        DataGrid1.SetFocus
    End If
End If
Exit Sub

ElErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub CmdGrabar_Click()
On Error GoTo GrErr
If Trim(Text1) = "" Then
    MsgBox "El Documento no puede ser vacio", vbInformation, "Mensaje"
    Text1.SetFocus: Exit Sub
End If
If Trim(Text3) = "" Then
    MsgBox "La Descripción no puede ser vacio", vbInformation, "Mensaje"
    Text3.SetFocus: Exit Sub
End If

If nTipOper = 1 Then
    If ValNumSer(Text1, Text4) = False Then
        MsgBox "Serie ya existe para el Tipo de Documento  & " & Text1 & "", vbInformation, "Mensaje"
        Text4.SetFocus: Exit Sub
    End If

    csql = "Insert Into Num_Documentos (CtnCodigo,CtnNumser,CtnNumIni,"
    csql = csql & "CtnNumFin,CtnNumero,CtnFecha,CtnRuc,CtnRazSoci,"
    csql = csql & "CtnDirecci,CtnOrdComp,CtnCotizac,CtnPedido,CtnAlmacen,"
    csql = csql & "CtnModArti,CTNFORMATO,CTNLINEAS,CTMUEVSTOCK,CTPTO,CTSERNUM,CTCAMBIO,CTSTOCK,CTIMPRESORA,CTCONTROLADOR,CTPUERTO) Values ("
    csql = csql & "'" & Text1 & "',"
    csql = csql & "'" & Text4 & "'," & Text5 & "," & Text6 & ","
    csql = csql & "" & Text7 & ","
    If ChFecha.Value = 1 Then
        csql = csql & "'S',"
    Else
        csql = csql & "'N',"
    End If
    If ChRuc.Value = 1 Then
        csql = csql & "'S',"
    Else
        csql = csql & "'N',"
    End If
    If ChRazon.Value = 1 Then
        csql = csql & "'S',"
    Else
        csql = csql & "'N',"
    End If
    If ChDireccion.Value = 1 Then
        csql = csql & "'S',"
    Else
        csql = csql & "'N',"
    End If
    If ChOrden.Value = 1 Then
        csql = csql & "'S',"
    Else
        csql = csql & "'N',"
    End If
    If ChCotizacion.Value = 1 Then
        csql = csql & "'S',"
    Else
        csql = csql & "'N',"
    End If
    If ChPedido.Value = 1 Then
        csql = csql & "'S',"
    Else
        csql = csql & "'N',"
    End If
    If Chalmacen.Value = 1 Then
        csql = csql & "'S',"
    Else
        csql = csql & "'N',"
    End If
    If ChModo.Value = 1 Then
        csql = csql & "'S',"
    Else
        csql = csql & "'N',"
    End If
    csql = csql & "'" & Text9 & "','" & Text10 & "',"
    If ChStock.Value = 1 Then
        csql = csql & "'S',"
    Else
        csql = csql & "'N',"
    End If
    If ChPto.Value = 1 Then
        csql = csql & "'S',"
    Else
        csql = csql & "'N',"
    End If
    If ChSerie.Value = 1 Then
        csql = csql & "'S',"
    Else
        csql = csql & "'N',"
    End If
    If ChCambio.Value = 1 Then
        csql = csql & "'S',"
    Else
        csql = csql & "'N',"
    End If
    If ChControl.Value = 1 Then
        csql = csql & "'S',"
    Else
        csql = csql & "'N',"
    End If
    csql = csql & "'" & TxImpresora & "','" & TxControlador & "','" & TxPuerto & "')"
    
    
ElseIf nTipOper = 2 Then
    csql = "Update Num_Documentos Set CtnCodigo ='" & Text1 & "' ,"
    csql = csql & "CtnNumser='" & Text4 & "',CtnNumIni=" & Text5 & ","
    csql = csql & "CtnNumFin=" & Text6 & ",CtnNumero=" & Text7 & ","
    If ChFecha.Value = 1 Then
        csql = csql & "CtnFecha='S',"
    Else
        csql = csql & "CtnFecha='N',"
    End If
    If ChRuc.Value = 1 Then
        csql = csql & "CtnRuc='S',"
    Else
        csql = csql & "CtnRuc='N',"
    End If
    If ChRazon.Value = 1 Then
        csql = csql & "CtnRazSoci='S',"
    Else
        csql = csql & "CtnRazSoci='N',"
    End If
    If ChDireccion.Value = 1 Then
        csql = csql & "CtnDirecci='S',"
    Else
        csql = csql & "CtnDirecci='N',"
    End If
    If ChOrden.Value = 1 Then
        csql = csql & "CtnOrdComp='S',"
    Else
        csql = csql & "CtnOrdComp='N',"
    End If
    If ChCotizacion.Value = 1 Then
        csql = csql & "CtnCotizac='S',"
    Else
        csql = csql & "CtnCotizac='N',"
    End If
    If ChPedido.Value = 1 Then
        csql = csql & "CtnPedido='S',"
    Else
        csql = csql & "CtnPedido='N',"
    End If
    If Chalmacen.Value = 1 Then
        csql = csql & "CtnAlmacen='S',"
    Else
        csql = csql & "CtnAlmacen='N',"
    End If
    If ChModo.Value = 1 Then
        csql = csql & "CtnModArti='S',"
    Else
        csql = csql & "CtnModArti='N',"
    End If
    csql = csql & "CTNFORMATO = '" & Text9 & "',CTNLINEAS = '" & Text10 & "' "
    If ChStock.Value = 1 Then
        csql = csql & ",CTMUEVSTOCK= 'S',"
    Else
        csql = csql & ",CTMUEVSTOCK = 'N',"
    End If
    If ChPto.Value = 1 Then
        csql = csql & "CTPTO= 'S',"
    Else
        csql = csql & "CTPTO = 'N',"
    End If
    If ChSerie.Value = 1 Then
        csql = csql & "CTSERNUM= 'S',"
    Else
        csql = csql & "CTSERNUM = 'N',"
    End If
    If ChCambio.Value = 1 Then
        csql = csql & "CTCAMBIO= 'S',"
    Else
        csql = csql & "CTCAMBIO = 'N',"
    End If
    If ChControl.Value = 1 Then
        csql = csql & "CTSTOCK= 'S',"
    Else
        csql = csql & "CTSTOCK = 'N',"
    End If
    csql = csql & "CTIMPRESORA = '" & TxImpresora & "',CTCONTROLADOR = '" & TxControlador & "',CTPUERTO = '" & TxPuerto & "'   "
    csql = csql & "  Where CtnCodigo = '" & cTip & "' and CtnNumSer = '" & cSER & "'"
End If

nTra = 1
VGCNx.BeginTrans
VGCNx.Execute csql
VGCNx.CommitTrans
nTra = 0
Dim CSQL2 As String
Dim adoreg As ADODB.Recordset
If Text9 <> "" Then
  'if Nuevo
  Set adoreg = New ADODB.Recordset
   adoreg.Open "select nom_rep from formato where cod_emp='" & VGCODEMPRESA & "' and tipo_doc ='" & Text1 & "' ", VGconfig, adOpenStatic
   If adoreg.RecordCount > 0 Then
      CSQL2 = "Update FORMATO  Set NOM_REP = '" & Text9 & ".rpt" & "',COD_FOR ='" & Text9 & "', NOMBRE = '" & IIf(Trim(VGparametros.RucEmpresa) = "", " ", SupCadSQL(VGparametros.RucEmpresa)) & "'  Where "
      CSQL2 = CSQL2 & " COD_EMP = '" & VGCODEMPRESA & "' and TIPO_DOC = '" & Text1 & "'"
   Else
      CSQL2 = "Insert Into FORMATO (COD_EMP,COD_FOR,TIPO_DOC,NOM_REP,NOMBRE) Values "
      CSQL2 = CSQL2 & " ('" & VGCODEMPRESA & "','" & Text9 & "','" & Text1 & "','" & Text9 & ".rpt" & "','" & IIf(Trim(VGparametros.RucEmpresa) = "", " ", SupCadSQL(VGparametros.RucEmpresa)) & "')"
   End If
   VGconfig.BeginTrans
   VGconfig.Execute CSQL2
   VGconfig.CommitTrans
End If
adodc1.Requery
Limpiar
If nTipOper = 1 Then
    Text1.SetFocus
ElseIf nTipOper = 2 Then
    CmdSalir2_Click
End If

Exit Sub
GrErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub
Public Sub CmdImprimir_Click()

Dim CADENA As String
Dim cNomRepor  As String

cNomRepor = "al_defdocumentos.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Documentos"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport + "\" + cNomRepor
   
    CrystalReport1.Connect = VGcadenareport2
    CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
    
    CrystalReport1.formulas(0) = "emp ='" & VGparametros.RucEmpresa & "'"
    
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowState = crptMaximized
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
Else
    MsgBox "No existe el nombre del Reporte, verifique en Formatos", vbInformation, "Información"
    Exit Sub
End If

End Sub

Private Sub CmdIng_Click()
OculObj (False)
nTipOper = 1
Limpiar
Frame1.Caption = "Ingreso de Documento"
Frame1.Visible = True
Cmdgrabar.Visible = True
CmdSalir2.Visible = True
End Sub

Private Sub CmdModi_Click()
If adodc1.RecordCount > 0 Then
    nPos = adodc1.Bookmark
    If Not IsNull(adodc1("CtnCodigo")) Then cTip = adodc1("CtnCodigo")
    If Not IsNull(adodc1("CtnNumSer")) Then cSER = adodc1("CtnNumSer")
    OculObj (False)
    nTipOper = 2
    Limpiar
    Mostrar
    Text1.Enabled = False
    Frame1.Caption = "Modificación de Documento"
    Frame1.Visible = True
    Cmdgrabar.Visible = True
    CmdSalir2.Visible = True
    If nPos <> 0 Then adodc1.AbsolutePosition = nPos
End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSalir2_Click()
Text1.Enabled = True
Frame1.Visible = False
InhObj (True)
OculObj (True)
Cmdgrabar.Visible = False
CmdSalir2.Visible = False
DataGrid1.SetFocus
If nTipOper = 2 Then If nPos <> 0 Then adodc1.AbsolutePosition = nPos
End Sub

Private Sub Form_Activate()
If DataGrid1.Enabled And DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub Form_Load()
central Me
Set adodc1 = New ADODB.Recordset
adodc1.Open "Select CtnCodigo,TDO_CODSUN,TDO_DESCRI,CtnNumser,CtnNumero From Num_Documentos A Inner Join Tipo_Docu B on A.CTNCODIGO = B.TDO_TIPDOC order by CtnCodigo", VGCNx, adOpenStatic
'Adodc1.Open "Select CtnCodigo,TDO_CODSUN,TDO_DESCRI,CtnNumser,CtnNumero From Tipo_Docu B  Inner Join Num_Documentos A  on A.CTNCODIGO = B.TDO_TIPDOC order by CtnCodigo", Vgcnx, adOpenStatic
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh
End Sub

Private Sub Text1_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT TDO_TIPDOC,TDO_DESCRI,TDO_CODCON, TDO_CODSUN  FROM TIPO_DOCU ", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT TDO_TIPDOC,TDO_DESCRI,TDO_CODCON, TDO_CODSUN  FROM TIPO_DOCU"
frmReferencia.Label1.Caption = "Tipo de Documento"
frmReferencia.Show vbModal
Adodc3.Close
If vGUtil(1) <> "" Then Text1 = (vGUtil(1))
If vGUtil(3) <> "" Then Text2 = (vGUtil(3))
If vGUtil(2) <> "" Then Text3 = (vGUtil(2))
If vGUtil(4) <> "" Then Text8 = (vGUtil(4))

vGUtil(1) = ""
vGUtil(3) = ""
vGUtil(2) = ""
vGUtil(4) = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text1_DblClick
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Text1) = "" Then
        MsgBox "El Documento no puede ser vacio", vbInformation, "Mensaje"
        Text1.SetFocus: Exit Sub
    Else
        If Existe(1, Text1, "Tipo_Docu", "TDO_TIPDOC", False) Then
            Text3 = Devolver_Dato(1, Text1, "Tipo_Docu", "TDO_TIPDOC", False, "TDO_DESCRI")
            Text2 = Devolver_Dato(1, Text1, "Tipo_Docu", "TDO_TIPDOC", False, "TDO_CODCON")
            Text8 = Devolver_Dato(1, Text1, "Tipo_Docu", "TDO_TIPDOC", False, "TDO_CODSUN")
            SendKeys "{tab}"
        Else
            MsgBox "No existe existe este tipo de Documento", vbInformation, "Sistema de Ventas"
            Text1.SetFocus: Exit Sub
        End If
    End If
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub Text10_GotFocus()
Enfoque Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text2_GotFocus()
Enfoque Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub Text3_GotFocus()
Enfoque Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4.SetFocus
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub Text4_GotFocus()
Enfoque Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If nTipOper = 1 Then
        If ValNumSer(Text1, Text4) = False Then
            MsgBox "Serie ya existe para el Tipo de Documento  " & Text1 & "", vbInformation, "Mensaje"
            Text4.SetFocus: Exit Sub
        Else
            Text5.SetFocus
        End If
    Else
        SendKeys "{tab}"
    End If
End If
End Sub

Private Sub Text5_GotFocus()
Enfoque Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If NumSpto(KeyAscii) Then
    If KeyAscii = 13 Then Text6.SetFocus
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text6_GotFocus()
Enfoque Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If NumSpto(KeyAscii) Then
    If KeyAscii = 13 Then Text7.SetFocus
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text7_GotFocus()
Enfoque Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If NumSpto(KeyAscii) Then
    If KeyAscii = 13 Then SendKeys "{tab}"
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text8_GotFocus()
Enfoque Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If NumSpto(KeyAscii) Then
    If KeyAscii = 13 Then Text9.SetFocus
Else
    KeyAscii = 0
End If
End Sub

Private Sub OculObj(bT As Boolean)
DataGrid1.Visible = bT
CmdIng.Visible = bT
CmdModi.Visible = bT
cmdEli2.Visible = bT
CmdSalir.Visible = bT
CmdCon.Visible = bT
'cmdImprimir = bT

End Sub

Private Sub Limpiar()
Text1 = "": Text2 = " ": Text3 = " ": Text4.Enabled = True
Text4 = " ": Text5 = 0: Text6 = 0: TxImpresora = "": TxControlador = "": TxPuerto = ""
Text7 = 0: Text8 = 0: Text9 = " ": Text10 = 0
ChFecha.Value = 0
ChRazon.Value = 0
ChOrden.Value = 0
ChPedido.Value = 0
ChRuc.Value = 0
ChDireccion.Value = 0
ChCotizacion.Value = 0
Chalmacen.Value = 0
ChModo.Value = 0
ChPto.Value = 0
ChSerie.Value = 0
ChCambio.Value = 0
ChStock.Value = 0
ChControl.Value = 0
End Sub

Private Function ValNumSer(TipDoc As String, Serie As String) As Boolean
Dim cS As String, cR As ADODB.Recordset

cS = "Select * from Num_Documentos where cTncodigo = '" & TipDoc & "' and CtnNumser = '" & Serie & "'"
Set cR = New ADODB.Recordset
cR.Open cS, VGCNx, adOpenStatic
If cR.RecordCount > 0 Then
   ValNumSer = False
Else
    ValNumSer = True
End If
cR.Close
End Function

Private Sub Mostrar()
Dim cS1 As String, cR1 As ADODB.Recordset

cS1 = "Select * from Num_Documentos Where CtnCodigo = '" & cTip & "' and CtnNumSer = '" & cSER & "'"
Set cR1 = New ADODB.Recordset
cR1.Open cS1, VGCNx, adOpenStatic
If cR1.RecordCount > 0 Then
    Text1 = cR1("CtnCodigo")
    If cR1("Ctnfecha") = "S" Then
        ChFecha.Value = 1
    Else
        ChFecha.Value = 0
    End If
    
    If cR1("CtnRuc") = "S" Then
        ChRuc.Value = 1
    Else
        ChRuc.Value = 0
    End If
    
    If cR1("CtnRazSoci") = "S" Then
        ChRazon.Value = 1
    Else
        ChRazon.Value = 0
    End If
    
    If cR1("CtnDirecci") = "S" Then
        ChDireccion.Value = 1
    Else
        ChDireccion.Value = 0
    End If
    
    If cR1("CtnOrdComp") = "S" Then
        ChOrden.Value = 1
    Else
        ChOrden.Value = 0
    End If
     
    If cR1("CtnCotizac") = "S" Then
        ChCotizacion.Value = 1
    Else
        ChCotizacion.Value = 0
    End If
    
    If cR1("CtnPedido") = "S" Then
        ChPedido.Value = 1
    Else
        ChPedido.Value = 0
    End If
    
    If cR1("CtnAlmacen") = "S" Then
        Chalmacen.Value = 1
    Else
        Chalmacen.Value = 0
    End If
    
    If cR1("CtnModArti") = "S" Then
        ChModo.Value = 1
    Else
        ChModo.Value = 0
    End If
    
    If cR1("CTMUEVSTOCK") = "S" Then
        ChStock.Value = 1
    Else
        ChStock.Value = 0
    End If
    
    If cR1("CTPTO") = "S" Then
        ChPto.Value = 1
    Else
        ChPto.Value = 0
    End If
    If cR1("CTSERNUM") = "S" Then
        ChSerie.Value = 1
    Else
        ChSerie.Value = 0
    End If
    If cR1("CTCAMBIO") = "S" Then
        ChCambio.Value = 1
    Else
        ChCambio.Value = 0
    End If
    If cR1("CTSTOCK") = "S" Then
        ChControl.Value = 1
    Else
        ChControl.Value = 0
    End If
    
    If Not IsNull(cR1("CtnNumSer")) Then Text4 = cR1("CtnNumSer")
    If Not IsNull(cR1("CtnNumIni")) Then Text5 = cR1("CtnNumIni")
    If Not IsNull(cR1("CtnNumFin")) Then Text6 = cR1("CtnNumFin")
    If Not IsNull(cR1("CtnNumero")) Then Text7 = cR1("CtnNumero")
    If Not IsNull(cR1("CTNFORMATO")) Then Text9 = cR1("CTNFORMATO")
    If Not IsNull(cR1("CTNLINEAS")) Then Text10 = cR1("CTNLINEAS")
    If Not IsNull(cR1("CtIMPRESORA")) Then TxImpresora = cR1("CtIMPRESORA")
    If Not IsNull(cR1("CTCONTROLADOR")) Then TxControlador = cR1("CTCONTROLADOR")
    If Not IsNull(cR1("CTPUERTO")) Then TxPuerto = cR1("CTPUERTO")

    Text3 = Devolver_Dato(1, Text1, "Tipo_Docu", "TDO_TIPDOC", False, "TDO_DESCRI")
    Text2 = Devolver_Dato(1, Text1, "Tipo_Docu", "TDO_TIPDOC", False, "TDO_CODCON")
    Text8 = Devolver_Dato(1, Text1, "Tipo_Docu", "TDO_TIPDOC", False, "TDO_CODSUN")
    
End If
cR1.Close
End Sub

Private Sub InhObj(bT As Boolean)
Text1.Enabled = bT
Text4.Enabled = bT
Text5.Enabled = bT
Text6.Enabled = bT
Text7.Enabled = bT
Text9.Enabled = bT
Text10.Enabled = bT
ChFecha.Enabled = bT
ChRazon.Enabled = bT
ChOrden.Enabled = bT
ChPedido.Enabled = bT
ChRuc.Enabled = bT
ChDireccion.Enabled = bT
ChCotizacion.Enabled = bT
Chalmacen.Enabled = bT
ChModo.Enabled = bT
ChSerie.Enabled = bT
ChPto.Enabled = bT
ChCambio.Enabled = bT
ChStock.Enabled = bT
ChControl.Enabled = bT
End Sub

Private Sub Text9_GotFocus()
Enfoque Text9
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text10.SetFocus
End Sub


Private Sub TxControlador_GotFocus()
Enfoque TxControlador
End Sub

Private Sub TxControlador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxImpresora_GotFocus()
Enfoque TxImpresora
End Sub

Private Sub TxImpresora_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxPuerto_GotFocus()
Enfoque TxPuerto
End Sub

Private Sub TxPuerto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SSTab1.Tab = 0
End Sub

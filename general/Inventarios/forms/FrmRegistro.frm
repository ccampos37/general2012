VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmRegistro 
   Caption         =   "Form1"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Comentarios"
      Height          =   3030
      Left            =   1695
      TabIndex        =   53
      Top             =   3405
      Visible         =   0   'False
      Width           =   8745
      Begin VB.TextBox Text12 
         Height          =   2295
         Left            =   600
         MultiLine       =   -1  'True
         TabIndex        =   56
         Top             =   360
         Width           =   5655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Grabar"
         Height          =   495
         Left            =   6600
         TabIndex        =   55
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   6600
         TabIndex        =   54
         Top             =   1680
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   210
      TabIndex        =   48
      Top             =   5970
      Width           =   11745
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   10020
         TabIndex        =   50
         Top             =   240
         Width           =   1425
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1410
         TabIndex        =   49
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label Label16 
         Caption         =   "Total  Cantidad"
         Height          =   195
         Index           =   0
         Left            =   8760
         TabIndex        =   52
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label Label16 
         Caption         =   "Total  Items"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   51
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   11850
      Begin VB.TextBox TxTransa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1335
         MaxLength       =   2
         TabIndex        =   23
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5505
         MaxLength       =   11
         TabIndex        =   22
         Top             =   195
         Width           =   1275
      End
      Begin VB.TextBox TxTProveedor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   21
         Top             =   990
         Width           =   1320
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   20
         Top             =   1710
         Width           =   375
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   1350
         Width           =   1320
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8220
         MaxLength       =   20
         TabIndex        =   18
         Top             =   1350
         Width           =   1995
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7950
         MaxLength       =   3
         TabIndex        =   17
         Top             =   570
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   2400
         Width           =   1320
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   15
         Top             =   2085
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Valorizado"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   2955
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   288
         Left            =   10350
         MaxLength       =   15
         TabIndex        =   13
         Top             =   225
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Cmddetalle 
         Caption         =   "<<      Insertar producto(s)     >>"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6435
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2700
         Width           =   5295
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11325
         TabIndex        =   10
         Top             =   1350
         Width           =   405
      End
      Begin VB.TextBox tx_ordfab 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7740
         TabIndex        =   9
         Top             =   2085
         Visible         =   0   'False
         Width           =   1416
      End
      Begin VB.TextBox tx_codmaq 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         TabIndex        =   8
         Top             =   2535
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CheckBox ChkTalla 
         Alignment       =   1  'Right Justify
         Caption         =   "Ingresos por Tallas"
         Height          =   225
         Left            =   8370
         TabIndex        =   7
         Top             =   225
         Visible         =   0   'False
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1365
         TabIndex        =   12
         Top             =   225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   98369537
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin TextFer.TxFer TxNdoc 
         Height          =   375
         Left            =   6600
         TabIndex        =   24
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Appearance      =   0
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
      End
      Begin TextFer.TxFer TxSerie 
         Height          =   375
         Left            =   5760
         TabIndex        =   25
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Appearance      =   0
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   4
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
      End
      Begin VB.Label LblCC 
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
         Left            =   8730
         TabIndex        =   47
         Top             =   2070
         Width           =   2730
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Doc. :"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   150
         TabIndex        =   46
         Top             =   285
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Transaccion :"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   150
         TabIndex        =   45
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Num. Doc :"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   4590
         TabIndex        =   44
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor :"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   150
         TabIndex        =   43
         Top             =   1035
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tip Doc Ref :"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   150
         TabIndex        =   42
         Top             =   1740
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Cod. Cliente :"
         Height          =   195
         Left            =   150
         TabIndex        =   41
         Top             =   1410
         Width           =   945
      End
      Begin VB.Label Label8 
         Caption         =   "Orden Compra"
         ForeColor       =   &H80000006&
         Height          =   210
         Left            =   6915
         TabIndex        =   40
         Top             =   1410
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Autorizacion"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   6975
         TabIndex        =   39
         Top             =   615
         Width           =   1350
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Centro Ref :"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   150
         TabIndex        =   38
         Top             =   2475
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Almacen :"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   2145
         Width           =   705
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label 13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2700
         TabIndex        =   36
         Top             =   990
         Width           =   6630
      End
      Begin VB.Label Label14 
         Caption         =   "Num. Ref"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   4890
         TabIndex        =   35
         Top             =   1785
         Width           =   810
      End
      Begin VB.Label lbltrans 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbltrans"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1770
         TabIndex        =   34
         Top             =   600
         Width           =   5025
      End
      Begin VB.Label lbltipref 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbltipref"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1755
         TabIndex        =   33
         Top             =   1740
         Width           =   3015
      End
      Begin VB.Label lblClie 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblclie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2715
         TabIndex        =   32
         Top             =   1350
         Width           =   4080
      End
      Begin VB.Label lblauto 
         Caption         =   "lblauto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8580
         TabIndex        =   31
         Top             =   600
         Width           =   1965
      End
      Begin VB.Label lblalmacen 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblalmacen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1845
         TabIndex        =   30
         Top             =   2085
         Width           =   4290
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda"
         Height          =   210
         Left            =   10350
         TabIndex        =   29
         Top             =   1395
         Width           =   900
      End
      Begin VB.Label Label12 
         Caption         =   "Orden Fabricación"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   6240
         TabIndex        =   28
         Top             =   2145
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Equip./Maqui. :"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   3030
         TabIndex        =   27
         Top             =   2580
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LbltComp 
         Height          =   255
         Left            =   7335
         TabIndex        =   26
         Top             =   225
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Adicionar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6750
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6750
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   5655
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6750
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6750
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2580
      Left            =   120
      TabIndex        =   0
      Top             =   3465
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   4551
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FormatString    =   $"FrmRegistro.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   195
      Top             =   6075
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    vgRegEnt = 0 significa salida
'    vgregent = 1 significa ingreso
'    VGSeleccion = 1 Significa que es seleccion con frame de tipo de cambio
'    VGSeleccion = 2 Significa que es seleccion sin frame de tipo de cambio para modificar el contenido
'    VGSeleccion = 3 Significa que es seleccion sin frame de tipo de cambio para agregar item
'    VGform significa con formulario esta trabajando
'     text9    autorizado
'     text10  cencos
'     text11  almacen
Option Explicit
'Dim db As Database
Dim VGDllGeneral As New dllgeneral.dll_general
Dim nument As Long
Dim precioprom As Double
Dim CANTIDAD As Double
Dim canttemp As Double
Dim Campo As String * 2
Dim contador As Integer
Dim auxdisp As Integer
Dim num As Integer
Dim TT_CONTADOR As Integer
Dim estadocosto As Integer
Dim cadena As String
Dim alma As String
Dim tipo As String * 2
Dim dato As String
Dim empresaorigen As String
Dim NumDoc As String
Dim Codigo2 As String
Dim Comenta  As Boolean
Dim WithEvents Conex As ADODB.Connection
Attribute Conex.VB_VarHelpID = -1
Dim Completo As Boolean
Dim Nimprimir As Integer
Public CENTROCOSTO As Integer

Dim analitico As Integer
Dim xccosto As String
Dim xordfab As String
Dim xequip As String
Dim xfiltroequip As String
'***********************************
'**************RMM  07/07/2001
Dim rsSTKART As New ADODB.Recordset

'Ingreso
Private Sub Command1_Click()
If Check1.Value = 0 Then
   VGSeleccion = 1
   FrmCreacionSin.Caption = "Ingreso del Articulo"
   buscar_trans
   mostrardetalle
   FrmCreacionSin.Show 1
Else
   If MSFlexGrid1.Rows = 1 Then
      VGValnuevo = True
      VGSeleccion = 1
   Else
      VGSeleccion = 3
   End If
   FormCreacion.Caption = "Ingreso del Articulo"
   FormCreacion.Show 1
End If
End Sub

Private Sub Command2_Click()
If MSFlexGrid1.Rows = 1 Then
    MsgBox "No hay registros para Modificar", vbInformation, "Información"
    Exit Sub
End If
VGSeleccion = 2
If Check1.Value = 0 Then
    buscar_trans
    FrmCreacionSin.Caption = "Modificación del Detalle"
    FrmCreacionSin.Show 1
Else
    FormCreacion.Caption = "Modificación del Detalle"
    FormCreacion.Show 1
End If
End Sub
'Eliminar
Private Sub Command3_Click()
Dim i As Integer

If MSFlexGrid1.Rows = 1 Then
    MsgBox "No hay registros para Eliminar", vbInformation, "Información"
    Exit Sub
End If
If MsgBox("Desea Eliminar el Registro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
    
    i = MSFlexGrid1.RowSel
    If MSFlexGrid1.Rows > 2 Then
        MSFlexGrid1.RemoveItem i
    Else
        MSFlexGrid1.Clear
        MSFlexGrid1.Rows = 1
        MSFlexGrid1.Row = 0
        inicializaFG
        Command7.SetFocus
    End If
End If
End Sub

Private Sub Cmddetalle_Click()
'VALIDA LA CABECERA DE LA GUIA   >>
 Dim rf As New ADODB.Recordset
 Dim contitem As Integer
 contitem = 0
 
 If TxTProveedor <> "" Then
     If prove(TxTProveedor) <> "" Then
        contitem = contitem + 1
     Else
        TxTProveedor.SetFocus
        Exit Sub
     End If
 Else
     If TxTProveedor.Enabled And TxTProveedor.Visible Then
         MsgBox "falta llenar el Codigo del proveedor", vbExclamation, mensaje1
         Enfoque TxTProveedor
         Exit Sub
     End If
 End If
 If Text6.Enabled And Text6.Visible Then
     contitem = contitem + 1
 End If
 If Trim(Text10) <> "" Then
    contitem = contitem + 1
 Else
    If Text10.Enabled And Text10.Visible Then
         MsgBox "falta llenar el Codigo del Centro de Costo", vbExclamation, mensaje1
         Text10.SetFocus
         Exit Sub
    End If
 End If
 If Text8.Visible And Text8.Enabled Then
    contitem = contitem + 1
 End If
 If Trim(Text7) <> "" Then
    Set rf = VGCNx.Execute("select * from vt_cliente where clientecodigo='" & Text7 & "'")
    If rf.RecordCount = 0 Then
    'If Existe(1, Text7, "Maecli", "ccodcli", False) = False Then
        MsgBox "El Cliente no existe", vbInformation, "Información"
        Text7.SetFocus: Exit Sub
    End If
    Set rf = Nothing
    contitem = contitem + 1
 Else
     If Text7.Enabled And Text7.Visible Then
         MsgBox "falta llenar el Codigo del Cliente", vbExclamation, mensaje1
         Text7.SelStart = 0: Text7.SelLength = Len(Text7)
         Text7.SetFocus
         Exit Sub
     End If
 End If
 If TxTProveedor.Enabled And TxTProveedor.Visible Then
   If Text6.text = "" And VGRegEnt = 1 Then
     If Text6.text = "" Then
        MsgBox "Digite Tipo de Documento", vbInformation, "Información"
        If Text6.Enabled = True Then Text6.SetFocus: Exit Sub
      ElseIf TxSerie.text = "" Then
             MsgBox "Digite Numero de serie", vbInformation, "Información"
              TxSerie.SetFocus: Exit Sub
           ElseIf TxNdoc.text = "" Then
             MsgBox "Digite Numero de Numero de Documento", vbInformation, "Información"
            TxNdoc.SetFocus: Exit Sub
     End If
   End If
 End If
 If Trim(Text9) <> "" Then
     If Trim(validarautorizado(Text9)) = "" Then
        MsgBox "El Autorizado no existe", vbInformation, "Información"
        Text9.SetFocus: Exit Sub
     End If
     contitem = contitem + 1
 Else
     If Text9.Enabled And Text9.Visible Then
         MsgBox "Falta llenar el Codigo del Autorizado", vbExclamation, mensaje1
         Text9.SetFocus
         Exit Sub
     End If
 End If
 If Trim(Text11) <> "" Then
     contitem = contitem + 1
 Else
     If Text11.Enabled And Text11.Visible Then
         MsgBox "falta llenar el Codigo del almacen", vbExclamation, mensaje1
         Text11.SetFocus
     End If
 End If

     Cmddetalle.Enabled = False
     Check1.Enabled = False
     Text2 = "01"
     muestra
Text1.text = TxSerie.text + TxNdoc.text
End Sub

Private Sub Command4_Click()
' GRABA EL COMENTARIO DE LA GUIA
 Dim RSQL As String
 Dim rpta As String
 On Error GoTo Err
 RSQL = "Update MovAlmCab set CAGLOSA = '" & Text12 & "' "
 RSQL = RSQL & "Where CAALMA = '" & VGAlma & "'AND  CATD= '" & tipo & "' AND CANUMDOC = '" & Trim(Text4) & "'" '
 VGCNx.Execute RSQL
 Frame2.Visible = False
 crtlvisible (True)
' inicializar
 rpta = MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso")
 If rpta = vbYes Then
    imprimir
 End If
 inicializar
 inicializaFG
 Exit Sub
Err:
   MsgBox Err.Description
End Sub

Private Sub command5_Click()
' CANCELA EL COMENTARIO
  Dim rpta As Integer
  Frame2.Visible = False
  crtlvisible (True)
  inicializar
  rpta = MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso")
  If rpta = vbYes Then
     imprimir
  End If
End Sub
'****************************** Graba la NI ,NS ****************
Private Sub Command7_Click()
 ' Dim adodc2 As ADODB.Recordset
  Dim Data2 As New ADODB.Recordset
  Dim criterio As String
  Dim cadena As String
  Dim cadena1 As String
  Dim cadena2 As String
  Dim rpta As Integer
  Dim merma As Integer
    Dim FACTOR As Double
  Dim uSql As String
  On Error GoTo GrabErr
  
    
   CANTIDAD = 0
   If MSFlexGrid1.Rows = 1 Then
     MsgBox "No se puede grabar,debe adicionar registro", vbInformation, mensaje1
     Exit Sub
   End If
   If Not IsNumeric(Text4) Then
     MsgBox "Numero de Documento no consecutivo", vbExclamation, "Aviso"
     Exit Sub
   End If
   Text4 = Format(Text4, String(11, "0"))
' cambio en el control de numero de documentos
   Dim J As Integer
   Dim vgxregent As Integer
   Dim xdato As String
   Dim xtipo As String
   Dim X As Boolean
Set Data2 = Nothing
Set Data2 = Nothing
Data2.Open "movalmdet", VGCNx, adOpenDynamic, adLockOptimistic
  J = 0
  vgxregent = VGRegEnt
  xdato = dato
  xtipo = tipo
  For J = 1 To 2
  Nimprimir = 0
  If J = 2 Then
    VGRegEnt = 2
    dato = "S"
    tipo = "NS"
    TxTransa.text = "90"
'   Else
'    Exit For
  End If
   If J = 1 Then X = existe_numdoc(Text4, tipo)
'   Screen.MousePointer = 11
   If J = 1 Then grabacabecera
   FACTOR = 1    ' factor de conversion
   contador = 1  ' Contador de item
   'graba detalle
   NumDoc = Text4
   merma = 0
   While MSFlexGrid1.Rows > contador
     If (IIf(VGRegEnt = 1, True, True)) Then      'verificastk
       cadena = MSFlexGrid1.TextMatrix(contador, 0)
       CANTIDAD = 0
       If Not VGActualizar Then
              If J = 1 Or J = 2 And Val(MSFlexGrid1.TextMatrix(contador, 14)) <> 0 Then
                 Data2.AddNew
                 If J = 2 Then X = existe_numdoc(Text4, tipo)
                    NumDoc = Text4
                 If J = 2 Then grabacabecera
              End If
       Else
              criterio = "DECODIGO = '" & UCase(cadena) & "'"
              criterio = criterio + " and  DEALMA = '" & Text11 & "'"  ' VGAlma & "'"
              Data2.Find criterio
              If Data2.RecordCount = 0 Then
                MsgBox " No encontrado...!!!", vbInformation, "AVISO"
                Data2.Close
                Set Data2 = Nothing
                Exit Sub
              End If
       End If
      If J = 1 Or J = 2 And Val(MSFlexGrid1.TextMatrix(contador, 14)) <> 0 Then
         Data2("DEALMA") = Text11    'VGAlma
         Data2("DETD") = tipo ' "NS ,NI"
         Data2("DENUMDOC") = Text4.text
         Data2("DEITEM") = contador
         Data2("DECODIGO") = UCase(MSFlexGrid1.TextMatrix(contador, 0))   ' Format(MSFlexGrid1.TextMatrix(contador, 0), "00000000")
         Data2("DEDESCRI") = MSFlexGrid1.TextMatrix(contador, 1) 'Antes no se debe grababa se consulta a MAEART
         If J = 1 Then
            CANTIDAD = Val(MSFlexGrid1.TextMatrix(contador, 6))
          Else
            CANTIDAD = Val(MSFlexGrid1.TextMatrix(contador, 14))
         End If
         Data2("DECANTID") = CANTIDAD
         Data2("DECODMON") = Text2  'antes no se graba en detalle se consultaba a la cabecera
         Data2("DEUNIDAD") = MSFlexGrid1.TextMatrix(contador, 4) 'Antes no se debe grababa se consulta a MAEART
         Data2("DECANREF1") = "" & IIf(MSFlexGrid1.TextMatrix(contador, 14) = "", 0, MSFlexGrid1.TextMatrix(contador, 14))
         If MSFlexGrid1.TextMatrix(contador, 3) <> "" Then
            grabastk
            If MSFlexGrid1.TextMatrix(contador, 5) <> "" Then    'si tiene precio de costo
                Data2("DEPRECIO") = Val(MSFlexGrid1.TextMatrix(contador, 7)) ' * VGTipCamb '******el precio
                Data2("DETIPCAM") = MSFlexGrid1.TextMatrix(contador, 15) 'DevolverTCambio(DTPicker1.Value)
            ElseIf (estadocosto = 1 And VGRegEnt = 0) Or Text10.Visible Then  'SALIDA VALORIZADA  0 - SALIDA,1 - ENTRADA, text10 indica salida x C
                Data2("DEPRECIO") = precioprom  '******'valorizacion de precio prom
            Else
                Data2("DEPRECIO") = 0
            End If
            Data2("DECENCOS") = MSFlexGrid1.TextMatrix(contador, 11)
            Data2("DEORDFAB") = MSFlexGrid1.TextMatrix(contador, 12)
            Data2("DEQUIPO") = MSFlexGrid1.TextMatrix(contador, 13)
            alma = Text11 '' VGAlma  'indica el almacen
            'mejorar a una funcion
            If MSFlexGrid1.TextMatrix(contador, 10) = "S" Then
                grabaserie alma, cadena
                Data2("DESERIE") = MSFlexGrid1.TextMatrix(contador, 2)
            End If
            If MSFlexGrid1.TextMatrix(contador, 10) = "N" Then
                grabalote alma, cadena
                Data2("DELOTE") = MSFlexGrid1.TextMatrix(contador, 2)
            End If
         End If
         Data2.Update
       End If
     End If
     contador = contador + 1
   Wend
   'data2.Refresh
   
   Dim cad As String
   If Text11.text <> "" And (TxTransa = "TD" Or TxTransa = "SD") Then
     contador = 1
     While MSFlexGrid1.Rows > contador
        CANTIDAD = Val(MSFlexGrid1.TextMatrix(contador, 6))
        cad = insertar1
        Completo = False
        Conex.BeginTrans
        Conex.Execute cad
        Conex.CommitTrans
        Do
          DoEvents
        Loop Until Completo
        
        grabastk1                'graba en la tabla stk del otro almacen
        alma = Text11          'codigo del almacen
        tipo = "NI"                'cuando se realiza otra traansaccion
        If MSFlexGrid1.TextMatrix(contador, 10) = "S" Then grabaserie Text11, cadena
        If MSFlexGrid1.TextMatrix(contador, 10) = "N" Then grabalote Text11, cadena
        tipo = "NS"
        contador = contador + 1
     Wend
   End If
   
  'Activa el menu en las opciones reporte y consulta
  If Comenta And Nimprimir = 1 Then
     rpta = MsgBox("Desea Agregar Comentarios", vbYesNo + vbQuestion, "Aviso")
  Else
     rpta = vbNo
  End If
  If rpta = vbYes Then
     crtlvisible (False)
     Frame2.Visible = True
     Text12.SetFocus
  Else
   '  TxTransa.Enabled = True
     If Nimprimir = 1 Then
        rpta = MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso")
        If rpta = vbYes Then
           imprimir
        End If
     End If
 End If
Next
VGRegEnt = vgxregent
dato = xdato
tipo = xtipo
inicializar
inicializaFG
 VGSoles = True
 VGTipCamb = 1
 Screen.MousePointer = 1
 Exit Sub
GrabErr:
 'Resume
 MsgBox Err.Description, vbExclamation, "Error"
'Resume
 Screen.MousePointer = 1
 Exit Sub
 Resume
End Sub

Private Sub Command8_Click()
'*********************************** SALIR
Dim rpta As Integer
   If MSFlexGrid1.Rows > 1 Then
     rpta = MsgBox("Desea Grabar", vbYesNo + vbQuestion, "Aviso")
     If rpta = vbYes Then
       Command7_Click
     End If
   End If
   Label13.Caption = ""
   lbltrans = ""
   lbltipref = ""
   VGval = False
   Text1.Enabled = True
   Text6.Enabled = True
   TxTProveedor.Enabled = True
   Text8.Visible = True
   Label8.Visible = True
   Text8.Enabled = True
   Check1.Enabled = True
   VGForm = 5
   Unload Me
End Sub

Private Sub Check1_Click()
   VGval = True   'Para toda la valorizacion'
   VGValnuevo = True   'Para la pantalla de inicio'
   VGForm = 1
   SendKeys "{tab}"
End Sub

Private Sub Conex_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
  Completo = True
End Sub

Private Sub DTPicker1_Change()
If DTPicker1.Value > VGParamSistem.fechatrabajo Then
      MsgBox "Fecha de documento mayor a fecha de trabajo", vbInformation, "Mensaje"
      DTPicker1.SetFocus
End If
DTPicker1.Value = UltimoCierreFech(DTPicker1.Value)
VGTipCamb = DevolverTCambio(DTPicker1.Value)
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub Form_Activate()
   Dim J, kTotal As Double
   VGtipocreacion = 1
   
   If MSFlexGrid1.Rows > 1 Then
      Text5 = Format(MSFlexGrid1.Rows - 1, "##,###,##0.00")
      kTotal = 0
      For J = 1 To MSFlexGrid1.Rows - 1
        kTotal = kTotal + CDbl(MSFlexGrid1.TextMatrix(J, 3))
      Next
      Text3 = Format(kTotal, "##,###,##0.00")
   Else
      Text5 = Format(0, "##,###,##0.00")
      Text3 = Format(0, "##,###,##0.00")
   End If
   If VGAutomatico Then
     Text4.Enabled = False
   End If
End Sub

Private Sub Form_Load()
   Dim rs As New ADODB.Recordset
   Dim RSQL As String
   Dim numsal As String
   Me.Left = (Screen.Width - Me.Width) / 2
   Me.Top = 800
   DoEvents
   
    VGSeleccion = 1               'Indica el modo de apertura = 1 y modificacion=2
    VGtipocreacion = 1
    VGActualizar = False
    VGSoles = True
    VGForm = 5
    LIMPIACABECERA
    DTPicker1.MaxDate = VGParamSistem.fechatrabajo
    DTPicker1.Value = UltimoCierreFech(CDate(Format(VGParamSistem.fechatrabajo, "dd/MM/yyyy")))
    VGTipCamb = DevolverTCambio(DTPicker1.Value)
    
    RSQL = "select  TANUMENT, TANUMSAL from TabAlm  "
    Set rs = VGCNx.Execute(RSQL)
    If rs.RecordCount() = 0 Then
       MsgBox ("No existe registro del almacen en tabla de almacenes")
       GoTo salir
    End If
    nument = IIf(IsNull(rs(0)), 1, rs(0))
    numsal = IIf(IsNull(rs(1)), 1, rs(1))
    VGCNx.Execute RSQL
    
    If VGRegEnt = 1 Then
      Text4.text = Format(Val(nument), "00000000000")
      FrmRegistro.Caption = "Registro de Entrada"
      dato = "I"
      tipo = "NI"
      Codigo2 = "NOTA DE INGRESO"
      Text2.Visible = True
      ChkTalla.Caption = "Ingreso por Tallas"
      ocultarlabel
      
    Else
       ChkTalla.Caption = "Salida por Tallas"
       FrmRegistro.Caption = "Registro de Salida"
       dato = "S"
       tipo = "NS"
       Text2.Visible = False
       Label1.Visible = False
       Codigo2 = "NOTA DE SALIDA"
       Check1.Visible = False
       Text4.text = Format(Val(numsal) + 1, "00000000000")
    End If
    VGval = False
    habilitado (False)
    inicializaFG
    Text4.Enabled = False
    
    Command1.Picture = MDIPrincipal.ImageList2.ListImages("Adicionar").Picture
    Command2.Picture = MDIPrincipal.ImageList2.ListImages("Modificar").Picture
    Command3.Picture = MDIPrincipal.ImageList2.ListImages("Eliminar").Picture
    Command7.Picture = MDIPrincipal.ImageList2.ListImages("Grabar").Picture
    Command8.Picture = MDIPrincipal.ImageList2.ListImages("Retornar").Picture
    
    Exit Sub
salir:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        VGTipCamb = DevolverTCambio(VG_FecTrab)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 '************************** NUM REF
  If Text7 = "" And KeyAscii = 13 Then
     SendKeys "{tab}"
     KeyAscii = 0
     Exit Sub
  End If
 
 If KeyAscii = 13 And Text1.text <> "" Then
    If Not IsNumeric(Text1) And (Text6 = "BV") Then
       MsgBox "Ingrese el Numero de  la Boleta", vbOKOnly, "Aviso"
       Exit Sub
    End If
    If Text6.text = "FT" And Check1.Value = 1 Then
       FormCreacion.Text6.text = Text1
       FormCreacion.Text6.Enabled = False
    End If
       If Text8.Enabled Then
             Text8.SetFocus
       ElseIf Text7.Enabled Then
             Text7.SetFocus
       ElseIf Text9.Enabled Then
             Text9.SetFocus
       ElseIf Text10.Enabled Then
             Text10.SetFocus
       ElseIf Text11.Enabled Then
             Text11.SetFocus
       Else
             Cmddetalle_Click
       End If
 Else
    If Text6 = "BV" Then
        If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And KeyAscii <> 8 Then KeyAscii = 0
    End If
 End If
 Set Conex = New ADODB.Connection
 
End Sub

Private Sub Text10_DblClick()
  Dim Adodc3 As ADODB.Recordset   'Centro de Costos
  Set Adodc3 = New ADODB.Recordset
  If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
        Adodc3.Open "SELECT cencost_codigo,cencost_descripcion FROM centro_costos where  len(cencost_codigo) = '6' ", VGcnxCT, adOpenStatic, adLockOptimistic
  Else
        Adodc3.Open "SELECT cencost_codigo,cencost_descripcion FROM centro_costos ", VGCNx, adOpenStatic, adLockOptimistic
  End If
  
        frmReferencia.Conectar Adodc3, "SELECT cencost_codigo,cencost_descripcion FROM centro_costos  "
        frmReferencia.Label1.Caption = "Centro de Costos"
        frmReferencia.Show vbModal
        Adodc3.Close
        If vGUtil(1) <> "" Then
                 Text10 = vGUtil(1)
                 LblCC = vGUtil(2)
        End If
        If Text10 <> "" Then Text10_KeyPress (13)
 
End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   Text10_DblClick
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
'**********************CENTRO COSTO
If KeyAscii = 13 And Text10.text <> "" Then
  If Trim(Text10.text) <> "" Then
     If Existe(1, Text10, "CENTRO_COSTOS", "cencost_codigo", False) = False Then
              MsgBox "Centro de Costo no existe", vbInformation, "Mensaje"
             Text10.SetFocus: Exit Sub
     End If
     If Text11.Enabled Then
          Text11.SetFocus
      Else
          Tabula (KeyAscii)
          'Cmddetalle_Click
      End If
   Else
      MsgBox "Ingrese el numero de Centro de Costo", vbInformation, mensaje1
      Text10.SetFocus
   End If
End If
End Sub

Private Sub Text11_DblClick()
Dim Adodc3 As ADODB.Recordset    'Almacen Destino
Set Adodc3 = New ADODB.Recordset

'where empresacodigo='" & VGparametros.empresacodigo & "' and puntovtacodigo='" & VGparametros.puntovta & "'
Set Adodc3 = VGCNx.Execute("SELECT TAALMA,TADESCRI, empresacodigo FROM TABALM ")
frmReferencia.Conectar Adodc3, "SELECT TAALMA,TADESCRI FROM TABALM where empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "' "
frmReferencia.Label1.Caption = "Almacenes"
frmReferencia.Show vbModal
If vGUtil(1) <> "" Then Text11 = (vGUtil(1))
VGAlma = Trim(Text11)
Set Adodc3 = Nothing
Set Adodc3 = VGCNx.Execute("select * from stkart where stalma='" & VGAlma & "' and stskdis > 0 ")
If Adodc3.RecordCount = 0 Then
   MsgBox (" No existe saldos para el almacen Nro. --> " & VGAlma & "")
   Exit Sub
End If
Set Adodc3 = VGCNx.Execute("select * from tabalm where taalma='" & VGAlma & "'")
  If Adodc3.RecordCount > 0 Then
    If VGRegEnt = 1 Then
      Text4 = Format(Adodc3!tanument, "00000000000")
    Else
      Text4 = Format(Adodc3!tanumsal, "00000000000")
    End If
    empresaorigen = Adodc3!empresacodigo
  End If
If Text11 <> "" Then Text11_KeyPress (13)
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text11_DblClick
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Len(TxTransa.text) = 2 Then
      If TxTransa = "TD" And Text11 = VGAlma Then
           MsgBox "No se puede transferir al mismo almacen", vbExclamation, "Error"
           Text11.SetFocus
      Else
        lblalmacen = existe_almacen(Text11)
        If lblalmacen = "" Then Exit Sub
          Cmddetalle_Click
      End If
   End If
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
       Text9_DblClick
   End If
End Sub





Private Sub TxNdoc_KeyPress(KeyAscii As Integer)

Dim RSQL As New ADODB.Recordset
If KeyAscii = 13 Then
If IsNumeric(RTrim(TxNdoc.text)) Then TxNdoc.text = Right("0000000000" & RTrim(TxNdoc.text), TxNdoc.MaxLength)
   If Len(TxNdoc.text) > 0 And VGRegEnt = 1 Then
       SQL = " select * from movalmcab where empresacodigo='" & VGParametros.empresacodigo & "' and CACODPRO ='" & RTrim(TxTProveedor.text) & " '"
       SQL = SQL & " and carftdoc='" & Text6.text & "' and carfndoc='" & TxSerie.text & TxNdoc.text & "'"
       Set RSQL = VGCNx.Execute(SQL)
       If RSQL.RecordCount > 0 Then
          MsgBox (" Ya existe ingresado Numero de documento ")
          TxNdoc.SetFocus
          Exit Sub
      End If
   End If
End If
End Sub

Private Sub TxNdoc_LostFocus()
TxNdoc_KeyPress (13)
End Sub

Private Sub TxSerie_LostFocus()
If IsNumeric(RTrim(TxSerie.text)) Then TxSerie.text = Format(TxSerie.text, "0000")
End Sub

Private Sub TxTransa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    TxTransa_DblClick
ElseIf KeyCode = 46 Then
    lbltrans = ""
End If
  
End Sub

Private Sub TxTransa_KeyPress(KeyAscii As Integer)
  '****************** TRANSACCIONES
     If KeyAscii = 13 And Len(TxTransa.text) = 2 Then
           buscar_trans
           lbltrans = Mid(lbltrans, 1, 30)
           If lbltrans = "" Then
                Enfoque TxTransa
                Exit Sub
           End If
     Else       'habilitado (False)
        If KeyAscii = 8 Then
           lbltrans = ""
           habilitado (True)
           LIMPIACABECERA
           habilitado (False)
        Else
           KeyAscii = Asc(UCase(Chr(KeyAscii)))
       End If
       If Cmddetalle.Enabled Then habilitado (False)
    End If
End Sub

Private Sub TxTransa_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT TT_CODMOV,TT_DESCRI,tt_clie FROM Tabtransa where TT_CODTRANS_AUTO='' AND TT_tipmov = '" & IIf(VGRegEnt = 1, "I", "S") & "'", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT TT_CODMOV,TT_DESCRI,tt_clie FROM Tabtransa where TT_CODTRANS_AUTO='' AND TT_tipmov = '" & IIf(VGRegEnt = 1, "I", "S") & "'"
frmReferencia.Label1.Caption = "Transacciones"
frmReferencia.Show vbModal
If vGUtil(1) <> "" Then
    TxTransa = vGUtil(1)
    buscar_trans
    lbltrans = Mid(vGUtil(2), 1, 30)
End If
If TxTransa.text <> "" Then TxTransa_KeyPress (13)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If TxTransa.Enabled = False And KeyAscii = 13 And Text4 <> "" Then
   Text4 = Format(Text4, "00000000000")
   If Command7.Visible = True Then
     Command7.SetFocus
   End If
End If
End Sub

'**********     PROVEEDOR ****************
Private Sub TxtProveedor_DblClick()
  VGForm1 = 12
  FormAyuProv.Show 1
  If Trim(TxTProveedor) <> "" Then
     siguiente_tx5
  End If
End Sub

Private Sub TxtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    TxtProveedor_DblClick
ElseIf KeyCode = 46 Then
    Label13 = ""
End If
End Sub

Private Sub TxtProveedor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Trim(TxTProveedor) <> "" Then
         siguiente_tx5
   Else
      If KeyAscii = 8 Then Label13 = ""
   End If
End Sub

Private Sub siguiente_tx5()
          Label13 = Mid(prove(TxTProveedor), 1, 20)
          If Label13 = "" Then
             Enfoque TxTProveedor
             Exit Sub
          End If
          If Text6.Enabled Then
             Text6.SetFocus
          ElseIf Text8.Enabled Then
             Text8.SetFocus
          ElseIf Text7.Enabled Then
             Text7.SetFocus
          ElseIf Text9.Enabled Then
             Text9.SetFocus
          ElseIf Text10.Enabled Then
             Text10.SetFocus
          ElseIf Text11.Enabled Then
             Text11.SetFocus
          Else
              Cmddetalle_Click
          End If
End Sub
'**************** num ref *********************
Private Sub Text6_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
If VGRegEnt = 2 Then
   Adodc3.Open "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU WHERE TDO_TIPDOC<>'GR' ", VGCNx, adOpenStatic, adLockOptimistic
   frmReferencia.Conectar Adodc3, "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU WHERE TDO_TIPDOC<>'GR'"
Else
   Adodc3.Open "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU  ", VGCNx, adOpenStatic, adLockOptimistic
   frmReferencia.Conectar Adodc3, "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU "
End If
frmReferencia.Label1.Caption = "Tipo de Documentos"
frmReferencia.Show vbModal
If vGUtil(1) <> "" Then Text6 = (vGUtil(1))
If vGUtil(2) <> "" Then lbltipref = (vGUtil(2))
If Text6 <> "" Then
   If Text6 = "FT" And Check1.Value = 1 Then
         FormCreacion.Text6 = Text1
   End If
   TxSerie.SetFocus
Else
    Text6.SetFocus
End If
End Sub
Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Text6_DblClick
ElseIf KeyCode = 46 Then
     lbltipref = ""
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Len(Text6) = 2 Then
           Text6 = UCase(Text6)
           lbltipref = Mid(ValidarDoc(Text6), 1, 15)
           If lbltipref = "" Then
              Enfoque Text6
              Exit Sub
           End If
           TxSerie.SetFocus
    Else
           Text6 = ""
           SendKeys "{tab}"
           KeyAscii = 0
    End If
  Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If
  If KeyAscii = 8 Then lbltipref = ""
End Sub

Private Sub Text7_DblClick()
Dim acliente As New ADODB.Recordset
  FrmAyuCliente.Show 1
  Text7 = FrmAyuCliente.cCod
  lblClie = FrmAyuCliente.cNom
  siguiente_tx7
If analitico = 1 Then
   SQL = " clientecodigo='" & Text7 & "' and proyectocierre=0 and tipoanaliticocodigo='" & VGParamSistem.tipoanaliticocodigo & "'"
   Set acliente = VGCNx.Execute(" select * from gr_proyectos where " & SQL)
   If acliente.RecordCount = 0 Then
      MsgBox ("No existe proyectos activos para este cliente ")
      Text7.SetFocus
      xequip = "N"
      Exit Sub
    Else
      xequip = "S"
      xfiltroequip = SQL
    End If
End If
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 112 Then   'Cliente
    Text7_DblClick
  ElseIf KeyCode = 46 Then
    lblClie = ""
  End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     lblClie.Enabled = True
     Text7 = LTrim(Text7)
     lblClie = existe_clie(Text7)
     If lblClie = "" Then
           MsgBox "No existe el codigo del Cliente", vbInformation, mensaje1
           Exit Sub
      End If
     siguiente_tx7
  Else
     If KeyAscii = 8 Then lblClie = ""
  End If
End Sub
Private Sub siguiente_tx7()
   'lblClie = Mid(lblClie, 1, 10)
   lblClie = lblClie.Caption
   If Text7 <> "" Then
          If Text9.Enabled And Text9.Visible And Trim(Text9) = "" Then
             Text9.SetFocus
          ElseIf Text10.Enabled And Text10.Visible And Trim(10) = "" Then
             Text10.SetFocus
          ElseIf Text11.Enabled And Text11.Visible Then
             Text11.SetFocus
          Else
              Cmddetalle_Click
          End If
   End If
End Sub
 '***** Orden de compra
Private Sub Text8_KeyPress(KeyAscii As Integer)
  Dim criterio As String
  If KeyAscii = 13 Then
        Text8 = Trim(Text8)
        If Text8 <> "" Then
            criterio = "CANUMORD = '" & Text8.text & "' AND  CACODPRO ='" & TxTProveedor.text & "'"
'            Data1.Recordset.FindFirst criterio
            If VGDllGeneral.VerificaDatoExistente(VGCNx, "select * from movalmcab where " & criterio) = 1 Then
              MsgBox "La Orden de Compra ya fue registrada !", vbExclamation, mensaje1
              Exit Sub
            End If
        End If
        If Text7.Enabled And Text7.Visible Then
           Text7.SetFocus
        ElseIf Text9.Enabled And Text9.Visible Then
           Text9.SetFocus
        ElseIf Text10.Enabled And Text10.Visible Then
           Text10.SetFocus
        ElseIf Text11.Enabled And Text11.Visible Then
           Text11.SetFocus
        Else
           Cmddetalle_Click
        End If
End If
  
End Sub

Private Sub ocultarlabel()
    label7.Visible = False
    Text7.Visible = False
    Label9.Visible = False
    Text9.Visible = False
    Label10.Visible = False
    Text10.Visible = False
    Label11.Visible = False
    Text11.Visible = False
End Sub

Private Sub Text9_DblClick()
  FormAyuda.Show 1
  If Text10.Enabled And Text10 <> "" Then
        Text10.SetFocus
  ElseIf Text11.Enabled And Text11 <> "" Then
        Text11.SetFocus
  ElseIf TxSerie.text <> "" Then
         TxNdoc.SetFocus
  Else
        SendKeys "{tab}"
        
  End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then          'Autorizado
            If Trim(Text9) <> "" Then
                    If Trim(validarautorizado(Text9)) = "" Then
                            MsgBox "No existe el Autorizado", vbInformation, "Mensaje"
                            If Text9.Enabled And Text9.Visible Then Text9.SetFocus
                            Exit Sub
                    End If
                    lblauto = Mid(validarautorizado(Text9), 1, 10)
                    SendKeys "{tab}"
            ElseIf Text11.Enabled And Text11.Visible Then
                    Text11.SetFocus
            Else
                    Cmddetalle.SetFocus
            End If
       End If
End Sub

Private Sub muestra()
     Dim numfil As Integer
    ' Dim nument As Long
     Dim numsal As String
     Dim rs As New ADODB.Recordset
     Dim RSQL As String
     On Error GoTo errormuestra
     If Trim(Text11) <> "" Then
        VGAlma = Text11
        RSQL = "select  TANUMENT, TANUMSAL from TabAlm  WHERE TAALMA='" & Text11 & "'"  ' VGAlma & "' "
        'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
        Set rs = VGCNx.Execute(RSQL)
        
        Command1.Visible = True
        Command2.Visible = True
        Command3.Visible = True
        Command7.Visible = True
        mostrardetalle
        If Check1.Value = 0 Then
            VGSeleccion = 1
            buscar_trans
            'Fernando: 06/09/2001:

            If ChkTalla.Value = 0 Then
                FrmCreacionSin.Caption = "Ingreso del Detalle"
                FrmCreacionSin.Show 1
              Else
                'Call Load(FrmCreacionSin)
                'FrmIngTallas.Show 1
            End If
            '***
         Else
            VGSeleccion = 1
            FormCreacion.Caption = "Ingreso del Detalle"
            FormCreacion.Show 1
        End If
     Else
        MsgBox "No ningún Almacen Activo", vbInformation, "Información"
     End If
     Exit Sub
errormuestra:
Resume Next
 MsgBox Err.Description
End Sub

Public Function insertar1() As String

  Dim cad As String
  If MSFlexGrid1.TextMatrix(contador, 10) = "S" Then
          cad = "insert into MovAlmDet (DEALMA,DETD,DENUMDOC,DECODIGO,DECANTID,DEPRECIO,DEITEM,DESERIE,DECODMON,DECENCOS,DEORDFAB,DEQUIPO) values ('" & Text11 & "','" & Campo & "','" & Format(nument, "0000000000") & "','" & UCase(MSFlexGrid1.TextMatrix(contador, 0)) & "'," & CANTIDAD & "," & Val(precioprom) & "," & contador & ",'" & MSFlexGrid1.TextMatrix(contador, 2) & "','" & Text2 & "','" & MSFlexGrid1.TextMatrix(contador, 11) & "','" & MSFlexGrid1.TextMatrix(contador, 12) & "','" & MSFlexGrid1.TextMatrix(contador, 13) & "') "
  ElseIf MSFlexGrid1.TextMatrix(contador, 10) = "N" Then
          cad = "insert into MovAlmDet (DEALMA,DETD,DENUMDOC,DECODIGO,DECANTID,DEPRECIO,DEITEM,DELOTE,DECODMON,DECENCOS,DEORDFAB,DEQUIPO) values ('" & Text11 & "','" & Campo & "','" & Format(nument, "0000000000") & "','" & UCase(MSFlexGrid1.TextMatrix(contador, 0)) & "'," & CANTIDAD & "," & Val(precioprom) & "," & contador & ",'" & MSFlexGrid1.TextMatrix(contador, 2) & "','" & Text2 & "','" & MSFlexGrid1.TextMatrix(contador, 11) & "','" & MSFlexGrid1.TextMatrix(contador, 12) & "','" & MSFlexGrid1.TextMatrix(contador, 13) & "')"
  Else
          cad = "insert into MovAlmDet (DEALMA,DETD,DENUMDOC,DECODIGO,DECANTID,DEPRECIO,DEITEM,DECODMON,DECENCOS,DEORDFAB,DEQUIPO) values ('" & Text11 & "','" & Campo & "','" & Format(nument, "0000000000") & "','" & UCase(MSFlexGrid1.TextMatrix(contador, 0)) & "'," & CANTIDAD & "," & Val(precioprom) & "," & contador & ",'" & Text2 & "','" & MSFlexGrid1.TextMatrix(contador, 11) & "','" & MSFlexGrid1.TextMatrix(contador, 12) & "','" & MSFlexGrid1.TextMatrix(contador, 13) & "') "
  End If
  insertar1 = cad
End Function

Public Sub grabaalmacen()
 'proceso para una transferencia
  Dim uSql As String
  Dim insertar1 As String
  Dim Adodc3 As ADODB.Recordset
  Set Adodc3 = New ADODB.Recordset
  
  Adodc3.Open "select  TANUMENT from tabAlm where TAALMA =  '" & Text11 & " '", VGCNx, adOpenStatic, adLockOptimistic
  'Set rS = db.OpenRecordset(rSql, dbOpenSnapshot)
  If Adodc3.EOF Then
     MsgBox "No se ha declarado la numeracion para el almacen destino", vbInformation, "Aviso"
     Adodc3.Close
     Exit Sub
  End If
  nument = Adodc3(0) + 1
  Campo = "NI" 'verifica que el numero sea consecutivo
     
     Set Adodc3 = New ADODB.Recordset
     Adodc3.Open "SELECT  CANUMDOC from MOVALMCAB where CAALMA ='" & Text11 & "' AND  CATD = '" & Campo & "' and CANUMDOC =  '" & Format(nument, "0000000000") & "' ", VGCNx, adOpenStatic, adLockOptimistic
     If Not Adodc3.EOF Then
       Set Adodc3 = New ADODB.Recordset
       Adodc3.Open "SELECT MAX (CANUMDOC) from MOVALMCAB where CAALMA ='" & Text11 & "' AND  CATD = '" & Campo & "' ", VGCNx, adOpenStatic, adLockOptimistic
       nument = Adodc3(0) + 1
     End If
     Adodc3.Close
    
  insertar1 = "insert into MovAlmCab (CAALMA,CATD,CANUMDOC,CACODMOV,CAFECDOC,CATIPMOV,CASITGUI,CARFTDOC,CARFNDOC,CARFALMA,CAHORA,CACODPRO,CANOMPRO,CACODCLI,CANOMCLI,CACODMON) "
  insertar1 = insertar1 & " values ('" & Text11 & "','" & Campo & "','" & Format(nument, "0000000000") & "','03','" & DTPicker1 & "','I','V','NS','" & Text4 & "','01','" & Time & "','" & SupCadSQL(Trim(UCase$(TxTProveedor.text))) & "','" & SupCadSQL(LTrim(Label13.Caption)) & "','" & SupCadSQL(Mid$(UCase$(Text7.text), 1, 11)) & "','" & SupCadSQL(LTrim(lblClie.Caption)) & "','" & Text2 & "')"
  VGCNx.Execute insertar1
  uSql = "Update TabAlm set TANUMENT = " & nument & " where TAALMA='" & Text11 & "' "
  VGCNx.Execute uSql
 
    
End Sub

Public Sub grabastk()
  Dim acmd As New ADODB.Command
  Dim cadena As String
  Dim criterio As String
  Dim entrada As Boolean
  On Error GoTo GrabErr
   
cadena = MSFlexGrid1.TextMatrix(contador, 0)
Set rsSTKART = New ADODB.Recordset
rsSTKART.Open "Select * from STKART ", VGCNx, adOpenDynamic, adLockOptimistic
criterio = " STCODIGO = '" & cadena & "' and  STALMA ='" & Text11 & "'"
rsSTKART.Filter = criterio

If Not rsSTKART.EOF Then      'si existe el articulo
  
                canttemp = IIf(IsNull(rsSTKART("STSKDIS")), 0, rsSTKART("STSKDIS"))  ' revisar si validar en creacion
                rsSTKART("STKFECULT") = DTPicker1.Value
                If VGRegEnt = 1 Then
                    If LbltComp.Caption = 1 Then
                        rsSTKART("STSKCOM") = rsSTKART("STSKCOM") - CANTIDAD
                    Else
                        rsSTKART("STSKDIS") = rsSTKART("STSKDIS") + CANTIDAD
                    End If
                   'aqui actualiza
                   If Not IsNull(rsSTKART("STKPREPRO")) Then
                      precioprom = rsSTKART("STKPREPRO")
                      If MSFlexGrid1.TextMatrix(contador, 5) <> "" Then
                         rsSTKART("STKPREULT") = Val(MSFlexGrid1.TextMatrix(contador, 7)) * VGTipCamb 'el precio
                         If VGval And (canttemp + CANTIDAD) <> 0 Then
                          'valorizaAnte                          'valorizaActual                                                  saldoActu
                            rsSTKART("STKPREPRO") = Round(((precioprom * canttemp) + CANTIDAD * Val(Val(MSFlexGrid1.TextMatrix(contador, 7)) * VGTipCamb)) / (canttemp + CANTIDAD), 6)
                         End If
                      End If
                    Else
                      precioprom = 0
                      If MSFlexGrid1.TextMatrix(contador, 5) <> "" Then
                         rsSTKART("STKPREPRO") = Round(Val(MSFlexGrid1.TextMatrix(contador, 7)) * VGTipCamb, 6) 'el precio
                         If VGval Then
                            rsSTKART("STKPREULT") = Round(Val(MSFlexGrid1.TextMatrix(contador, 7)) * VGTipCamb) 'el precio
                            rsSTKART("STKFECULT") = DTPicker1.Value
                         End If
                      End If
                   End If
                Else
                  'para la salida
                   rsSTKART("STSKDIS") = rsSTKART("STSKDIS") - CANTIDAD
                   'aqui actualiza
                   If Not IsNull(rsSTKART("STKPREPRO")) Then
                      precioprom = Round(rsSTKART("STKPREPRO"), 6)
                    Else
                      precioprom = 0
                   End If
               End If
       Else
            rsSTKART.AddNew                   'existe
            rsSTKART("STALMA") = Text11    'VGAlma   '"01"
            rsSTKART("STCODIGO") = MSFlexGrid1.TextMatrix(contador, 0)
            rsSTKART("STKFECULT") = DTPicker1.Value
            If VGRegEnt Then
                rsSTKART("STSKDIS") = CANTIDAD
                rsSTKART("STKPREULT") = Val(MSFlexGrid1.TextMatrix(contador, 7)) * VGTipCamb    'el costo de ingreso
                If MSFlexGrid1.TextMatrix(contador, 5) <> "" Then
                      rsSTKART("STKPREPRO") = Round(Val(MSFlexGrid1.TextMatrix(contador, 7)) * VGTipCamb, 6) '******el  costo = costo prom
               End If
            End If
          'Grabamos en Facturacion
          Set acmd.ActiveConnection = VGGeneral
          acmd.CommandText = "al_actualizaproducto_pro"
          acmd.CommandType = adCmdStoredProc
          acmd.Prepared = True
          With acmd
            .Parameters("@baseini") = VGCNx.DefaultDatabase
            .Parameters("@basefin") = VGCNx.DefaultDatabase
            .Parameters("@almacen") = VGAlma
            .Parameters("@articulo") = MSFlexGrid1.TextMatrix(contador, 0)
            .Parameters("@tipo") = "1"
         End With
         acmd.Execute
         Set acmd = Nothing
         entrada = IIf(VGRegEnt = 1, True, False)
         Call ValMes(VGAlma, entrada) 'para la valorizacion
 End If
 rsSTKART.Update
 rsSTKART.Close
 Exit Sub
GrabErr:
 MsgBox Err.Description
 Exit Sub
 Resume
End Sub

Public Sub grabastk1()
   Dim criterio As String
   Dim cadena As String
   Dim acmd As New ADODB.Command
   
   On Error GoTo GrabErr
   cadena = MSFlexGrid1.TextMatrix(contador, 0)
   criterio = " STCODIGO ='" & cadena & "' and  STALMA ='" & Text11 & "'"
   rsSTKART.Filter = criterio
   If rsSTKART.EOF Then
     rsSTKART.AddNew
     rsSTKART("STSKDIS") = CANTIDAD
     rsSTKART("STKPREPRO") = Round(precioprom, 6)
     rsSTKART("STALMA") = Text11  '"01"
     rsSTKART("STCODIGO") = MSFlexGrid1.TextMatrix(contador, 0)
     
      Set acmd.ActiveConnection = VGCNx
       acmd.CommandText = "al_actualizaproducto_pro"
        acmd.CommandType = adCmdStoredProc
        acmd.Prepared = True
        With acmd
            .Parameters("@baseini") = VGCNx.DefaultDatabase
            .Parameters("@basefin") = VGBase2
            .Parameters("@almacen") = Text11
            .Parameters("@articulo") = MSFlexGrid1.TextMatrix(contador, 0)
            .Parameters("@tipo") = "1"
        End With
        acmd.Execute
        Set acmd = Nothing
   Else
     
     auxdisp = rsSTKART("STSKDIS")
     If rsSTKART("STKPREPRO") <> 0 And (canttemp + auxdisp) <> 0 Then 'no se registrado algun precio
       rsSTKART("STKPREPRO") = Round((precioprom * canttemp + auxdisp * rsSTKART("STKPREPRO")) / (canttemp + auxdisp), 6)
       rsSTKART("STKFECULT") = DTPicker1.Value
       rsSTKART("stkultfechacompra") = DTPicker1.Value
     End If
      rsSTKART("STSKDIS") = rsSTKART("STSKDIS") + CANTIDAD
   End If
   rsSTKART.Update
'   Data3.Refresh
   Call ValMes(Text11, True)  'para la valorizacion
   Exit Sub
GrabErr:
    MsgBox Err.Description
    'Resume
End Sub

Public Sub buscar_trans()
  Dim criterio As String
  Dim rs As New ADODB.Recordset
  Dim RSQL As String
  analitico = 0
   On Error GoTo GrabErrR
    
    TxTransa = UCase(LTrim(TxTransa))
    'Busco la transaccion
    RSQL = "select * from TabTransa where TT_CODMOV ='" & TxTransa.text & "' and TT_TIPMOV ='" & dato & "'"
    'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    If rs.RecordCount = 0 Then
       MsgBox "El tipo de transaccion no existe !", vbOKOnly, "Error"
       LIMPIACABECERA
       habilitado (False)
       TxTransa.SetFocus
       Exit Sub
    End If
    habilitado (True)
    lbltrans = Mid(rs("TT_DESCRI"), 1, 30)
    If Not IsNull(rs("TT_CONT")) Then
       TT_CONTADOR = rs("TT_CONT")
    Else
       MsgBox "El tipo de transacción no esta inicializara !" & Chr(13) & "Para inicializarla ir a la tabla de Transacción", vbOKOnly + vbExclamation, "Error"
       habilitado (False)
       Exit Sub
    End If
    estadocosto = ESNULO(rs("estadocosto"), 0)
    If rs("TT_PRV") = "N" Then
       TxTProveedor.Visible = False
       Label13.Visible = False
     Else
      TxTProveedor.Visible = True
      Label13.Visible = True
    End If
    If rs("tt_alma") = "S" Then
      Text11.Visible = True
      Label11.Visible = True
    End If
    If rs("TT_DR") = "N" Then
       Text6.Visible = False
       lbltipref.Visible = False
       Text1.Visible = False
     Else
       Text6.Enabled = True
       lbltipref.Visible = True
       Text1.Enabled = True
    End If
    
    If rs("TT_AT") = "N" Then
       Text9.Enabled = False
       Label9.Visible = False
       Text9.Visible = False
    Else
       Label9.Visible = True
       Text9.Visible = True
       Text9.Enabled = True
    End If
    CENTROCOSTO = 0
    Text10.Visible = False
    Text10.Enabled = False
    xccosto = rs("TT_CC")
    If rs("TT_CC") = "N" Then
       Text10.Enabled = False
       Label10.Visible = False
       Text10.Visible = False
       Check1.Enabled = True
    Else
       CENTROCOSTO = 1
       Label10.Visible = True
    '   Text10.Visible = True
    '   Text10.Enabled = True
       Check1.Enabled = False
       Check1.Value = 0
    End If
        
    If rs("TT_OC") = "N" Then
       Text8.Visible = False
       Label8.Visible = False
     Else
       Text8.Visible = True
       Label8.Visible = True
    End If
    If rs("TT_CLIE") = "S" Then
        Label8.Visible = False
        Text8.Visible = False
        label7.Visible = True
        Text7.Visible = True
        Text7.Enabled = True
   Else
        Label8.Visible = True
        Text8.Visible = True
        Text7.Enabled = False
        label7.Visible = False
        Text7.Visible = False
   End If
xordfab = rs("TT_ORDFAB")
If rs("TT_ORDFAB") = "S" Then
   '   tx_ordfab.Visible = True
   '   Label12.Visible = True
   Else
      tx_ordfab.Visible = False
      Label12.Visible = False
   End If
   xequip = rs("TT_EQUIP")
   If rs("TT_EQUIP") = "S" Then
      analitico = 1
      tx_codmaq.Visible = True
      Label15.Visible = True
   Else
      tx_codmaq.Visible = False
      Label15.Visible = False

  End If
  If rs("ingresosfuturos") = "S" Then
            LbltComp.Caption = 1
   Else
      LbltComp.Caption = 0
  End If
     
'*RMM*************************
   Comenta = IIf(rs("TT_CO") = "S", True, False)
   lbltrans = Mid(lbltrans, 1, 31)
   Cmddetalle.Enabled = True
   Exit Sub
GrabErrR:
 MsgBox Err.Description, vbInformation, "Aviso"

End Sub

Private Sub grabacabecera()
  Dim criterio As String
  Dim cadena As String
  Dim FACTOR As Double
  Dim uSql As String
  Dim Data1 As New ADODB.Recordset
  VGCNx.BeginTrans
  Set Data1 = Nothing
  Data1.Open "movalmcab", VGCNx, adOpenDynamic, adLockOptimistic
   On Error GoTo GrabErr
  'Desea grabar el registro
   If Text4.text <> "" Then
      VGAlma = "" & Trim(Text11)
      If Not VGActualizar Then
         Data1.AddNew
         Data1("empresacodigo") = VGParametros.empresacodigo
         Data1("CAALMA") = VGAlma
         Data1("CANUMDOC") = Mid$(UCase$(Text4.text), 1, 12)
      Else
         criterio = " CANUMDOC ='" & Text4 & "'"
         criterio = criterio + " and  CAALMA ='" & VGAlma & "'"
         Data1.Find criterio
      End If
      Data1("CATIPMOV") = dato
      Data1("CATD") = tipo
      Data1("CAHORA") = Format(Time, "hh:mm:ss")
      Data1("CAFECDOC") = DTPicker1.Value            ' CDate(Text2.text)
      Data1("CACOTIZA") = IIf(Len(Trim(tx_ordfab)) = 0, " ", tx_ordfab)
      
      If Trim(Text1.text) <> "" Then
         Data1("CARFNDOC") = SupCadSQL(Trim(Text1.text))
      Else
         Data1("CARFNDOC") = " "
      End If
      If TxTransa.text <> "" Then
         Data1("CACODMOV") = SupCadSQL(Mid$(UCase$(TxTransa.text), 1, 2))
      Else
         Data1("CACODMOV") = " "
      End If
      Text4 = Trim(UCase$(Text4.text))
      Data1("CANUMDOC") = Text4
      If Trim(TxTProveedor.text) <> "" Then
         Data1("CACODPRO") = SupCadSQL(Trim(UCase$(TxTProveedor.text)))
         Data1("CANOMPRO") = SupCadSQL(LTrim(Label13.Caption))
      Else
         Data1("CACODPRO") = " "
      End If
      Data1("CAFECACT") = Now
      If Trim(Text6) <> "" Then
         Data1("CARFTDOC") = SupCadSQL(Mid$(UCase$(Text6.text), 1, 2))
      Else
         Data1("CARFTDOC") = " "
      End If
      If Text7.Visible And Text7.text <> "" Then
         Data1("CACODCLI") = SupCadSQL(Mid$(UCase$(Text7.text), 1, 11))
      Else
         Data1("CACODCLI") = " "
      End If
      
     If Trim(Text8.text) <> "" And VGRegEnt = 1 Then
         Data1("CANUMORD") = SupCadSQL(Trim(UCase$(Text8.text)))
      Else
         Data1("CANUMORD") = " "
      End If
      If Text9.Visible And Trim(Text9) <> "" Then
         Data1("CASOLI") = Mid$(UCase$(Text9.text), 1, 3)
      Else
         Data1("CASOLI") = " "
      End If
      Data1("CAUSUARI") = UCase(VGUsuario)
      If Text10.Visible And Trim(Text10.text) <> "" Then
         Data1("CACENCOS") = Text10.text
      Else
         Data1("CACENCOS") = " "
      End If
      If Text11.Visible And Trim(Text11.text) <> "" Then
         Data1("CARFALMA") = Mid$(UCase$(Text11.text), 1, 2)
      Else
         Data1("CARFALMA") = " "
      End If
      Data1("CACODMON") = Text2
      'Data1.Recordset("CATIPCAM") = VGTipCamb
      Data1("CATIPCAM") = DevolverTCambio(DTPicker1.Value)
      VGCodMon = Text2
      Data1("CASITGUI") = "V"
      'Data1.Recordset("CASITUA") = "V"
      Data1("CAESTIMP") = "V"
      Data1("empresacodigo") = empresaorigen
      Data1.Update
   End If
   Data1.Close
   VGCNx.CommitTrans
   Nimprimir = 1
   Exit Sub
GrabErr:
       MsgBox Err.Description
       VGCNx.RollbackTrans
       Exit Sub
       Resume
End Sub
Function ValidarDoc(txt As TextBox) As String
  
  Dim rs As New ADODB.Recordset
  Dim RSQL As String
  
    RSQL = "select TDO_DESCRI  from TIPO_DOCU  where TDO_TIPDOC='" & SupCadSQL(txt.text) & "'"
  '  Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    If rs.EOF Then
       MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
       ValidarDoc = ""
       txt.SetFocus
       Exit Function
    End If
    ValidarDoc = rs(0)
    rs.Close

End Function

Function transa(text As TextBox) As String
 Dim rs As Recordset
 Dim RSQL As String
  RSQL = "select  TT_DESCRI FROM TabTransa where TT_CODMOV= '" & text & "' and TT_TIPMOV ='" & dato & "'" '

  Set rs = VGCNx.Execute(RSQL)
  If Not rs.EOF Then
    transa = rs(0)
  Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
    transa = ""
  End If
   rs.Close
End Function
Function tipref(text As TextBox) As String
 Dim rs As Recordset
 Dim RSQL As String
  RSQL = "select  TDO_DESCRI FROM TIPO_DOCU where TDO_TIPDOC= '" & text & "'" '
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
  Set rs = VGCNx.Execute(RSQL)
  If Not rs.EOF Then
    tipref = rs(0)
  Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
    tipref = ""
  End If
  rs.Close
End Function

Function prove(txt As TextBox) As String
 Dim rs As New ADODB.Recordset
 Dim RSQL As String
   RSQL = "select clienterazonsocial as PRVCNOMBRE FROM cp_proveedor where clientecodigo= '" & SupCadSQL(txt.text) & "'" '

   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
     prove = rs(0)
   Else
     MsgBox "El codigo del proveedor no existe !", vbExclamation, "Error"
     prove = ""
  End If
  rs.Close
End Function

Private Sub LIMPIACABECERA()
   TxTProveedor = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   Text10 = ""
   Text11 = ""
   Text1 = ""
   lbltrans = ""
   lbltipref = ""
   lblClie = ""
   lblauto = ""
   LblCC = ""
   lblalmacen = ""
   Label13 = ""
   Text2 = ""
   Label13.Caption = ""  'nombre
End Sub

Private Sub habilitado(bol As Boolean)
   TxTProveedor.Enabled = bol
   Text6.Enabled = bol
   Text8.Enabled = bol
   Text7.Enabled = bol
   Text9.Enabled = bol
   Text10.Enabled = bol
   Text11.Enabled = bol
   Text1.Enabled = bol
   Cmddetalle.Enabled = bol
End Sub
Private Sub inicializar()

  TxTransa.text = ""
  Text4.text = ""
  Check1.Value = 0
'  TxTransa.Enabled = True
  ocultarlabel
  Text12 = ""
  MSFlexGrid1.Clear
  MSFlexGrid1.Rows = 1
 ' inicializaFG
  Command1.Visible = False
  Command2.Visible = False
  Command3.Visible = False
  Command7.Visible = False
  'inicializar
  If Text6.text = "F" Then
    FormCreacion.Text6.Enabled = True
    FormCreacion.Text6.text = ""
  End If
  habilitado (True)
  LIMPIACABECERA
  habilitado (False)
  VGval = False
  Check1.Enabled = True
  Cmddetalle.Enabled = True
 
End Sub


Private Sub ValMes(almacen As String, entrada As Boolean)
  Dim cadena As String
  Dim criterio As String
  Dim adoreg As ADODB.Recordset
  Dim RSQL As String
  Dim Cantent As Double
  Dim Cantsal As Double
  Dim mespro As String
  Dim uSql As String
  On Error GoTo Err
   mespro = Year(DTPicker1) & Format(Month(DTPicker1), "00")
   cadena = MSFlexGrid1.TextMatrix(contador, 0) 'codigo del art
   RSQL = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & almacen & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & cadena & "'"  '
   Set adoreg = New ADODB.Recordset
   adoreg.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
    If Not adoreg.EOF Then 'existe
      If entrada Then
        Cantent = adoreg(0) + CANTIDAD
        uSql = "Update MoResMes set SMCANENT = " & Cantent & " where SMALMA='" & almacen & "'  and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
       Else
        Cantsal = adoreg(1) + CANTIDAD
        uSql = "Update MoResMes set SMCANSAL = " & Cantsal & " where SMALMA='" & almacen & "' and   SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
       End If
   Else
      If entrada Then
        Cantent = CANTIDAD
        Cantsal = 0
      Else
        Cantsal = CANTIDAD
        Cantent = 0
      End If
       uSql = "insert into MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMMNPREUNI,SMUSPREUNI,SMSALDOINI) VALUES ('" & almacen & "','" & cadena & "','" & mespro & "' ," & Cantent & "," & Cantsal & "," & Val(cNull(rsSTKART("STKPREPRO"))) & ",0,0) "

   End If
   VGCNx.Execute uSql
  Exit Sub
Err:
   MsgBox Err.Description
   
End Sub

Private Sub crtlvisible(dato As Boolean)
   MSFlexGrid1.Visible = dato
   Command1.Visible = dato
   Command2.Visible = dato
   Command3.Visible = dato
   Command7.Visible = dato
   Command8.Visible = dato

End Sub

Private Sub grabalote(alma As String, codigo As String)
Dim uSql As String
Dim Lote As String
Dim nuevo_stk As Double
Dim RSQL As String
Dim rs As Recordset
Dim fecfab As Date
Dim fecven As Date
    If (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" Then
      fecfab = MSFlexGrid1.TextMatrix(contador, 9)
    End If
    If (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
      fecven = MSFlexGrid1.TextMatrix(contador, 8)
    End If
    Lote = MSFlexGrid1.TextMatrix(contador, 2)
    RSQL = "select STSLKDIS FROM STKLOTE where  STSALMA= '" & alma & "' and  STSCODIGO= '" & codigo & "' and STSLOTE= '" & Lote & "'" '

    Set rs = VGCNx.Execute(RSQL)
    
    If Not rs.EOF Then
       If tipo = "NI" Then
         nuevo_stk = rs(0) + CANTIDAD
       Else
         nuevo_stk = rs(0) - CANTIDAD
       End If
       
       uSql = "Update STKLOTE set STSLKDIS = " & nuevo_stk & " WHERE  STSALMA='" & alma & "' and STSCODIGO='" & codigo & "'AND STSLOTE='" & Lote & "'"
    Else
        If (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) = "__/__/____" Then
            fecfab = MSFlexGrid1.TextMatrix(contador, 9)
            uSql = "insert into STKLOTE (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB,stscodprov) VALUES ('" & alma & "','" & codigo & "','" & Lote & "'," & CANTIDAD & ",'" & fecfab & "','" & TxTProveedor & "' ) "
        ElseIf (MSFlexGrid1.TextMatrix(contador, 9)) = "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
            fecven = MSFlexGrid1.TextMatrix(contador, 8)
            uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECVEN,stscodprov)  VALUES ('" & alma & "','" & codigo & "','" & Lote & "' ," & CANTIDAD & " ,'" & fecven & "','" & TxTProveedor & "') " 'SIN FECFAB
        ElseIf (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
            uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB,STSFECVEN,STSCODPROV)  VALUES ('" & alma & "','" & codigo & "','" & Lote & "' ," & CANTIDAD & " ,'" & fecfab & "','" & Format(fecven, "MM/DD/YYYY") & "','" & TxTProveedor & "') "
        Else
            uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSCODPROV)  VALUES ('" & alma & "','" & codigo & "','" & Lote & "' ," & CANTIDAD & ",'" & TxTProveedor & "') "
        End If
    End If
    VGCNx.Execute uSql
       
End Sub

Private Sub grabaserie(alma As String, codigo As String)
Dim uSql As String
Dim Serie As String
Dim valor As Integer
Dim rs As Recordset
Dim RSQL As String
Dim fecfab As Date
Dim fecven As Date
    Serie = MSFlexGrid1.TextMatrix(contador, 2)
    RSQL = "select STSSKDIS FROM STKSERI where   STSALMA= '" & alma & "' and  STSCODIGO= '" & codigo & "' and STSSERIE= '" & Serie & "'" '
    
    Set rs = VGCNx.Execute(RSQL)
    If Not rs.EOF Then
       valor = IIf(tipo = "NI", 1, 0)
       uSql = "update STKSERI set STSSKDIS = " & valor & " WHERE  STSALMA='" & alma & "' and STSCODIGO='" & codigo & "'AND STSSERIE='" & Serie & "'"
    Else
       uSql = "insert into STKSERI (STSALMA,STSCODIGO,STSSERIE,STSSKDIS)   VALUES ('" & alma & "','" & codigo & "','" & Serie & "',1) "
    End If
    rs.Close
    
    Set rs = Nothing
    VGCNx.Execute uSql
       
End Sub

Private Sub inicializaFG()
     Dim rf As New ADODB.Recordset
    
     MSFlexGrid1.Clear
     MSFlexGrid1.Cols = 16
     MSFlexGrid1.Row = 0
     MSFlexGrid1.ColWidth(0) = 2000
     MSFlexGrid1.ColWidth(1) = 2500
     MSFlexGrid1.ColWidth(2) = 1500
     MSFlexGrid1.ColWidth(3) = 1200
     MSFlexGrid1.ColWidth(4) = 900
     MSFlexGrid1.ColWidth(5) = 1200
     MSFlexGrid1.ColWidth(6) = 1000
     MSFlexGrid1.ColWidth(7) = 1000
     MSFlexGrid1.ColWidth(8) = 5
     MSFlexGrid1.ColWidth(9) = 5
     MSFlexGrid1.ColWidth(10) = 5
     
     MSFlexGrid1.ColWidth(11) = 1100
     MSFlexGrid1.ColWidth(12) = 1100
     MSFlexGrid1.ColWidth(13) = 1100
     MSFlexGrid1.ColWidth(14) = 1100
          
     MSFlexGrid1.TextMatrix(0, 0) = " CODIGO "
     MSFlexGrid1.TextMatrix(0, 1) = " DESCRIPCION"
     MSFlexGrid1.TextMatrix(0, 2) = " SERIE \LOT"
     MSFlexGrid1.TextMatrix(0, 3) = " CANTIDAD ING"
     MSFlexGrid1.TextMatrix(0, 4) = " UNIDAD"
     MSFlexGrid1.TextMatrix(0, 5) = " COSTO UNIT"
     MSFlexGrid1.TextMatrix(0, 6) = " CANT INF"
     MSFlexGrid1.TextMatrix(0, 7) = " COST0 INF"
     MSFlexGrid1.TextMatrix(0, 8) = " FECV"
     MSFlexGrid1.TextMatrix(0, 9) = " FECF"
     MSFlexGrid1.TextMatrix(0, 10) = " FS"
     
     MSFlexGrid1.TextMatrix(0, 11) = "Cent.Costo "
     MSFlexGrid1.TextMatrix(0, 12) = "Ord.Fabri  "
     MSFlexGrid1.TextMatrix(0, 13) = "Maqu./Equi."
     MSFlexGrid1.TextMatrix(0, 14) = "Cant.Ref"
     MSFlexGrid1.TextMatrix(0, 15) = "T.Cambio"
     
     
     MSFlexGrid1.ColAlignment(0) = 1
     MSFlexGrid1.ColAlignment(2) = 1

 
End Sub
Function existe_numdoc(text As TextBox, stipo As String) As Boolean
Dim numsal As String
Dim rs As New ADODB.Recordset
Dim RSQL As String
VGAlma = Text11
If Trim(Text11) <> "" Then
VGCNx.BeginTrans

   rs.Open "select  TANUMENT, TANUMSAL from TabAlm  WHERE TAALMA='" & Text11 & "'", VGCNx, adOpenDynamic, adLockOptimistic

   nument = IIf(IsNull(rs(0)), 1, rs(0))
   numsal = IIf(IsNull(rs(1)), 1, rs(1))
   If VGRegEnt = 1 Then
      Text4.text = Format(nument, "00000000000")
      rs("tanument") = rs("tanument") + 1
      nument = Text4.text
    Else
      Text4.text = Format(numsal, "00000000000")
      rs("tanumsal") = rs("tanumsal") + 1
      numsal = Text4.text
   End If
   rs.Update
   rs.Close
VGCNx.CommitTrans
End If
existe_numdoc = False
End Function
Function existe_ordcom(text As TextBox) As Boolean
Dim criterio As String
Dim RSQ As New ADODB.Recordset

 If Text8 <> "" And TxTProveedor <> "" Then
    criterio = "CANUMORD = '" & Text8.text & "' AND  CACODPROV ='" & TxTProveedor.text & "'"

    Set RSQ = VGCNx.Execute("select * from movalmcab where " & criterio)
    If RSQ.RecordCount > 0 Then
        MsgBox "El Numero documento ya ha sido registrado !", vbExclamation, "Error"
        existe_ordcom = True
        Exit Function
    End If
  End If
  existe_ordcom = False
End Function
Function existe_almacen(text As TextBox) As String
  Dim RSQL As String
  Dim rs As New ADODB.Recordset
  
   RSQL = "SELECT TADESCRI FROM TabAlm where  TAALMA= '" & text & "' and empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "'"
   'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then 'existe
     existe_almacen = rs(0)
   Else
     MsgBox "El codigo del almacen no existe !", vbOKOnly + vbInformation, "Error"
     existe_almacen = ""
   End If
   rs.Close
End Function

Function existe_clie(text As TextBox) As String
  Dim RSQL As String
  Dim rs As New ADODB.Recordset
  RSQL = "SELECT CNOMCLI FROM maecli where CCODCLI= '" & Trim(text) & "'"
  Set rs = VGCNx.Execute(RSQL)
  If rs.RecordCount > 0 Then 'existe
     existe_clie = rs(0)
  Else
     existe_clie = ""
  End If
  rs.Close
End Function

Function validarautorizado(text As TextBox) As String
  Dim RSQL As String
  Dim rs As Recordset
  Dim codayu As String
  codayu = 12
  RSQL = "Select TCLAVE,TDESCRI from TABAYU  where TCOD= '" & codayu & "' and  Tclave ='" & Trim(text) & "'"
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then 'existe
     validarautorizado = rs(1)
   Else
     validarautorizado = ""
  End If
  rs.Close
End Function

'******************************************************
'Procedimiento que permite verificar antes de grabar
Function verificastk() As Boolean
  Dim cadena As String
   cadena = MSFlexGrid1.TextMatrix(contador, 0)
   If MSFlexGrid1.TextMatrix(contador, 10) = "S" Then
     verificastk = IIf(existe_serie(cadena), True, False)
   ElseIf MSFlexGrid1.TextMatrix(contador, 10) = "N" Then
      verificastk = IIf(existe_lote(cadena), True, False)
   ElseIf consulta_stk Then
     verificastk = True
   Else
     verificastk = False
  End If
End Function

'Las siguientes consultas verifican si existe stock antes de grabar
'solo si esta saliendo mercaderia se hace la consulta
Function consulta_stk() As Boolean
Dim RSQL As String
Dim rs As Recordset
Dim cadena As String
   cadena = MSFlexGrid1.TextMatrix(contador, 0)
   RSQL = "select  stskdis from stkart  WHERE STALMA='" & VGAlma & "'  and stcodigo ='" & cadena & "'"
   'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
     If CANTIDAD > rs(0) Then
       consulta_stk = False
     Else
       consulta_stk = True
     End If
   End If
   rs.Close
End Function

Function existe_lote(text As String) As Boolean
Dim rs As Recordset
Dim RSQL As String
Dim Lote As String

   Lote = MSFlexGrid1.TextMatrix(contador, 2)
   RSQL = "select  STSLKDIS from STKLOTE where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & text & "' and STSLOTE = '" & Lote & "'"
'   Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
     If CANTIDAD > rs(0) Then
       MsgBox "No hay stock del" & text & "lote:" & Lote, vbInformation, "Aviso"
       existe_lote = False
     Else
       existe_lote = True
     End If
   End If
   rs.Close
End Function

Function existe_serie(text As String) As Boolean
Dim rs As Recordset
Dim RSQL As String
Dim Serie As String
   Serie = MSFlexGrid1.TextMatrix(contador, 2)
   RSQL = "select STSSKDIS from STKSERI where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & text & "' and STSSERIE = '" & Serie & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
     If CANTIDAD > rs(0) Then
       MsgBox "No hay stock " & text & " serie: " & Serie, vbInformation, "Aviso"
       existe_serie = False
     Else
       existe_serie = True
     End If
   End If
   rs.Close
End Function
Private Sub imprimir()
    Dim cadena As String
    Dim cFormato As String
    Dim cDireccion As String
    Dim cRuc As String
    Dim cNomRepor  As String
    Dim aBusca As New ADODB.Recordset
    
                           CrystalReport1.Reset
                            cNomRepor = "REPNOTAING.rpt"
                            CrystalReport1.ReportFileName = VGParamSistem.RutaReport & cNomRepor
               
                            CrystalReport1.Connect = VGCadenaReport2
                            CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
                            CrystalReport1.StoredProcParam(1) = VGAlma
                            CrystalReport1.StoredProcParam(2) = tipo
                            CrystalReport1.StoredProcParam(3) = Text4.text
                            
                            CrystalReport1.DiscardSavedData = True
                            CrystalReport1.Destination = crptToWindow
                            ''CrystalReport1.SelectionFormula = cadena
                            ''CrystalReport1.Formulas(0) = "Empresa = '" & VGparametros.RucEmpresa & "'"
                            ''CrystalReport1.Formulas(1) = "Direccion = '" & cDireccion & "' "
                            ''CrystalReport1.Formulas(2) = "Ruc = '" & cRuc & "' "
                            CrystalReport1.formulas(0) = "fecha='" & DTPicker1.Value & "'"
                            
                            
                            CrystalReport1.formulas(1) = "xtrans = '" & lbltrans.Caption & "' "
                            CrystalReport1.formulas(2) = "xtd = '" & Trim(tipo) & "' "
                            CrystalReport1.formulas(3) = "xndoc = '" & Text4.text & "' "
                            
                            
                            If tipo = "NI" Then
                                CrystalReport1.WindowTitle = "RepNotaIng -- Impresion de Notas de Ingreso"
                                CrystalReport1.formulas(4) = "Xnalma = '" & Text10.text & "' "
                                CrystalReport1.formulas(5) = "Dalma = '" & LblCC.Caption & "' "
                                CrystalReport1.formulas(6) = "AlmaDes = '" & VGAlma & "' "
                                CrystalReport1.formulas(7) = "Dalmades = '" & lblalmacen.Caption & "' "
                            
                            ElseIf tipo = "NS" Then
                                CrystalReport1.WindowTitle = "RepNotaIng -- Impresion de Notas de Salida"
                                CrystalReport1.formulas(4) = "Xnalma = '" & VGAlma & "' "
                                CrystalReport1.formulas(5) = "Dalma = '" & lblalmacen.Caption & "' "
                                CrystalReport1.formulas(6) = "AlmaDes = '" & Text10.text & "' "
                                CrystalReport1.formulas(7) = "Dalmades = '" & LblCC.Caption & "' "
                        
                            End If
                            
                            CrystalReport1.formulas(8) = "NRef = '" & Text1.text & "' "
                            CrystalReport1.formulas(9) = "DocRef = '" & Text6.text & "' "
                            CrystalReport1.formulas(10) = "TTrans = '" & TxTransa.text & "' "
                            CrystalReport1.formulas(11) = "emp = '" & VGParametros.RucEmpresa & "'"
                            CrystalReport1.WindowShowPrintBtn = True
                            CrystalReport1.WindowShowRefreshBtn = True
                            CrystalReport1.WindowShowSearchBtn = True
                            CrystalReport1.WindowShowPrintSetupBtn = True
                            CrystalReport1.WindowState = crptMaximized
                            
                            
                            If CrystalReport1.Status <> 2 Then
                                CrystalReport1.Action = 1
                                VGCNx.Execute "Update MovAlmCab Set CaEstImp = 'I' Where CATD = '" & tipo & "' and CANUMDOC = '" & Text4.text & "'"
                            End If
        Exit Sub
ErrImp:
     MsgBox Err.Description
     Resume Next
End Sub



Private Sub imprimirBK()
Dim cadena As String
If TxTransa = "DP" Then
   CrystalReport1.WindowTitle = "Inv520 -- Control de Inventarios"
   CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "\inv520.rpt"
Else
   CrystalReport1.WindowTitle = "Inv043 -- Control de Inventarios"
   CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "\inv043.rpt"
End If
Ubi_Tab CrystalReport1
cadena = "{MOVALMCAB.CAALMA} = '" & VGAlma & "'  and {MOVALMCAB.CATD} = '" & tipo & "' and {MOVALMCAB.CANUMDOC} = '" & NumDoc & "'"
CrystalReport1.DiscardSavedData = True
CrystalReport1.Destination = crptToWindow
CrystalReport1.WindowTitle = " Control de Inventarios"
CrystalReport1.ReplaceSelectionFormula (cadena)
CrystalReport1.WindowShowPrintBtn = True
CrystalReport1.WindowShowRefreshBtn = True
CrystalReport1.WindowShowSearchBtn = True
CrystalReport1.WindowShowPrintSetupBtn = True
CrystalReport1.formulas(0) = "empresa ='" & VGParametros.RucEmpresa & "'"
CrystalReport1.formulas(1) = "nota ='" & Codigo2 & "'"
CrystalReport1.formulas(2) = "hora ='" & Time & "'"
If VGRegEnt = 0 Then
    CrystalReport1.formulas(3) = "Tipo = 'S'"
Else
    CrystalReport1.formulas(3) = "Tipo = 'I'"
End If
CrystalReport1.Action = 1

If VGRegEnt <> 1 And TxTransa = "TD" Then
    If vbOK = MsgBox(" Desea imprimir la nota de Ingreso", vbInformation + vbOKCancel, "Aviso") Then
        CrystalReport1.WindowTitle = "Inv043 -- Control de Inventarios"
        CrystalReport1.ReportFileName = RUTA & "reportes\inv043.rpt"
        Ubi_Tab CrystalReport1
        cadena = "{MOVALMCAB.CAALMA} = '" & Text11 & "'  and {MOVALMCAB.CATD} = '" & Campo & "' and {MOVALMCAB.CANUMDOC} = '" & Format(nument, "0000000000") & "'"
        CrystalReport1.DiscardSavedData = True
        CrystalReport1.Destination = crptToWindow
        CrystalReport1.WindowTitle = " Control de Inventarios"
        CrystalReport1.ReplaceSelectionFormula (cadena)
        CrystalReport1.WindowShowPrintBtn = True
        CrystalReport1.WindowShowRefreshBtn = True
        CrystalReport1.WindowShowSearchBtn = True
        CrystalReport1.WindowShowPrintSetupBtn = True
        CrystalReport1.formulas(0) = "empresa ='" & VGParametros.RucEmpresa & "'"
        CrystalReport1.formulas(1) = "nota ='NOTA DE INGRESO'"
        CrystalReport1.formulas(2) = "hora ='" & Time & "'"
        CrystalReport1.formulas(3) = "Tipo = 'S'"
        CrystalReport1.Action = 1
   End If
End If
End Sub
Private Sub mostrardetalle()
      If xccosto = "N" Then
         FrmCreacionSin.txccosto.Enabled = False
         FrmCreacionSin.lblccosto.Enabled = False
         FrmCreacionSin.txccosto.Enabled = True
       Else
         FrmCreacionSin.lblccosto.Enabled = True
         FrmCreacionSin.txccosto = Text10
      End If
      If xordfab = "S" Then
         FrmCreacionSin.lblordfab.Enabled = True
         FrmCreacionSin.TxordFab.Enabled = True
        Else
         FrmCreacionSin.lblordfab.Enabled = False
         FrmCreacionSin.TxordFab.Enabled = False
      End If
      If xequip = "S" Then
         FrmCreacionSin.Ctr_AyuAnalitico.Enabled = True
         FrmCreacionSin.Ctr_AyuAnalitico.filtro = xfiltroequip
       Else
         FrmCreacionSin.Ctr_AyuAnalitico.Enabled = False
      End If
End Sub



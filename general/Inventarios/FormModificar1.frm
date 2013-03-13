VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormModificar 
   Caption         =   "Modificar"
   ClientHeight    =   5550
   ClientLeft      =   1455
   ClientTop       =   1275
   ClientWidth     =   9510
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   ScaleHeight     =   5550
   ScaleWidth      =   9510
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   1260
      TabIndex        =   30
      Top             =   4575
      Width           =   4230
      Begin VB.CommandButton Command6 
         Caption         =   "&Adicionar"
         Height          =   735
         Left            =   480
         Picture         =   "FormModificar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   120
         Width           =   840
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1800
         Picture         =   "FormModificar.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   120
         Width           =   840
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   735
         Left            =   3075
         Picture         =   "FormModificar.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton Cmdgrabar 
         Caption         =   "&Grabar"
         Height          =   735
         Left            =   3075
         Picture         =   "FormModificar.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   5685
      Picture         =   "FormModificar.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4680
      Width           =   885
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Consultar"
      Height          =   735
      Left            =   3195
      Picture         =   "FormModificar.frx":1412
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4680
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle del Documento"
      ForeColor       =   &H80000007&
      Height          =   4455
      Left            =   165
      TabIndex        =   5
      Top             =   120
      Width           =   9255
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6480
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   2115
         Width           =   525
      End
      Begin VB.CommandButton CmdGrabarCab 
         Caption         =   ">>"
         Height          =   255
         Left            =   8595
         TabIndex        =   16
         Top             =   2130
         Width           =   495
      End
      Begin VB.TextBox TxClie 
         Height          =   285
         Left            =   1815
         MaxLength       =   11
         TabIndex        =   13
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox TxAut 
         Height          =   285
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   14
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox TxCenCos 
         Height          =   285
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   15
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox TxOrdCom 
         Height          =   285
         Left            =   6480
         MaxLength       =   10
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TxAlmDes 
         Height          =   285
         Left            =   6480
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "Text6"
         Top             =   1440
         Width           =   495
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   6495
         TabIndex        =   35
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36521
      End
      Begin VB.TextBox TxDocNroRef 
         Height          =   285
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   1425
         Width           =   1575
      End
      Begin VB.TextBox TxDocRef 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox TxProv 
         Height          =   285
         Left            =   1800
         MaxLength       =   11
         TabIndex        =   8
         Top             =   1095
         Width           =   1395
      End
      Begin VB.TextBox TxTransa 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox TxDoc 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid MsFlexGrid1 
         Height          =   1815
         Left            =   240
         TabIndex        =   29
         Top             =   2520
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   1
         Cols            =   11
         FormatString    =   $"FormModificar.frx":1854
      End
      Begin VB.Label Label11 
         Caption         =   "Moneda"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   5205
         TabIndex        =   53
         Top             =   2145
         Width           =   1095
      End
      Begin VB.Label LblCosto 
         Caption         =   "lblCosto"
         Height          =   225
         Left            =   2850
         TabIndex        =   51
         Top             =   2190
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000B&
         Caption         =   "Cod. Cliente"
         Height          =   255
         Left            =   480
         TabIndex        =   41
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Autorizacion"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   5220
         TabIndex        =   40
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Centro Costo"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblClie 
         Caption         =   "lblclie"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2880
         TabIndex        =   38
         Top             =   1845
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Orden Compra"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   5205
         TabIndex        =   37
         Top             =   1095
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Almacen Dest."
         Height          =   255
         Left            =   5220
         TabIndex        =   36
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblProv 
         Caption         =   "Label8"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3330
         TabIndex        =   25
         Top             =   1095
         Width           =   1845
      End
      Begin VB.Label Lbltransa 
         Caption         =   "Label7"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label6 
         Caption         =   "Num"
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   5760
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Lblnumdoc 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   21
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Transaccion"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Doc Referencial"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   1440
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   195
      TabIndex        =   0
      Top             =   120
      Width           =   9090
      Begin VB.OptionButton Option4 
         Caption         =   "Todos"
         Height          =   225
         Left            =   3840
         TabIndex        =   52
         Top             =   3555
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Guias"
         Height          =   240
         Left            =   3840
         TabIndex        =   3
         Top             =   2955
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nota de Salida"
         Height          =   270
         Left            =   3840
         TabIndex        =   2
         Top             =   2355
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nota de Ingreso"
         Height          =   255
         Left            =   3840
         TabIndex        =   1
         Top             =   1785
         Value           =   -1  'True
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   255
         Left            =   6795
         TabIndex        =   42
         Top             =   3735
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36704
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   6795
         TabIndex        =   43
         Top             =   3375
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36704
      End
      Begin VB.Label Label8 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   5835
         TabIndex        =   45
         Top             =   3735
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Desde"
         Height          =   255
         Left            =   5835
         TabIndex        =   44
         Top             =   3375
         Width           =   735
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   1305
         Picture         =   "FormModificar.frx":1938
         Stretch         =   -1  'True
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Seleccione un Tipo de Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   1080
         Width           =   4530
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   3090
      Left            =   240
      TabIndex        =   28
      Top             =   1305
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   5450
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
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
      Height          =   975
      Left            =   1275
      TabIndex        =   46
      Top             =   180
      Width           =   6735
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormModificar.frx":1EA62
         Left            =   4440
         List            =   "FormModificar.frx":1EA6F
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox TxtBuscar 
         Height          =   285
         Left            =   1320
         TabIndex        =   47
         Top             =   375
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "Indice"
         Height          =   255
         Left            =   3480
         TabIndex        =   50
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   360
         TabIndex        =   49
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "FormModificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim valorizado As Boolean
Public tipo As String
Dim cantidad As Double

Dim Rs2 As Recordset
Dim rsstock As Recordset
Dim entrodetalle As Boolean
Public contador As Integer
Dim codigo As String
Dim serie_lote As String
Public numitem As Integer

Private Sub CmdEliminar_Click()
Dim rpta, item As Integer
Dim fila As Integer
Dim csql As String
  If MSFlexGrid1.Rows = 1 Then Exit Sub
  rpta = MsgBox("Seguro que desea eliminar" & Chr(13) & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0), vbInformation + vbOKCancel, "Confirmacion")
  If rpta = vbOK Then
          contador = MSFlexGrid1.Row
          codigo = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
          cantidad = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
          actualizastk (codigo)
          csql = "delete from movalmdet where  DEALMA ='" & VGAlma & "' AND DETD = '" & TxDoc & "' AND DENUMDOC ='" & Lblnumdoc & "' AND DECODIGO ='" & codigo & "'  and  DEITEM =" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) & ""
          cConexCom.Execute csql
          'falta eliminar el item del flex
         fila = MSFlexGrid1.RowSel
         If MSFlexGrid1.Rows > 2 Then
             MSFlexGrid1.RemoveItem fila
         Else
             inicializarFlex
        End If
End If
End Sub

Private Sub CmdGrabarCab_Click()
Dim rSql As String
Dim Rsql1 As String
  ' validar
  If TxDoc = "NI" Or TxDoc = "NS" Or TxDoc = "GS" Then
  
  Else
        MsgBox "Ingreso incorrecto de codigo de transacción", vbExclamation, "Error"
        TxDoc.SetFocus
        Exit Sub
  End If
  If Trim(TxTransa) <> "" Then
        Lbltransa = transa1(TxTransa)
        If Lbltransa = "" Then
             TxTransa.SetFocus
             Exit Sub
        End If
        Lbltransa = Mid(Lbltransa, 1, 18)
  Else
       MsgBox "Ingrese el codigo de transacción", vbExclamation, "Error"
       TxTransa.SetFocus
       Exit Sub
  End If
  If Trim(TxProv) <> "" Then
       lblProv = prove1(TxProv)
       If lblProv = "" Then
             TxProv.SetFocus
             Exit Sub
        End If
        lblProv = Mid(lblProv, 1, 18)
  End If
  grabacabecera
  rSql = "select * from STKART"
  Set rsstock = VGBaseDatos.OpenRecordset(rSql, dbOpenDynaset)          ', dbDenyWrite, dbPessimistic)
  entrodetalle = True
  muestradetalle
  numitem = MSFlexGrid1.Rows
  CmdGrabarCab.Enabled = False
End Sub

Private Sub Command1_Click()
 Dim precio As Double
  Dim cantidad As Double
  Dim contador As Integer
  Dim Rsql1 As String
  Dim rSql As String
  Dim rS As Recordset
  Dim Adoreg1 As ADODB.Recordset
  Dim ser_lot As String
  Dim dato As String
  limpia
  If Frame2.Visible Then
        If Option3.Value Then
          tipo = "GS"
         ElseIf Option2.Value Then
           tipo = "NS"
         ElseIf Option1.Value Then
              tipo = "NI"
              
         Else
             tipo = "XX"
             rSql = "select  m.CATD, m.CANUMDOC, m.CACODMOV ,m.CAFECDOC, m.CACODPRO,m.CACODCLI, m.CARFTDOC ,m.CARFNDOC,m.CACIERRE from MovAlmCab m where m.CASITGUI<>'A' and m.CAALMA ='" & VGAlma & "'   and m.CATD  IN  ('NI','NS','GS')  and  m.cafecdoc  between " & DateSQL(DTPicker2.Value) & " and " & DateSQL(DTPicker3.Value) & "  ORDER BY m.CANUMDOC"    '
         End If
         If tipo <> "XX" Then
            '  *********************  No se puede modificar guia ya impresa           ********************************************************************************
            '  Tambien cuando viene de transferecia y su documento es de tipo NI ,caestimp estado de impresion pueder  nulo
          
           rSql = "select  m.CATD, m.CANUMDOC, m.CACODMOV ,m.CAFECDOC, m.CACODPRO,m.CACODCLI, m.CARFTDOC ,m.CARFNDOC,m.CACIERRE " & _
               "from MOVALMCAB m where m.CASITGUI<>'A' and m.CAALMA ='" & VGAlma & "' and m.CATD='" & tipo & "' and not m.CACIERRE AND m.cafecdoc  between " & DateSQL(DTPicker2.Value) & " and " & DateSQL(DTPicker3.Value) & "  ORDER BY m.CANUMDOC" '
        End If
        Set Adoreg1 = New ADODB.Recordset
        Adoreg1.Open rSql, cConexCom, adOpenDynamic, adLockOptimistic
        If Adoreg1.RecordCount = 0 Then
             MsgBox "No hay documentos a modificar", vbInformation, "Aviso"
             Exit Sub
        End If
        FG.Visible = False
        FG.Rows = 1
        Adoreg1.MoveFirst
        While Not Adoreg1.EOF
            FG.AddItem (Adoreg1(0) & vbTab & Adoreg1(1) & vbTab & Adoreg1(2) & vbTab & Adoreg1(3) & vbTab & Adoreg1(4) & vbTab & Adoreg1(5) & vbTab & Adoreg1(6) & vbTab & Adoreg1(7) & vbTab & IIf(Adoreg1(8), "*", " "))
            Adoreg1.MoveNext
        Wend
        Adoreg1.Close
        CmdGrabarCab.Enabled = True
        FG.Visible = True
        Frame2.Visible = False
        Command1.Caption = "Consultar"
        Exit Sub
  End If
  If Frame1.Visible Then
      Frame1.Visible = False
  Else
     If FG.TextMatrix(FG.Row, 8) = "*" Then
         MsgBox "No se puede modificar,documento ya procesado", vbInformation, "Modificar"
         Exit Sub
     End If
     If Trim(FG.TextMatrix(FG.Row, 3)) <> "" Then
         DTPicker1 = Format(FG.TextMatrix(FG.Row, 3), "dd/mm/yyyy") 'fecha
     End If
     TxTransa = FG.TextMatrix(FG.Row, 2)  'tras
     Lbltransa = transa
     TxDoc = FG.TextMatrix(FG.Row, 0) ' tipo de doc
     If TxTransa = "TD" Then
         MsgBox "No se puede modificar el documento por tipo de transacion", vbInformation, "Aviso"
         Exit Sub
     End If
     Command1.Visible = False
     Frame1.Visible = True
     Frame3.Visible = True
     Command1.Caption = "&Aceptar"
     Lblnumdoc = FG.TextMatrix(FG.Row, 1) ' cod de doc
     TxProv = FG.TextMatrix(FG.Row, 4)  'proveedor
     If TxProv <> "" Then lblProv = prove
     If lblProv <> "" Then lblProv = Mid(lblProv, 1, 18)
     TxClie = FG.TextMatrix(FG.Row, 5)
     If TxClie <> "" Then lblClie = existe_clie(TxClie)
     If lblClie <> "" Then lblClie = Mid(lblClie, 1, 18)
     TxDocRef = FG.TextMatrix(FG.Row, 6)  'doc ref
     TxDocNroRef = FG.TextMatrix(FG.Row, 7)  'proveedor
     rSql = "select CANUMORD,CACENCOS,CASOLI  from MOVALMCAB  where   CAALMA ='" & VGAlma & "' and CATD='" & tipo & "' AND CANUMDOC='" & Lblnumdoc & "'" '
     Set Adoreg1 = New ADODB.Recordset
     Adoreg1.Open rSql, cConexCom, adOpenDynamic, adLockOptimistic
     If Adoreg1.RecordCount <> 0 Then
        TxOrdCom = IIf(IsNull(Adoreg1(0)), "", Adoreg1(0))
        TxCenCos = IIf(IsNull(Adoreg1(1)), "", Adoreg1(1))
        TxAut = IIf(IsNull(Adoreg1(2)), "", Adoreg1(2))
      End If
     Adoreg1.Close
     CmdGrabarCab.Enabled = True
     CmdGrabarCab.SetFocus
  End If
End Sub

Private Sub CmdGrabar_Click()
  VGEstadomodi = False
  entrodetalle = False
  Unload Me
End Sub

'Modificar
Private Sub command5_Click()
     If MSFlexGrid1.Rows = 1 Then Exit Sub
      VGSeleccion = 2
      contador = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
      codigo = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
      cantidad = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3))
      numitem = contador
      VGRegEnt = IIf(Trim(TxDoc) = "NI", 1, 0)
      actualizastk (codigo)                 'descarga
      VGEstadomodi = True
      VGSeleccion = 2
      VGtipocreacion = 2
      If Not valorizado Then
          FormCreacionSin.Show 1
      Else
'          MsgBox "En esta opción no se puede modificar los precios", vbInformation, "Modificar"
          FormCreacionSin.Show 1
      End If
      VGEstadomodi = False
End Sub

Private Sub Command6_Click()
      If Trim(TxDoc) = "NI" Then
          VGRegEnt = 1
      Else
          VGRegEnt = 0
          valorizado = False
      End If
      numitem = MSFlexGrid1.Rows
      If Not valorizado Then
         VGtipocreacion = 2
         If VGtipocreacion = 2 Then VGSeleccion = 1
         FormCreacionSin.Show 1
      Else
         VGSeleccion = 3
         VGtipocreacion = 2
         FormCreacion.Show 1
     End If
End Sub

Private Sub Command7_Click()
  If Frame1.Visible Then
     limpia
     Frame1.Visible = False
     Frame3.Visible = False
     'Command1_Click
     Command1.Visible = True
  Else
     Unload Me
   End If
End Sub

Private Sub Form_Load()
 Dim db As Database
 Dim rS As Recordset
 Dim rSql As String

  limpia
  central Me
  entrodetalle = False
  VGSoles = True
  VGtipocreacion = 2      'para identificar de que pantalla viene
  DTPicker3 = Date
  DTPicker2 = DateAdd("m", -2, Date)
  Inicializa
  LblCosto = ""
End Sub

Public Sub limpia()
  Text1 = ""
  TxDoc = ""
  TxTransa = ""
  TxProv = ""
  TxDocRef = ""
  TxDocNroRef = ""
  TxAlmDes = ""
  TxOrdCom = ""
  lblClie = ""
  Lblnumdoc = ""
  lblProv = ""
  Lbltransa = ""
  TxAut = ""
  TxCenCos = ""
  
  inicializarFlex
End Sub

Private Sub inicializarFlex()
  MSFlexGrid1.Clear
  MSFlexGrid1.Rows = 1
  MSFlexGrid1.TextMatrix(0, 0) = " CODIGO "
  MSFlexGrid1.TextMatrix(0, 1) = " DESCRIPCION"
  MSFlexGrid1.TextMatrix(0, 2) = " SERIE \ LOT"
  MSFlexGrid1.TextMatrix(0, 3) = " CANTIDAD"
  MSFlexGrid1.TextMatrix(0, 4) = " UNIDAD ING"
  MSFlexGrid1.TextMatrix(0, 5) = " COSTO UNIT"
  MSFlexGrid1.TextMatrix(0, 6) = " NRO "
  MSFlexGrid1.ColAlignment(0) = 1
End Sub

Function transa() As String
 Dim db As Database
 Dim rS As Recordset
 Dim rSql As String
'  VGAlma = "01"
  rSql = "select  TT_DESCRI FROM TabTransa where TT_CODMOV= '" & FG.TextMatrix(FG.Row, 2) & "'" '
  Set db = Workspaces(0).OpenDatabase(cRuta2)
  Set rS = db.OpenRecordset(rSql, dbOpenSnapshot)
  If Not rS.EOF Then
    transa = rS(0)
  End If
  
End Function
Function prove() As String

 Dim db As Database
 Dim rS As Recordset
 Dim rSql As String
'  VGAlma = "01"
  rSql = "select PRVCNOMBRE FROM maeprov where PRVCCODIGO= '" & FG.TextMatrix(FG.Row, 4) & "'" '
  Set db = Workspaces(0).OpenDatabase(cRuta2)
  Set rS = db.OpenRecordset(rSql, dbOpenSnapshot)
  If Not rS.EOF Then
    prove = rS(0)
  End If
End Function

Private Sub grabacabecera()
 Dim criterio As String
 Dim uSql As String
 Dim rSql As String
 Dim adodc1 As ADODB.Recordset
 Dim Adoreg1 As ADODB.Recordset
   
  'Desea grabar el registro
  On Error GoTo GrabErr
   If TxDoc.text <> "" Then
        Set adodc1 = New ADODB.Recordset
        adodc1.Open "Select * From MovAlmCab Where CANUMDOC = '" & FG.TextMatrix(FG.Row, 1) & "' And  CATD = '" & tipo & "' And CAALMA = '" & VGAlma & "'", cConexCom, adOpenDynamic, adLockOptimistic
        If adodc1.RecordCount > 0 Then
          adodc1("CAFECDOC") = DTPicker1.Value
          If Trim(TxDocNroRef.text) <> "" Then adodc1("CARFNDOC") = Trim(TxDocNroRef.text)
          If Trim(TxTransa) <> "" Then adodc1("CACODMOV") = Mid$(UCase$(TxTransa.text), 1, 2)
          If tipo = "NI" Then
             If Not IsNull(adodc1("CACODMON")) Then Text1 = Trim(adodc1("CACODMON"))
          Else
              Text1.Visible = False
              Label11.Visible = False
          End If
          
          If Trim(TxProv) <> "" Then
             adodc1("CACODPRO") = Trim(UCase$(TxProv.text))
          Else
             adodc1("CACODPRO") = " "
          End If
    
          If Trim(TxDocRef) <> "" Then
             adodc1("CARFTDOC") = LTrim(UCase$(TxDocRef))
          Else
             adodc1("CARFTDOC") = " "
          End If
           If Trim(TxDocNroRef) <> "" Then
             adodc1("CARFNDOC") = LTrim(UCase$(TxDocNroRef))
          Else
             adodc1("CARFNDOC") = " "
          End If
          
          If Trim(TxClie) <> "" Then adodc1("CACODCLI") = LTrim(UCase$(TxClie))
             rSql = "Select CCodCli,CNumRuc,CNomCli from MaeCli where CCodCli ='" & Trim(TxClie) & "'"
             Set Adoreg1 = New ADODB.Recordset
             Adoreg1.Open rSql, cConexCom, adOpenDynamic, adLockOptimistic
             If Adoreg1.RecordCount <> 0 Then
                If Not IsNull(Adoreg1("CCODCLI")) Then adodc1("CACODCLI") = Adoreg1("CCODCLI")
                If Not IsNull(Adoreg1("CNumRuc")) Then adodc1("CARUC") = Adoreg1("CNumRuc")
                If Not IsNull(Adoreg1("CCODCLI")) Then adodc1("CANOMCLI") = Adoreg1("CCODCLI")
             End If
             Adoreg1.Close

         If Trim(TxOrdCom.text) <> "" Then
             adodc1("CANUMORD") = LTrim(UCase$(TxOrdCom.text))
          Else
             adodc1("CANUMORD") = " "
          End If
          If Trim(TxAut) <> "" Then
             adodc1("CASOLI") = LTrim(UCase$(TxAut))
          Else
             adodc1("CASOLI") = " "
          End If
          ' Adodc1("CAHORA") = Format(Time, "hh:mm:ss")
          If Trim(TxCenCos) <> "" Then adodc1("CACENCOS") = TxCenCos
          If Trim(TxAlmDes) <> "" Then
             adodc1("CARFALMA") = Mid$(UCase$(TxAlmDes), 1, 2)
             'grabaalmacen
          End If
          adodc1.Update
       End If
   adodc1.Requery
   End If
   adodc1.Close
'   Set adodc1 = Nothing
'   IF FechaAnt<>dtpicker3.Month
'   Set adodc1 = New ADODB.Recordset
'   adodc1.Open "Select * From MovAlmdet Where DENUMDOC = '" & FG.TextMatrix(FG.Row, 1) & "' And  DETD = '" & tipo & "' And DEALMA = '" & VGAlma & "'", cConexCom, adOpenDynamic, adLockOptimistic
'   While Not adodc1.EOF
     
      
  
   Exit Sub
GrabErr:
    MsgBox err.Description
End Sub
Private Sub grabadetalle()
 Dim rS As Recordset
 Dim rSql As String
 Dim codigo As String
 Dim item As Integer
 Dim criterio As String
 Dim cadena As String
 
 On Error GoTo GrabErr
   If MSFlexGrid1.Rows = 1 Then Exit Sub
   contador = 1
   While MSFlexGrid1.Rows > contador
         codigo = MSFlexGrid1.TextMatrix(contador, 0)
         item = contador
         criterio = "DECODIGO = " & Chr$(34) + codigo + Chr$(34)
         criterio = criterio + "and  DETD = " & Chr$(34) + TxDoc + Chr$(34)
         criterio = criterio + "and  DEITEM = " & Chr$(34) + item + Chr$(34)
         criterio = criterio + "and  DENUMDOC = " & Chr$(34) + Lblnumdoc + Chr$(34)
         criterio = criterio + "and  DEALMA = " & Chr$(34) + VGAlma + Chr$(34)
         Rs2.FindFirst criterio
         If Rs2.NoMatch Then
            Rs2.Edit
         Else
             Rs2.AddNew
             Rs2("dealma") = VGAlma
             Rs2("DEITEM") = numitem + 1
             Rs2("DECODIGO") = MSFlexGrid1.TextMatrix(contador, 0)   ' Format(MSFlexGrid1.TextMatrix(contador, 0), "00000000")
             'rs("DEDESCRI") = MSFlexGrid1.TextMatrix(contador, 1)
             Rs2("detd") = TxDoc
             Rs2("denumdoc") = Lblnumdoc
         End If
         cantidad = MSFlexGrid1.TextMatrix(contador, 3)
         Rs2("decantid") = cantidad
         grabastk
        If MSFlexGrid1.TextMatrix(contador, 5) <> "" Then    'si tiene precio
             Rs2("DEPRECIO") = Val(MSFlexGrid1.TextMatrix(contador, 5)) * VGTipCamb '******el precio
'        ElseIf TT_VALOR = "F" Then
'          rs2("DEPRECIO") = precioprom  '******'valorizacion de precio prom
        Else
             Rs2("DEPRECIO") = 0
        End If
        'mejorar a una funcion
        
        If MSFlexGrid1.TextMatrix(contador, 10) = "S" Then
             grabaserie VGAlma, cadena
             Rs2("DESERIE") = MSFlexGrid1.TextMatrix(contador, 2)
        End If
        If MSFlexGrid1.TextMatrix(contador, 10) = "N" Then
             grabalote VGAlma, cadena
             Rs2("DELOTE") = MSFlexGrid1.TextMatrix(contador, 2)
        End If
        Rs2.Update
        contador = contador + 1
   Wend
   rsstock.Close

GrabErr:
 Exit Sub
End Sub
Function existe_clie(text As TextBox) As String
  Dim rSql As String
  Dim rS As Recordset
   rSql = "SELECT CNOMCLI FROM maecli where CCODCLI= '" & text & "'"
   Set rS = VGBaseDatos.OpenRecordset(rSql, dbOpenSnapshot)
   If Not rS.EOF Then 'existe
      existe_clie = rS(0)
   Else
      existe_clie = ""
  End If
  rS.Close
End Function
Private Sub actualizastk(codigo As String)
Dim criterio As String
Dim canttemp As Double
On Error GoTo err
   criterio = " STCODIGO = " & Chr$(34) + codigo + Chr$(34)
   criterio = criterio + "and  STALMA = " & Chr$(34) + VGAlma + Chr$(34)
   rsstock.FindFirst criterio
   rsstock.Edit
   If tipo = "NI" Then
      'descuenta lo que hay en stock
      canttemp = rsstock("stskdis")
      VGTipCamb = 1
      rsstock("stskdis") = rsstock("stskdis") - cantidad  'para actualizar el precio prom
      If IsNumeric(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)) > 0 Then rsstock("stkprepro") = (rsstock("stkprepro") * canttemp - cantidad * Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) * VGTipCamb)) / (canttemp - cantidad)
   Else
     ' aumenta el stock  debido a que es una salida a modificar
      rsstock("stskdis") = rsstock("stskdis") + cantidad
   End If
   rsstock.Update
   If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) <> "" Then
      serie_lote = MSFlexGrid1.TextMatrix(contador, 2)
      If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10) = "S" Then actserie (codigo)
      If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10) = "N" Then actlote (codigo)
   End If
   actvalmes 'descarga al moresmes
   Exit Sub
err:
'Resume Next
    MsgBox err.Description
End Sub

Private Sub actlote(codigo As String)
Dim uSql As String
Dim nuevo_stk As Double
Dim rSql As String
Dim rS As Recordset

    rSql = "select STSLKDIS FROM STKLOTE where   STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & codigo & "' and STSLOTE= '" & serie_lote & "'" '
    Set rS = VGBaseDatos.OpenRecordset(rSql, dbOpenSnapshot)
    If Not rS.EOF Then
          nuevo_stk = IIf(tipo = "NI", rS(0) - cantidad, rS(0) + cantidad)
          uSql = "Update STKLOTE set STSLKDIS = " & nuevo_stk & " WHERE  STSALMA= '" & VGAlma & "' and STSCODIGO='" & codigo & "'AND STSLOTE='" & serie_lote & "'"
          cConexCom.Execute uSql
    End If
     rS.Close
End Sub

Private Sub actserie(codigo As String)
Dim uSql As String
Dim Serie As String
Dim Valor As Integer
Dim rS As Recordset
Dim rSql As String

    rSql = "select STSSKDIS FROM STKSERI where   STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & codigo & "' and STSSERIE= '" & serie_lote & "'" '
    Set rS = VGBaseDatos.OpenRecordset(rSql, dbOpenSnapshot)
    If Not rS.EOF Then
       Valor = IIf(tipo = "NI", 0, 1)
       uSql = "update STKSERI set STSSKDIS = " & Valor & " WHERE  STSALMA='" & VGAlma & "' and STSCODIGO='" & codigo & "'AND STSSERIE='" & serie_lote & "'"
       cConexCom.Execute uSql
    End If
   
End Sub

Private Sub muestradetalle()
Dim Rsql1 As String
Dim Adoreg1 As ADODB.Recordset
Dim ser_lot As String
Dim dato As String
Dim n As String
     '                   0      1        2       3         4         5       6      7
     Rsql1 = "select DECODIGO,deDESCRI, AUNIDAD, DECANTID, DEPRECIO,DESERIE,DELOTE,deitem  from MovAlmDet n ,MaeArt where  DEALMA ='" & VGAlma & "' AND DETD = '" & TxDoc & "' AND ACODIGO= DECODIGO AND DENUMDOC ='" & Lblnumdoc & "'  ORDER BY n.DEITEM "  '
     Set Adoreg1 = New ADODB.Recordset
     Adoreg1.Open Rsql1, cConexCom, adOpenDynamic, adLockOptimistic
     If Adoreg1.RecordCount = 0 Then
         MsgBox "No se grabo ningun detalle", vbInformation, "Aviso"
         Exit Sub
     End If
     Adoreg1.MoveFirst
     MSFlexGrid1.Rows = 1
     If Adoreg1("DEPRECIO") <> 0 Then
          valorizado = True
     Else
          valorizado = False
     End If
     'N = 0
     While Not Adoreg1.EOF
        If IsNull(Adoreg1(5)) And IsNull(Adoreg1(6)) Then
              ser_lot = ""
              dato = "X"
        ElseIf Not IsNull(Adoreg1(5)) Then
              ser_lot = Adoreg1(5)
              dato = "S"
        Else
              ser_lot = Adoreg1(6)
              dato = "N"
        End If
        'N = N + 1                'COD                           DES                         SER                            CANT                                                               UND                                 PRECIO                              6             7            8           9            10
        MSFlexGrid1.AddItem (Adoreg1(0) & vbTab & Adoreg1(1) & vbTab & ser_lot & vbTab & Format(Adoreg1(3), "##0.00") & vbTab & Adoreg1(2) & vbTab & Format(Adoreg1(4), "####0.0000") & vbTab & Adoreg1(7) & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & dato & vbTab)
        
        Adoreg1.MoveNext
     Wend
     Adoreg1.Close
End Sub

Private Sub actvalmes()
  Dim cadena As String
  Dim criterio As String
 
  Dim Adoreg1 As ADODB.Recordset
  Dim rSql As String
  Dim Cantent As Double
  Dim Cantsal As Double
  Dim mespro As String
  Dim uSql As String
  On Error GoTo err
   mespro = Year(DTPicker1) & Format(Month(DTPicker1), "00")
   cadena = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) 'codigo del art
   rSql = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & cadena & "'"  '
   
   Set Adoreg1 = New ADODB.Recordset
   Adoreg1.Open rSql, cConexCom, adOpenDynamic, adLockOptimistic
   If Adoreg1.RecordCount <> 0 Then
     'descargo en la cuenta respectiva
      If tipo = "NI" Then
          Cantent = Adoreg1(0) - cantidad
          uSql = "Update MoResMes set SMCANENT = " & Cantent & " where SMALMA='" & VGAlma & "'  and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
      Else
         Cantsal = Adoreg1(1) - cantidad
         uSql = "Update MoResMes set SMCANSAL = " & Cantsal & " where SMALMA='" & VGAlma & "' and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
      End If
   Else
      If tipo = "NI" Then
          Cantent = cantidad
          Cantsal = 0
      Else
          Cantsal = cantidad
          Cantent = 0
      End If
      uSql = "insert into MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL) VALUES ('" & VGAlma & "','" & cadena & "','" & mespro & "' ," & Cantent & "," & Cantsal & ") "
   End If
   cConexCom.Execute uSql
   Adoreg1.Close
   Exit Sub
err:
     MsgBox err.Description
End Sub


Public Sub grabastk()
  Dim cadena As String
  Dim criterio As String
  Dim canttemp
  Dim precioprom
  On Error GoTo GrabErr
   cadena = MSFlexGrid1.TextMatrix(contador, 0)
   criterio = " STCODIGO = " & Chr$(34) + cadena + Chr$(34)
   criterio = criterio + "and  STALMA = " & Chr$(34) + VGAlma + Chr$(34)
   rsstock.FindFirst criterio
   If Not rsstock.NoMatch Then      'si existe el articulo
            rsstock.Edit
            canttemp = rsstock("STSKDIS")     ' revisar si validar en creacion
            If tipo = "NI" Then
                rsstock("STSKDIS") = rsstock("STSKDIS") + cantidad
                'aqui actualiza
                If Not IsNull(rsstock("STKPREPRO")) Then
                   precioprom = rsstock("STKPREPRO")
                   If MSFlexGrid1.TextMatrix(contador, 5) <> "" Then rsstock("STKPREPRO") = (precioprom * canttemp + cantidad * Val(MSFlexGrid1.TextMatrix(contador, 7) * VGTipCamb)) / (canttemp + cantidad)
                End If
            End If
     Else
         If tipo = "NI" Then
            rsstock.AddNew                  'existe
            rsstock("STALMA") = VGAlma   '"01"
            rsstock("STCODIGO") = MSFlexGrid1.TextMatrix(contador, 0)
            rsstock("STSKDIS") = cantidad
            If MSFlexGrid1.TextMatrix(contador, 5) <> "" Then rsstock("STKPREPRO") = Val(MSFlexGrid1.TextMatrix(contador, 7))    '******el precio
          End If
  End If
  rsstock.Update
  ValMes  'para la valorizacion
  Exit Sub
GrabErr:
    MsgBox err.Description
End Sub

Private Sub ValMes()
  Dim cadena As String
  Dim criterio As String
  Dim Adoreg1 As ADODB.Recordset
  Dim rSql As String
  Dim Cantent As Double
  Dim Cantsal As Double
  Dim mespro As String
  Dim uSql As String
  On Error GoTo err
   mespro = Year(DTPicker1) & Format(Month(DTPicker1), "00")
   cadena = MSFlexGrid1.TextMatrix(contador, 0) 'codigo del art
   rSql = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & cadena & "'"  '
   
   Set Adoreg1 = New ADODB.Recordset
   Adoreg1.Open rSql, cConexCom, adOpenDynamic, adLockOptimistic
    If Adoreg1.RecordCount <> 0 Then
       If tipo = "NI" Then
          Cantent = Adoreg1(0) + cantidad
          uSql = "Update MoResMes set SMCANENT = " & Cantent & " where SMALMA='" & VGAlma & "' and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
       Else
          Cantsal = Adoreg1(1) + cantidad
          uSql = "Update MoResMes set SMCANSAL = " & Cantsal & " where SMALMA='" & VGAlma & "' and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
       End If
   Else
       If tipo = "NI" Then
          Cantent = cantidad
          Cantsal = 0
       Else
         Cantsal = cantidad
         Cantent = 0
       End If
       uSql = "insert into MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL) VALUES ('" & VGAlma & "','" & cadena & "','" & mespro & "' ," & Cantent & "," & Cantsal & ") "
   End If
   cConexCom.Execute uSql
   Adoreg1.Close
   Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub grabalote(alma As String, codigo As String)
Dim uSql As String
Dim lote As String
Dim nuevo_stk As Double
Dim rSql As String
Dim rS As Recordset
Dim fecfab As Date
Dim fecven As Date
    On Error GoTo err
    If (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" Then
      fecfab = MSFlexGrid1.TextMatrix(contador, 9)
    End If
    If (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
      fecven = MSFlexGrid1.TextMatrix(contador, 8)
    End If
    lote = MSFlexGrid1.TextMatrix(contador, 2)
    rSql = "select STSLKDIS FROM STKLOTE where  STSALMA= '" & alma & "' and  STSCODIGO= '" & codigo & "' and STSLOTE= '" & lote & "'" '
    Set rS = VGBaseDatos.OpenRecordset(rSql, dbOpenSnapshot)
    If Not rS.EOF Then
       If tipo = "NI" Then
         nuevo_stk = rS(0) + cantidad
       Else
         nuevo_stk = rS(0) - cantidad
       End If
       uSql = "Update STKLOTE set STSLKDIS = " & nuevo_stk & " WHERE  STSALMA='" & alma & "' and STSCODIGO='" & codigo & "'AND STSLOTE='" & lote & "'"
    Else
            If (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) = "__/__/____" Then
                 fecfab = MSFlexGrid1.TextMatrix(contador, 9)
                 uSql = "insert into STKLOTE (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB) VALUES ('" & alma & "','" & codigo & "','" & lote & "' ," & cantidad & " ,#" & Format(fecfab, "MM/DD/YYYY") & "#) "
             ElseIf (MSFlexGrid1.TextMatrix(contador, 9)) = "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
                 fecven = MSFlexGrid1.TextMatrix(contador, 8)
                 uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECVEN)  VALUES ('" & alma & "','" & codigo & "','" & lote & "' ," & cantidad & " ,#" & Format(fecven, "MM/DD/YYYY") & "#) " 'SIN FECFAB
             ElseIf (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
                 uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB,STSFECVEN)  VALUES ('" & alma & "','" & codigo & "','" & lote & "' ," & cantidad & " ,#" & Format(fecfab, "MM/DD/YYYY") & "#,#" & Format(fecven, "MM/DD/YYYY") & "#) "
             Else
                 uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS)  VALUES ('" & alma & "','" & codigo & "','" & lote & "' ," & cantidad & ") "
             End If
    End If
    cConexCom.Execute uSql
    Exit Sub
err:
     MsgBox err.Description
End Sub

Private Sub grabaserie(alma As String, codigo As String)
Dim uSql As String
Dim Serie As String
Dim Valor As Integer
Dim rS As Recordset
Dim rSql As String
Dim fecfab As Date
Dim fecven As Date
On Error GoTo err
    'fecfab = " " '  MSFlexGrid1.TextMatrix(contador, 8)
    'fecven = " " 'MSFlexGrid1.TextMatrix(contador, 9)
    Serie = MSFlexGrid1.TextMatrix(contador, 2)
    rSql = "select STSSKDIS FROM STKSERI where  STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & codigo & "' and STSSERIE= '" & Serie & "'" '
    Set rS = VGBaseDatos.OpenRecordset(rSql, dbOpenSnapshot)
    If Not rS.EOF Then
       Valor = IIf(tipo = "NI", 1, 0)
       uSql = "update STKSERI set STSSKDIS = " & Valor & " WHERE  STSALMA='" & alma & "' and STSCODIGO='" & codigo & "'AND STSSERIE='" & Serie & "'"
    Else
       uSql = "insert into STKSERI (STSALMA,STSCODIGO,STSSERIE,STSSKDIS)  VALUES ('" & alma & "','" & codigo & "','" & Serie & "' ,' ',' ','1') "
    End If
    cConexCom.Execute uSql
     Exit Sub
err:
      MsgBox err.Description, vbExclamation, "Aviso"
End Sub


Private Sub TxAlmDes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}'"
  KeyAscii = 0
End If
End Sub


Private Sub TxAut_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}'"
  KeyAscii = 0
End If
End Sub


Private Sub TxCenCos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}'"
  KeyAscii = 0
End If
End Sub

Private Sub TxClie_DblClick()
   FrmAyuCliente.Show 1
  TxClie = FrmAyuCliente.cCod
  lblClie = FrmAyuCliente.cNom
End Sub

Private Sub TxClie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    TxClie_DblClick
ElseIf KeyCode = 8 Then
    lblClie = ""
End If
End Sub

Private Sub TxClie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}'"
  KeyAscii = 0
End If
End Sub

Private Sub TxDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}'"
  KeyAscii = 0
End If
End Sub

Private Sub TxDocNroRef_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}'"
  KeyAscii = 0
End If
End Sub


Private Sub TxOrdCom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}'"
  KeyAscii = 0
End If
End Sub

Private Sub TxProv_DblClick()
    Dim Adodc3 As ADODB.Recordset
    Set Adodc3 = New ADODB.Recordset
    Adodc3.Open "SELECT PRVCCODIGO,PRVCNOMBRE FROM MAEPROV", cConexCom, adOpenStatic, adLockOptimistic
    frmReferencia.conectar Adodc3, "SELECT PRVCCODIGO,PRVCNOMBRE FROM MAEPROV"
    frmReferencia.Label1.Caption = "Proveedores"
    frmReferencia.Show vbModal
    Adodc3.Close
    If vGUtil(1) <> "" Then TxProv = (vGUtil(1))
    If vGUtil(2) <> "" Then lblProv = (vGUtil(2))
    
End Sub

Private Sub TxProv_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    TxProv_DblClick
ElseIf KeyCode = 8 Then
    lblProv = ""
End If
End Sub

Private Sub TxProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
   KeyAscii = 0
End If
End Sub

Function prove1(txt As TextBox) As String
 Dim Adoreg As ADODB.Recordset
   Set Adoreg = New ADODB.Recordset
   Adoreg.Open "select PRVCNOMBRE FROM maeprov where PRVCCODIGO= '" & txt & "'", cConexCom, adOpenDynamic, adLockOptimistic

   If Not Adoreg.EOF Then
      prove1 = Adoreg(0)
   Else
     MsgBox "El codigo del proveedor no existe !", vbExclamation, "Error"
     prove1 = ""
  End If
  Adoreg.Close
End Function

Function transa1(text As TextBox) As String
 Dim Adoreg As ADODB.Recordset
  Set Adoreg = New ADODB.Recordset
  
  If Trim(tipo) = UCase("xx") Then
     Adoreg.Open "select  TT_DESCRI FROM TabTransa where TT_CODMOV= '" & text & "'", cConexCom, adOpenDynamic, adLockOptimistic
  Else
      Adoreg.Open "select  TT_DESCRI FROM TabTransa where TT_CODMOV= '" & text & "' and TT_TIPMOV ='" & IIf(tipo = "NI", "I", "S") & "'", cConexCom, adOpenDynamic, adLockOptimistic
  End If
  If Not Adoreg.EOF Then
    transa1 = Adoreg(0)
  Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
    transa1 = ""
  End If
   Adoreg.Close
End Function



Private Sub TxTransa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Then
   Lbltransa = ""
End If
End Sub

Private Sub TxTransa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxTransa = UCase(TxTransa)
    SendKeys "{tab}'"
    KeyAscii = 0
End If
End Sub

Private Sub Combo1_Click()
    FG.Col = Combo1.ListIndex + 1
    FG.Sort = 5
End Sub

Private Sub Txtbuscar_Change()
Dim i As Integer
Dim n As Integer
   n = Combo1.ListIndex + 1
   If TxtBuscar <> "" Then
      For i = 1 To FG.Rows - 1
          If UCase(Left(FG.TextMatrix(i, n), Len(TxtBuscar))) = UCase(Trim(TxtBuscar)) Then
              Exit For
          End If
      Next i
      If i >= FG.Rows Then
            FG.HighLight = flexHighlightNever
      Else
            FG.HighLight = flexHighlightAlways
            FG.TopRow = i
            FG.Row = i
            FG.Col = 0
            FG.ColSel = FG.Cols - 1
      End If
   End If
End Sub

Private Sub Inicializa()
    FG.FormatString = "Tipo Doc.|Numero de Doc| Tr| Fecha | Proveedor|Cliente|Td REF|Num.Doc Ref."
    FG.Row = 0
    FG.Cols = 9
    FG.ColWidth(0) = 800
    FG.ColWidth(1) = 1500
    FG.ColWidth(2) = 800
    FG.ColWidth(3) = 1000
    FG.ColWidth(4) = 1300
    FG.ColWidth(5) = 1300
    FG.ColWidth(6) = 800
    FG.ColWidth(7) = 1500
    FG.ColWidth(8) = 2
    'Grilla del detalle
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Cols = 12
    MSFlexGrid1.ColWidth(0) = 2000
    MSFlexGrid1.ColWidth(1) = 4500
    MSFlexGrid1.ColWidth(2) = 1100
    MSFlexGrid1.ColWidth(3) = 1100
    MSFlexGrid1.ColWidth(4) = 1300
    MSFlexGrid1.ColWidth(5) = 1200
    MSFlexGrid1.ColWidth(6) = 1200
    MSFlexGrid1.ColWidth(7) = 2
    MSFlexGrid1.ColWidth(8) = 2
    MSFlexGrid1.ColWidth(9) = 2
    MSFlexGrid1.ColWidth(10) = 2
    MSFlexGrid1.ColWidth(11) = 2
    Frame3.Visible = False
    Frame1.Visible = False

End Sub

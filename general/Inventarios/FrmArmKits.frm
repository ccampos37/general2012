VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmArmKits 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Armado de Kids"
   ClientHeight    =   5928
   ClientLeft      =   2160
   ClientTop       =   3288
   ClientWidth     =   8568
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5928
   ScaleWidth      =   8568
   Begin VB.CommandButton Command2 
      Caption         =   "&Agregar"
      Height          =   735
      Left            =   1080
      Picture         =   "FrmArmKits.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4905
      Width           =   720
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   405
      Top             =   5115
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Eliminar"
      Height          =   735
      Left            =   3285
      Picture         =   "FrmArmKits.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox TxCantidad 
      Height          =   288
      Left            =   1548
      TabIndex        =   26
      Text            =   "TxCantidad"
      Top             =   2700
      Visible         =   0   'False
      Width           =   768
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Limpiar"
      Height          =   735
      Left            =   4395
      Picture         =   "FrmArmKits.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   735
      Left            =   2160
      Picture         =   "FrmArmKits.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   8295
      Begin VB.TextBox TxCambio 
         Height          =   285
         Left            =   5880
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox CmbMoneda 
         Height          =   315
         ItemData        =   "FrmArmKits.frx":0FD0
         Left            =   1680
         List            =   "FrmArmKits.frx":0FDA
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox TxTransa 
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   5880
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   3
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1320
         Width           =   1995
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         MaxLength       =   3
         TabIndex        =   6
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5880
         MaxLength       =   10
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2561
         _ExtentY        =   508
         _Version        =   393216
         Format          =   24903681
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo de Cambio"
         Height          =   255
         Left            =   4680
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda "
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Doc."
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Transaccion"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Num. Doc"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   4680
         TabIndex        =   23
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Tip Doc Ref"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Orden Compra"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Autorizacion"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   4680
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Num. Ref"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   4680
         TabIndex        =   19
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label lbltrans 
         Caption         =   "lbltrans"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   630
         Width           =   5055
      End
      Begin VB.Label lbltipref 
         Caption         =   "lbltipref"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   990
         Width           =   2295
      End
      Begin VB.Label lblauto 
         Caption         =   "lblauto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   16
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label LblCC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3195
         TabIndex        =   15
         Top             =   2460
         Width           =   2145
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   5490
      Picture         =   "FrmArmKits.frx":0FE6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4935
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2172
      Left            =   192
      TabIndex        =   9
      ToolTipText     =   "Doble Click o F1 para la ayuda"
      Top             =   2412
      Width           =   8292
      _ExtentX        =   14626
      _ExtentY        =   3831
      _Version        =   393216
      RowHeightMin    =   290
      Appearance      =   0
   End
End
Attribute VB_Name = "FrmArmKits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    VGSeleccion = 1 Significa que es seleccion con frame de tipo de cambio
'    VGSeleccion = 2 Significa que es seleccion sin frame de tipo de cambio para modificar el contenido
'    VGSeleccion = 3 Significa que es seleccion sin frame de tipo de cambio para agregar item
'    VGform significa con formulario esta trabajando
'     text9    autorizado
'     text10  cencos
'     text11  almacen
Option Explicit
Dim Adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim Adodc22 As ADODB.Recordset
Dim Adodc3 As ADODB.Recordset
Dim ContaSalida As String

Dim db As Database
Dim nument As Long
Dim precioprom As Double
Dim cantidad As Double
Dim canttemp As Double
Dim Campo As String * 2
Dim contador As Integer
Dim num As Integer
Dim TT_CONTADOR As Integer
Dim TT_VALOR As String * 1
Dim cadena As String
Dim alma As String
Dim Tipo As String * 2
Dim dato As String
Dim NumDoc As String
Dim Codigo2 As String

Private Sub CmbMoneda_Click()
If CmbMoneda.ListIndex = 0 Then
    VGSoles = True
    TxCambio.Enabled = False
    VGTipCamb = 1
Else
    VGSoles = False
    TxCambio.Enabled = True
End If
End Sub

Private Sub CmbMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

'Eliminar
Private Sub Command1_Click()
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
    End If
End If
End Sub

Private Sub Command2_Click()
 MSFlexGrid1_DblClick
End Sub

'Limpiar
Private Sub Command3_Click()
TxTransa = "": Text6 = "": Text8 = "": Text4 = ""
Text1 = "": Text9 = "": TxCantidad = "": TxCambio = "0": TxCambio.Enabled = False
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 1
inicializaFG
End Sub

'****************************** Graba LA GUIA****************
Private Sub CmdGrabar_Click()
Dim criterio As String, cadena As String, cadena1 As String, cadena2 As String
Dim rpta As Integer, FACTOR As Double, uSql As String, nIt As Integer
Dim cSel1 As ADODB.Recordset
On Error GoTo GrabErr

If Trim(TxTransa) = "" Then
    MsgBox "Debe Ingresar el Movimiento", vbInformation, "Mensaje"
    TxTransa.SetFocus
    Exit Sub
End If
If TxCambio.Enabled Then
    If Val(TxCambio) = 0 Then
        MsgBox "Ingrese Tipo de Cambio", vbInformation, "Mensaje"
        TxCambio.SetFocus: Exit Sub
    Else
        VGTipCamb = TxCambio
    End If
End If
cantidad = 0
If MSFlexGrid1.Rows = 1 Then
      MsgBox "No se puede grabar,debe adicionar registro", vbInformation, mensaje1
      Exit Sub
End If
'Numeracion
If Trim(Text4) = "" Then muestra

If Not IsNumeric(Text4) Then
     MsgBox "Numero de Documento no consecutivo", vbExclamation, "Aviso"
     Exit Sub
End If
Text4 = Format(Text4, String(10, "0"))
If existe_numdoc(Text4) Then Exit Sub
Screen.MousePointer = 11

'***
'Verificar I/S
'RMM****************************************
 Call VerificarSTOCK
'RMM****************************************

Call grabacabecera("I", "NI", Text4, 1) 'Entrada
Call grabacabecera("S", "NS", ContaSalida, 0) 'Salida
'***


FACTOR = 1
contador = 1
'Graba detalle
NumDoc = Text4
 nIt = 0
While MSFlexGrid1.Rows > contador
     cantidad = MSFlexGrid1.TextMatrix(contador, 2)
     If (IIf(VGRegEnt = 1, True, True)) Then      'verificastk
       cadena = MSFlexGrid1.TextMatrix(contador, 0)
       cantidad = 0
       If Not VGActualizar Then
              Adodc2.AddNew
       Else
              criterio = "DECODIGO = '" & cadena & "'"
              criterio = criterio + " and  DEALMA = '" & VGAlma & "'"
              If Not Adodc2.EOF Then Adodc2.Filter = criterio
       End If
       Adodc2("DEALMA") = VGAlma
       Adodc2("DETD") = "NI"
       Adodc2("DENUMDOC") = Text4.text
       Adodc2("DEITEM") = contador
       Adodc2("DECODIGO") = MSFlexGrid1.TextMatrix(contador, 0)
       Adodc2("DEDESCRI") = MSFlexGrid1.TextMatrix(contador, 1)
       cantidad = Val(MSFlexGrid1.TextMatrix(contador, 2))
       Adodc2("DECANTID") = cantidad
       'Data2.Recordset("DEUNIDAD") = MsFlexGrid1.TextMatrix(contador, 4)
       If MSFlexGrid1.TextMatrix(contador, 2) <> "" Then 'Cantidad Ingresada
            Call grabastk(MSFlexGrid1.TextMatrix(contador, 0), 1, Val(MSFlexGrid1.TextMatrix(contador, 3)))
            
            If Trim(MSFlexGrid1.TextMatrix(contador, 3)) <> "" Then    'si tiene precio
                Adodc2("DEPRECIO") = Val(MSFlexGrid1.TextMatrix(contador, 3)) '* VGTipCamb '******el precio
                Adodc2("DETIPCAM") = VGTipCamb
            'ElseIf (TT_VALOR = "V" And VGRegEnt = 0) Or Text10.Visible Then  'SALIDA VALORIZADA  0 - SALIDA,1 - ENTRADA, text10 indica salida x CC
            '    Adodc2("DEPRECIO") = precioprom  '******'valorizacion de precio prom
            Else
                Adodc2("DEPRECIO") = 0
            End If
            alma = VGAlma
            '****
            Set cSel1 = New ADODB.Recordset
            cSel1.Open "Select * From Kits Where CodKit = '" & MSFlexGrid1.TextMatrix(contador, 0) & "'", cConexCom, adOpenStatic
            If cSel1.RecordCount > 0 Then
               
                Do While Not cSel1.EOF
                    nIt = nIt + 1
                    If Existe(1, cSel1("codart"), "kits", "codkit", False) Then
                      MsgBox "No se puede armar Kits de Kits", vbInformation, "Aviso"
                       nIt = nIt - 1
                    Else
                      cantidad = cSel1("canart") * Val(MSFlexGrid1.TextMatrix(contador, 2))
                      'Call grabastk(cSel1("codart"), 2, 0)
                      Call GrabarDetKit(cSel1("CODART"), "NS", nIt, Devolver_Dato(1, cSel1("CODART"), "MaeArt", "Acodigo", False, "Adescri"), (cSel1("CanArt") * Val(cantidad)), 0)
                    End If
                    cSel1.MoveNext
                    If cSel1.EOF Then Exit Do
                Loop
            End If
            cSel1.Close
            
       End If
       Adodc2.Update
       Adodc2.Filter = ""
     End If
     contador = contador + 1
Wend
Adodc2.Requery

Dim rSql As String
Dim rs As Recordset
rSql = "select  stcodigo from  StkArt  where  STALMA = '" & VGAlma & "' "
Set db = Workspaces(0).OpenDatabase(cRuta2)
Set rs = db.OpenRecordset(rSql, dbOpenSnapshot)
If Not rs.EOF Then
     FormPrincipal.Men_TraCor = True
     FormPrincipal.Men_TraVal = True
     FormPrincipal.mnucons = True
     FormPrincipal.mnurep = True
End If
rpta = MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso")
If rpta = vbYes Then
    imprimir
End If
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 1
inicializar
VGSoles = True
VGTipCamb = 1
Screen.MousePointer = 1
Exit Sub
GrabErr:
  MsgBox Err.Description, vbExclamation, "Error"
  Screen.MousePointer = 1
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub Form_Activate()
If VGAutomatico Then Text4.Enabled = False
End Sub

Private Sub Form_Load()
central FrmArmKits

Set Adodc1 = New ADODB.Recordset
Adodc1.Open "Select * From MovAlmCab", cConexCom, adOpenDynamic, adLockOptimistic

Set Adodc2 = New ADODB.Recordset
Adodc2.Open "Select * From MovAlmDet", cConexCom, adOpenDynamic, adLockOptimistic

Set Adodc22 = New ADODB.Recordset
Adodc22.Open "Select * From MovAlmDet", cConexCom, adOpenDynamic, adLockOptimistic

Set Adodc3 = New ADODB.Recordset
Adodc3.Open "Select * From StkArt", cConexCom, adOpenDynamic, adLockOptimistic
     
VGActualizar = False
VGSoles = True
'VGForm = 5
lbltrans = "": lbltipref = "": lblauto = ""
DTPicker1 = CDate(Format(Date, "dd/MM/yyyy"))

VGRegEnt = 1
If VGRegEnt = 1 Then
    FrmArmKits.Caption = "Registro de Armado de Kits "
    dato = "I"
    Tipo = "NI"
    Codigo2 = "NOTA DE INGRESO"
Else
    FrmArmKits.Caption = "Registro de Desarmado de Kits"
    dato = "S"
    Tipo = "NS"
    Codigo2 = "NOTA DE SALIDA"
End If
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 1
inicializar
inicializaFG
Set db = Workspaces(0).OpenDatabase(cRuta2)
End Sub

Private Sub MSFlexGrid1_DblClick()
VGRegEnt = 1: VGForm1 = 4
FormAyuArtKid.Show 1
inicializaFG
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then MSFlexGrid1_DblClick
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
Alinear
'If MSFlexGrid1.Col = 2 Or MSFlexGrid1.Col = 3 Then
'    TxCantidad.Visible = True
'    TxCantidad.SetFocus
'
'    If KeyAscii <> 13 And KeyAscii <> 27 And KeyAscii <> 9 And KeyAscii <> 8 Then
'        TxCantidad.text = TxCantidad.text & Chr(KeyAscii)
'    End If
'
'    TxCantidad.SelStart = Len(TxCantidad.text)
'    TxCantidad.SelLength = 0
'End If
If (MSFlexGrid1.Col = 2 And Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)) >= 0) Or (MSFlexGrid1.Col = 3 And Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)) >= 0) Then
        If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 46 Then
            TxCantidad.FontName = MSFlexGrid1.CellFontName
            TxCantidad.FontSize = MSFlexGrid1.CellFontSize
            TxCantidad.Width = MSFlexGrid1.CellWidth
            TxCantidad.Height = 276 'MSFlexGrid1.CellHeight
            TxCantidad.Left = MSFlexGrid1.Left + MSFlexGrid1.CellLeft
            TxCantidad.Top = MSFlexGrid1.Top + MSFlexGrid1.CellTop
            TxCantidad.Visible = True
            TxCantidad = Chr(KeyAscii)
            TxCantidad.SelStart = 1
            TxCantidad.SetFocus
        End If
        
 End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'************************** NUM REF
If KeyAscii = 13 And Text1.text <> "" Then
   If Not IsNumeric(Text1) And (Text6 = "BV") Then
      MsgBox "Ingrese el Numero de  la Boleta", vbOKOnly, "Aviso"
      Exit Sub
   Else
      MSFlexGrid1.SetFocus
      'Text8.SetFocus
   End If
Else
 If Text6 = "BV" Then
   If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And KeyAscii <> 8 Then KeyAscii = 0
 End If
End If
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 112 Then Text9_DblClick
End Sub
Private Sub TxCambio_KeyPress(KeyAscii As Integer)
If NumPto(KeyAscii) Then
    If KeyAscii = 13 Then SendKeys "{tab}"
Else
    KeyAscii = 0
End If
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
            Enfoque Text6
       End If
       MSFlexGrid1.SetFocus
       Exit Sub
 Else
     If KeyAscii = 8 Then
         lbltrans = ""
         LIMPIACABECERA
    End If
End If
End Sub

Private Sub TxTransa_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT TT_CODMOV,TT_DESCRI FROM Tabtransa where  TT_tipmov = '" & IIf(VGRegEnt = 1, "I", "S") & "'", cConexCom, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT TT_CODMOV,TT_DESCRI FROM Tabtransa where  TT_tipmov = '" & IIf(VGRegEnt = 1, "I", "S") & "'"
frmReferencia.Label1.Caption = "Transacciones"
frmReferencia.Show vbModal
Adodc3.Close
If vGUtil(1) <> "" Then
    If vGUtil(1) <> "IK" Then
        MsgBox "Debe ser Transferencia por Elaboración de Kids", vbInformation, "Mensaje"
        TxTransa.SetFocus: Exit Sub
    End If
    TxTransa = vGUtil(1)
    lbltrans = Mid(vGUtil(2), 1, 30)
End If
If TxTransa.text <> "" Then TxTransa_KeyPress (13)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text4 <> "" Then
   Text4 = Format(Text4, "0000000000")
   
Else
   TxTransa.SetFocus
End If
End Sub

'**************** num ref *********************
Private Sub Text6_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU", cConexCom, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU"
frmReferencia.Label1.Caption = "Tipo de Documentos"
frmReferencia.Show vbModal
Adodc3.Close
If vGUtil(1) <> "" Then Text6 = (vGUtil(1))
If vGUtil(2) <> "" Then lbltipref = (vGUtil(2))
If Text6 <> "" Then
   Text1.SetFocus
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
         Text1.SetFocus
  Else
         Text6 = ""
         SendKeys "{tab}"
         KeyAscii = 0
  End If
End If
If KeyAscii = 8 Then lbltipref = ""
End Sub

 '***** Orden de compra
Private Sub Text8_KeyPress(KeyAscii As Integer)

Dim criterio As String
If KeyAscii = 13 Then
      Text8 = Trim(Text8)
      If Text8 <> "" Then
            'criterio = "CANUMORD = " & Chr$(34) + Text8.text + Chr$(34) & "AND  CACODPRO = " & Chr$(34) + TxtProveedor.text + Chr$(34)
            'Data1.Recordset.FindFirst criterio
            'If Not Data1.Recordset.NoMatch Then
            '  MsgBox "La Orden de Compra ya ha sido registrada !", vbExclamation, mensaje1
            '  Exit Sub
            'End If
      End If
      Text9.SetFocus
End If
End Sub

Private Sub Text9_DblClick()
FormAyuda.Show 1
End Sub
'Autorizado
Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      If Trim(Text9) <> "" Then
              'muestra
              If Trim(validarautorizado(Text9)) = "" Then
                      MsgBox "No existe el Autorizado", vbInformation, "Mensaje"
                      If Text9.Enabled And Text9.Visible Then Text9.SetFocus
                      Exit Sub
              End If
              lblauto = Mid(validarautorizado(Text9), 1, 10)
              CmbMoneda.SetFocus
      End If
 End If
End Sub
'Numeracion
Private Sub muestra()
Dim nument As Long, numsal As String
Dim rs As Recordset, rSql As String
If Trim(VGAlma) <> "" Then
    rSql = "select  TANUMENT, TANUMSAL from TabAlm  WHERE TAALMA='" & VGAlma & "' "
    Set rs = db.OpenRecordset(rSql, dbOpenSnapshot)
    nument = IIf(IsNull(rs(0)), 1, rs(0))
    numsal = IIf(IsNull(rs(1)), 1, rs(1))
    If VGRegEnt = 1 Then
        Text4.text = Format(Val(nument) + 1, "0000000000")
        ContaSalida = Format(Val(numsal) + 1, "0000000000")
    Else
        Text4.text = Format(Val(numsal) + 1, "0000000000")
        ContaSalida = Format(Val(nument) + 1, "0000000000")
    End If
    Command1.Visible = True
    CmdGrabar.Visible = True
    Command3.Visible = True
    Command7.Visible = True
Else
   MsgBox "No ningún Almacen Activo", vbInformation, "Información"
End If
End Sub

Public Sub grabastk(Art As String, RegEs As Integer, PreUni As Double)
Dim cadena As String, criterio As String
Dim entrada As Boolean
On Error GoTo GrabErr
cadena = Art
criterio = " STCODIGO = '" & cadena & "'"
criterio = criterio + "  and  STALMA = '" & VGAlma & "'"
If Not Adodc3.EOF Then Adodc3.Filter = criterio
  
If Not Adodc3.EOF Then      'si existe el articulo
    canttemp = Adodc3("STSKDIS")     ' revisar si validar en creacion
    If RegEs = 1 Then 'Entrada
        Adodc3("STKFECULT") = DTPicker1.Value
        Adodc3("STSKDIS") = Adodc3("STSKDIS") + cantidad
        'aqui actualiza
        If Not IsNull(Adodc3("STKPREPRO")) Then
            precioprom = Adodc3("STKPREPRO")
            If PreUni <> 0 Then
                   Adodc3("STKPREULT") = PreUni * VGTipCamb 'el precio
                   If PreUni <> 0 And (canttemp + cantidad) <> 0 Then   'Valorizacion
                      Adodc3("STKPREPRO") = Round((precioprom * canttemp + cantidad * Val(Val(PreUni) * VGTipCamb)) / (canttemp + cantidad), 6)
                   End If
              End If
         Else
              precioprom = 0
              If PreUni <> 0 Then
                 Adodc3("STKPREPRO") = Round(Val(PreUni) * VGTipCamb, 6) 'el precio
                 If PreUni <> 0 Then 'Valorzacion
                   Adodc3("STKPREULT") = Round(Val(PreUni) * VGTipCamb) 'el precio
                   Adodc3("STKFECULT") = DTPicker1.Value
                 End If
              End If
           End If
    Else 'para la salida
    
         Adodc3("STSKDIS") = Adodc3("STSKDIS") - cantidad
         'aqui actualiza
         If Not IsNull(Adodc3("STKPREPRO")) Then
            precioprom = Round(Adodc3("STKPREPRO"), 6)
         Else
            precioprom = 0
         End If
    End If
Else
       Adodc3.AddNew                  'existe
       Adodc3("STALMA") = VGAlma   '"01"
       Adodc3("STCODIGO") = Art
       Adodc3("STKFECULT") = DTPicker1.Value
       If RegEs = 1 Then 'entrda
           Adodc3("STSKDIS") = cantidad
           Adodc3("STKPREULT") = Val(PreUni) '* VGTipCamb    'el costo de ingreso
           If PreUni <> 0 Then
                 Adodc3("STKPREPRO") = Round(Val(PreUni), 6) '******el  costo = costo prom
           End If
       End If
End If
Adodc3.Update
Adodc3.Filter = ""
entrada = IIf(RegEs = 1, True, False)
Call ValMes(VGAlma, entrada) 'para la valorizacion
Exit Sub
GrabErr:
 MsgBox Err.Description
End Sub

Private Sub buscar_trans()
Dim criterio As String
Dim rs As Recordset
Dim rSql As String
TxTransa = UCase(LTrim(TxTransa))
If TxTransa = "TD" And VGRegEnt Then
  MsgBox "El tipo de transaccion no puede ser usado para registrar !", vbOKOnly, "Error"
  lbltrans = ""
  TxTransa.SetFocus
  Exit Sub
End If
'Busco la transaccion
rSql = "select  *  from TabTransa  where TT_CODMOV ='" & TxTransa.text & "' and TT_TIPMOV ='" & dato & "'"
Set rs = db.OpenRecordset(rSql, dbOpenSnapshot)
If rs.EOF Then
   MsgBox "El tipo de transaccion no existe !", vbOKOnly, "Error"
   LIMPIACABECERA
   TxTransa.SetFocus
   Exit Sub
End If
lbltrans = Mid(rs("TT_DESCRI"), 1, 30)
       
If Not IsNull(rs("TT_CONT")) Then
    TT_CONTADOR = rs("TT_CONT")
Else
    MsgBox "El tipo de transacción no esta inicializara !" & Chr(13) & "Para inicializarla ir a la tabla de Transacción", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If
End Sub

Private Sub grabacabecera(Dat As String, Tip As String, num As String, RegEs As Integer)
Dim criterio As String, cadena As String
Dim FACTOR As Double, uSql As String

On Error GoTo GrabErr
'Desea grabar el registro

If num <> "" Then
    If Not VGActualizar Then
      Adodc1.AddNew
      Adodc1("CAALMA") = VGAlma     '"01"
      Adodc1("CANUMDOC") = Mid$(UCase$(num), 1, 10)
    Else
      criterio = " CANUMDOC = '" & num & "' AND CATD = '" & Tip & "' "
      criterio = criterio + " and  CAALMA = '" & VGAlma & "'"
      If Not Adodc1.EOF Then Adodc1.Filter = criterio
    End If
    Adodc1("CATIPMOV") = Dat
    Adodc1("CATD") = Tip
    Adodc1("CAHORA") = Format(Time, "hh:mm:ss")
    Adodc1("CAFECDOC") = DTPicker1.Value            ' CDate(Text2.text)
   
    If Trim(Text1.text) <> "" Then
      Adodc1("CARFNDOC") = Trim(Text1.text)
    Else
      Adodc1("CARFNDOC") = " "
    End If
    If TxTransa.text <> "" Then
      Adodc1("CACODMOV") = Mid$(UCase$(TxTransa.text), 1, 2)
    Else
      Adodc1("CACODMOV") = " "
    End If
    num = Trim(UCase$(num))
    Adodc1("CANUMDOC") = num
    If RegEs = 1 Then
       uSql = "Update TabAlm set TANUMENT= '" & num & "' where TAALMA='" & VGAlma & "' " ',TANUMSAL= '" & ContaSalida & "'
    Else
       uSql = "Update TabAlm set TANUMSAL= '" & num & "'  where TAALMA='" & VGAlma & "' " ',TANUMENT= '" & ContaSalida & "'
    End If
    cConexCom.Execute uSql
   
    If Trim(Text6) <> "" Then
      Adodc1("CARFTDOC") = Mid$(UCase$(Text6.text), 1, 2)
    Else
      Adodc1("CARFTDOC") = " "
    End If
    If Trim(Text8.text) <> "" And RegEs = 1 Then
      Adodc1("CANUMORD") = Mid$(UCase$(Text8.text), 1, 10)
    Else
      Adodc1("CANUMORD") = " "
    End If
    If Text9.Visible And Trim(Text9) <> "" Then
      Adodc1("CASOLI") = Mid$(UCase$(Text9.text), 1, 3)
    Else
      Adodc1("CASOLI") = " "
    End If
    Adodc1("CAUSUARI") = UCase(VGUsua)
    Adodc1("CACODMON") = VGCodMon
    Adodc1("CASITGUI") = "V"
    'Data1.Recordset("CASITUA") = "V"
    Adodc1("CAESTIMP") = "V"
    Adodc1.Update
    Adodc1.Filter = ""
    Adodc1.Requery
End If
Exit Sub
GrabErr:
    MsgBox Err.Description
End Sub
Function ValidarDoc(txt As TextBox) As String
Dim rs As Recordset, rSql As String
rSql = "select TDO_DESCRI  from TIPO_DOCU  where TDO_TIPDOC='" & txt.text & "'"
Set rs = db.OpenRecordset(rSql, dbOpenSnapshot)
If rs.EOF Then
    MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
    ValidarDoc = ""
    txt.SetFocus
    Exit Function
End If
ValidarDoc = rs(0)
rs.Close
End Function

Private Sub LIMPIACABECERA()
TxTransa = "": Text6 = "": Text8 = "": Text4 = "": Text1 = "": Text9 = ""
lbltrans = "": lbltipref = "": lblauto = "": TxCantidad = "": TxCambio.Enabled = False: TxCambio = "0"
CmbMoneda.ListIndex = 0
End Sub

Private Sub inicializar()
TxTransa = "": Text6 = "": Text8 = "": Text4 = "": Text1 = "": Text9 = "": TxCantidad = "": TxCambio = "0": TxCambio.Enabled = False
CmbMoneda.ListIndex = 0
inicializaFG
Command1.Visible = True
Command3.Visible = True
Command7.Visible = True
CmdGrabar.Visible = True
LIMPIACABECERA
End Sub

Private Sub ValMes(almacen As String, entrada As Boolean)
  Dim cadena As String
  Dim criterio As String
  Dim adoreg As ADODB.Recordset
  Dim rSql As String
  Dim Cantent As Double
  Dim Cantsal As Double
  Dim mespro As String
  Dim uSql As String
  On Error GoTo Err
  mespro = Year(DTPicker1) & Format(Month(DTPicker1), "00")
  cadena = MSFlexGrid1.TextMatrix(contador, 0) 'codigo del art
  rSql = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & almacen & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & cadena & "'"  '
  Set adoreg = New ADODB.Recordset
  adoreg.Open rSql, cConexCom, adOpenDynamic, adLockOptimistic
  If Not adoreg.EOF Then 'existe
      If entrada Then
          Cantent = adoreg(0) + cantidad
          uSql = "Update MoResMes set SMCANENT = " & Cantent & " where SMALMA='" & almacen & "'  and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
      Else
          Cantsal = adoreg(1) + cantidad
          uSql = "Update MoResMes set SMCANSAL = " & Cantsal & " where SMALMA='" & almacen & "' and   SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
      End If
  Else
      If entrada Then
          Cantent = cantidad
          Cantsal = 0
      Else
          Cantsal = cantidad
          Cantent = 0
      End If
      uSql = "insert into MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMSALDOINI) VALUES ('" & almacen & "','" & cadena & "','" & mespro & "' ," & Cantent & "," & Cantsal & ",0) "
   End If
   cConexCom.Execute uSql
  Exit Sub
Err:
   MsgBox Err.Description
End Sub
'Solo para lote, arreglar
Private Sub grabalote(alma As String, codigo As String)
Dim uSql As String
Dim lote As String
Dim nuevo_stk As Double
Dim rSql As String
Dim rs As Recordset
Dim fecfab As Date
Dim fecven As Date
    If (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" Then
      fecfab = MSFlexGrid1.TextMatrix(contador, 9)
    End If
    If (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
      fecven = MSFlexGrid1.TextMatrix(contador, 8)
    End If
    lote = MSFlexGrid1.TextMatrix(contador, 2)
    rSql = "select STSLKDIS FROM STKLOTE where  STSALMA= '" & alma & "' and  STSCODIGO= '" & codigo & "' and STSLOTE= '" & lote & "'" '
    Set rs = db.OpenRecordset(rSql, dbOpenSnapshot)
    If Not rs.EOF Then
       If Tipo = "NI" Then
         nuevo_stk = rs(0) + cantidad
       Else
         nuevo_stk = rs(0) - cantidad
       End If
       
       uSql = "Update STKLOTE set STSLKDIS = " & nuevo_stk & " WHERE  STSALMA='" & alma & "' and STSCODIGO='" & codigo & "'AND STSLOTE='" & lote & "'"
    Else
    If (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) = "__/__/____" Then
        fecfab = MSFlexGrid1.TextMatrix(contador, 9)
        uSql = "insert into STKLOTE (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB) VALUES ('" & alma & "','" & codigo & "','" & lote & "'," & cantidad & ",#" & Format(fecfab, "MM/DD/YYYY") & "#) "
    ElseIf (MSFlexGrid1.TextMatrix(contador, 9)) = "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
        fecven = MSFlexGrid1.TextMatrix(contador, 8)
        uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECVEN)  VALUES ('" & alma & "','" & codigo & "','" & lote & "' ," & cantidad & " ,#" & Format(fecven, "MM/DD/YYYY") & "#) " 'SIN FECFAB
    ElseIf (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
        uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB,STSFECVEN)  VALUES ('" & alma & "','" & codigo & "','" & lote & "' ," & cantidad & " ,#" & Format(fecfab, "MM/DD/YYYY") & "#,#" & Format(fecven, "MM/DD/YYYY") & "#) "
    Else
        uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS)  VALUES ('" & alma & "','" & codigo & "','" & lote & "' ," & cantidad & ") "
    End If
    
    End If
    db.Execute uSql
       
End Sub
'Solo para serie arreglar
Private Sub grabaserie(alma As String, codigo As String)
Dim uSql As String
Dim Serie As String
Dim VALOR As Integer
Dim rs As Recordset
Dim rSql As String
Dim fecfab As Date
Dim fecven As Date
    'fecfab = " " '  MSFlexGrid1.TextMatrix(contador, 8)
    'fecven = " " 'MSFlexGrid1.TextMatrix(contador, 9)
    Serie = MSFlexGrid1.TextMatrix(contador, 2)
    rSql = "select STSSKDIS FROM STKSERI where   STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & codigo & "' and STSSERIE= '" & Serie & "'" '
    Set rs = db.OpenRecordset(rSql, dbOpenSnapshot)
    If Not rs.EOF Then
       VALOR = IIf(Tipo = "NI", 1, 0)
       uSql = "update STKSERI set STSSKDIS = " & VALOR & " WHERE  STSALMA='" & alma & "' and STSCODIGO='" & codigo & "'AND STSSERIE='" & Serie & "'"
    Else
       uSql = "insert into STKSERI (STSALMA,STSCODIGO,STSSERIE,STSSKDIS)   VALUES ('" & alma & "','" & codigo & "','" & Serie & "',1) "
    End If
    cConexCom.Execute uSql
       
End Sub

Private Sub inicializaFG()
MSFlexGrid1.FormatString = "  Codigo|   Descripcion|  Cantidad Ing.|  Costo Unit. "
MSFlexGrid1.ColWidth(0) = 1500
MSFlexGrid1.ColWidth(1) = 3500
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 1500
MSFlexGrid1.ColAlignment(0) = 1
MSFlexGrid1.ColAlignment(1) = 1
'MSFlexGrid1.Rows = 1
End Sub
Function existe_numdoc(text As TextBox) As Boolean
Dim criterio As String
  criterio = " CANUMDOC = '" & Format(text, String(10, "0")) & "' "
  criterio = criterio + " and  CAALMA = '" & VGAlma & "'"
  criterio = criterio + " and  CATD = '" & Tipo & "'"
  If Not Adodc1.EOF Then Adodc1.Filter = criterio
  If Not Adodc1.EOF Then
         MsgBox "El Número del documento ya ha sido registrado: " & Format(text, String(10, "0")) & " !", vbExclamation, "Error"
         existe_numdoc = True
  Else
         existe_numdoc = False
  End If
  Adodc1.Filter = ""
End Function

Function validarautorizado(text As TextBox) As String
  Dim rSql As String
  Dim rs As Recordset
  Dim codayu As String
  codayu = 12
  rSql = "Select TCLAVE,TDESCRI from TABAYU  where TCOD= '" & codayu & "' and  Tclave ='" & Trim(text) & "'"
  Set rs = db.OpenRecordset(rSql, dbOpenSnapshot)
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
Dim rSql As String
Dim rs As Recordset
Dim cadena As String
   cadena = MSFlexGrid1.TextMatrix(contador, 0)
   rSql = "select  stskdis from stkart  WHERE STALMA='" & VGAlma & "'  and stcodigo ='" & cadena & "'"
   Set rs = db.OpenRecordset(rSql, dbOpenSnapshot)
   If Not rs.EOF Then
     If cantidad > rs(0) Then
       consulta_stk = False
     Else
       consulta_stk = True
     End If
   End If
   rs.Close
End Function

Function existe_lote(text As String) As Boolean
Dim rs As Recordset
Dim rSql As String
Dim lote As String
   lote = MSFlexGrid1.TextMatrix(contador, 2)
   rSql = "select  STSLKDIS from STKLOTE where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & text & "' and STSLOTE = '" & lote & "'"
   Set rs = VGBaseDatos.OpenRecordset(rSql, dbOpenSnapshot)
   If Not rs.EOF Then
     If cantidad > rs(0) Then
       MsgBox "No hay stock del" & text & "lote:" & lote, vbInformation, "Aviso"
       existe_lote = False
     Else
       existe_lote = True
     End If
   End If
   rs.Close
End Function

Function existe_serie(text As String) As Boolean
Dim rs As Recordset
Dim rSql As String
Dim Serie As String
   Serie = MSFlexGrid1.TextMatrix(contador, 2)
   rSql = "select STSSKDIS from STKSERI where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & text & "' and STSSERIE = '" & Serie & "'"
   Set rs = VGBaseDatos.OpenRecordset(rSql, dbOpenSnapshot)
   If Not rs.EOF Then
     If cantidad > rs(0) Then
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
CrystalReport1.WindowTitle = "Inv043 -- Control de Inventarios"
CrystalReport1.ReportFileName = cRutP & "inv043.rpt"
Ubi_Tab CrystalReport1
cadena = "{MOVALMCAB.CAALMA} = '" & VGAlma & "'  and {MOVALMCAB.CATD} = '" & Tipo & "' and {MOVALMCAB.CANUMDOC} = '" & NumDoc & "'"
CrystalReport1.DiscardSavedData = True
CrystalReport1.Destination = crptToWindow
 CrystalReport1.WindowTitle = " Control de Inventarios"
CrystalReport1.SelectionFormula = cadena
CrystalReport1.WindowShowPrintBtn = True
CrystalReport1.WindowShowRefreshBtn = True
CrystalReport1.WindowShowSearchBtn = True
CrystalReport1.WindowShowPrintSetupBtn = True
CrystalReport1.Formulas(0) = "empresa ='" & VGNemp & "'"
CrystalReport1.Formulas(1) = "nota ='" & Codigo2 & "'"
CrystalReport1.Formulas(2) = "hora ='" & Time & "'"
If VGRegEnt = 0 Then
    CrystalReport1.Formulas(3) = "Tipo = 'S'"
Else
    CrystalReport1.Formulas(3) = "Tipo = 'I'"
End If
CrystalReport1.Action = 1

If VGRegEnt <> 1 And TxTransa = "TD" Then
    If vbOK = MsgBox(" Desea imprimir la nota de Ingreso", vbInformation, "Aviso") Then
        CrystalReport1.WindowTitle = "Inv043 -- Control de Inventarios"
        CrystalReport1.ReportFileName = cRutP & "inv043.rpt"
        Ubi_Tab CrystalReport1
        cadena = "{MOVALMCAB.CAALMA} = '" & VGAlma & "'  and {MOVALMCAB.CATD} = '" & Campo & "' and {MOVALMCAB.CANUMDOC} = '" & Format(nument, "0000000000") & "'"
        CrystalReport1.DiscardSavedData = True
        CrystalReport1.Destination = crptToWindow
        CrystalReport1.WindowTitle = " Control de Inventarios"
        CrystalReport1.SelectionFormula = cadena
         CrystalReport1.WindowShowPrintBtn = True
         CrystalReport1.WindowShowRefreshBtn = True
         CrystalReport1.WindowShowSearchBtn = True
         CrystalReport1.WindowShowPrintSetupBtn = True
        CrystalReport1.Formulas(0) = "empresa ='" & VGNemp & "'"
        CrystalReport1.Formulas(1) = "nota ='NOTA DE INGRESO'"
        CrystalReport1.Formulas(2) = "hora ='" & Time & "'"
        CrystalReport1.Formulas(3) = "Tipo = 'S'"
        CrystalReport1.Action = 1
   End If
End If
End Sub

Private Sub GrabarDetKit(Art As String, Tip As String, Cont As Integer, Descr As String, CANT As Double, Preu As Double)
Dim criterio As String
If Not VGActualizar Then
    Adodc22.AddNew
Else
    criterio = "DECODIGO = '" & Art & "' "
    criterio = criterio + " and  DEALMA = '" & VGAlma & "' "
    If Not Adodc22.EOF Then Adodc22.Filter = criterio
End If
Adodc22("DEALMA") = VGAlma
Adodc22("DETD") = Tip
Adodc22("DENUMDOC") = ContaSalida
Adodc22("DEITEM") = Cont
Adodc22("DECODIGO") = Art
Adodc22("DEDESCRI") = Descr
Adodc22("DECANTID") = CANT
If Preu <> 0 Then    'si tiene precio
    Adodc22("DEPRECIO") = Val(Preu) '* VGTipCamb '******el precio
    Adodc22("DETIPCAM") = VGTipCamb
Else
    Adodc22("DEPRECIO") = 0
End If
alma = VGAlma

'mejorar a una funcion
'If MSFlexGrid1.TextMatrix(contador, 10) = "S" Then
'     grabaserie alma, cadena
'     Data2.Recordset("DESERIE") = MSFlexGrid1.TextMatrix(contador, 2)
'End If
'If MSFlexGrid1.TextMatrix(contador, 10) = "N" Then
'    grabalote alma, cadena
'    Data2.Recordset("DELOTE") = MSFlexGrid1.TextMatrix(contador, 2)
'End If
Adodc22.Update
Adodc22.Filter = ""
Adodc22.Requery
End Sub

Sub Alinear()
TxCantidad.Width = MSFlexGrid1.CellWidth
TxCantidad.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
TxCantidad.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
TxCantidad.Height = MSFlexGrid1.CellHeight
End Sub

Private Sub TxCantidad_KeyPress(KeyAscii As Integer)
If NumPto(KeyAscii) Then
    Select Case KeyAscii
      Case Is = 13
         MSFlexGrid1.text = TxCantidad.text
         TxCantidad.Visible = False
         TxCantidad.text = ""
         MSFlexGrid1.SetFocus
      Case Is = 27
         TxCantidad.Visible = False
         TxCantidad.text = ""
    End Select
Else
    KeyAscii = 0
End If
End Sub

Function VerificarSTOCK() As Boolean
Dim Xstock As Double
For n = 1 To MSFlexGrid1.Rows - 1
    Dim rs As New ADODB.Recordset
    rs.Open "Select codart,canart from kits Where codkit='" & MSFlexGrid1.TextMatrix(n, 0) & "'", cConexCom, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
       Xstock = ClsTock.SaldoArti(VGAlma, rs!codart, cConexCom)
       If rs!Canart >= Xstock Then
          
       End If
    End If
Next
End Function

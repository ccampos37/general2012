VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FormCreacion 
   Caption         =   "Detalle  Valorizado"
   ClientHeight    =   4695
   ClientLeft      =   2565
   ClientTop       =   2115
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "FormCreacion"
   ScaleHeight     =   4695
   ScaleWidth      =   7905
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   6864
      Picture         =   "FormCreacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Limpiar"
      Height          =   735
      Left            =   5712
      Picture         =   "FormCreacion.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Enviar"
      Height          =   735
      Left            =   4524
      Picture         =   "FormCreacion.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3720
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Moneda"
      Height          =   2535
      Left            =   240
      TabIndex        =   27
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   2040
         TabIndex        =   37
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FormCreacion.frx":0CC6
         Left            =   1560
         List            =   "FormCreacion.frx":0CD0
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormCreacion.frx":0CE4
         Left            =   1560
         List            =   "FormCreacion.frx":0CF1
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   345
         TabIndex        =   28
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Conversion"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Tipo Cambio"
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   1440
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Salida 
      Height          =   2535
      Left            =   5160
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   4560
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7800
      Begin VB.TextBox Txtcantref 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   2640
         Width           =   1380
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1215
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5895
         TabIndex        =   6
         Top             =   2175
         Width           =   1695
      End
      Begin VB.TextBox TxtArticulo 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   0
         Top             =   285
         Width           =   2280
      End
      Begin VB.TextBox TxtCantidad 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   2160
         Width           =   1380
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         TabIndex        =   9
         Top             =   2775
         Width           =   1380
      End
      Begin VB.TextBox TxtPrecioUnit 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   3240
         Width           =   1380
      End
      Begin VB.TextBox TxtTotal 
         Height          =   285
         Left            =   5880
         MaxLength       =   10
         TabIndex        =   11
         Top             =   3240
         Width           =   1695
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   285
         Left            =   5880
         TabIndex        =   5
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483634
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   288
         Left            =   1680
         TabIndex        =   4
         Top             =   1668
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   -2147483634
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label22 
         Caption         =   "Merma"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Num de Item:"
         Height          =   255
         Left            =   4680
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "L17"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   38
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "S/."
         Height          =   255
         Left            =   5520
         TabIndex        =   36
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "S/."
         Height          =   255
         Left            =   1440
         TabIndex        =   35
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "Serie \ Lote"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Unidad de Ingreso"
         Height          =   495
         Left            =   4800
         TabIndex        =   24
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1680
         TabIndex        =   1
         Top             =   768
         Width           =   5928
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Estandar"
         Height          =   375
         Left            =   4800
         TabIndex        =   22
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Vcto."
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Factura"
         Height          =   255
         Left            =   4800
         TabIndex        =   19
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Precio Unitario"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Fab."
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Total"
         Height          =   255
         Left            =   4470
         TabIndex        =   16
         Top             =   3270
         Width           =   1095
      End
   End
   Begin VB.Label Label21 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   2760
      TabIndex        =   41
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "FormCreacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cant As Double
Dim I As Integer
Dim fin As Integer
Dim flagserie As String * 1
Dim flaglote As String * 1
Dim xserie As String
Dim graba As Boolean
Dim varform As Form
Dim ya_grabo_det As Boolean
'Dim db As Database
'***********************************
'**************RMM  07/07/2001
Dim rsSTKART As New ADODB.Recordset



Private Sub Combo2_Click()

If Combo2.text = "Soles" Then
        VGTipCamb = 1
        VGSoles = True
        Label19 = "S/."
        Label20 = "S/."
        Frame2.Visible = False
        Frame1.Visible = True
        bloquearc1_c2 (True)
        Me.Height = 5500
        Me.Width = 9000
        If varform.Caption <> "Modificar" Then FrmRegistro.Text2 = "01"
        central FormCreacion
  Else
        If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
            Text8 = Val(Devolver_Dato(3, FrmRegistro.DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
        ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
            Text8 = Val(Devolver_Dato(1, FrmRegistro.DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
        End If
       Text8.Enabled = True
       Label19 = "$"
       Label20 = "$"
       If varform.Caption <> "Modificar" Then FrmRegistro.Text2 = "02"
       VGSoles = False
       VGTipCamb = Val(Text8)
  End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Combo2.text = "Soles" Then
         Frame2.Visible = False
   Else
        Combo2.SetFocus
   End If
End If
End Sub


Private Sub Command1_Click()
   'Enviar
Dim criterio As String
Dim abrev As String
'Dim db As Database
Dim rs As Recordset
Dim RSQL As String
Dim FACTOR As Double
Dim ncombo As Integer
graba = True

If TxtPrecioUnit = "" Or Not IsNumeric(TxtPrecioUnit) Then
    MsgBox " Ingrese el precio del articulo", vbInformation, " Aviso"
    TxtPrecioUnit.SetFocus
    Exit Sub
End If

If (TxtCantidad = "") Or Not IsNumeric(TxtCantidad) Then
    MsgBox " Ingrese cantidad ", vbInformation, " Aviso"
    TxtCantidad.SetFocus
    Exit Sub
End If

If Trim(TxtArticulo) = "" Then
    MsgBox " Ingrese el codigo del articulo", vbInformation, " Información"
    TxtArticulo.SetFocus
    Exit Sub
End If

Dim rst_Art As ADODB.Recordset
Set rst_Art = New ADODB.Recordset

rst_Art.Open "Select ADESCRI From Maeart Where ACODIGO='" & TxtArticulo & "' Order by ACODIGO;", VGCNx, adOpenForwardOnly, adLockReadOnly
If rst_Art.EOF Then
    MsgBox "Codigo de Articulo es inválido, no existe. ", vbInformation, " Información"
    TxtArticulo.SetFocus
    Set rst_Art = Nothing
    Exit Sub
Else
    TxtArticulo = UCase(TxtArticulo)
    Label13 = rst_Art!adescri
End If
Set rst_Art = Nothing

If flagserie = "S" And Trim(Text3) = "" Then  'And Not Combo1.Enabled
     MsgBox "Ingrese el Número de serie", vbOKOnly + vbExclamation, "Error"
     Text3.SetFocus
     Exit Sub
End If

If flaglote = "S" And (Trim(Text3) = "") Then 'And Not Combo1.Enabled
     MsgBox "Ingrese el Número de Lote", vbOKOnly + vbExclamation, "Error"
     Text3.SetFocus
     Exit Sub
End If
If (flagserie = "S") And Text3 <> "" Then
    If FrmRegistro.MSFlexGrid1.Rows <> 1 Then
        For ncombo = 1 To FrmRegistro.MSFlexGrid1.Rows - 1
          If UCase(Text3) = UCase(FrmRegistro.MSFlexGrid1.TextMatrix(ncombo, 2)) Then
            MsgBox "Ya se ingreso la serie", vbInformation, "Aviso"
            Text3.SetFocus
            Exit Sub
          End If
        Next ncombo
    End If
End If
'Solo valida cuando es salida
If (flagserie = "S") And VGRegEnt <> 1 Then
    If existe_serie(Text3) Then Exit Sub
End If

If VGabrev <> Label14 Then                          'CONSULTA POR DEFECTO MODIFICAR
    RSQL = "select  p.EQCANTEQUI from TabEqui p where p.EQUNIPRI = '" & VGabrev & "'   and p.EQUNIEQUI = '" & Label14.Caption & "'"
    'Set db = Workspaces(0).OpenDatabase(cRuta2)
    'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    If rs.RecordCount > 0 Then
         rs.MoveFirst
         FACTOR = rs.Fields(0)
    End If
    'db.Close
Else
    FACTOR = 1 'factor por defecto
End If

'PrecioUni = val(TxtPrecioUnit) / FACTOR

If varform.Caption = "Modificar" Then
     MsgBox "Documento a modificar", vbInformation, "Aviso"
     VGtipocreacion = 2
     If VGtipocreacion = 2 Then
        ya_grabo_det = False
        grabadetalle
        If ya_grabo_det Then
             MsgBox "Se grabo satisfactoriamente", vbInformation, "Aviso"
        Else
             Unload Me
             Exit Sub
        End If
    End If
End If
If VGSeleccion = 2 Then
   FACTOR = 1 'Inicialmente sin conversion
   varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 2) = Text3.text  'serie lote
   varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 3) = Format(Val(TxtCantidad.text), "##0.00") 'cantidad
   varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 4) = VGabrev
   varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 5) = Format(Val(TxtPrecioUnit.text), "###0.0000")  'Precio
   If VGtipocreacion = 1 Then
        varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 6) = Format(Val(TxtCantidad) * FACTOR, "###0.00") 'Label14  'unidad principal
   Else
        varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 6) = Val(FrmModificar.numitem)    'Cantidad informada
   End If
   varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 7) = Format(Val(TxtPrecioUnit) / FACTOR, "###0.0000") 'unidad principal
   If flagserie = "S" Then
        varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "S"
        varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 8) = MaskEdBox1
        varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 9) = MaskEdBox1
        xserie = "S"
   ElseIf flaglote = "S" Then
        varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "N"
        xserie = "N"
   Else
        varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "X"
        xserie = "X"
   End If
  
Else
    xserie = "X" 'por defecto ningun articulo tiene serie
    pro_xserie  'incica si el producto tiene serie o lote
    'actualizo_stock
    If FACTOR = 0 Then FACTOR = 1 'Factor indica factor de conversion
    If VGtipocreacion = 2 Then
             varform.MSFlexGrid1.AddItem (Trim(TxtArticulo.text) & vbTab & Label13 & vbTab & Text3.text & vbTab & Format(Val(TxtCantidad.text), "###0.00") & vbTab & VGabrev & vbTab & Format(Val(TxtPrecioUnit.text), "##0.0000") & vbTab & Val(FrmModificar.numitem - 1) & vbTab & Format(Val(TxtPrecioUnit) / FACTOR, "###0.0000") & vbTab & MaskEdBox1 & vbTab & MaskEdBox2 & vbTab & xserie & vbTab & Txtcantref.text & vbTab & VGTipCamb)
    Else
             varform.MSFlexGrid1.AddItem (Trim(TxtArticulo.text) & vbTab & Label13 & vbTab & Text3.text & vbTab & Format(Val(TxtCantidad.text), "###0.00") & vbTab & VGabrev & vbTab & Format(Val(TxtPrecioUnit.text), "##0.0000") & vbTab & Format(Val(TxtCantidad) * FACTOR, "###0.00") & vbTab & Format(Val(TxtPrecioUnit) / FACTOR, "###0.0000") & vbTab & MaskEdBox1 & vbTab & MaskEdBox2 & vbTab & xserie & vbTab & vbTab & vbTab & vbTab & Txtcantref.text & vbTab & VGTipCamb)
    End If
     'varform.MSFlexGrid1.AddItem (Trim(TxtArticulo.text) & vbTab & Label13 & vbTab & Text3.text & vbTab & Format(Val(TxtCantidad.text), "###0.00") & vbTab & VGabrev & vbTab & Format(Val(TxtPrecioUnit.text), "##0.0000") & vbTab & Val(FrmModificar.numitem - 1) & vbTab & Format(Val(TxtPrecioUnit) / FACTOR, "###0.0000") & vbTab & MaskEdBox1 & vbTab & MaskEdBox2 & vbTab & xserie)
    'Cuando es llamado del modulo de modificar
    If VGtipocreacion = 2 Then
            Unload Me
            Exit Sub
    End If
End If
limpia
Label16.Visible = True  'num item
Label17.Visible = True
Label17 = I
I = I + 1
If I < fin Then
   DisplayDisp
   If flagserie = "S" Or flaglote = "S" Then
       Text3.Enabled = True
       Text3.SetFocus
   Else
       TxtCantidad.Enabled = True
       TxtCantidad.SetFocus
   End If
Else
   label7.Visible = False
   Command1.Enabled = False
   Command7.SetFocus
End If
If VGEstadomodi Then
    Unload Me
ElseIf VGSeleccion = 2 Then
    Unload Me
End If
End Sub

Private Sub Command2_Click()


If Combo2.ListIndex = 1 Then
    If Not IsNumeric(Text8) Or Val(Text8) = 0 Then
            VGCodMon = "02"
            MsgBox "Ingrese el tipo de cambio", vbExclamation, "Aviso"
            Text8.SetFocus
    Else
            If Val(Text8) <> 0 And IsNumeric(Text8) Then
               VGTipCamb = Val(Text8)
           Else
               VGTipCamb = Val(Text8) 'obtengo el tipo de cambio
            End If
            VGCodMon = "01"
            Frame2.Visible = False
            Frame1.Visible = True
            bloquearc1_c2 (True)
            Me.Height = 5500
            Me.Width = 9000
            central FormCreacion
          TxtArticulo.SetFocus
  End If
Else
            VGTipCamb = 1
            Frame2.Visible = False
            Frame1.Visible = True
            bloquearc1_c2 (True)
            Me.Height = 5500
            Me.Width = 9000
            central FormCreacion
            TxtArticulo.SetFocus
End If
End Sub

Private Sub Command3_Click()
limpia
TxtArticulo.Enabled = True
TxtArticulo.SetFocus
End Sub

Private Sub Command4_Click()
 VGSeleccion = 1
 Unload Me
End Sub


Private Sub Command7_Click()
    VGSeleccion = 1
    limpia
    Unload Me
End Sub

Private Sub MaskEdBox1_GotFocus()
  MaskEdBox1.SelStart = 0: MaskEdBox1.SelLength = Len(MaskEdBox1)
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MaskEdBox1 = "__/__/____" Then
            MsgBox "Ingrese  Fecha ", vbExclamation + vbOKOnly, "Advertencia"
'            MaskEdBox1.SetFocus
'            Exit Sub
    End If
    MaskEdBox2.SetFocus
End If
End Sub

Private Sub MaskEdBox1_LostFocus()
Dim cValor As String

If MaskEdBox1 = "__/__/____" Then
Else
    cValor = ValidFecha(MaskEdBox1)
    If cValor = "" Then
       MsgBox "Ingrese la Fecha Correctamente", vbExclamation + vbOKOnly, "Advertencia"
       MaskEdBox1 = "__/__/____"
       MaskEdBox1.SetFocus
       Exit Sub
    Else
      MaskEdBox1 = cValor
    End If
    If CDate(MaskEdBox1) < Date Then
        MsgBox "El articulo ya vencio", vbExclamation + vbOKOnly, "Error"
        MaskEdBox1.SetFocus
    End If
End If
End Sub

Private Sub MaskEdBox2_GotFocus()
  MaskEdBox1.SelStart = 0: MaskEdBox1.SelLength = Len(MaskEdBox1)
End Sub

Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If MaskEdBox2 = "__/__/____" Then
      MsgBox "Ingrese  Fecha ", vbExclamation + vbOKOnly, "Advertencia"
'      MaskEdBox2.SetFocus
'      Exit Sub
  End If
  SendKeys "{TAB}"
  KeyAscii = 0
End If
End Sub

Private Sub MaskEdBox2_LostFocus()
Dim cValor As String

If MaskEdBox2 = "__/__/____" Then
Else
    cValor = ValidFecha(MaskEdBox2)
    If cValor = "" Then
       MsgBox "Ingrese la Fecha Correctamente", vbExclamation + vbOKOnly, "Advertencia"
       MaskEdBox2 = "__/__/____"
       MaskEdBox2.SetFocus
    ElseIf MaskEdBox1 <> "__/__/____" Then
       If CDate(MaskEdBox1) < CDate(MaskEdBox2) And CDate(MaskEdBox2) > Date Then
'        MsgBox "Ingrese fecha Valida", vbExclamation + vbOKOnly, "Error"
'        MaskEdBox2.SetFocus
        End If
    Else
       MaskEdBox2 = cValor
    End If
End If
End Sub

Private Sub Form_Activate()
Dim criterio As String
Dim Data1 As New ADODB.Recordset

  If varform.Caption = "Modificar" Then
    Frame1.Visible = True
    Frame2.Visible = False
    If TxtArticulo.Enabled Then
       TxtArticulo.SetFocus
    Else
       If TxtCantidad.Enabled Then TxtCantidad.SetFocus
    End If
  End If
  If Command2.Visible Or VGValnuevo Then
    Frame1.Visible = False
    Frame2.Visible = True
    Command2.Enabled = True
    Command2.SetFocus
    Me.Height = 3250
    Me.Width = 4700
    central FormCreacion
    VGValnuevo = False
  End If
  If VGSeleccion = 2 Then 'significa que viene de modificar
  
     Set Data1 = VGCNx.Execute("SELECT ACODIGO FROM MaeArt ")
     TxtArticulo.text = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 0)
     Label13 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 1)
     Text3 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 2)
     Label14 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 4)
     MaskEdBox1 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 8)
     MaskEdBox2 = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 9)
     TxtCantidad = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 3)
     If varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "X" Then
             Text3.Enabled = False
             Text3.BackColor = &H8000000F
             MaskEdBox1.BackColor = &H8000000F
             MaskEdBox2.BackColor = &H8000000F
             MaskEdBox1.Enabled = False
             MaskEdBox2.Enabled = False
             TxtCantidad.SetFocus
     Else
             Text3.Enabled = True
             If varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 10) = "S" Then
                  MaskEdBox1.BackColor = &H8000000F
                  MaskEdBox2.BackColor = &H80000009
                  MaskEdBox1.Enabled = False
                  MaskEdBox2.Enabled = True
                  TxtCantidad = "1"
                  TxtCantidad.Enabled = False
             Else
                  MaskEdBox1.Enabled = True
                  MaskEdBox2.Enabled = True
             End If
             Text3.SetFocus
     End If
     TxtPrecioUnit.text = varform.MSFlexGrid1.TextMatrix(varform.MSFlexGrid1.Row, 5)
     TxtTotal.text = Val(TxtCantidad.text) * Val(TxtPrecioUnit.text)
     criterio = "ACODIGO = " & "'" + TxtArticulo.text + "'"
     TxtArticulo.Enabled = False
  End If

End Sub

Private Sub Form_Load()
  '**************************************
  Set rsSTKART = New ADODB.Recordset
  rsSTKART.Open "Select * from STKART WHERE STALMA='" & VGAlma & "'", VGCNx, adOpenDynamic, adLockOptimistic
  '**************************************
  
  VGForm1 = 1
  'data1.DatabaseName = cRuta2
  'Data2.DatabaseName = cRuta2
  central FormCreacion
  limpia
  Text8.Enabled = False
  
  'If UCase(Dir$(cRuta4)) = ucase(cNomBd4) Then
  '    Text8 = val(Devolver_Dato(3, Date, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
  'ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
  '    Text8 = val(Devolver_Dato(1, Date, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
  'End If
   bloquearc1_c2 (False) 'significa que dos estan bloqueados

   Select Case VGtipocreacion
   Case 1
           Set varform = FrmRegistro
           Command1.Caption = "Enviar"
           FrmRegistro.Text2 = IIf(VGSoles, "01", "02")
   Case 2
           Set varform = FrmModificar
           If Trim(FrmModificar.moneda) <> "" Then
               Text8 = FrmModificar.TipoDeCambio
               VGSoles = False
               If Trim(FrmModificar.moneda) = "01" Then VGSoles = True
           End If
           Command1.Caption = "Grabar"
           VGValnuevo = False
  End Select
   'PrecioUni.Visible = True
  If VGSoles Then
     Label19 = "S/."
     Label20 = "S/."
  Else
     Label19 = "$"
     Label20 = "$"
  End If
  'VGSeleccion = 3 significa que es nuevo y tiene tipo de cambio proviene de formulario de modificar
  ' Or VGSeleccion = 2 proviene del formulario de de registro
  If VGSeleccion = 2 Or VGSeleccion = 3 Then
     Frame1.Visible = True
     Frame2.Visible = False
  Else
     Frame1.Visible = False
     Frame2.Visible = True
     Combo1.ListIndex = 0
     Combo2.ListIndex = 0
  End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{tab}"
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Text3_DblClick
ElseIf KeyCode = vbKeyTab Then
    Text3_KeyPress (13)
End If
End Sub

Private Sub TxtArticulo_DblClick()
   I = 1
   FormAyuArt.Show 1
   fin = Salida.Rows
   If Salida.Row = 0 Then
     TxtArticulo.SetFocus
     Exit Sub
   End If
   Label16.Visible = True
   Label17.Visible = True
   Label17 = I
   DisplayDisp
   TxtArticulo.Enabled = False
   Text2 = Label14
   If flagserie = "S" Or flaglote = "S" Then
        Text3.Enabled = True
        Text3.SetFocus
    Else
        Text3.Enabled = False
        TxtCantidad.SetFocus
   End If
End Sub

Private Sub TxtArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then
    TxtArticulo_DblClick
 ElseIf KeyCode = 46 Then
     Label13 = ""
 End If
End Sub

Private Sub TxtArticulo_KeyPress(KeyAscii As Integer)
  Dim rpta As Integer
  Dim criterio As String
  Dim Data1 As New ADODB.Recordset
  If KeyAscii = 13 Then
         If Trim(TxtArticulo.text) = "" Then
            TxtArticulo_DblClick
         End If
         criterio = "ACODIGO = " & "'" + TxtArticulo.text + "'"
         Data1.Open "MAEART", VGCNx, adOpenDynamic, adLockOptimistic
         Data1.Find criterio
         If Data1.RecordCount > 0 Then
                    Label13.Caption = Data1.Fields("ADESCRI")
                    Label14.Caption = Data1.Fields("AUNIDAD")
                    flagserie = IIf(Not IsNull(Data1.Fields("AFSERIE")), Data1.Fields("AFSERIE"), "N")
                    flaglote = IIf(Not IsNull(Data1.Fields("AFLOTE")), Data1.Fields("AFLOTE"), "N")
                    Text2.Enabled = True
                    Text2 = Label14
                    TxtArticulo.Enabled = True
                    TxtCantidad.SetFocus
                    VGabrev = Text2
                    ver_serie_lote
                    If flagserie = "S" Or flaglote = "S" Then
                            Text3.Enabled = True
                            Text3.SetFocus
                    Else
                            TxtCantidad.SetFocus
                    End If
            Else
                    If TxtArticulo <> "" Then
                      MsgBox "El Código de Articulo no existe ", vbInformation, mensaje1
                    End If
        '            rpta = MsgBox("Desea registrar un nuevo articulo", vbYesNo + vbQuestion, "Crear un nuevo Articulo")
        '            If rpta = vbYes Then
        '               VGcrea = True
        '               FormArticulos.show 1
        '               VGcrea = False
        '               'Text3.Enabled = True
        '
        '            End If
                    TxtArticulo.SetFocus
           End If
           fin = 0
 Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
 
 End If
End Sub

Private Sub TxtTotal_Change()
Dim real As Double
' If Not IsNumeric(TxtCantidad) And TxtCantidad <> "" Then
'           MsgBox "Ingrese un la cantidad", vbOKOnly + vbExclamation, "Error"
'           TxtCantidad.SetFocus
'           'MOMENTO DE MODIFICAR
' ElseIf Not IsNumeric(TxtTotal.text) And TxtTotal <> "" Then
'           MsgBox "Ingrese un Numero", vbOKOnly + vbExclamation, "Error"
'           TxtTotal.SetFocus
' Else
'     If Trim(TxtTotal) <> "" And Trim(TxtCantidad) <> "" Then
'      real = val(TxtTotal) / val(TxtCantidad)
'      TxtPrecioUnit.text = Format(real, "###0.0000")
''      Label8 = val(TxtCantidad) * FACTOR
''      If Label9 = "" Then Label9 = lblUniEst
'       Command1.Enabled = True
'     End If
'   End If
End Sub

Private Sub TxtTotal_KeyPress(KeyAscii As Integer)
Dim real As Double
If KeyAscii = 13 Then
      If IsNumeric(TxtTotal.text) Then
            real = Val(TxtTotal.text) / Val(TxtCantidad.text)
            TxtPrecioUnit.text = Format(real, "###0.0000")
           Command1.Enabled = True
           Command1.SetFocus
      Else
           MsgBox "Ingrese un numero", vbOKOnly + vbExclamation, "Error"
           TxtTotal.SetFocus
   End If
End If
End Sub

Private Sub Text2_DblClick()
  FrmMntUnidMedida.Show
End Sub

Private Sub Text3_DblClick()
 VGcod = TxtArticulo
 If flaglote = "S" Then
   FormAyuLote.Show 1
 End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim I As Integer
 If KeyAscii = 13 Then
          Text3 = Trim(Text3)
          If Text3 <> "" Then
                If flaglote = "S" Then
                         existe_lote Text3
                         MaskEdBox1.SetFocus
                Else
                         If flagserie = "S" Then
                            If FrmRegistro.MSFlexGrid1.Rows <> 1 Then
                                For I = 1 To FrmRegistro.MSFlexGrid1.Rows - 1
                                  If UCase(Text3) = UCase(FrmRegistro.MSFlexGrid1.TextMatrix(I, 2)) Then
                                    MsgBox "Ya se ingreso la serie", vbInformation, "Aviso"
                                    Exit Sub
                                  End If
                                Next I
                            End If
                            If existe_serie(Text3) Then Exit Sub
                            TxtCantidad = "1"
                            TxtCantidad.Enabled = False
                            Text3.Enabled = False
                            Command1.Enabled = True
                         End If
                End If
        Else
              If flaglote = "S" Then
                          MsgBox "Ingrese el número de Lote", vbInformation, "Aviso"
             Else
                          MsgBox "Ingrese el número de Serie", vbInformation, "Aviso"
             End If
             Exit Sub
      End If
      SendKeys "{tab}"
      KeyAscii = 0
 End If
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
 Dim criterio1 As String
 
   If KeyAscii = 13 Then
      If Not IsNumeric(TxtCantidad.text) Then
           MsgBox "Ingrese un Numero", vbOKOnly + vbExclamation, "Error"
           TxtCantidad.SetFocus
      Else
           If VGRegEnt = 2 Then
             If cant < Val(TxtCantidad.text) Then      ' revisar si validar en creacion
                    MsgBox "No hay stock suficente", 48, "Aviso"
                    TxtCantidad.SetFocus
                    Exit Sub
             End If
           End If
           Txtcantref.SetFocus
      End If
   Else
     If Chr(KeyAscii) = "." And IsNumeric(TxtCantidad) Then Exit Sub
     If ((Chr$(KeyAscii) < "0" Or Chr(KeyAscii) > "9")) And KeyAscii <> 8 Then KeyAscii = 0
   End If
End Sub

Private Sub TxtCantref_KeyPress(KeyAscii As Integer)
 Dim criterio1 As String
 
   If KeyAscii = 13 Then
           If IsNumeric(TxtPrecioUnit) <> 0 Then
                    TxtTotal = Val(TxtCantidad.text) * Val(TxtPrecioUnit.text)
                    TxtTotal.text = Format(TxtTotal, "###0.0000")
                   Command1.Enabled = True
                   TxtPrecioUnit.SelStart = 0
                   TxtPrecioUnit.SelLength = Len(TxtPrecioUnit.text)
                   TxtPrecioUnit.SetFocus
                   'Command1.SetFocus
               Exit Sub
           End If
           TxtPrecioUnit.SelStart = 0
           TxtPrecioUnit.SelLength = Len(TxtPrecioUnit.text)
           Txtcantref.SetFocus
   Else
     If Chr(KeyAscii) = "." And IsNumeric(TxtCantidad) Then Exit Sub
     If ((Chr$(KeyAscii) < "0" Or Chr(KeyAscii) > "9")) And KeyAscii <> 8 Then KeyAscii = 0
   End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      If Not IsNumeric(Text6.text) Then
           MsgBox "Ingrese un Numero", vbOKOnly + vbExclamation, "Error"
           Text6.SetFocus
      Else
           TxtPrecioUnit.SetFocus
      End If
   Else
      If ((Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0")) And Chr$(KeyAscii) = "." And KeyAscii <> 8 Then KeyAscii = 0
   End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsNumeric(Text8.text) Then
           MsgBox "Ingrese un Numero", vbOKOnly + vbExclamation, "Error"
           Text8.SetFocus
      Else
           Command2.Enabled = True
           Command2.SetFocus
      End If
   Else
      If Chr$(KeyAscii) = "." Then Exit Sub
      If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And KeyAscii <> 8 Then KeyAscii = 0
   End If
End Sub



Private Sub TxtPrecioUnit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And IsNumeric(TxtPrecioUnit) Then
      If Not IsNumeric(TxtPrecioUnit.text) Then
           MsgBox "Ingrese un Numero", vbOKOnly + vbExclamation, "Error"
           TxtPrecioUnit.SetFocus
      Else
           TxtTotal = Val(TxtCantidad.text) * Val(TxtPrecioUnit.text)
           TxtTotal.text = Format(TxtTotal, "###0.0000")
           Command1.Enabled = True
           Command1.SetFocus
      End If
   Else
     If Chr(KeyAscii) = "." And IsNumeric(TxtCantidad) Then Exit Sub
     If ((Chr$(KeyAscii) < "0" Or Chr(KeyAscii) > "9")) And KeyAscii <> 8 Then KeyAscii = 0
   End If
End Sub

Private Sub limpia()
   Label16.Visible = False   'num item
   Label17.Visible = False
   TxtArticulo.text = ""
   Text2 = ""
   TxtCantidad.Enabled = True
   TxtCantidad.text = ""
   Text8.text = ""
   TxtPrecioUnit.text = ""
   TxtTotal.text = ""
   Text3 = ""
   TxtArticulo.Enabled = True
   Text8.Enabled = True
   Label14 = ""
   Label13 = ""
   MaskEdBox1 = "__/__/____"
   MaskEdBox2 = "__/__/____"
   Text3.BackColor = &H80000009
   MaskEdBox1.BackColor = &H80000009
   MaskEdBox2.BackColor = &H80000009
End Sub

Public Sub DisplayDisp()
  Dim criterio As String
   TxtArticulo = Salida.TextMatrix(I, 0)
   Label13 = Salida.TextMatrix(I, 1)
   Label14 = Salida.TextMatrix(I, 2)
   flagserie = Salida.TextMatrix(I, 3) 'serie
   flaglote = Salida.TextMatrix(I, 4) 'lote
   criterio = " STCODIGO ='" & TxtArticulo.text & "' and STALMA ='" & VGAlma & "'"
   'RMM 16/07/2001 ****************************************************
   rsSTKART.Filter = criterio
   'RMM ****************************************************
   If Not rsSTKART.EOF Then
      cant = rsSTKART("STSKDIS")
   Else
      cant = 0
   End If
   TxtPrecioUnit = UltimoPrecio(TxtArticulo.text, IIf(VGSoles, "01", "02"))
   Text2.text = Label14
   ver_serie_lote
   TxtArticulo.Enabled = False
   VGabrev = Text2
End Sub

Private Sub ver_serie_lote()
 If flaglote = "S" Then
       Text3.Enabled = True
       MaskEdBox1.Enabled = True
       MaskEdBox2.Enabled = True
       MaskEdBox1.BackColor = &H80000009
       MaskEdBox2.BackColor = &H80000009
       TxtCantidad.Enabled = True
ElseIf flagserie = "S" Then
       MaskEdBox1.Enabled = False
       MaskEdBox2.Enabled = True
       MaskEdBox1.BackColor = &H8000000F
       MaskEdBox2.BackColor = &H80000009
       TxtCantidad = "1"
       TxtCantidad.Enabled = False
Else
      Text3.Enabled = False
      Text3.BackColor = &H8000000F
      TxtCantidad.Enabled = True
      MaskEdBox1.BackColor = &H8000000F
      MaskEdBox2.BackColor = &H8000000F
      MaskEdBox1.Enabled = False
      MaskEdBox2.Enabled = False
End If
End Sub

Private Sub bloquearc1_c2(dato As Boolean)
Command1.Enabled = dato
Command2.Enabled = dato
End Sub

Private Sub pro_xserie()
  If flagserie = "S" Then
      xserie = "S"
  ElseIf flaglote = "S" Then
      xserie = "N"
  Else
      xserie = "X"
  End If
End Sub

Function Existe(text As TextBox) As Boolean
 Dim RSQL As String
 Dim rs As Recordset
 RSQL = "select stscodigo from STKSERI where STSALMA ='" & VGAlma & "'   and STSCODIGO ='" & TxtArticulo & "'AND  STSSERIE ='" & Text3 & "'"
 Existe = True
 'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
 Set rs = VGCNx.Execute(RSQL)
 If rs.EOF Then
   Existe = False
 End If
 rs.Close
End Function
Function existe_serie(text As TextBox) As Boolean
Dim rs As Recordset
Dim RSQL As String
existe_serie = False
RSQL = "select STSSERIE from STKSERI where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & TxtArticulo & "' and STSSERIE = '" & text & "'"
'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = VGCNx.Execute(RSQL)
If Not rs.EOF Then
    If Not graba Then MsgBox "Serie Registrada en Almacen", vbInformation, "Aviso"
    existe_serie = True
Else
    If Not graba Then MsgBox "Serie  No Registrada en Almacen", vbInformation, "Aviso"
    existe_serie = False
End If
rs.Close
End Function


'***************************************    graba  detalle
Private Sub grabadetalle()
 Dim Adoreg1 As ADODB.Recordset
 Dim AdoReg2 As ADODB.Recordset
 Dim Rsql1 As String
 Dim criterio As String
 Dim item As Integer
 On Error GoTo Erra
  TxtArticulo = Trim(TxtArticulo)
  If VGSeleccion = 2 Then   'indica es el dato se modifica
     Rsql1 = "select * from movalmdet where DETD= '" & FrmModificar.TxDoc & "' AND DENUMDOC= '" & FrmModificar.Lblnumdoc & "'  AND DEALMA= '" & VGAlma & "' and  DEITEM = " & FrmModificar.contador & " and DECODIGO = '" & TxtArticulo & "'"
  Else
     Rsql1 = "select * from movalmdet where DETD='TT'"
  End If
  Set Adoreg1 = New ADODB.Recordset
  Adoreg1.Open Rsql1, VGCNx, adOpenDynamic, adLockOptimistic
  ' Si es nuevo adicciono los datos primary key
  If Adoreg1.RecordCount = 0 Then
        Adoreg1.AddNew
        Adoreg1("dealma") = VGAlma
Retor:
        Rsql1 = "select * from movalmdet where DETD= '" & FrmModificar.TxDoc & "' AND DENUMDOC= '" & FrmModificar.Lblnumdoc & "'  AND DEALMA= '" & VGAlma & "' and  DEITEM = " & FrmModificar.numitem & ""
        Set AdoReg2 = New ADODB.Recordset
        AdoReg2.Open Rsql1, VGCNx, adOpenDynamic, adLockOptimistic
        If Not AdoReg2.EOF Then
           FrmModificar.numitem = AdoReg2("deitem") + 1
           FrmModificar.contador = FrmModificar.numitem
           FormCreacion.Label11 = FrmModificar.numitem
           GoTo Retor
        End If
        AdoReg2.Close
        Adoreg1("deitem") = FrmModificar.numitem
        Adoreg1("decodigo") = Trim(TxtArticulo)   ' Format(MSFlexGrid1.TextMatrix(contador, 0), "00000000")
        Adoreg1("dedescri") = Label13
        Adoreg1("detd") = FrmModificar.TxDoc
        Adoreg1("denumdoc") = FrmModificar.Lblnumdoc
        FrmModificar.numitem = FrmModificar.numitem + 1
  End If
  ' adicciono la nueva cantidad, serie y lote
     Adoreg1("decantid") = Val(TxtCantidad)
     Adoreg1("decodmon") = IIf(Label19 = "S/.", "01", "02")
     '*RMM**************CONVERSION  NO APLICABLE*****************************************************
     Adoreg1("deprecio") = CDbl(TxtPrecioUnit) '* IIf(Label19 = "S/.", 1, CDbl(Text8))
     '*RMM**************
     Adoreg1("detipcam") = Val(Text8)
     
     If xserie = "S" Then
         actserie
         Adoreg1("DESERIE") = Trim(Text3)
     ElseIf xserie = "N" Then
         grabalote
         Adoreg1("DELOTE") = Trim(Text3)
     End If
    Adoreg1.Update
    Adoreg1.Close
    'nuevodet = True   'para que no actualice dos veces
    'RMM*************************************************
    actualizastk TxtArticulo  'actualizando el stocK
    'RMM*************************************************
    ya_grabo_det = True
    Exit Sub
Erra:
    'MsgBox "HOLA"
    MsgBox Err.Description
    'hubo_error = True
End Sub

Private Sub actualizastk(codigo As String)
  Dim criterio As String
  Dim canttemp As Double
  Dim adoreg As ADODB.Recordset
  Dim RSQL As String
   RSQL = "select * from STKART where  STCODIGO= '" & TxtArticulo & "' and STALMA = '" & VGAlma & "'  "
   Set adoreg = New ADODB.Recordset
   adoreg.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
   If adoreg.RecordCount = 0 Then
     adoreg.AddNew
     adoreg("stalma") = VGAlma
     adoreg("stcodigo") = TxtArticulo
   Else
     canttemp = adoreg("stskdis")
   End If
   'Adoreg("stskdis") = IIf(FrmModificar.tipo = "NI", Adoreg("stskdis") + Val(TxtCantidad), Adoreg("stskdis") - Val(TxtCantidad))
   adoreg("stskdis") = IIf(FrmModificar.tipo = "NI", canttemp + Val(TxtCantidad), canttemp - Val(TxtCantidad))
   adoreg.Update
   ValMes
   adoreg.Close
End Sub

Public Sub grabalote()
Dim uSql As String
Dim Lote As String
Dim nuevo_stk As Double
Dim RSQL As String
Dim rs As Recordset
Dim fecfab As Date
Dim fecven As Date
    On Error GoTo Erra
    RSQL = "select STSLKDIS FROM STKLOTE where  STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & TxtArticulo & "' and STSLOTE= '" & Text3 & "'" '
    'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    If Not rs.EOF Then
       nuevo_stk = IIf(FrmModificar.tipo = "NI", rs(0) + Val(TxtCantidad), rs(0) - Val(TxtCantidad))
       uSql = "Update STKLOTE set STSLKDIS = " & nuevo_stk & " WHERE  STSALMA='" & VGAlma & "' and STSCODIGO='" & TxtArticulo & "'AND STSLOTE='" & Lote & "'"
    Else
        If MaskEdBox1 <> "__/__/____" And (MaskEdBox2 = "__/__/____") Then
            uSql = "insert into STKLOTE (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB)  VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & Text3 & "' ," & Val(TxtCantidad) & " ,'" & DateSQL(MaskEdBox2) & "') "
        ElseIf MaskEdBox1 = "__/__/____" And MaskEdBox2 <> "__/__/____" Then
            uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECVEN)VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & Text3 & "' ," & Val(TxtCantidad) & " ,' ','" & DateSQL(MaskEdBox1) & "') "  'SIN FECFAB
        ElseIf MaskEdBox1 <> "__/__/____" And MaskEdBox2 <> "__/__/____" Then
            uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB,STSFECVEN)  VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & Text3 & "' ," & Val(TxtCantidad) & " ,'" & DateSQL(MaskEdBox2) & "','" & DateSQL(MaskEdBox1) & "') "
        Else
            uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS)  VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & Text3 & "' ," & Val(TxtCantidad) & " ,' ',' ') "
        End If
    
    End If
    VGCNx.Execute uSql
    Exit Sub
Erra:
    MsgBox "Base de datos bloqueada"
    'hubo_error = False
End Sub

Public Sub actserie()
On Error GoTo salir

Dim uSql As String
Dim Serie As String
Dim valor As Integer
Dim rs As Recordset
Dim RSQL As String
Dim fecfab As Date
Dim fecven As Date

    RSQL = "select STSSKDIS FROM STKSERI where   STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & TxtArticulo & "' and STSSERIE= '" & Text3 & "'" '
    'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    If Not rs.EOF Then
       valor = IIf(FrmModificar.tipo = "NI", 1, 0)
       uSql = "update STKSERI set STSSKDIS = " & valor & " WHERE  STSALMA='" & VGAlma & "' and STSCODIGO='" & TxtArticulo & "'AND STSSERIE='" & Text3 & "'"
    Else
       uSql = "insert into STKSERI (STSALMA,STSCODIGO,STSSERIE,STSSKDIS) VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & Text3 & "' ,'1') "
    End If
    VGCNx.Execute uSql
    Exit Sub
salir:
 Exit Sub
'   MsgBox  err.  , vbExclamation, "Aviso"
End Sub

Private Sub ValMes()
 
  Dim criterio As String
  Dim Adoreg1 As ADODB.Recordset
  Dim RSQL As String
  Dim Cantent As Double
  Dim Cantsal As Double
  Dim mespro As String
  Dim uSql As String
  
   mespro = Year(FrmModificar.DTPicker1) & Format(Month(FrmModificar.DTPicker1), "00")
   RSQL = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & TxtArticulo & "'"  '
   
   Set Adoreg1 = New ADODB.Recordset
   Adoreg1.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
   If Adoreg1.RecordCount <> 0 Then
      If FrmModificar.tipo = "NI" Then
        Cantent = Adoreg1(0) + Val(TxtCantidad)        '
        uSql = "Update MoResMes set SMCANENT = " & Cantent & "  where SMALMA='" & VGAlma & "'  and  SMCODIGO ='" & TxtArticulo & "' AND SMMESPRO ='" & mespro & "' "
       Else
        Cantsal = Adoreg1(1) + Val(TxtCantidad)   'INICIALMENTE DESCARGA, ACTUALIZA LA NUEVA CANTIDAD
        uSql = "Update MoResMes set SMCANSAL = " & Cantsal & "  where SMALMA='" & VGAlma & "' and  SMCODIGO ='" & TxtArticulo & "' AND SMMESPRO ='" & mespro & "' "
       End If
   Else
      If FrmModificar.tipo = "NI" Then
        Cantent = Val(TxtCantidad)
        Cantsal = 0
      Else
        Cantsal = Val(TxtCantidad)
        Cantent = 0
      End If
      uSql = "insert into MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMSALDOINI) VALUES ('" & VGAlma & "','" & TxtArticulo & "','" & mespro & "' ," & Cantent & "," & Cantsal & ",0) "
   End If
   VGCNx.Execute uSql
   Adoreg1.Close
End Sub

Private Sub existe_lote(text As TextBox)
Dim rs As Recordset
Dim RSQL As String
   RSQL = "select STSLOTE, STSLKDIS,STSFECVEN,STSFECFAB from STKLOTE where  STSALMA ='" & VGAlma & "' and STSCODIGO = '" & TxtArticulo & "' and STSLOTE = '" & text & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
     MsgBox "Lote Registrado en Almacen", vbInformation, "Aviso"
    ' lbcantstk = rS(1)
     MaskEdBox1 = IIf(IsNull(rs(2)), "__/__/____", rs(2))
     MaskEdBox2 = IIf(IsNull(rs(3)), "__/__/____", rs(3))
   End If
End Sub



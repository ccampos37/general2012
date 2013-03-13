VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmEmisionOC_detalle1 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículo"
   ClientHeight    =   6516
   ClientLeft      =   1380
   ClientTop       =   1860
   ClientWidth     =   7056
   ControlBox      =   0   'False
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6516
   ScaleWidth      =   7056
   Begin VB.CommandButton CmdOK1 
      Caption         =   "&Ok"
      Height          =   684
      Left            =   2688
      Picture         =   "frmEmisionOC_detalle.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5712
      Width           =   972
   End
   Begin VB.Frame Frame4 
      Height          =   1020
      Left            =   96
      TabIndex        =   35
      Top             =   4656
      Width           =   6780
      Begin VB.TextBox txtCo1 
         Height          =   384
         Left            =   1128
         TabIndex        =   11
         Top             =   336
         Width           =   5244
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Comentario"
         Height          =   288
         Left            =   48
         TabIndex        =   36
         Top             =   384
         Width           =   984
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Precios"
      ForeColor       =   &H8000000D&
      Height          =   2124
      Left            =   96
      TabIndex        =   19
      Top             =   2544
      Width           =   6780
      Begin VB.TextBox txtPIg 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5736
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   312
         Width           =   735
      End
      Begin VB.TextBox txtPUn 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   288
         Width           =   1335
      End
      Begin VB.TextBox txtPDe 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3576
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   312
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Total Neto"
         Height          =   192
         Left            =   5712
         TabIndex        =   34
         Top             =   1392
         Width           =   756
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Porcen. IGV."
         Height          =   192
         Left            =   4704
         TabIndex        =   33
         Top             =   408
         Width           =   912
      End
      Begin VB.Label lblTCo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   288
         Left            =   1416
         TabIndex        =   32
         Top             =   1680
         Width           =   1332
      End
      Begin VB.Label lblIgv 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   336
         Left            =   3192
         TabIndex        =   31
         Top             =   1656
         Width           =   1212
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Importe IGV."
         Height          =   192
         Left            =   3336
         TabIndex        =   30
         Top             =   1368
         Width           =   888
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   192
         Left            =   6576
         TabIndex        =   29
         Top             =   312
         Width           =   120
      End
      Begin VB.Label lblTNe 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   288
         Left            =   5208
         TabIndex        =   28
         Top             =   1680
         Width           =   1332
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total Compra"
         Height          =   192
         Left            =   1584
         TabIndex        =   27
         Top             =   1344
         Width           =   1236
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Precio Neto"
         Height          =   192
         Left            =   5472
         TabIndex        =   26
         Top             =   768
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Porcen. Dsct."
         Height          =   192
         Left            =   2544
         TabIndex        =   25
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Precio Unit."
         Height          =   192
         Left            =   144
         TabIndex        =   24
         Top             =   288
         Width           =   876
      End
      Begin VB.Label lblDes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   288
         Left            =   3096
         TabIndex        =   23
         Top             =   936
         Width           =   1212
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Importe Dsct."
         Height          =   192
         Left            =   3216
         TabIndex        =   22
         Top             =   696
         Width           =   948
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   192
         Left            =   4416
         TabIndex        =   21
         Top             =   360
         Width           =   120
      End
      Begin VB.Label lblPNe 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   288
         Left            =   5160
         TabIndex        =   20
         Top             =   1008
         Width           =   1332
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos ITEM"
      Height          =   1020
      Left            =   96
      TabIndex        =   12
      Top             =   1536
      Width           =   6828
      Begin VB.TextBox txtordfab 
         Height          =   336
         Left            =   1104
         MaxLength       =   20
         TabIndex        =   4
         Top             =   192
         Width           =   1740
      End
      Begin VB.TextBox txtCan 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3768
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   216
         Width           =   1095
      End
      Begin VB.TextBox txtURe 
         Height          =   285
         Left            =   3768
         TabIndex        =   6
         Top             =   612
         Width           =   735
      End
      Begin VB.TextBox txtRef 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5844
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   588
         Width           =   756
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Orden fab."
         Height          =   240
         Left            =   192
         TabIndex        =   18
         Top             =   240
         Width           =   744
      End
      Begin VB.Label lblUni 
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   5880
         TabIndex        =   17
         Top             =   204
         Width           =   732
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Unidad"
         Height          =   192
         Left            =   4944
         TabIndex        =   16
         Top             =   264
         Width           =   516
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   192
         Left            =   2976
         TabIndex        =   15
         Top             =   276
         Width           =   636
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Unid. Ref."
         Height          =   192
         Left            =   3000
         TabIndex        =   14
         Top             =   672
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
         Height          =   192
         Left            =   4944
         TabIndex        =   13
         Top             =   696
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   675
      Left            =   4320
      Picture         =   "frmEmisionOC_detalle.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5712
      Width           =   775
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_TipoArt 
      Height          =   312
      Left            =   912
      TabIndex        =   1
      Top             =   192
      Width           =   3336
      _ExtentX        =   5884
      _ExtentY        =   550
      XcodMaxLongitud =   2
      xcodwith        =   100
      NomTabla        =   "al_tipoarticulo"
      TituloAyuda     =   "Tipo de Articulo"
      ListaCampos     =   "Tipoarticulocodigo(1),tipoarticuloDescripcion(2)"
      XcodCampo       =   "tipoarticulocodigo"
      XListCampo      =   "tipoarticulodescripcion"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "Tipoarticulocodigo,tipoarticuloDescripcion"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Art 
      Height          =   348
      Left            =   912
      TabIndex        =   2
      Top             =   576
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   614
      XcodMaxLongitud =   0
      xcodwith        =   1500
      NomTabla        =   "maeart"
      ListaCampos     =   "acodigo(1),adescri(1)"
      XcodCampo       =   "acodigo"
      XListCampo      =   "adescri"
      ListaCamposDescrip=   "Codigo,descripcion"
      ListaCamposText =   "acodigo,adescri"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ccostos 
      Height          =   312
      Left            =   912
      TabIndex        =   3
      Top             =   960
      Width           =   3312
      _ExtentX        =   5842
      _ExtentY        =   550
      XcodMaxLongitud =   10
      xcodwith        =   1000
      NomTabla        =   "ct_centrocosto"
      TituloAyuda     =   "Ayuda de Centro de Costos"
      ListaCampos     =   "centrocostocodigo(1), centrocostodescripcion(2)"
      XcodCampo       =   "centrocostocodigo"
      XListCampo      =   "centrocostodescripcion"
      ListaCamposDescrip=   "Codigo,Descripcion"
      ListaCamposText =   "centrocostocodigo,centrocostodescripcion"
      Requerido       =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "C.Costos"
      Height          =   192
      Left            =   192
      TabIndex        =   42
      Top             =   1008
      Width           =   648
   End
   Begin VB.Label lblFab 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   5148
      TabIndex        =   41
      Top             =   204
      Width           =   1452
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   192
      Left            =   192
      TabIndex        =   40
      Top             =   672
      Width           =   492
   End
   Begin VB.Label lblUnidad 
      AutoSize        =   -1  'True
      Caption         =   "Fabricante"
      Height          =   192
      Left            =   4320
      TabIndex        =   39
      Top             =   240
      Width           =   756
   End
   Begin VB.Label lbltipo 
      AutoSize        =   -1  'True
      Caption         =   "Tipo "
      Height          =   192
      Left            =   192
      TabIndex        =   38
      Top             =   288
      Width           =   360
   End
End
Attribute VB_Name = "frmEmisionOC_detalle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public activado As Boolean
Public cancelado As Boolean
Public Igv As Single
Public Tipo As String
Dim Mensaje As String

Private Sub cmdCancel_Click()
    cancelado = True
    Unload Me
End Sub

Private Sub CmdOK1_Click()
    If txtCod = "" Then
        Mensaje = "Debe ingresar Código de Artículo"
        MsgBox Mensaje, vbExclamation, "Error"
        txtCod.SetFocus
        Exit Sub
    Else
        If Not txtDes.Enabled Then
            If Not Existe(1, txtCod, "maeart", "acodigo", False) And Not Existe(1, txtCod, "Servicios", "ser_codigo", False) Then
                Mensaje = "Código de Artículo no válido"
                MsgBox Mensaje, vbExclamation, "Error"
                txtCod.SetFocus
                Exit Sub
            Else
                txtCod_KeyPress 13
                CmdOK1.SetFocus
            End If
        End If
    End If
    
If Tipo = "B" Then
    If Val(txtCan) = 0 Then
        Mensaje = "Debe especificar Cantidad"
        MsgBox Mensaje, vbExclamation, "Error"
        txtCan.SetFocus
        Exit Sub
    End If
 End If
    
    If txtURe <> "" Then
        If Not txtRef.Enabled Then
            If Not Existe(1, txtURe, "tabunimed", "um_abrev", False) Then
                Mensaje = "Unidad de referencia no válida"
                MsgBox Mensaje, vbExclamation, "Error"
                txtURe.SetFocus
                Exit Sub
            Else
                txtURe_KeyPress 13
                CmdOK1.SetFocus
            End If
        End If
        If Val(txtRef) = 0 Then
            Mensaje = "Debe especificar Orden de FabricacionccionReferencia"
            MsgBox Mensaje, vbExclamation, "Error"
            txtRef.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txtPUn) = 0 Then
        Mensaje = "Debe especificar Precio Unitario"
        MsgBox Mensaje, vbExclamation, "Error"
        txtPUn.SetFocus
        Exit Sub
    End If
    
    cancelado = False
    txtCod.Enabled = True
    txtCod.SetFocus
    Me.Hide
End Sub

Private Sub Form_Activate()
    Igv = Val(txtPIg)
End Sub

Private Sub Form_Load()
    Call Ctr_TipoArt.Conexion(cn)
    Call Ctr_Art.Conexion(cn)
    Call Ctr_Ccostos.Conexion(cn)
    Ctr_TipoArt.xclave = "": Ctr_TipoArt.xnombre = ""
    central Me
End Sub

Private Sub txtCan_Change()
    Calculo_Automatico
End Sub

Private Sub txtCan_GotFocus()
    Enfoque txtCan
End Sub

Private Sub txtCan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtCan) > 0 Then
            txtURe.SetFocus
            txtURe = ""
        Else
            Enfoque txtCan
        End If
    End If
    Reales_Positivos KeyAscii, txtCan
End Sub

Private Sub txtCan_LostFocus()
    txtCan = Format(Val(txtCan), "0.00")
End Sub

Private Sub txtordfab_GotFocus()
    Enfoque txtordfab
End Sub

Private Sub txtordfab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCo1.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtco1_GotFocus()
    Enfoque txtCo1
End Sub

Private Sub txtCo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdOK1.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtPDe_Change()
    Calculo_Automatico
End Sub

Private Sub txtPDe_GotFocus()
    Enfoque txtPDe
End Sub

Private Sub txtPDe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPIg.SetFocus
    End If
    Reales_Positivos KeyAscii, txtPDe
End Sub

Private Sub txtPDe_LostFocus()
    txtPDe = Format(Val(txtPDe), "0.00")
End Sub

Private Sub txtPIg_Change()
    Calculo_Automatico
End Sub

Private Sub txtPIg_GotFocus()
    Enfoque txtPIg
End Sub

Private Sub txtPIg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtordfab.SetFocus
    End If
    Reales_Positivos KeyAscii, txtPIg
End Sub

Private Sub txtPIg_LostFocus()
    txtPIg = Format(Val(txtPIg), "0.00")
End Sub

Private Sub txtPUn_Change()
    If Val(txtPUn) = 0 Then
        txtPDe = "0.00"
        txtPDe.Enabled = False
        txtPIg = Format(Igv, "0.00")
        txtPIg.Enabled = False
    Else
        txtPDe.Enabled = True
        txtPIg.Enabled = True
    End If
    Calculo_Automatico
End Sub

Private Sub txtPUn_GotFocus()
    Enfoque txtPUn
End Sub

Private Sub txtPUn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtPUn) > 0 Then
            txtPDe.SetFocus
        Else
            txtPUn.SetFocus
        End If
    End If
    Reales_Positivos KeyAscii, txtPUn
End Sub

Private Sub txtPUn_LostFocus()
    txtPUn = Format(Val(txtPUn), "0.00")
End Sub

Private Sub txtRef_Change()
    If Val(txtRef) = 0 Then
        If Me.ActiveControl.name <> "txtURe" Then
           txtPUn.Enabled = False
        End If
    Else
        txtPUn.Enabled = True
    End If
End Sub

Private Sub txtRef_GotFocus()
    Enfoque txtRef
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtRef) > 0 Then
            txtPUn.SetFocus
        Else
            txtRef.SetFocus
        End If
    End If
    Reales_Positivos KeyAscii, txtRef
End Sub

Private Sub txtRef_LostFocus()
    txtRef = Format(Val(txtRef), "0.00")
End Sub

Private Sub txtURe_Change()
    If txtURe = "" Then
       txtRef.Enabled = False
       txtPUn.Enabled = True
    Else
       txtRef.Enabled = True
    End If
    txtRef = ""
    Calculo_Automatico
End Sub

Private Sub txtURe_DblClick()
    Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT um_abrev,um_nombre FROM tabunimed"
    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Unidades de Medida"
    frmReferencia.Inicio
    frmReferencia.Show vbModal
    Adodc2.Close
    
    If vGUtil(1) <> "" Then
        txtURe = vGUtil(1)
        txtURe_KeyPress 13
    End If
End Sub

Private Sub txtURe_GotFocus()
    Enfoque txtURe
End Sub

Private Sub txtURe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then txtURe_DblClick
End Sub

Private Sub txtURe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtURe = Trim(txtURe)
        If txtURe <> "" Then
            If Not Existe(1, txtURe, "tabunimed", "um_abrev", False) Then
                Mensaje = "La Unidad de medida de Referencia no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                txtURe.SetFocus
            Else
                If Not txtRef.Enabled Then
                    txtRef = "0.00"
                    txtRef.Enabled = True
                End If
                txtRef.SetFocus
            End If
        Else
            txtPUn.SetFocus
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Sub Calculo_Automatico()
    If Not activado Then Exit Sub
    
    If Not txtRef.Enabled Then
        lblDes = Format(Val(txtPUn) * Val(txtPDe) / 100, "0.00")
        lblPNe = Format(Val(txtPUn) - Val(lblDes), "0.00")
        If Tipo = "S" Then
          lblTCo = Format(1 * Val(lblPNe), "0.00")
        Else
         lblTCo = Format(Val(txtCan) * Val(lblPNe), "0.00")
        End If
    Else
        lblDes = Format(Val(txtPUn) * Val(txtPDe) / 100, "0.00")
        lblPNe = Format(Val(txtPUn) - Val(lblDes), "0.00")
        lblTCo = Format(Val(txtRef) * Val(lblPNe), "0.00")
    End If

    lblIgv = Format(Val(lblTCo) * Val(txtPIg) / 100, "0.00")
    lblTNe = Format(Val(lblTCo) + Val(lblIgv), "0.00")
End Sub
